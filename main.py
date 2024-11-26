"""
Teams Helper
Version: 1.0
Author: Tomasz Witke
License: Apache License 2.0

Description:
A desktop application for managing audio recordings during Microsoft Teams meetings.
"""

import os
import sys
import threading
import logging
import signal
import json
import time
from datetime import datetime
from tkinter import Tk, messagebox, filedialog, StringVar, Label, Entry, Button

import numpy as np
import sounddevice as sd
import lameenc
from pystray import Icon, Menu, MenuItem
from PIL import Image
from websocket import create_connection, WebSocketConnectionClosedException
import configparser


class TeamsHelperRecorder:
    """
    Main recorder class for Teams Helper.
    Handles audio recording, settings management, and integration with Teams API.
    """
    def __init__(self, sample_rate=44100, channels=2, buffer_size=1024):
        self.sample_rate = sample_rate
        self.channels = channels
        self.buffer_size = buffer_size
        self.recording = False
        self.stop_flag = False
        self.frames = []
        self.ws = None
        self.can_toggle_mute = False
        self.record_all_meetings = True  # Default to True

        # Paths and settings
        self.settings_file = os.path.join(os.getenv("APPDATA"), "teamshelper", "settings.ini")
        self.output_dir = self.load_settings()
        os.makedirs(self.output_dir, exist_ok=True)
        self.tray_icon = None

        # Logging setup
        self.configure_logging()

        # LAME MP3 encoder setup
        self.encoder = lameenc.Encoder()
        self.encoder.set_bit_rate(128)
        self.encoder.set_in_sample_rate(self.sample_rate)
        self.encoder.set_channels(self.channels)
        self.encoder.set_quality(2)

    def get_icon_path(self):
        """
        Get the path to the `icon.ico` file.
        Handles both development and compiled (PyInstaller) environments.
        """
        if hasattr(sys, "_MEIPASS"):
            return os.path.join(sys._MEIPASS, "icon.ico")
        return os.path.join(os.path.dirname(__file__), "icon.ico")

    def set_tray_icon(self, icon):
        """
        Set the tray icon object for the application.
        """
        self.tray_icon = icon
        self.update_tray_title()

    def update_tray_title(self):
        """
        Update the tray icon's title based on the recording state.
        """
        if self.tray_icon:
            state = "Recording" if self.recording else "Idle"
            self.tray_icon.title = f"Teams Helper - {state}"

    def load_settings(self):
        """
        Load application settings from the configuration file.
        If the file does not exist, create it with default values.
        """
        config = configparser.ConfigParser()
        default_folder = os.path.join(os.path.expanduser("~"), "Downloads", "Recordings")
        if os.path.exists(self.settings_file):
            config.read(self.settings_file)
            return config.get("Settings", "output_dir", fallback=default_folder)
        else:
            os.makedirs(os.path.dirname(self.settings_file), exist_ok=True)
            self.save_settings(default_folder)
            return default_folder

    def save_settings(self, output_dir):
        """
        Save application settings to the configuration file.
        """
        config = configparser.ConfigParser()
        config["Settings"] = {"output_dir": output_dir}
        with open(self.settings_file, "w") as configfile:
            config.write(configfile)
        self.output_dir = output_dir
        os.makedirs(self.output_dir, exist_ok=True)
        self.configure_logging()

    def configure_logging(self):
        """
        Configure logging to save logs in the output directory.
        """
        self.log_file = os.path.join(self.output_dir, "app.log")
        logging.basicConfig(
            filename=self.log_file,
            level=logging.DEBUG,
            format="%(asctime)s - %(levelname)s - %(message)s",
        )
        logging.getLogger("PIL").setLevel(logging.WARNING)
        logging.info("Teams Helper started. Recordings and logs will be saved in '%s'.", self.output_dir)

    def show_error(self, title, message):
        """
        Display an error message in a popup window.
        """
        logging.error(f"{title}: {message}")
        root = Tk()
        root.withdraw()  # Hide the main Tkinter window
        messagebox.showerror(title, message)
        root.destroy()

    def show_settings_window(self):
        """
        Display the settings window for changing the output directory.
        """
        def settings_window():
            def browse_folder():
                """
                Open a dialog to select a new folder.
                """
                folder = filedialog.askdirectory(initialdir=self.output_dir)
                if folder:
                    folder_var.set(folder)

            def save_settings():
                """
                Save the selected folder and close the settings window.
                """
                selected_folder = folder_var.get()
                if os.path.isdir(selected_folder):
                    self.save_settings(selected_folder)
                    messagebox.showinfo("Settings", f"Folder changed to: {selected_folder}")
                    root.destroy()
                else:
                    messagebox.showerror("Error", "The selected folder does not exist. Please select a valid folder.")

            # Create the settings window
            root = Tk()
            root.title("Settings")
            root.geometry("500x150")
            root.iconbitmap(self.get_icon_path())

            # Current folder label and entry
            folder_var = StringVar(value=self.output_dir)
            Label(root, text="Current Recording Folder:", anchor="w").pack(fill="x", padx=10, pady=5)
            Entry(root, textvariable=folder_var, state="readonly", width=60).pack(fill="x", padx=10, pady=5)

            # Buttons
            Button(root, text="Browse...", command=browse_folder).pack(side="left", padx=10, pady=10)
            Button(root, text="Save", command=save_settings).pack(side="right", padx=10, pady=10)

            root.mainloop()

        threading.Thread(target=settings_window, daemon=True).start()

    def toggle_record_all_meetings(self):
        """
        Toggle the setting to record all meetings automatically.
        """
        self.record_all_meetings = not self.record_all_meetings
        state = "enabled" if self.record_all_meetings else "disabled"
        print(f"[INFO] Record All Meetings is now {state}.")
        logging.info("Record All Meetings toggled to: %s", state)

    def connect_to_teams(self):
        """
        Establish a WebSocket connection to the Teams API.
        Monitors meeting events and triggers recording.
        """
        base_url = "ws://localhost:8124"
        url = f"{base_url}?protocol-version=2.0.0&manufacturer=Kyvaith&device=TeamsHelper&app=TeamsHelper&app-version=1.0"

        print("[INFO] Connecting to Microsoft Teams API...")
        logging.info("Connecting to Microsoft Teams API...")

        try:
            self.ws = create_connection(url)
            self.ws.settimeout(1)  # Non-blocking mode
            logging.info("Connected to Microsoft Teams WebSocket.")
            print("[INFO] Connected to Microsoft Teams WebSocket.")

            while not self.stop_flag:
                try:
                    message = self.ws.recv()
                    self.handle_teams_update(message)
                except WebSocketConnectionClosedException:
                    logging.warning("WebSocket connection closed.")
                    print("[WARN] WebSocket connection closed.")
                    break
                except Exception as e:
                    time.sleep(0.1)

            if self.ws:
                self.ws.close()
                print("[INFO] WebSocket connection closed.")
                logging.info("WebSocket connection closed.")

        except Exception as e:
            self.show_error("Teams API Connection Error", f"Failed to connect to Teams API: {e}")

    def handle_teams_update(self, message):
        """
        Process updates from Microsoft Teams WebSocket API.
        """
        try:
            logging.debug("Received WebSocket message: %s", message)
            update = json.loads(message)

            # Extract the meeting permissions and status
            meeting_permissions = update.get("meetingUpdate", {}).get("meetingPermissions", {})
            can_toggle_mute = meeting_permissions.get("canToggleMute", False)

            # Start or stop recording based on meeting state
            if self.record_all_meetings:
                if can_toggle_mute and not self.can_toggle_mute:
                    self.can_toggle_mute = True
                    self.start_recording()
                elif not can_toggle_mute and self.can_toggle_mute:
                    self.can_toggle_mute = False
                    self.stop_recording()

        except Exception as e:
            self.show_error("Teams API Update Error", str(e))

    def start_recording(self):
        """
        Start the audio recording process.
        """
        print("[INFO] Starting recording...")
        logging.info("Starting recording...")
        self.recording = True
        self.update_tray_title()
        self.recording_thread = threading.Thread(target=self.record_audio)
        self.recording_thread.start()

    def record_audio(self):
        """
        Record audio from both Stereo Mix and Microphone.
        """
        try:
            stereo_mix_device = self.get_device_by_name("Stereo Mix")
            microphone_device = self.get_device_by_name("Microphone")

            if stereo_mix_device is None or microphone_device is None:
                raise RuntimeError("Required audio devices not found.")

            formatted_time = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
            mp3_file = os.path.join(self.output_dir, f"Recording {formatted_time}.mp3")
            print(f"[INFO] Recording directly to {mp3_file}")

            with open(mp3_file, "wb") as mp3:
                with sd.InputStream(samplerate=self.sample_rate,
                                    channels=self.channels,
                                    dtype='int16',
                                    device=stereo_mix_device,
                                    blocksize=self.buffer_size) as stereo_stream, \
                     sd.InputStream(samplerate=self.sample_rate,
                                    channels=self.channels,
                                    dtype='int16',
                                    device=microphone_device,
                                    blocksize=self.buffer_size) as mic_stream:

                    while self.recording:
                        stereo_data, _ = stereo_stream.read(self.buffer_size)
                        mic_data, _ = mic_stream.read(self.buffer_size)
                        combined_data = np.add(stereo_data, mic_data, dtype=np.int32)

                        # Normalize audio if it exceeds 16-bit range
                        max_value = np.max(np.abs(combined_data))
                        if max_value > 32767:
                            scaling_factor = 32767 / max_value
                            combined_data = (combined_data * scaling_factor).astype(np.int16)
                        else:
                            combined_data = combined_data.astype(np.int16)

                        mp3_data = self.encoder.encode(combined_data.tobytes())
                        mp3.write(mp3_data)

                mp3.write(self.encoder.flush())

            print(f"[INFO] Recording saved as {mp3_file}")
        except Exception as e:
            self.show_error("Recording Error", str(e))

    def stop_recording(self):
        """
        Stop the audio recording process.
        """
        print("[INFO] Stopping recording...")
        logging.info("Stopping recording...")
        self.recording = False
        self.update_tray_title()
        if hasattr(self, "recording_thread"):
            self.recording_thread.join()

    def get_device_by_name(self, device_name):
        """
        Find an audio device by its name.
        """
        devices = sd.query_devices()
        for i, device in enumerate(devices):
            if device_name.lower() in device.get("name", "").lower():
                return i
        return None


def create_tray_icon(recorder):
    """
    Create the system tray icon with context menu.
    """
    def exit_app(icon, item):
        recorder.stop_flag = True
        icon.stop()
        sys.exit(0)

    menu = Menu(
        MenuItem(
            "Record All Meetings",
            lambda icon, item: recorder.toggle_record_all_meetings(),
            checked=lambda item: recorder.record_all_meetings,
        ),
        MenuItem("Settings", lambda icon, item: recorder.show_settings_window()),
        MenuItem("Exit", exit_app),
    )

    icon_path = recorder.get_icon_path()
    icon_image = Image.open(icon_path)
    icon = Icon("Teams Helper", icon_image, "Teams Helper - Idle", menu)
    recorder.set_tray_icon(icon)
    return icon


def main():
    """
    Main entry point for the application.
    """
    recorder = TeamsHelperRecorder()
    tray_icon = create_tray_icon(recorder)

    # Start Teams WebSocket connection in a separate thread
    threading.Thread(target=recorder.connect_to_teams, daemon=True).start()

    signal.signal(signal.SIGINT, lambda sig, frame: tray_icon.stop())

    tray_icon.run()


if __name__ == "__main__":
    main()