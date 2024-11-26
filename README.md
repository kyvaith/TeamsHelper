
# Microsoft Teams Helper App

## Overview
The **Microsoft Teams Helper App** is a desktop application that integrates with the Microsoft Teams client to record audio during meetings. It provides seamless audio recording by combining inputs from your microphone and the Stereo Mix.

### Key Features
- **Automatic Audio Recording**: Starts recording automatically when the meeting state changes.
- **Changeable Recording Folder**: Modify the default recording folder through the settings menu.
- **Lightweight Integration**: Uses the local Teams client API to avoid additional permissions or configurations.
- **Logs**: Logs are saved in the recording folder for troubleshooting.
- **Mouse Jiggler**: Enable the "Keep Available Status" option in the application's context menu to maintain a green status in Teams.

---

## Setup
1. Download TeamsHelper-v1.x.exe from ([Releases](https://github.com/kyvaith/TeamsHelper/releases)) to your Windows computer that runs the New Teams client
2. In Microsoft Teams, enable the Third-Party
  API ([see Microsoft documentation](https://support.microsoft.com/en-us/office/connect-to-third-party-devices-in-microsoft-teams-aabca9f2-47bb-407f-9f9b-81a104a883d6?storagetype=live))
3. Enable Stereo Mix in your system
4. Run the application. You can find it in the system tray
 
## Enabling Stereo Mix on Windows

To ensure proper functionality, you may need to enable Stereo Mix on your system:
1. Right-click on the speaker icon in the system tray.
2. Select **"Sounds"**.
3. Navigate to the **"Recording"** tab.
4. If **"Stereo Mix"** or **"Sound Mixer"** is listed but disabled, right-click it and choose **"Enable"**.
5. If it's not listed:
   - Right-click in an empty space and check both **"Show Disabled Devices"** and **"Show Disconnected Devices"**.
   - Look for **"Stereo Mix"** or **"Sound Mixer"**, right-click, and select **"Enable"**.
6. Click **"OK"** to save your changes.

---

## Notices
- **Pull Requests, Issues, and Feature Requests are welcomed!**
- This integration supports only the **New Teams (2.0 client)**.
- **Logs** are stored in `app.log` in the recordings folder.
- The **recordings folder** can be changed through the settings menu.
- The app uses the **local Teams client API** instead of Azure/M365 for simplicity.

### Advantages of Local Teams API
- **No permissions or app registrations** required in Azure.
- Organizations don't need to approve additional exceptions for security policies.

### Limitations of Local Teams API
- The app cannot detect meetings joined via other clients, such as mobile or web.

---

## API Details

### Example Connections
1. **Before Token Acquisition**:
   ```
   ws://localhost:8124?protocol-version=2.0.0&manufacturer=Kyvaith&device=TeamsHelper&app=TeamsHelper&app-version=1.0
   ```
2. **After Token Acquisition**:
   ```
   ws://localhost:8124?token=<TOKEN>&protocol-version=2.0.0&manufacturer=Kyvaith&device=TeamsHelper&app=TeamsHelper&app-version=1.0
   ```

### Example Updates and Requests
1. **Teams -> Client Update**:
   ```json
   {
     "meetingUpdate": {
       "meetingState": {
         "isMuted": false,
         "isVideoOn": false,
         "isHandRaised": false,
         "isInMeeting": true,
         "isRecordingOn": false,
         "isBackgroundBlurred": false,
         "isSharing": false,
         "hasUnreadMessages": false
       },
       "meetingPermissions": {
         "canToggleMute": true,
         "canToggleVideo": true,
         "canToggleHand": true,
         "canToggleBlur": false,
         "canLeave": true,
         "canReact": true,
         "canToggleShareTray": true,
         "canToggleChat": true,
         "canStopSharing": false,
         "canPair": false
       }
     }
   }
   ```

2. **Client -> Teams Request Toggle Mute**:
   ```json
   {
     "requestId": 1,
     "apiVersion": "2.0.0",
     "action": "toggle-mute"
   }
   ```

3. **Teams -> Client Request Confirmation**:
   ```json
   {
     "requestId": 2,
     "response": "Success"
   }
   ```

4. **Client -> Teams Token Refresh**:
   ```json
   {
     "tokenRefresh": "132cf8d9-403f-4d71-9067-8390160224cb"
   }
   ```

### Reference Document (Legacy Teams)
For more details on the legacy Teams API, refer to the [Microsoft Teams WebSocket API Documentation](https://lostdomain.notion.site/Microsoft-Teams-WebSocket-API-5c042838bc3e4731bdfe679e864ab52a).

---

## License

This project is licensed under the **Apache License 2.0**. See the [LICENSE](LICENSE) file for details.

---

## Contributing

Contributions are welcome! Please open issues or pull requests for new features, bug fixes, or documentation improvements.
