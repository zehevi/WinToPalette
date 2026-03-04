# WinToPalette

A Windows 11 application that launches the PowerToys Command Palette when you press the Windows key, replacing the default Start Menu.

## Features

- ✅ Detects Windows key press at OS level using kernel-mode input interception
- ✅ Suppresses the default Start Menu
- ✅ Launches PowerToys Command Palette directly
- ✅ Auto-starts on system boot

## Requirements

- **Windows 11** (Windows 10 may work with modifications)
- **PowerToys** installed
- **Administrator privileges** (application requests UAC elevation)
- **Interception kernel driver** (auto-installed during setup)

## Installation

### Quick Install (Recommended)

1. Download the latest release from the [Releases](https://github.com/zehevi/WinToPalette/releases) page
2. Extract the ZIP file
3. Double-click `Install-WinToPalette.bat`
4. Click "Yes" when prompted for administrator privileges
5. Press `y` to confirm installation
6. **Reboot** your computer (required for Interception driver activation)

### PowerShell Installation

1. Extract the release ZIP file
2. Open PowerShell as Administrator
3. Navigate to the extracted folder
4. Run: `.\Install-WinToPalette.ps1`
5. **Reboot** your computer

## Usage

Simply press the **Windows key** at your desktop - PowerToys Command Palette will launch instead of the Start Menu.

## Uninstallation

1. Double-click `Uninstall-WinToPalette.bat`
2. Click "Yes" when prompted for administrator privileges
3. Press `y` to confirm uninstallation
4. Optionally remove log files when prompted
5. **Reboot** your computer

The uninstaller will:
- Stop any running WinToPalette processes
- Remove the startup task
- Delete application files
- Optional: Clean up log files

## Troubleshooting

### Windows key doesn't work
- **Solution**: Check that PowerToys is installed, reboot your computer, and verify Interception driver was installed
- **Logs**: Check `%APPDATA%\WinToPalette\` for error messages

### PowerToys isn't detected
- **Solution**: Install PowerToys from the Microsoft Store or [GitHub](https://github.com/microsoft/PowerToys)

### Installation fails
- **Solution**: Ensure you're running the installer as Administrator and have a working internet connection

### Application won't start
- **Solution**: Right-click the .exe and select "Run as administrator", or re-run the installer

## Logging

Application logs are saved to `%APPDATA%\WinToPalette\`. You can view them in PowerShell:
```powershell
Get-Content "$env:APPDATA\WinToPalette\*.log" -Tail 50
```

## Building from Source

### Prerequisites
- Visual Studio 2022 or VS Code
- .NET 8.0 SDK

### Build Steps
```powershell
git clone https://github.com/zehevi/WinToPalette.git
cd WinToPalette
dotnet build -c Release
```

## Technical Details

This application uses the **Interception kernel driver** to:
- Intercept keyboard input at the OS level (before Windows processes it)
- Detect when the Windows key (VK_LWIN/VK_RWIN) is pressed
- Suppress the keystroke (prevent the Start Menu from appearing)
- Trigger the PowerToys Command Palette launch

Windows 11 processes the Windows key deeply in the input stack, so traditional user-mode hooks cannot reliably intercept it. The kernel-level Interception driver provides true suppression.

## Architecture

**Interception Driver** → Kernel-level input interception  
**InterceptionManager** → Manages driver connection and keyboard events  
**PowerToysLauncher** → Detects and launches PowerToys  
**StartupManager** → Handles UAC elevation and auto-launch registration

## Disclaimer

This application requires kernel-level access to input devices. Use at your own risk.

## References

- [Interception Library](https://github.com/oblitum/Interception) - Kernel driver for keyboard interception
- [PowerToys](https://github.com/microsoft/PowerToys) - Microsoft productivity tools
- [Windows API - Keyboard Input](https://docs.microsoft.com/en-us/windows/win32/inputdev/keyboard-input)
