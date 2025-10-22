# CMV Local Printing Agent - Installation Guide

Complete installation instructions for the CMV Local Printing Agent Windows service.

## Table of Contents

- [Quick Start](#quick-start)
- [Automated Installation (Recommended)](#automated-installation-recommended)
- [Manual Installation](#manual-installation)
- [Service Management](#service-management)
- [Troubleshooting](#troubleshooting)
- [Uninstallation](#uninstallation)

---

## Quick Start

**For IT Teams - Fastest Method:**

```powershell
# 1. Download the 
Invoke-WebRequest -Uri "https://raw.githubusercontent.com/franchise-factory/cmv-local-printing-agent-releases/master/download-and-install.ps1" -OutFile "download-and-install.ps1"

# 2. Unblock the script
Unblock-File -Path .\download-and-install.ps1

# 3. Run as Administrator
.\download-and-install.ps1
```

That's it! The script will download, extract, and install the latest version automatically.

---

## Automated Installation (Recommended)

### Prerequisites

- Windows 10/11 or Windows Server 2016+
- PowerShell 5.1 or later
- Internet connectivity
- Administrator privileges

### Step-by-Step Instructions

#### 1. Download the Installation Script

**Option A: Direct download from GitHub**
```powershell
Invoke-WebRequest -Uri "https://raw.githubusercontent.com/franchise-factory/cmv-local-printing-agent-releases/master/download-and-install.ps1" -OutFile "download-and-install.ps1"
```

**Option B: Download from releases**
- Go to: https://github.com/franchise-factory/cmv-local-printing-agent-releases/releases
- Download the latest release zip
- Extract and use the included `download-and-install.ps1`

#### 2. Unblock the Script

**IMPORTANT:** Windows blocks downloaded scripts by default. You must unblock it first:

```powershell
Unblock-File -Path .\download-and-install.ps1
```

**Why?** PowerShell's execution policy blocks unsigned scripts downloaded from the internet for security. This script is not digitally signed.

#### 3. Open PowerShell as Administrator

- Right-click **PowerShell** or **Windows Terminal**
- Select **"Run as Administrator"**
- Navigate to the directory containing the script

#### 4. Run the Installation Script

**Install latest version:**
```powershell
.\download-and-install.ps1
```

**Install specific version:**
```powershell
.\download-and-install.ps1 -Version "v1.2.3"
```

**Install to custom location:**
```powershell
.\download-and-install.ps1 -InstallPath "C:\CMV-Agent"
```

**Test download without installing:**
```powershell
.\download-and-install.ps1 -SkipInstaller
```

#### 5. What the Script Does

The automated installer will:
1. Query GitHub for the latest (or specified) release
2. Download the release package
3. Validate all required files are present
4. Extract to the installation directory
5. Run the Windows service installer
6. Install and start the CMV Local Printing Agent service

#### Installation Output

```
CMV Local Printing Agent Installer
===================================

  >> Fetching release information...
  [OK] Found version: v1.2.3

  >> Downloading package...
  [OK] Download complete (2.45 MB)

  >> Extracting files...
  [OK] Extraction complete

  >> Verifying installation files...
  [OK] All files verified

  >> Running installer...

===============================================
  Deployment Summary
===============================================
  Version:    v1.2.3
  Location:   .\cmv-agent
  Downloaded: 2.45 MB
===============================================
```

---

## Manual Installation

If you prefer manual installation or automated download is not available:

### 1. Download Release Package

Go to the releases page:
https://github.com/franchise-factory/cmv-local-printing-agent-releases/releases

Download the latest `cmv-local-printing-agent-{version}-windows.zip`

### 2. Extract Files

```powershell
Expand-Archive -Path .\cmv-local-printing-agent-v1.2.3-windows.zip -DestinationPath C:\CMV-Agent
```

### 3. Unblock the Installer Script

```powershell
cd C:\CMV-Agent
Unblock-File -Path .\windows-service-install.ps1
```

### 4. Run Installer as Administrator

```powershell
.\windows-service-install.ps1
```

### 5. Select Installation Option

You'll see an interactive menu:

```
===============================================
  CMV Local Printing Agent - Service Manager
===============================================

Select an option:

  1. Install Windows Service
  2. Uninstall Windows Service
  3. View Logs
  4. View Configuration
  5. Exit

Enter your choice (1-5):
```

Select **option 1** to install the service.

### Installation Process

The installer will:
1. Check for administrator privileges
2. Verify all required files (.exe, .dll)
3. Install the Windows service
4. Configure the service for automatic startup
5. Start the service
6. Show installation summary

---

## Service Management

### Interactive Menu

Run the installer script to access service management:

```powershell
.\windows-service-install.ps1
```

**Available Options:**

1. **Install Windows Service** - First-time installation or reinstall
2. **Uninstall Windows Service** - Complete removal (with confirmation)
3. **View Logs** - Browse and follow service logs
4. **View Configuration** - Display service configuration and status
5. **Exit** - Close the menu

### Using Windows Services

**Check service status:**
```powershell
Get-Service CMVLocalPrintingAgent
```

**Start the service:**
```powershell
Start-Service CMVLocalPrintingAgent
```

**Stop the service:**
```powershell
Stop-Service CMVLocalPrintingAgent
```

**Restart the service:**
```powershell
Restart-Service CMVLocalPrintingAgent
```

### Log Files

Logs are stored in:
- **Service Log:** `C:\Windows\System32\config\systemprofile\AppData\Local\CMVLocalPrintingAgent\data\logs\service.log`
- **Emergency Log:** `C:\Windows\System32\config\systemprofile\AppData\Local\CMVLocalPrintingAgent\emergency.log`
- **Bootstrap Log:** `C:\Windows\System32\config\systemprofile\AppData\Local\CMVLocalPrintingAgent\bootstrap.log`

**View logs interactively:**
```powershell
.\windows-service-install.ps1
# Select option 3: View Logs
```

**View logs manually:**
```powershell
# Last 50 lines
Get-Content "C:\Windows\System32\config\systemprofile\AppData\Local\CMVLocalPrintingAgent\data\logs\service.log" -Tail 50

# Follow in real-time
Get-Content "C:\Windows\System32\config\systemprofile\AppData\Local\CMVLocalPrintingAgent\data\logs\service.log" -Wait
```

---

## Troubleshooting

### Script Execution Errors

#### "File cannot be loaded because running scripts is disabled"

**Cause:** PowerShell execution policy is too restrictive.

**Solution:**
```powershell
# Check current policy
Get-ExecutionPolicy

# Set to RemoteSigned (recommended)
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
```

#### "File is blocked"

**Cause:** Script was downloaded from the internet and needs to be unblocked.

**Solution:**
```powershell
Unblock-File -Path .\download-and-install.ps1
Unblock-File -Path .\windows-service-install.ps1
```

### Download Issues

#### "GitHub API rate limit exceeded"

**Cause:** Too many requests to GitHub API (60 per hour for unauthenticated requests).

**Solution:**
- Wait one hour before retrying
- Use manual installation instead
- Download directly from browser: https://github.com/franchise-factory/cmv-local-printing-agent-releases/releases

#### "Version not found"

**Cause:** Specified version doesn't exist.

**Solution:**
- Check available versions: https://github.com/franchise-factory/cmv-local-printing-agent-releases/releases
- Use `latest` instead: `.\download-and-install.ps1 -Version "latest"`
- Verify version format (e.g., `v1.2.3` not `1.2.3`)

#### Network connectivity issues

**Solution:**
- Check internet connection
- Verify GitHub is accessible: https://www.githubstatus.com/
- Check firewall/proxy settings
- Try manual download from browser

### Installation Issues

#### "Service failed to start"

**Cause:** Missing dependencies or configuration issues.

**Solution:**
1. Check service status:
   ```powershell
   Get-Service CMVLocalPrintingAgent | Format-List *
   ```

2. Check logs:
   ```powershell
   .\windows-service-install.ps1
   # Select option 3: View Logs
   ```

3. Verify `libusb-1.0.dll` is present in installation directory

4. Check Windows Event Viewer:
   - Open Event Viewer
   - Navigate to: Windows Logs → Application
   - Look for CMVLocalPrintingAgent errors

#### "Access Denied" errors

**Cause:** Not running as Administrator.

**Solution:**
- Close PowerShell
- Right-click PowerShell or Windows Terminal
- Select **"Run as Administrator"**
- Run the script again

#### "Cannot find path" errors

**Cause:** Working directory or file paths incorrect.

**Solution:**
- Verify you're in the correct directory:
  ```powershell
  Get-Location
  ```
- Change to the directory containing the script:
  ```powershell
  cd C:\path\to\cmv-agent
  ```

### Service Runtime Issues

#### Service stops unexpectedly

**Solution:**
1. Check service logs for errors
2. Verify configuration in `config.toml`
3. Check Windows Event Viewer
4. Restart the service:
   ```powershell
   Restart-Service CMVLocalPrintingAgent
   ```

#### USB printer not detected

**Solution:**
1. Verify WinUSB driver is installed for the printer
2. Check service logs for USB errors
3. Ensure printer is connected before starting service
4. Check Device Manager for driver issues

---

## Uninstallation

### Automated Uninstall

```powershell
# Run the installer script
.\windows-service-install.ps1

# Select option 2: Uninstall Windows Service
# Confirm when prompted
```

### Manual Uninstall

```powershell
# Stop the service
Stop-Service CMVLocalPrintingAgent

# Remove the service
sc.exe delete CMVLocalPrintingAgent

# Remove installation directory
Remove-Item -Path "C:\Windows\System32\config\systemprofile\AppData\Local\CMVLocalPrintingAgent" -Recurse -Force
```

### What Gets Removed

**Automatically removed:**
- Windows service registration
- Service executable and libraries
- Configuration files (`config.toml`)
- Log files (service.log, emergency.log, bootstrap.log)
- Database files (telemetry.db, queue.db)
- Installation directory

**NOT removed (intentionally):**
- WinUSB drivers - To avoid breaking other USB devices that may use them

---

## Advanced Configuration

### Custom Installation Parameters

The installer supports custom configuration:

```powershell
# Custom API URL
.\windows-service-install.ps1 -APIURL "https://custom-api.example.com"

# Custom log level
.\windows-service-install.ps1 -LogLevel "debug"

# Custom executable path (if needed)
.\windows-service-install.ps1 -ExePath "C:\Custom\Path\cmv-local-printing-agent.exe"
```

### Configuration File

After installation, configuration is stored in:
`C:\Windows\System32\config\systemprofile\AppData\Local\CMVLocalPrintingAgent\config.toml`

View configuration:
```powershell
.\windows-service-install.ps1
# Select option 4: View Configuration
```

---

## Support

### Getting Help

- **Installation issues:** Review this README and troubleshooting section
- **Service errors:** Check log files for detailed error messages
- **Feature requests:** Contact your IT department or development team

### Version Information

To check installed version:
```powershell
.\windows-service-install.ps1
# Select option 4: View Configuration
# Version shown in service information
```

To check available versions:
- Visit: https://github.com/franchise-factory/cmv-local-printing-agent-releases/releases

---

## Package Contents

Each release package contains:

| File | Purpose |
|------|---------|
| `cmv-local-printing-agent.exe` | Main service executable |
| `libusb-1.0.dll` | USB communication library |
| `windows-service-install.ps1` | Interactive service installer/manager |

---

## System Requirements

- **Operating System:** Windows 10/11 or Windows Server 2016+
- **PowerShell:** Version 5.1 or later
- **Privileges:** Administrator access required
- **Disk Space:** ~50 MB for installation
- **Network:** Internet connectivity for downloads (automated installation only)

---

## Security Notes

### Script Signing

The PowerShell scripts are **not digitally signed**. You must unblock them before execution:

```powershell
Unblock-File -Path .\download-and-install.ps1
Unblock-File -Path .\windows-service-install.ps1
```

### Execution Policy

Recommended PowerShell execution policy:
```powershell
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
```

This allows locally created scripts and signed remote scripts to run.

### Administrator Privileges

Installation requires administrator privileges because:
- Windows service registration requires admin rights
- Installation to system directories requires elevation
- WinUSB driver operations require administrative access

---

## License & Distribution

- **Private Source Code:** Development occurs in private repository
- **Public Releases:** Compiled binaries available at https://github.com/franchise-factory/cmv-local-printing-agent-releases
- **Distribution:** Authorized for internal use only

---

**Last Updated:** 2025-10-22
**Installation Guide Version:** 2.0
