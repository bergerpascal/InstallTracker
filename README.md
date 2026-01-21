# InstallTracker

A PowerShell tool for tracking and analyzing system changes on Windows computers. Creates detailed snapshots of installed programs, services, tasks, registry entries, and other system components.

## 📋 Overview

**InstallTracker** is a configuration management system that documents and analyzes system changes. It enables you to perform "before-and-after" comparisons and quickly identify what changes were made to a system.

## 🎯 Key Features

### InstallTracker.ps1

The main script with graphical user interface:

- **Create Snapshots**: Captures the current system state in CSV and JSON format
- **Configurable Tracking**: Choose which components to monitor:
  - 📦 Installed Programs (Uninstall Registry)
  - 🔧 Windows Services
  - 📅 Scheduled Tasks
  - 🔐 Registry Run Keys
  - 🚀 Shortcuts and Startup Programs
  - ✅ Version Checks

- **Custom Paths**: Define root paths for tracking custom installations
  - Supports environment variables (%USERPROFILE%, %APPDATA%, etc.)
  - Persistent configuration in JSON
  
- **Comprehensive Reports**: Detailed output files in CSV and JSON format
  - Timestamp-based file names
  - Structured data for easy analysis

- **Automatic Updates**: Checks GitHub for new versions and installs updates automatically

### InstallTracker-TestData.ps1

Helper script for generating test data:

- **Create Test Data**: Generates files and directories at various system locations:
  - `C:\Program Files\` simulations
  - `%LOCALAPPDATA%\` directory structures
  - Registry entries for versions
  - Shortcuts and startup programs

- **Delete Test Data**: Cleans up all created test data and restores the clean state

- **Automatic Updates**: Like InstallTracker with GitHub integration

## 🚀 Getting Started

### Requirements

- Windows PowerShell 5.1 or higher
- Administrator rights (for full access to registry and services)
- .NET Framework 4.7.2+ (for JSON processing)


### Basic Usage

#### Start InstallTracker

```powershell
# With GUI
powershell -NoProfile -ExecutionPolicy Bypass -File "InstallTracker.ps1"
```

The GUI will be displayed with the following options:

1. **Enable Categories**: Check boxes for desired tracking categories
2. **Configure Root Paths**: 
   - Click "SETTINGS" button
   - Add new paths (supports environment variables)
   - Save
3. **Create Snapshot**: Click "PRE" or "POST" button to generate reports

Reports are saved in the `_Snapshots/` directory.

#### Generate Test Data

```powershell
# Create test data
powershell -NoProfile -ExecutionPolicy Bypass -File "InstallTracker-TestData.ps1" -Action Create

# Delete test data
powershell -NoProfile -ExecutionPolicy Bypass -File "InstallTracker-TestData.ps1" -Action Delete
```

## 📁 Directory Structure

```
InstallTracker/
├── InstallTracker.ps1              # Main application with GUI
├── InstallTracker-TestData.ps1     # Test data generator
├── InstallTracker-Config.json      # Configuration file (auto-created)
└── _Snapshots/                     # Generated reports
    └── Reports_Pre/                # Snapshot files
        ├── services_*.csv          # Services export
        ├── services_*.json
        ├── runkeys_*.csv           # Registry exports
        ├── runkeys_*.json
        ├── tasks_*.csv             # Tasks exports
        ├── tasks_*.json
        ├── uninstall_*.csv         # Programs exports
        └── uninstall_*.json
```

## ⚙️ Configuration

### InstallTracker-Config.json

The configuration is automatically created and saved:

```json
{
  "TrackingPaths": {
    "RootPaths": ["%USERPROFILE%", "%APPDATA%", "%LOCALAPPDATA%"],
    "IncludeSubfolders": true
  },
  "Categories": {
    "TrackServices": true,
    "TrackTasks": true,
    "TrackRunKeys": true,
    "TrackUninstallKeys": true,
    "TrackShortcuts": true,
    "CheckVersions": true
  }
}
```

### Environment Variables in Paths

The system supports automatic resolution of environment variables:

- `%USERPROFILE%` → C:\Users\Username
- `%APPDATA%` → C:\Users\Username\AppData\Roaming
- `%LOCALAPPDATA%` → C:\Users\Username\AppData\Local
- `%ProgramFiles%` → C:\Program Files
- `%ProgramFiles(x86)%` → C:\Program Files (x86)

## 🔄 Comparison and Analysis

After creating multiple snapshots, you can compare them:

1. Snapshots from different time points available in `_Snapshots/Reports_Pre/`
2. Open CSV or JSON files in an editor or Excel
3. Compare entries to identify changes

**Example Analysis:**
- Search for new programs in `uninstall_*.csv`
- Identify new services in `services_*.csv`
- Check changes in registry keys

## 🔄 Auto-Update

Both scripts can automatically check GitHub for newer versions:

- Version is checked at startup
- If a newer version is available, an update is offered
- Updates are automatically downloaded, installed, and the script is restarted
- Prevents infinite loops using environment variables

## 📊 Output Formats

### CSV Format
Standard Comma-Separated-Values files, openable in:
- Excel
- Notepad
- All text editors
- Database tools

### JSON Format
Structured data for:
- Programmatic analysis
- Import into databases
- API integration
- Long-term archiving

## 🛠️ Troubleshooting

### "Execution Policy" Error
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### Administrator Rights Required
```powershell
# Open PowerShell with Administrator rights
powershell -NoProfile -ExecutionPolicy Bypass -File "InstallTracker.ps1"
```

### Reset Old Configuration
```powershell
Remove-Item "InstallTracker-Config.json"
# Restart script - new config will be created
```

**Version:** 1.0.1  
**Last Updated:** 2026-01-21
