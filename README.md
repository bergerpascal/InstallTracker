# InstallTracker

A PowerShell tool for tracking and analyzing system changes on Windows computers. Creates detailed snapshots of installed programs, services, tasks, registry entries, file systems, and other system components.

## 🎨 Created with Vibe Coding

This entire project was created using **Vibe Coding** - no lines of code were manually written. All development, debugging, and optimization was done using VS Code and AI Agents. This demonstrates the power of modern AI-assisted development workflows.

## 📋 Overview

**InstallTracker** is a configuration management system that documents and analyzes system changes. It enables you to perform "before-and-after" comparisons and quickly identify what changes were made to a system.

## 🎯 Core Applications

### 1. InstallTracker.ps1 (Main Application)

The primary script with a modern graphical user interface for system change tracking:

#### Features:
- **Create Snapshots**: Captures the current system state in CSV and JSON format
- **Pre/Post Analysis**: Create a PRE snapshot, make changes, then create a POST snapshot to see what changed
- **Comprehensive Tracking** of:
  - 📦 Installed Programs (Uninstall Registry Keys)
  - 🔧 Windows Services
  - 📅 Scheduled Tasks
  - 🔐 Registry Run/RunOnce Keys
  - 📁 Folders and Directory Structure
  - 🚀 Shortcuts and Startup Programs
  - 📄 All Files with detailed metadata

#### Advanced Features:
- **Custom Scan Paths**: Define any root paths for tracking custom installations
  - Supports environment variables (%USERPROFILE%, %APPDATA%, etc.)
  - Handles hidden folders (AppData, System folders, etc.)
  - Persistent configuration in JSON format
  - GUI-based path management
  
- **Detailed Change Reports**: 
  - Automatic before/after comparison
  - Formatted HTML reports with all changes listed
  - Separate counts for added/removed items by category
  - Timestamp-based file organization
  
- **Automatic Version Checking**: 
  - Checks GitHub for new versions at startup
  - Automatic download and installation
  - Seamless update process with backup

- **Modern UI**: 
  - WPF-based graphical interface
  - Color-coded status messages
  - Real-time progress updates
  - Settings management panel

---

### 2. InstallTracker-TestData.ps1 (Testing Helper)

A companion script for generating and cleaning up test data:

#### Purpose:
- **Development Testing**: Generate test files and registry entries for development
- **Training**: Create sample installations to test InstallTracker functionality
- **Validation**: Verify that InstallTracker correctly tracks changes

#### Capabilities:
- **Create Test Data**: Generates realistic system changes:
  - Test folders in `C:\Program Files\`
  - Test files in `%LOCALAPPDATA%\`
  - Registry entries simulating installations
  - Test shortcuts and startup programs
  - Sample service entries

- **Delete Test Data**: Cleans up all created test data:
  - Removes test folders and files
  - Cleans up registry entries
  - Restores system to clean state

- **Safe Operation**:
  - Only creates/removes data it created
  - Confirmation prompts before deletion
  - Error handling and rollback capabilities

---

### 3. InstallTracker.bat (Launcher Script)

A convenient batch file to launch InstallTracker with proper permissions:

#### Features:
- **Easy Execution**: Just Double-click or Right-click and select "Run as administrator"
- **ExecutionPolicy Bypass**: Handles PowerShell execution policy automatically
- **Error Handling**: Checks for required files and shows helpful error messages
- **Same Directory Detection**: Automatically finds InstallTracker.ps1 in the same folder

#### Usage:
1. Place `InstallTracker.bat` in the same folder as `InstallTracker.ps1`
2. Double-click on `InstallTracker.bat` or right- click and Select **"Run as administrator"**
3. InstallTracker will start

---

## 🚀 Getting Started

### Requirements

- Windows PowerShell 5.1 or higher
- .NET Framework 4.7.2+ (for JSON processing and WPF UI)
- Windows 10 or later
- Optional Administrator rights (for full access to registry, services, and system folders)

### Installation

1. Download all files to a folder:
   - `InstallTracker.ps1` (main application)
   - `InstallTracker.bat` (launcher) **← Easiest way to start**
   - `InstallTracker-TestData.ps1` (optional - for testing)

2. **Option A - Using Batch File (Recommended)**:
   - Right-click `InstallTracker.bat` → **Run as administrator**
   - InstallTracker will start

3. **Option B - Using PowerShell Directly**:
   ```powershell
   powershell -NoProfile -ExecutionPolicy Bypass -File "C:\path\to\InstallTracker.ps1"
   ```

### Basic Usage

#### Using InstallTracker (Main Script)

1. **First Run - Create PRE Snapshot**:
   - Click the **"PRE"** button
   - System will scan and capture current state
   - Results saved to `_Snapshots\Reports_Pre\`

2. **Make System Changes**:
   - Install new software
   - Add/delete files
   - Modify settings
   - Any other changes you want to track

3. **Create POST Snapshot**:
   - Click the **"POST"** button
   - System will scan again and compare with PRE
   - Automatic report generated showing all changes

4. **Review Report**:
   - Report opens automatically (option to open)
   - Shows all added/removed items by category
   - Save report for documentation

#### Configuring Scan Paths and other settings

1. Click **"SETTINGS"** button
2. In the settings panel:
   - **Add Path**: Enter folder path and click "Add"
   - **Remove Path**: Select path and click "Remove"
   - **Enable/Disable Scans**: Check/uncheck scan options
   - **Version Check**: Enable automatic update checking
3. Click **"SAVE SETTINGS"** to persist changes

#### Using TestData Script (Optional)

```powershell
# Create test installations and files
& "C:\path\to\InstallTracker-TestData.ps1" -Action Create

# Later: Clean up all test data
& "C:\path\to\InstallTracker-TestData.ps1" -Action Delete
```

#### Quick Start with Batch File

The easiest way to start InstallTracker:

1. Locate `InstallTracker.bat` in your folder
2. Right-click on it
3. Select **"Run as administrator"**
4. InstallTracker starts!

The batch file handles:
- ✅ ExecutionPolicy bypass automatically
- ✅ Error checking for required files
- ✅ Proper exit code handling

## 📁 Directory Structure

```
InstallTracker/
├── InstallTracker.ps1                 # Main application (GUI)
├── InstallTracker-TestData.ps1        # Test data generator
├── InstallTracker-Config.json         # Configuration (auto-created)
├── README.md                          # This file
└── _Snapshots/                        # Generated reports
    ├── Reports_Pre/                   # PRE snapshot files
    │   ├── services_pre_*.csv/.json
    │   ├── tasks_pre_*.csv/.json
    │   ├── runkeys_pre_*.csv/.json
    │   ├── uninstall_pre_*.csv/.json
    │   ├── folders_pre_*.csv/.json
    │   ├── shortcuts_pre_*.csv/.json
    │   └── files_pre_*.csv/.json
    │
    ├── Reports_Post/                  # POST snapshot files
    │   ├── services_post_*.csv/.json
    │   ├── tasks_post_*.csv/.json
    │   └── ... (same structure as Pre)
    │
    └── ChangeReport_*.html            # Final comparison report
```

## ⚙️ Configuration

### InstallTracker-Config.json

Automatically created on first settings change save (if nothing is changed and saved the default values are used). Stores your preferences:

```json
{
  "rootPaths": [
    "%USERPROFILE%",
    "%APPDATA%\\Microsoft\\Windows\\Start Menu",
    "C:\\Program Files",
    "C:\\ProgramData",
    "C:\\Users\\Public\\Desktop",
    "C:\\Program Files (x86)"
  ],
  "scanOptions": {
    "services": true,
    "runKeys": true,
    "uninstallKeys": true,
    "startMenuShortcuts": true,
    "scheduledTasks": true
  },
  "checkVersions": true,
  "gitHubRepository": "bergerpascal/InstallTracker"
}
```

### Supported Environment Variables

The system automatically expands these environment variables:

| Variable | Example Value |
|----------|---------------|
| `%USERPROFILE%` | `C:\Users\Username` |
| `%APPDATA%` | `C:\Users\Username\AppData\Roaming` |
| `%LOCALAPPDATA%` | `C:\Users\Username\AppData\Local` |
| `%ProgramFiles%` | `C:\Program Files` |
| `%ProgramFiles(x86)%` | `C:\Program Files (x86)` |
| `%ProgramData%` | `C:\ProgramData` |
| `%WINDIR%` | `C:\Windows` |

## 📊 Output Formats

### CSV Format
- Comma-Separated Values
- Compatible with Excel, databases, analysis tools
- Easy to filter and sort
- Human-readable

### JSON Format
- Structured data format
- Ideal for programmatic analysis
- API integration ready
- Long-term archiving
- Deep nesting for complex data

### HTML Report
- Human-readable summary in browser format
- Organized by change type
- Includes counts and details
- Ready for printing or documentation

## 🔍 How It Works

### PRE Snapshot Process:
1. Scans all configured root paths
2. Recursively reads folders and files (including hidden folders)
3. Queries Windows Services registry
4. Reads Scheduled Tasks
5. Exports Run/RunOnce registry keys
6. Exports Uninstall registry keys
7. Captures all data to CSV and JSON files
8. Saves with timestamp for organization

### POST Snapshot Process:
1. Repeats the scanning process
2. Compares each category against PRE snapshot
3. Identifies added items (not in PRE)
4. Identifies removed items (not in POST)
5. Generates detailed comparison report
6. Shows summary statistics

### Comparison Logic:
- **Added**: Exists in POST but not in PRE
- **Removed**: Exists in PRE but not in POST
- **Unchanged**: Exists in both (not reported)

## 🛠️ Troubleshooting

### "Execution Policy" Error

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### Administrator Rights Required

```powershell
# Open PowerShell with Administrator rights, then:
& "C:\path\to\InstallTracker.ps1"
```

### AppData and Hidden Folders Not Being Scanned

- The script uses `-Force` parameter to include hidden folders
- If you still have issues, try adding explicit paths:
  - `%USERPROFILE%\AppData\Local`
  - `%USERPROFILE%\AppData\Roaming`

### Reset Configuration

```powershell
Remove-Item "InstallTracker-Config.json"
# Restart script - new config will be created with defaults
```

### Reports Not Generating

- Ensure `_Snapshots` folder is writable
- Check Administrator rights
- Verify sufficient disk space
- Check antivirus isn't blocking file operations

### Key Features Implemented:
- ✅ Full GUI with modern WPF interface
- ✅ Service, Task, Registry, and File scanning
- ✅ Recursive folder traversal with error handling
- ✅ Hidden folder support (AppData, System folders)
- ✅ Before/after comparison and reporting
- ✅ Configuration management via JSON
- ✅ Automatic version checking and updates
- ✅ Comprehensive status and progress reporting
- ✅ CSV and JSON export formats
- ✅ Error recovery and resilience

## 💡 Tips & Best Practices

1. **Regular Snapshots**: Take PRE snapshots regularly to establish baselines
2. **Clear Documentation**: Always save reports for later reference
3. **Safe Testing**: Use TestData script in controlled environments first
4. **Backup Important Data**: Before making system changes
5. **Run as Admin**: Always run with Administrator privileges
6. **Review Changes**: Always review the report before taking action

## 📄 License

Use freely for personal and commercial purposes.

---

**Created with Vibe Coding** - Entirely developed using VS Code and AI Agents

## 📝 Documentation History

**Version:** 1.0.1  
**Last Updated:** January 23, 2026  