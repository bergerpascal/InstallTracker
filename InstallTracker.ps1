<#
.SYNOPSIS
  Takes before/after snapshots and produces diffs for Windows Services,
  Scheduled Tasks, Run/RunOnce registry entries, folders, and shortcuts.

.NOTES
  PowerShell 5.1+; run as Administrator for best coverage.
#>

# Script version
$scriptVersion = "1.0.4"

# Determine script directory - works even when sourced
$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }

# Set default snapshot root directory
$SnapshotRoot = "_Snapshots"

# Load config from JSON if RootPaths not provided
$configFile = Join-Path $scriptDir "InstallTracker-Config.json"
if (Test-Path $configFile) {
  try {
    $config = Get-Content $configFile -Raw | ConvertFrom-Json
    $RootPaths = $config.rootPaths | ForEach-Object {
      # Expand environment variables in paths (e.g., %USERPROFILE% -> C:\Users\Username)
      [System.Environment]::ExpandEnvironmentVariables($_)
    }
    $ScanOptions = $config.scanOptions
    $CheckVersions = if ($null -ne $config.checkVersions) { $config.checkVersions } else { $true }
    $GitHubRepository = $config.gitHubRepository
  } catch {
    # Fallback if JSON parsing fails
    $RootPaths = @(
      "%USERPROFILE%",
      "%APPDATA%\Microsoft\Windows\Start Menu",
      "C:\Program Files",
      "C:\ProgramData",
      "C:\Users\Public\Desktop",
      "C:\Program Files (x86)"
    ) | ForEach-Object { [System.Environment]::ExpandEnvironmentVariables($_) }
    $ScanOptions = @{
      services = $true
      runKeys = $true
      uninstallKeys = $true
      startMenuShortcuts = $true
      scheduledTasks = $true
    }
    $CheckVersions = $true
    $GitHubRepository = $null
  }
} else {
  # Default values if InstallTracker-Config.json doesn't exist
  $RootPaths = @(
    "%USERPROFILE%",
    "%APPDATA%\Microsoft\Windows\Start Menu",
    "C:\Program Files",
    "C:\ProgramData",
    "C:\Users\Public\Desktop",
    "C:\Program Files (x86)"
  ) | ForEach-Object { [System.Environment]::ExpandEnvironmentVariables($_) }
  $ScanOptions = @{
    services = $true
    runKeys = $true
    uninstallKeys = $true
    startMenuShortcuts = $true
    scheduledTasks = $true
  }
  $CheckVersions = $true
  $GitHubRepository = "bergerpascal/InstallTracker"
}


# Function to check for version updates on GitHub
function Test-GitHubVersionCheck {
  param(
    [string]$Repository,
    [string]$CurrentVersion
  )
  
  if ([string]::IsNullOrWhiteSpace($Repository)) {
    return $null
  }
  
  try {
    $uri = "https://api.github.com/repos/$Repository/releases/latest"
    $response = Invoke-WebRequest -Uri $uri -UseBasicParsing -ErrorAction Stop
    $releaseInfo = $response.Content | ConvertFrom-Json
    
    if ($releaseInfo.tag_name) {
      $latestVersion = $releaseInfo.tag_name -replace '^v', ''
      
      # Compare versions
      if ([version]$latestVersion -gt [version]$CurrentVersion) {
        return @{
          LatestVersion = $latestVersion
          DownloadUrl = $releaseInfo.html_url
          ReleaseNotes = $releaseInfo.body
        }
      }
    }
  } catch {
    # Silently fail if update check doesn't work
    Write-Verbose "Update check failed: $_"
  }
  
  return $null
}

# Check if running as 64-bit process, if not, restart as 64-bit
if (-not [System.Environment]::Is64BitProcess) {
  # Prevent infinite loop with environment variable
  if (-not $env:SYSTEM_DIFF_64BIT_RESTART) {
    Write-Host "Restarting script as 64-bit process..."
    $scriptPath = $MyInvocation.MyCommand.Path
    $env:SYSTEM_DIFF_64BIT_RESTART = "1"
    & "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -File $scriptPath
    exit
  } else {
    Write-Host "Warning: Could not start as 64-bit process. Continuing with 32-bit."
  }
}

# Define Run/RunOnce registry keys to scan (x64 and x86)
$runKeyPaths = @(
  'HKLM:\Software\Microsoft\Windows\CurrentVersion\Run',
  'HKLM:\Software\Microsoft\Windows\CurrentVersion\RunOnce',
  'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Run',
  'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\RunOnce',
  'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run',
  'HKCU:\Software\Microsoft\Windows\CurrentVersion\RunOnce'
)

# Define Uninstall registry keys to scan (x64, x86, and user)
$uninstallKeyPaths = @(
  'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall',
  'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall',
  'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall'
)

# --- GUI with persistent window and status display ---
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore

# Function to show version check notification
function Show-VersionCheckDialog {
  param(
    [hashtable]$UpdateInfo
  )
  
  Add-Type -AssemblyName System.Windows.Forms
  [System.Windows.Forms.Application]::EnableVisualStyles()
  
  $form = New-Object System.Windows.Forms.Form
  $form.Text = "InstallTracker Version Check"
  $form.Size = New-Object System.Drawing.Size(500, 300)
  $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
  $form.TopMost = $true
  $form.Font = New-Object System.Drawing.Font("Segoe UI", 10)
  $form.BackColor = [System.Drawing.Color]::FromArgb(255, 248, 249, 250)
  
  # Title
  $titleLabel = New-Object System.Windows.Forms.Label
  $titleLabel.Text = "A new version is available!"
  $titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
  $titleLabel.Location = New-Object System.Drawing.Point(20, 20)
  $titleLabel.Size = New-Object System.Drawing.Size(450, 30)
  $form.Controls.Add($titleLabel)
  
  # Version info
  $infoLabel = New-Object System.Windows.Forms.Label
  $infoLabel.Text = "Current version: $scriptVersion`r`nAvailable version: $($UpdateInfo.LatestVersion)"
  $infoLabel.Location = New-Object System.Drawing.Point(20, 60)
  $infoLabel.Size = New-Object System.Drawing.Size(450, 50)
  $infoLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
  $form.Controls.Add($infoLabel)
  
  # Release notes
  $notesLabel = New-Object System.Windows.Forms.Label
  $notesLabel.Text = "Release Notes:"
  $notesLabel.Location = New-Object System.Drawing.Point(20, 120)
  $notesLabel.Size = New-Object System.Drawing.Size(450, 20)
  $notesLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
  $form.Controls.Add($notesLabel)
  
  # Release notes text box (readonly)
  $notesBox = New-Object System.Windows.Forms.TextBox
  $notesBox.Text = $UpdateInfo.ReleaseNotes
  $notesBox.ReadOnly = $true
  $notesBox.Multiline = $true
  $notesBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
  $notesBox.Location = New-Object System.Drawing.Point(20, 145)
  $notesBox.Size = New-Object System.Drawing.Size(450, 80)
  $notesBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
  $form.Controls.Add($notesBox)
  
  # Buttons
  $yesButton = New-Object System.Windows.Forms.Button
  $yesButton.Text = "Go to Release"
  $yesButton.Location = New-Object System.Drawing.Point(240, 240)
  $yesButton.Size = New-Object System.Drawing.Size(110, 30)
  $yesButton.BackColor = [System.Drawing.Color]::FromArgb(255, 59, 130, 246)
  $yesButton.ForeColor = [System.Drawing.Color]::White
  $yesButton.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
  $yesButton.DialogResult = [System.Windows.Forms.DialogResult]::Yes
  $form.Controls.Add($yesButton)
  
  $noButton = New-Object System.Windows.Forms.Button
  $noButton.Text = "Later"
  $noButton.Location = New-Object System.Drawing.Point(360, 240)
  $noButton.Size = New-Object System.Drawing.Size(110, 30)
  $noButton.BackColor = [System.Drawing.Color]::FromArgb(255, 209, 213, 219)
  $noButton.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
  $noButton.DialogResult = [System.Windows.Forms.DialogResult]::No
  $form.Controls.Add($noButton)
  
  $form.AcceptButton = $yesButton
  $form.CancelButton = $noButton
  
  $result = $form.ShowDialog()
  $form.Dispose()
  
  return $result -eq [System.Windows.Forms.DialogResult]::Yes
}

$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" 
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="InstallTracker (v$scriptVersion)" 
        WindowStartupLocation="CenterScreen"
        Height="650" Width="750"
        Background="#F8F9FA"
        FontFamily="Segoe UI"
        FontSize="12">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
        </Grid.RowDefinitions>
        
        <!-- Modern Header -->
        <Grid Grid.Row="0" Background="#1F2937">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <StackPanel Grid.Column="0" VerticalAlignment="Center" Margin="25,20,0,20">
                <TextBlock Text="InstallTracker" 
                           FontSize="20" FontWeight="Bold" Foreground="White"/>
                <TextBlock Text="Installation Change Tracker" 
                           FontSize="11" Foreground="#9CA3AF" 
                           Margin="0,5,0,0"/>
            </StackPanel>
            <TextBlock Grid.Column="1" Text="v$scriptVersion" 
                       FontSize="12" Foreground="#6B7280" FontWeight="SemiBold"
                       VerticalAlignment="Center" HorizontalAlignment="Right" Margin="0,0,25,0"/>
        </Grid>
        
        <!-- Control Panel -->
        <Border Grid.Row="1" Background="#FFFFFF" BorderThickness="0,1,0,0" BorderBrush="#E5E7EB">
            <StackPanel Margin="25,25,25,25">
                <TextBlock Text="Operation Mode" 
                           FontSize="13" FontWeight="SemiBold" Foreground="#1F2937" 
                           Margin="0,0,0,15"/>
                
                <!-- Button Grid with Modern Styling -->
                <Grid Margin="0,0,0,20">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="85"/>
                    </Grid.RowDefinitions>
                    
                    <!-- PRE Button -->
                    <Button Name="PreButton" Grid.Column="0" Margin="0,0,10,0" 
                            Content="PRE&#10;Before Installation"
                            Background="#10B981" Foreground="White" FontWeight="SemiBold"
                            FontSize="11" Padding="15,15" Cursor="Hand"
                            BorderThickness="0">
                        <Button.Style>
                            <Style TargetType="Button">
                                <Setter Property="Template">
                                    <Setter.Value>
                                        <ControlTemplate TargetType="Button">
                                            <Border CornerRadius="6" Background="{TemplateBinding Background}" 
                                                    BorderThickness="0" Padding="{TemplateBinding Padding}">
                                                <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" TextBlock.TextAlignment="Center"/>
                                            </Border>
                                            <ControlTemplate.Triggers>
                                                <Trigger Property="IsMouseOver" Value="True">
                                                    <Setter Property="Background" Value="#059669"/>
                                                </Trigger>
                                                <Trigger Property="IsPressed" Value="True">
                                                    <Setter Property="Background" Value="#047857"/>
                                                </Trigger>
                                            </ControlTemplate.Triggers>
                                        </ControlTemplate>
                                    </Setter.Value>
                                </Setter>
                            </Style>
                        </Button.Style>
                    </Button>
                    
                    <!-- POST Button -->
                    <Button Name="PostButton" Grid.Column="1" Margin="5,0,5,0" 
                            Content="POST&#10;After Installation"
                            Background="#3B82F6" Foreground="White" FontWeight="SemiBold"
                            FontSize="11" Padding="15,15" Cursor="Hand"
                            BorderThickness="0">
                        <Button.Style>
                            <Style TargetType="Button">
                                <Setter Property="Template">
                                    <Setter.Value>
                                        <ControlTemplate TargetType="Button">
                                            <Border CornerRadius="6" Background="{TemplateBinding Background}" 
                                                    BorderThickness="0" Padding="{TemplateBinding Padding}">
                                                <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" TextBlock.TextAlignment="Center"/>
                                            </Border>
                                            <ControlTemplate.Triggers>
                                                <Trigger Property="IsMouseOver" Value="True">
                                                    <Setter Property="Background" Value="#2563EB"/>
                                                </Trigger>
                                                <Trigger Property="IsPressed" Value="True">
                                                    <Setter Property="Background" Value="#1D4ED8"/>
                                                </Trigger>
                                            </ControlTemplate.Triggers>
                                        </ControlTemplate>
                                    </Setter.Value>
                                </Setter>
                            </Style>
                        </Button.Style>
                    </Button>
                    
                    <!-- SETTINGS Button -->
                    <Button Name="ConfigButton" Grid.Column="2" Margin="5,0,5,0" 
                            Content="SETTINGS&#10;Configuration"
                            Background="#F59E0B" Foreground="White" FontWeight="SemiBold"
                            FontSize="11" Padding="15,15" Cursor="Hand"
                            BorderThickness="0">
                        <Button.Style>
                            <Style TargetType="Button">
                                <Setter Property="Template">
                                    <Setter.Value>
                                        <ControlTemplate TargetType="Button">
                                            <Border CornerRadius="6" Background="{TemplateBinding Background}" 
                                                    BorderThickness="0" Padding="{TemplateBinding Padding}">
                                                <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" TextBlock.TextAlignment="Center"/>
                                            </Border>
                                            <ControlTemplate.Triggers>
                                                <Trigger Property="IsMouseOver" Value="True">
                                                    <Setter Property="Background" Value="#D97706"/>
                                                </Trigger>
                                                <Trigger Property="IsPressed" Value="True">
                                                    <Setter Property="Background" Value="#B45309"/>
                                                </Trigger>
                                            </ControlTemplate.Triggers>
                                        </ControlTemplate>
                                    </Setter.Value>
                                </Setter>
                            </Style>
                        </Button.Style>
                    </Button>
                    
                    <!-- EXIT Button -->
                    <Button Name="ExitButton" Grid.Column="3" Margin="10,0,0,0" 
                            Content="EXIT&#10;Close App"
                            Background="#EF4444" Foreground="White" FontWeight="SemiBold"
                            FontSize="11" Padding="15,15" Cursor="Hand"
                            BorderThickness="0">
                        <Button.Style>
                            <Style TargetType="Button">
                                <Setter Property="Template">
                                    <Setter.Value>
                                        <ControlTemplate TargetType="Button">
                                            <Border CornerRadius="6" Background="{TemplateBinding Background}" 
                                                    BorderThickness="0" Padding="{TemplateBinding Padding}">
                                                <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" TextBlock.TextAlignment="Center"/>
                                            </Border>
                                            <ControlTemplate.Triggers>
                                                <Trigger Property="IsMouseOver" Value="True">
                                                    <Setter Property="Background" Value="#DC2626"/>
                                                </Trigger>
                                                <Trigger Property="IsPressed" Value="True">
                                                    <Setter Property="Background" Value="#B91C1C"/>
                                                </Trigger>
                                            </ControlTemplate.Triggers>
                                        </ControlTemplate>
                                    </Setter.Value>
                                </Setter>
                            </Style>
                        </Button.Style>
                    </Button>
                </Grid>
                
                <!-- Help Text -->
                <Grid Margin="0,0,0,0">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    <TextBlock Grid.Column="0" Text="PRE: Snapshot system before installation" 
                               FontSize="10" Foreground="#6B7280" TextWrapping="Wrap" Margin="0,0,10,0"/>
                    <TextBlock Grid.Column="1" Text="POST: Compare and generate detailed change report" 
                               FontSize="10" Foreground="#6B7280" TextWrapping="Wrap"/>
                </Grid>
            </StackPanel>
        </Border>
        
        <!-- Status Display Area -->
        <Border Grid.Row="2" Background="#FFFFFF" Margin="25,20,25,25" CornerRadius="6" 
                BorderBrush="#E5E7EB" BorderThickness="1" Padding="0">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>
                
                <!-- Status Header -->
                <Border Grid.Row="0" Background="#F3F4F6" CornerRadius="6,6,0,0" Padding="20,15">
                    <TextBlock Text="Status and Progress" FontSize="12" FontWeight="SemiBold" 
                               Foreground="#1F2937"/>
                </Border>
                
                <!-- Status Box -->
                <TextBox Name="StatusBox" Grid.Row="1" IsReadOnly="True" 
                         VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto"
                         Text="Ready. Select a mode to begin." Foreground="#374151" 
                         FontFamily="Consolas" FontSize="10" 
                         Margin="0" TextWrapping="Wrap" Background="White" 
                         BorderThickness="0" Padding="20,15" AcceptsReturn="True"/>
            </Grid>
        </Border>
    </Grid>
</Window>
"@

[xml]$xmlReader = $xaml
$reader = New-Object System.Xml.XmlNodeReader $xmlReader
$window = [System.Windows.Markup.XamlReader]::Load($reader)

$preBtn = $window.FindName("PreButton")
$postBtn = $window.FindName("PostButton")
$exitBtn = $window.FindName("ExitButton")
$statusBox = $window.FindName("StatusBox")
$configBtn = $window.FindName("ConfigButton")

function Update-Status {
  param([string]$Message, [switch]$Append)
  
  $timestamp = Get-Date -Format "HH:mm:ss"
  $statusMsg = "[$timestamp] $Message"
  
  if ($Append) {
    $statusBox.Text = $statusBox.Text + "`n" + $statusMsg
  } else {
    $statusBox.Text = $statusMsg
  }
  
  # Auto-scroll to bottom
  $statusBox.ScrollToEnd()
  
  $window.Dispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Background, [action]{})
}

# --- Setup (folders, timestamps, etc.)
$ErrorActionPreference = 'Stop'

# Helper: export data as JSON and CSV with consistent naming
function Export-JsonCsv {
  param(
    [Parameter(Mandatory)][string]$BaseName,
    [Parameter(Mandatory)][object]$Data,
    [Parameter(Mandatory)][string]$Stage,
    [Parameter(Mandatory)][string]$TimeStamp,
    [Parameter(Mandatory)][string]$SnapshotDir
  )
  $json = Join-Path $SnapshotDir "$($BaseName)_$Stage_$TimeStamp.json"
  $csv  = Join-Path $SnapshotDir "$($BaseName)_$Stage_$TimeStamp.csv"
  $Data | ConvertTo-Json -Depth 6 | Set-Content -Encoding UTF8 $json -ErrorAction SilentlyContinue
  $Data | Export-Csv -NoTypeInformation -Encoding UTF8 $csv -ErrorAction SilentlyContinue
}

function Invoke-Snapshot {
  param([string]$SelectedMode)
  
  try {
    # Ensure environment variables in RootPaths are expanded
    $RootPaths = $RootPaths | ForEach-Object { [System.Environment]::ExpandEnvironmentVariables($_) }
    
    $startTime = Get-Date
    $ts = Get-Date -Format "yyyyMMdd_HHmmss"
    $stage = if ($SelectedMode -eq 'Pre') { 'pre' } else { 'post' }
    
    # Separate directories for Pre and Post
    if ($SelectedMode -eq 'Pre') {
      $ssDir = Join-Path (Join-Path $scriptDir $SnapshotRoot) "Reports_Pre"
    } else {
      $ssDir = Join-Path (Join-Path $scriptDir $SnapshotRoot) "Reports_Post"
    }
    
    $rpDir = $ssDir
    $snapshotRootDir = Join-Path $scriptDir $SnapshotRoot
    
    Update-Status "Starting $SelectedMode snapshot..."
    
    # Check if PRE snapshots exist when user selected POST mode
    if ($SelectedMode -eq 'Post') {
      $preReportsDir = Join-Path (Join-Path $scriptDir $SnapshotRoot) "Reports_Pre"
      Update-Status "Checking for PRE snapshots in: $preReportsDir" -Append
      
      if (-not (Test-Path $preReportsDir)) {
        Update-Status "PRE reports directory does not exist!" -Append
        [System.Windows.MessageBox]::Show(
          "No PRE snapshots found.`n`nPlease create a PRE snapshot first by clicking the PRE button.",
          "Missing PRE Snapshots",
          [System.Windows.MessageBoxButton]::OK,
          [System.Windows.MessageBoxImage]::Information
        ) | Out-Null
        Update-Status "POST cancelled - no PRE snapshots found." -Append
        return
      }
      
      Update-Status "PRE reports directory exists." -Append
      
      # List all files in the directory for debugging
      $allFiles = @(Get-ChildItem -Path $preReportsDir -File -ErrorAction SilentlyContinue)
      Update-Status "Total files in directory: $($allFiles.Count)" -Append
      foreach ($f in $allFiles) {
        Update-Status "  - $($f.Name)" -Append
      }
      
      $preSnapshotFiles = @(Get-ChildItem -Path $preReportsDir -Filter "*.json" -File -ErrorAction SilentlyContinue)
      Update-Status "Found $($preSnapshotFiles.Count) PRE snapshot files matching *.json" -Append
      
      if ($preSnapshotFiles.Count -eq 0) {
        [System.Windows.MessageBox]::Show(
          "No PRE snapshots found.`n`nPlease create a PRE snapshot first by clicking the PRE button.",
          "Missing PRE Snapshots",
          [System.Windows.MessageBoxButton]::OK,
          [System.Windows.MessageBoxImage]::Information
        ) | Out-Null
        Update-Status "POST cancelled - no PRE snapshots found." -Append
        return
      }
    }
    
    # Check if _Snapshots folder exists and user selected PRE mode
    if ($SelectedMode -eq 'Pre') {
      $snapshotRootDir = Join-Path $scriptDir $SnapshotRoot
      
      if (Test-Path $snapshotRootDir) {
        Update-Status "Existing _Snapshots folder found." -Append
        
        $msgBoxInput = [System.Windows.MessageBox]::Show(
          "Existing snapshots found.`n`nDo you want to delete them and create new ones?",
          "Delete Existing Snapshots?",
          [System.Windows.MessageBoxButton]::YesNo,
          [System.Windows.MessageBoxImage]::Question
        )
        
        if ($msgBoxInput -eq 'Yes') {
          Update-Status "Deleting old snapshots..." -Append
          Remove-Item -Path $snapshotRootDir -Recurse -Force -ErrorAction SilentlyContinue
          Update-Status "Old snapshots deleted." -Append
        } else {
          Update-Status "Cancelled - existing snapshots preserved." -Append
          return
        }
      }
    }
    
    # Check if Reports_Post folder exists and user selected POST mode
    $skipPostSnapshot = $false
    if ($SelectedMode -eq 'Post') {
      $postReportsDir = Join-Path (Join-Path $scriptDir $SnapshotRoot) "Reports_Post"
      
      if (Test-Path $postReportsDir) {
        Update-Status "Existing POST snapshots found." -Append
        
        $msgBoxInput = [System.Windows.MessageBox]::Show(
          "POST snapshots already exist.`n`nDelete existing POST snapshots and create NEW ones?`n`nYes = Delete & Create New`nNo = Keep & Regenerate Report Only",
          "POST Snapshots Found",
          [System.Windows.MessageBoxButton]::YesNoCancel,
          [System.Windows.MessageBoxImage]::Question
        )
        
        if ($msgBoxInput -eq 'Cancel') {
          Update-Status "POST cancelled - operation aborted." -Append
          return
        } elseif ($msgBoxInput -eq 'No') {
          Update-Status "Using existing POST snapshots, regenerating report only." -Append
          $skipPostSnapshot = $true
        } else {
          Update-Status "Deleting old POST snapshots..." -Append
          Remove-Item -Path $postReportsDir -Recurse -Force -ErrorAction SilentlyContinue
          Update-Status "Old POST snapshots deleted." -Append
        }
      }
    }
    
    New-Item -ItemType Directory -Force -Path $ssDir,$rpDir | Out-Null
    
    # Skip snapshot collection if we're reusing existing POST snapshots
    if ($skipPostSnapshot -eq $false) {
      # --- 1) Services
      if ($ScanOptions.services) {
        Update-Status "Collecting services..." -Append
        $services = Get-CimInstance Win32_Service | Select-Object Name,DisplayName,StartMode,State,PathName
        Export-JsonCsv -BaseName 'services' -Data $services -Stage $stage -TimeStamp $ts -SnapshotDir $ssDir
        Update-Status "Services: $($services.Count) items" -Append
      } else {
        Update-Status "Skipping services (disabled in settings)" -Append
      }
      
      # --- 2) Scheduled Tasks
      if ($ScanOptions.scheduledTasks) {
        Update-Status "Collecting scheduled tasks..." -Append
        $tasks = Get-ScheduledTask | ForEach-Object {
          [PSCustomObject]@{
            TaskName   = $_.TaskName
            TaskPath   = $_.TaskPath
            State      = $_.State
            LastRun    = ""
            NextRun    = ""
            Actions    = ($_.Actions | ForEach-Object { $_.Execute + ' ' + ($_.Arguments) }) -join ' | '
            Author     = $_.Principal.UserId
            RunLevel   = $_.Principal.RunLevel
          }
        }
        Export-JsonCsv -BaseName 'tasks' -Data $tasks -Stage $stage -TimeStamp $ts -SnapshotDir $ssDir
        Update-Status "Tasks: $($tasks.Count) items" -Append
      } else {
        Update-Status "Skipping scheduled tasks (disabled in settings)" -Append
      }
      
      # --- 3) Run/RunOnce registry
      if ($ScanOptions.runKeys) {
        Update-Status "Collecting Run/RunOnce entries..." -Append
        $runItems = foreach ($p in $runKeyPaths) {
          if (Test-Path $p) {
            Get-ItemProperty -Path $p | ForEach-Object {
              $_.PSObject.Properties |
                Where-Object { $_.Name -notin 'PSPath','PSParentPath','PSChildName','PSDrive','PSProvider' } |
                ForEach-Object {
                  [PSCustomObject]@{
                    HivePath = $p
                    Name     = $_.Name
                    Value    = $_.Value
                  }
                }
            }
          }
        }
        Export-JsonCsv -BaseName 'runkeys' -Data $runItems -Stage $stage -TimeStamp $ts -SnapshotDir $ssDir
        Update-Status "Run entries: $($runItems.Count) items" -Append
      } else {
        Update-Status "Skipping Run/RunOnce entries (disabled in settings)" -Append
      }
      
      # --- 3b) Uninstall registry keys
      if ($ScanOptions.uninstallKeys) {
        Update-Status "Collecting Uninstall entries..." -Append
        $uninstallItems = foreach ($p in $uninstallKeyPaths) {
          if (Test-Path $p) {
            Get-ChildItem -Path $p -ErrorAction SilentlyContinue | ForEach-Object {
              $subkeyName = $_.PSChildName
              $displayName = (Get-ItemProperty -Path $_.PSPath -Name "DisplayName" -ErrorAction SilentlyContinue).DisplayName
              $displayVersion = (Get-ItemProperty -Path $_.PSPath -Name "DisplayVersion" -ErrorAction SilentlyContinue).DisplayVersion
              [PSCustomObject]@{
                HivePath = $p
                SubkeyName = $subkeyName
                DisplayName = $displayName
                DisplayVersion = $displayVersion
              }
            }
          }
        }
        Export-JsonCsv -BaseName 'uninstall' -Data $uninstallItems -Stage $stage -TimeStamp $ts -SnapshotDir $ssDir
        Update-Status "Uninstall entries: $($uninstallItems.Count) items" -Append
      } else {
        Update-Status "Skipping Uninstall entries (disabled in settings)" -Append
      }
      
      # --- 4) Folders & 5) Shortcuts & 6) Files - Combined collection
      Update-Status "Collecting folders and files..." -Append
      $folderList = New-Object System.Collections.ArrayList
      $fileList = New-Object System.Collections.ArrayList
      
      foreach ($root in $RootPaths) {
        Update-Status "  Processing: $root" -Append
        if (Test-Path $root) {
          Get-ChildItem -Path $root -Recurse -ErrorAction SilentlyContinue |
            Where-Object { $_.FullName -notlike "*\_Snapshots*" } |
            ForEach-Object {
              if ($_.PSIsContainer) {
                # Folder
                [void]$folderList.Add([PSCustomObject]@{FullName = $_.FullName})
              } elseif ($_.Extension -ne ".lnk") {
                # File (not .lnk)
                [void]$fileList.Add([PSCustomObject]@{FullName = $_.FullName})
              }
            }
        } else {
          Update-Status "  Path not accessible: $root" -Append
        }
      }
      
      $folderItems = $folderList.ToArray()
      $fileItems = $fileList.ToArray()
      
      Export-JsonCsv -BaseName 'folders' -Data $folderItems -Stage $stage -TimeStamp $ts -SnapshotDir $ssDir
      Update-Status "Folders: $($folderItems.Count) items" -Append
      
      if ($ScanOptions.startMenuShortcuts) {
        Update-Status "Collecting shortcuts..." -Append
        $shortcutItems = foreach ($root in $RootPaths) {
          if (Test-Path $root) {
            Get-ChildItem -Path $root -Recurse -Filter *.lnk -ErrorAction SilentlyContinue |
              Where-Object { $_.FullName -notlike "*\_Snapshots*" } |
              Select-Object FullName
          }
        }
        Export-JsonCsv -BaseName 'shortcuts' -Data $shortcutItems -Stage $stage -TimeStamp $ts -SnapshotDir $ssDir
        Update-Status "Shortcuts: $($shortcutItems.Count) items" -Append
      } else {
        Update-Status "Skipping shortcuts (disabled in settings)" -Append
      }
      
      Export-JsonCsv -BaseName 'files' -Data $fileItems -Stage $stage -TimeStamp $ts -SnapshotDir $ssDir
      Update-Status "Files: $($fileItems.Count) items" -Append
    }
    
    # --- Diff generation (only for POST mode)
    if ($SelectedMode -eq 'Post') {
      Update-Status "Generating diff report..." -Append
      
      $preReportsDir = Join-Path (Join-Path $scriptDir $SnapshotRoot) "Reports_Pre"
      
      function LatestSnapshot([string]$base){ 
        Get-ChildItem -Path $preReportsDir -Filter "$($base)*.json" | Sort-Object LastWriteTime -Descending | Select-Object -First 1 
      }
      
      function Compare-Json {
        param([string]$base,[string[]]$keys)
        $pre = LatestSnapshot $base
        if (-not $pre) { return "No Pre-snapshot for '$base' found." }
        $post = Get-ChildItem -Path $ssDir -Filter "$($base)*.json" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        $preObj = Get-Content $pre.FullName | ConvertFrom-Json
        $postObj = Get-Content $post.FullName | ConvertFrom-Json

        $preMap  = @{}
        $postMap = @{}
        foreach ($r in $preObj) { $k = ($keys | ForEach-Object { "$($r.$_)" }) -join '|' ; $preMap[$k] = $r }
        foreach ($r in $postObj){ $k = ($keys | ForEach-Object { "$($r.$_)" }) -join '|' ; $postMap[$k] = $r }

        $added   = @()
        $removed = @()
        foreach ($k in $postMap.Keys) { if (-not $preMap.ContainsKey($k)) { $added += $postMap[$k] } }
        foreach ($k in $preMap.Keys)  { if (-not $postMap.ContainsKey($k)) { $removed += $preMap[$k] } }
        [PSCustomObject]@{Added=$added;Removed=$removed}
      }

      $reportTime = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
      $report = @()
      $report += "# System Changes - Report ($reportTime)"
      $report += ""

      $svcDiff = Compare-Json -base 'services' -keys @('Name')
      $report += "## New Services: $($svcDiff.Added.Count)"
      $report += @($svcDiff.Added | Sort-Object Name | Select-Object Name,DisplayName,StartMode,State,PathName | Format-Table | Out-String)
      $report += "## Removed Services: $($svcDiff.Removed.Count)"
      $report += @($svcDiff.Removed | Sort-Object Name | Select-Object Name,DisplayName,StartMode,State,PathName | Format-Table | Out-String)

      $tskDiff = Compare-Json -base 'tasks' -keys @('TaskPath','TaskName')
      $report += "## New Scheduled Tasks: $($tskDiff.Added.Count)"
      $report += @($tskDiff.Added | Sort-Object TaskPath,TaskName | Select-Object TaskPath,TaskName,State,Actions | Format-Table | Out-String)
      $report += "## Removed Scheduled Tasks: $($tskDiff.Removed.Count)"
      $report += @($tskDiff.Removed | Sort-Object TaskPath,TaskName | Select-Object TaskPath,TaskName,State,Actions | Format-Table | Out-String)

      $runDiff = Compare-Json -base 'runkeys' -keys @('HivePath','Name')
      $report += "## New Run/RunOnce Entries: $($runDiff.Added.Count)"
      $report += @($runDiff.Added | Sort-Object HivePath,Name | Select-Object HivePath,Name,Value | Format-Table | Out-String)
      $report += "## Removed Run/RunOnce Entries: $($runDiff.Removed.Count)"
      $report += @($runDiff.Removed | Sort-Object HivePath,Name | Select-Object HivePath,Name,Value | Format-Table | Out-String)

      $uninstallDiff = Compare-Json -base 'uninstall' -keys @('HivePath','SubkeyName')
      $report += "## New Uninstall Entries: $($uninstallDiff.Added.Count)"
      $report += @($uninstallDiff.Added | Sort-Object HivePath,SubkeyName | Select-Object HivePath,SubkeyName,DisplayName,DisplayVersion | Format-Table | Out-String)
      $report += "## Removed Uninstall Entries: $($uninstallDiff.Removed.Count)"
      $report += @($uninstallDiff.Removed | Sort-Object HivePath,SubkeyName | Select-Object HivePath,SubkeyName,DisplayName,DisplayVersion | Format-Table | Out-String)

      $fldDiff = Compare-Json -base 'folders' -keys @('FullName')
      $report += "## Newly Created Folders: $($fldDiff.Added.Count)"
      $report += @($fldDiff.Added | Sort-Object FullName | Select-Object FullName,CreationTime | Format-Table | Out-String)

      $lnkDiff = Compare-Json -base 'shortcuts' -keys @('FullName')
      Update-Status "Shortcuts comparison: Added=$($lnkDiff.Added.Count), Removed=$($lnkDiff.Removed.Count)" -Append
      $report += "## New Shortcuts: $($lnkDiff.Added.Count)"
      $report += @($lnkDiff.Added | Sort-Object FullName | Select-Object FullName,CreationTime | Format-Table | Out-String)

      $filDiff = Compare-Json -base 'files' -keys @('FullName')
      Update-Status "Files comparison: Added=$($filDiff.Added.Count), Removed=$($filDiff.Removed.Count)" -Append
      $report += "## New Files: $($filDiff.Added.Count)"
      $report += @($filDiff.Added | Sort-Object FullName | Select-Object FullName,CreationTime,Length | Format-Table | Out-String)

      $snapshotRootDir = Join-Path $scriptDir $SnapshotRoot
      $txtPath = Join-Path $snapshotRootDir ("ChangeReport_{0}.txt" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
      $report | Out-String -Width 120 | Set-Content -Encoding UTF8 $txtPath
      Update-Status "Report created: $txtPath" -Append
      
      # Ask user if they want to open the report
      $openReport = [System.Windows.MessageBox]::Show(
        "Report has been created. Do you want to open it?",
        "Open Report?",
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Question
      )
      
      if ($openReport -eq 'Yes') {
        Invoke-Item $txtPath
      }
    }
    
    $endTime = Get-Date
    $duration = $endTime - $startTime
    $durationStr = "{0:D2}:{1:D2}:{2:D2}" -f $duration.Hours, $duration.Minutes, $duration.Seconds
    Update-Status "SUCCESS: $SelectedMode snapshot completed successfully! (Duration: $durationStr)" -Append
  }
  catch {
    Update-Status "ERROR: $($_.Exception.Message)" -Append
  }
}

# Button event handlers
$preBtn.Add_Click({
  $preBtn.IsEnabled = $false
  $postBtn.IsEnabled = $false
  $configBtn.IsEnabled = $false
  $exitBtn.IsEnabled = $false
  
  Invoke-Snapshot -SelectedMode "Pre"
  
  $preBtn.IsEnabled = $true
  $postBtn.IsEnabled = $true
  $configBtn.IsEnabled = $true
  $exitBtn.IsEnabled = $true
})

$postBtn.Add_Click({
  $preBtn.IsEnabled = $false
  $postBtn.IsEnabled = $false
  $configBtn.IsEnabled = $false
  $exitBtn.IsEnabled = $false
  
  Invoke-Snapshot -SelectedMode "Post"
  
  $preBtn.IsEnabled = $true
  $postBtn.IsEnabled = $true
  $configBtn.IsEnabled = $true
  $exitBtn.IsEnabled = $true
})

$configBtn.Add_Click({
  $preBtn.IsEnabled = $false
  $postBtn.IsEnabled = $false
  $configBtn.IsEnabled = $false
  $exitBtn.IsEnabled = $false
  
  # Reload RootPaths from InstallTracker-Config.json before opening the window
  if (Test-Path $configFile) {
    try {
      $config = Get-Content $configFile -Raw | ConvertFrom-Json
      $RootPaths = $config.rootPaths
    } catch {
      # Keep existing RootPaths if reload fails
    }
  }
  
  # Config window XAML
  # Config window XAML
  $configXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" 
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="InstallTracker - Settings" 
        WindowStartupLocation="CenterOwner"
        Height="850" Width="750"
        Background="#F8F9FA"
        FontFamily="Segoe UI"
        FontSize="12"
        ShowInTaskbar="False"
        WindowStyle="ToolWindow">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <!-- Modern Header -->
        <Grid Grid.Row="0" Background="#1F2937">
            <StackPanel VerticalAlignment="Center" Margin="20,15">
                <TextBlock Text="Settings" FontSize="18" FontWeight="Bold" Foreground="White"/>
                <TextBlock Text="Configure scan paths and options" FontSize="10" Foreground="#9CA3AF" Margin="0,3,0,0"/>
            </StackPanel>
        </Grid>
        
        <!-- Content Area -->
        <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto" Margin="20,20,20,20">
            <StackPanel>
                <!-- Scan Paths Section -->
                <TextBlock Text="Scan Paths" Margin="0,0,0,12" FontSize="13" FontWeight="SemiBold" Foreground="#1F2937"/>
                <Border Margin="0,0,0,20" BorderBrush="#E5E7EB" BorderThickness="1" CornerRadius="6" Background="White" Padding="15">
                    <Grid>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="*"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        <ListBox Name="PathListBox" Background="White" Foreground="#374151" 
                                 Padding="8" SelectionMode="Single" Grid.Row="0" MinHeight="100">
                        </ListBox>
                        <TextBox Name="NewPathTextBox" Height="38" Padding="12,10" Grid.Row="1" Margin="0,12,0,0"
                                 VerticalAlignment="Center" Foreground="#374151" Background="White"
                                 BorderBrush="#D1D5DB" BorderThickness="1"/>
                        <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="0,10,0,0" Height="36">
                            <Button Name="AddPathBtn" Content="Add" Padding="24,8" 
                                    Background="#10B981" Foreground="White" FontWeight="SemiBold" FontSize="11"
                                    BorderThickness="0" Cursor="Hand" MinWidth="100">
                                <Button.Style>
                                    <Style TargetType="Button">
                                        <Setter Property="Template">
                                            <Setter.Value>
                                                <ControlTemplate TargetType="Button">
                                                    <Border CornerRadius="4" Background="{TemplateBinding Background}" BorderThickness="0">
                                                        <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                                    </Border>
                                                    <ControlTemplate.Triggers>
                                                        <Trigger Property="IsMouseOver" Value="True">
                                                            <Setter Property="Background" Value="#059669"/>
                                                        </Trigger>
                                                    </ControlTemplate.Triggers>
                                                </ControlTemplate>
                                            </Setter.Value>
                                        </Setter>
                                    </Style>
                                </Button.Style>
                            </Button>
                            <Button Name="RemovePathBtn" Content="Remove" Padding="24,8" Margin="8,0,0,0"
                                    Background="#EF4444" Foreground="White" FontWeight="SemiBold" FontSize="11"
                                    BorderThickness="0" Cursor="Hand" MinWidth="100">
                                <Button.Style>
                                    <Style TargetType="Button">
                                        <Setter Property="Template">
                                            <Setter.Value>
                                                <ControlTemplate TargetType="Button">
                                                    <Border CornerRadius="4" Background="{TemplateBinding Background}" BorderThickness="0">
                                                        <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                                    </Border>
                                                    <ControlTemplate.Triggers>
                                                        <Trigger Property="IsMouseOver" Value="True">
                                                            <Setter Property="Background" Value="#DC2626"/>
                                                        </Trigger>
                                                    </ControlTemplate.Triggers>
                                                </ControlTemplate>
                                            </Setter.Value>
                                        </Setter>
                                    </Style>
                                </Button.Style>
                            </Button>
                        </StackPanel>
                    </Grid>
                </Border>
                
                <!-- Scan Options Section -->
                <TextBlock Text="Scan Options" Margin="0,0,0,12" FontSize="13" FontWeight="SemiBold" Foreground="#1F2937"/>
                <Border Margin="0,0,0,20" BorderBrush="#E5E7EB" BorderThickness="1" CornerRadius="6" Background="White" Padding="15">
                    <StackPanel>
                        <CheckBox Name="ServicesCheckBox" Content="Windows Services" 
                                  Margin="0,0,0,10" Foreground="#374151" FontSize="11" VerticalAlignment="Center"/>
                        <CheckBox Name="TasksCheckBox" Content="Scheduled Tasks" 
                                  Margin="0,0,0,10" Foreground="#374151" FontSize="11" VerticalAlignment="Center"/>
                        <StackPanel Orientation="Horizontal" Margin="0,0,0,10">
                            <CheckBox Name="RunKeysCheckBox" Content="Run/RunOnce Registry Keys" 
                                      Foreground="#374151" FontSize="11" VerticalAlignment="Center"/>
                            <Button Name="RunKeysInfoBtn" Content="i" Width="24" Height="24" Margin="8,0,0,0"
                                    Background="#3B82F6" Foreground="White" FontWeight="Bold" FontSize="13"
                                    BorderThickness="0" Padding="0" VerticalAlignment="Center" 
                                    ToolTip="Click to see registry keys" Cursor="Hand">
                                <Button.Style>
                                    <Style TargetType="Button">
                                        <Setter Property="Template">
                                            <Setter.Value>
                                                <ControlTemplate TargetType="Button">
                                                    <Border Name="border" CornerRadius="3" Background="#3B82F6" BorderThickness="0">
                                                        <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                                    </Border>
                                                </ControlTemplate>
                                            </Setter.Value>
                                        </Setter>
                                    </Style>
                                </Button.Style>
                            </Button>
                        </StackPanel>
                        <StackPanel Orientation="Horizontal" Margin="0,0,0,10">
                            <CheckBox Name="UninstallKeysCheckBox" Content="Uninstall Registry Keys" 
                                      Foreground="#374151" FontSize="11" VerticalAlignment="Center"/>
                            <Button Name="UninstallKeysInfoBtn" Content="i" Width="24" Height="24" Margin="8,0,0,0"
                                    Background="#3B82F6" Foreground="White" FontWeight="Bold" FontSize="13"
                                    BorderThickness="0" Padding="0" VerticalAlignment="Center"
                                    ToolTip="Click to see registry keys" Cursor="Hand">
                                <Button.Style>
                                    <Style TargetType="Button">
                                        <Setter Property="Template">
                                            <Setter.Value>
                                                <ControlTemplate TargetType="Button">
                                                    <Border Name="border" CornerRadius="3" Background="#3B82F6" BorderThickness="0">
                                                        <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                                    </Border>
                                                </ControlTemplate>
                                            </Setter.Value>
                                        </Setter>
                                    </Style>
                                </Button.Style>
                            </Button>
                        </StackPanel>
                        <CheckBox Name="ShortcutsCheckBox" Content="Start Menu Shortcuts" 
                                  Margin="0,0,0,0" Foreground="#374151" FontSize="11" VerticalAlignment="Center"/>
                    </StackPanel>
                </Border>
                
                <!-- Version Check Section -->
                <TextBlock Text="Version Check" Margin="0,0,0,12" FontSize="13" FontWeight="SemiBold" Foreground="#1F2937"/>
                <Border Margin="0,0,0,20" BorderBrush="#E5E7EB" BorderThickness="1" CornerRadius="6" Background="White" Padding="15">
                    <StackPanel>
                        <CheckBox Name="CheckVersionsCheckBox" Content="Check for new versions on startup" 
                                  Margin="0,0,0,0" Foreground="#374151" FontSize="11" VerticalAlignment="Center"/>
                        <TextBlock Text="Repository: bergerpascal/InstallTracker" 
                                   Margin="0,12,0,0" Foreground="#6B7280" FontSize="10"/>
                    </StackPanel>
                </Border>
            </StackPanel>
        </ScrollViewer>
        
        <!-- Save/Cancel buttons -->
        <Border Grid.Row="2" Background="#FFFFFF" BorderThickness="1,1,0,0" BorderBrush="#E5E7EB" Padding="20">
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                <Button Name="ResetBtn" Content="Reset Defaults" Padding="16,10"
                        Background="#F59E0B" Foreground="White" FontWeight="SemiBold" FontSize="11"
                        BorderThickness="0" Cursor="Hand" MinWidth="160" Height="40">
                    <Button.Style>
                        <Style TargetType="Button">
                            <Setter Property="Template">
                                <Setter.Value>
                                    <ControlTemplate TargetType="Button">
                                        <Border CornerRadius="4" Background="#F59E0B" BorderThickness="0">
                                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                        </Border>
                                    </ControlTemplate>
                                </Setter.Value>
                            </Setter>
                        </Style>
                    </Button.Style>
                </Button>
                <Button Name="SaveBtn" Content="Save Settings" Padding="16,10" Margin="10,0,0,0"
                        Background="#3B82F6" Foreground="White" FontWeight="SemiBold" FontSize="11"
                        BorderThickness="0" Cursor="Hand" MinWidth="160" Height="40">
                    <Button.Style>
                        <Style TargetType="Button">
                            <Setter Property="Template">
                                <Setter.Value>
                                    <ControlTemplate TargetType="Button">
                                        <Border CornerRadius="4" Background="#3B82F6" BorderThickness="0">
                                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                        </Border>
                                    </ControlTemplate>
                                </Setter.Value>
                            </Setter>
                        </Style>
                    </Button.Style>
                </Button>
                <Button Name="CancelBtn" Content="Cancel" Padding="16,10" Margin="10,0,0,0"
                        Background="#9CA3AF" Foreground="White" FontWeight="SemiBold" FontSize="11"
                        BorderThickness="0" Cursor="Hand" MinWidth="140" Height="40">
                    <Button.Style>
                        <Style TargetType="Button">
                            <Setter Property="Template">
                                <Setter.Value>
                                    <ControlTemplate TargetType="Button">
                                        <Border CornerRadius="4" Background="#9CA3AF" BorderThickness="0">
                                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                        </Border>
                                    </ControlTemplate>
                                </Setter.Value>
                            </Setter>
                        </Style>
                    </Button.Style>
                </Button>
            </StackPanel>
        </Border>
    </Grid>
</Window>
"@

  [xml]$configXmlReader = $configXaml
  $configReader = New-Object System.Xml.XmlNodeReader $configXmlReader
  $configWindow = [System.Windows.Markup.XamlReader]::Load($configReader)
  
  $pathListBox = $configWindow.FindName("PathListBox")
  $newPathTextBox = $configWindow.FindName("NewPathTextBox")
  $addPathBtn = $configWindow.FindName("AddPathBtn")
  $removePathBtn = $configWindow.FindName("RemovePathBtn")
  $saveBtn = $configWindow.FindName("SaveBtn")
  $cancelBtn = $configWindow.FindName("CancelBtn")
  $resetBtn = $configWindow.FindName("ResetBtn")
  
  $servicesCheckBox = $configWindow.FindName("ServicesCheckBox")
  $tasksCheckBox = $configWindow.FindName("TasksCheckBox")
  $runKeysCheckBox = $configWindow.FindName("RunKeysCheckBox")
  $uninstallKeysCheckBox = $configWindow.FindName("UninstallKeysCheckBox")
  $shortcutsCheckBox = $configWindow.FindName("ShortcutsCheckBox")
  $runKeysInfoBtn = $configWindow.FindName("RunKeysInfoBtn")
  $uninstallKeysInfoBtn = $configWindow.FindName("UninstallKeysInfoBtn")
  
  # Info button click handler
  $runKeysInfoBtn.Add_Click({
    # Create custom WPF window dynamically from $runKeyPaths array
    $stackPanel = New-Object System.Windows.Controls.StackPanel
    $stackPanel.Margin = New-Object System.Windows.Thickness(0)
    
    # Add registry keys from array
    foreach ($keyPath in $runKeyPaths) {
      $keyTextBlock = New-Object System.Windows.Controls.TextBlock
      $keyTextBlock.Text = "- $keyPath"
      $keyTextBlock.Margin = New-Object System.Windows.Thickness(0,0,0,8)
      $keyTextBlock.FontSize = 11
      $keyTextBlock.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString("#4B5563"))
      $stackPanel.Children.Add($keyTextBlock) | Out-Null
    }
    
    # ScrollViewer
    $scrollViewer = New-Object System.Windows.Controls.ScrollViewer
    $scrollViewer.VerticalScrollBarVisibility = "Auto"
    $scrollViewer.Padding = New-Object System.Windows.Thickness(0,10,0,0)
    $scrollViewer.Content = $stackPanel
    
    # Main Grid for window
    $mainGrid = New-Object System.Windows.Controls.Grid
    $mainGrid.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString("#F8F9FA"))
    
    # Row definitions
    $rowDef1 = New-Object System.Windows.Controls.RowDefinition
    $rowDef1.Height = [System.Windows.GridLength]::Auto
    $rowDef2 = New-Object System.Windows.Controls.RowDefinition
    $rowDef2.Height = New-Object System.Windows.GridLength(1, [System.Windows.GridUnitType]::Star)
    $rowDef3 = New-Object System.Windows.Controls.RowDefinition
    $rowDef3.Height = [System.Windows.GridLength]::Auto
    
    $mainGrid.RowDefinitions.Add($rowDef1) | Out-Null
    $mainGrid.RowDefinitions.Add($rowDef2) | Out-Null
    $mainGrid.RowDefinitions.Add($rowDef3) | Out-Null
    
    # Add modern header to Row 0
    $headerGrid = New-Object System.Windows.Controls.Grid
    $headerGrid.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString("#1F2937"))
    $headerGrid.Margin = New-Object System.Windows.Thickness(0)
    
    $headerTextBlock = New-Object System.Windows.Controls.TextBlock
    $headerTextBlock.Text = "Windows Run/RunOnce Registry Keys"
    $headerTextBlock.FontSize = 15
    $headerTextBlock.FontWeight = [System.Windows.FontWeights]::Bold
    $headerTextBlock.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString("White"))
    $headerTextBlock.Margin = New-Object System.Windows.Thickness(20,12,20,12)
    $headerGrid.Children.Add($headerTextBlock) | Out-Null
    [System.Windows.Controls.Grid]::SetRow($headerGrid, 0)
    $mainGrid.Children.Add($headerGrid) | Out-Null
    
    # Add content border with scrollviewer to Row 1
    $contentBorder = New-Object System.Windows.Controls.Border
    $contentBorder.Margin = New-Object System.Windows.Thickness(20)
    $contentBorder.Padding = New-Object System.Windows.Thickness(0)
    [System.Windows.Controls.Grid]::SetRow($contentBorder, 1)
    $contentBorder.Child = $scrollViewer
    $mainGrid.Children.Add($contentBorder) | Out-Null
    
    # Add footer with button to Row 2
    $footerBorder = New-Object System.Windows.Controls.Border
    $footerBorder.BorderThickness = New-Object System.Windows.Thickness(0,1,0,0)
    $footerBorder.BorderBrush = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString("#E5E7EB"))
    $footerBorder.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString("White"))
    $footerBorder.Padding = New-Object System.Windows.Thickness(20)
    
    $closeBtn = New-Object System.Windows.Controls.Button
    $closeBtn.Content = "Close"
    $closeBtn.Width = 90
    $closeBtn.Height = 36
    $closeBtn.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString("#3B82F6"))
    $closeBtn.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString("White"))
    $closeBtn.FontWeight = [System.Windows.FontWeights]::SemiBold
    $closeBtn.FontSize = 11
    $closeBtn.Cursor = [System.Windows.Input.Cursors]::Hand
    $closeBtn.Padding = New-Object System.Windows.Thickness(20, 8, 20, 8)
    
    $footerStackPanel = New-Object System.Windows.Controls.StackPanel
    $footerStackPanel.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Right
    $footerStackPanel.Children.Add($closeBtn) | Out-Null
    $footerBorder.Child = $footerStackPanel
    
    [System.Windows.Controls.Grid]::SetRow($footerBorder, 2)
    $mainGrid.Children.Add($footerBorder) | Out-Null
    
    # Create window
    $infoWindow = New-Object System.Windows.Window
    $infoWindow.Title = "Registry Run Keys Information"
    $infoWindow.Width = 700
    $infoWindow.Height = 450
    $infoWindow.WindowStartupLocation = "CenterScreen"
    $infoWindow.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString("#F8F9FA"))
    $infoWindow.ResizeMode = "CanResizeWithGrip"
    $infoWindow.Content = $mainGrid
    
    $closeBtn.Add_Click({ $infoWindow.Close() })
    $infoWindow.ShowDialog() | Out-Null
  })
  
  # Uninstall Keys info button click handler
  $uninstallKeysInfoBtn.Add_Click({
    $stackPanel = New-Object System.Windows.Controls.StackPanel
    $stackPanel.Margin = New-Object System.Windows.Thickness(0)
    
    foreach ($keyPath in $uninstallKeyPaths) {
      $keyTextBlock = New-Object System.Windows.Controls.TextBlock
      $keyTextBlock.Text = "- $keyPath"
      $keyTextBlock.Margin = New-Object System.Windows.Thickness(0,0,0,8)
      $keyTextBlock.FontSize = 11
      $keyTextBlock.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString("#4B5563"))
      $stackPanel.Children.Add($keyTextBlock) | Out-Null
    }
    
    $scrollViewer = New-Object System.Windows.Controls.ScrollViewer
    $scrollViewer.VerticalScrollBarVisibility = "Auto"
    $scrollViewer.Padding = New-Object System.Windows.Thickness(0,10,0,0)
    $scrollViewer.Content = $stackPanel
    
    # Main Grid for window
    $mainGrid = New-Object System.Windows.Controls.Grid
    $mainGrid.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString("#F8F9FA"))
    
    $rowDef1 = New-Object System.Windows.Controls.RowDefinition
    $rowDef1.Height = [System.Windows.GridLength]::Auto
    $rowDef2 = New-Object System.Windows.Controls.RowDefinition
    $rowDef2.Height = New-Object System.Windows.GridLength(1, [System.Windows.GridUnitType]::Star)
    $rowDef3 = New-Object System.Windows.Controls.RowDefinition
    $rowDef3.Height = [System.Windows.GridLength]::Auto
    
    $mainGrid.RowDefinitions.Add($rowDef1) | Out-Null
    $mainGrid.RowDefinitions.Add($rowDef2) | Out-Null
    $mainGrid.RowDefinitions.Add($rowDef3) | Out-Null
    
    # Add modern header to Row 0
    $headerGrid = New-Object System.Windows.Controls.Grid
    $headerGrid.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString("#1F2937"))
    $headerGrid.Margin = New-Object System.Windows.Thickness(0)
    
    $headerTextBlock = New-Object System.Windows.Controls.TextBlock
    $headerTextBlock.Text = "Windows Uninstall Registry Keys"
    $headerTextBlock.FontSize = 15
    $headerTextBlock.FontWeight = [System.Windows.FontWeights]::Bold
    $headerTextBlock.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString("White"))
    $headerTextBlock.Margin = New-Object System.Windows.Thickness(20,12,20,12)
    $headerGrid.Children.Add($headerTextBlock) | Out-Null
    [System.Windows.Controls.Grid]::SetRow($headerGrid, 0)
    $mainGrid.Children.Add($headerGrid) | Out-Null
    
    # Add content border with scrollviewer to Row 1
    $contentBorder = New-Object System.Windows.Controls.Border
    $contentBorder.Margin = New-Object System.Windows.Thickness(20)
    $contentBorder.Padding = New-Object System.Windows.Thickness(0)
    [System.Windows.Controls.Grid]::SetRow($contentBorder, 1)
    $contentBorder.Child = $scrollViewer
    $mainGrid.Children.Add($contentBorder) | Out-Null
    
    # Add footer with button to Row 2
    $footerBorder = New-Object System.Windows.Controls.Border
    $footerBorder.BorderThickness = New-Object System.Windows.Thickness(0,1,0,0)
    $footerBorder.BorderBrush = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString("#E5E7EB"))
    $footerBorder.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString("White"))
    $footerBorder.Padding = New-Object System.Windows.Thickness(20)
    
    $closeBtn = New-Object System.Windows.Controls.Button
    $closeBtn.Content = "Close"
    $closeBtn.Width = 90
    $closeBtn.Height = 36
    $closeBtn.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString("#3B82F6"))
    $closeBtn.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString("White"))
    $closeBtn.FontWeight = [System.Windows.FontWeights]::SemiBold
    $closeBtn.FontSize = 11
    $closeBtn.Cursor = [System.Windows.Input.Cursors]::Hand
    $closeBtn.Padding = New-Object System.Windows.Thickness(20, 8, 20, 8)
    
    $footerStackPanel = New-Object System.Windows.Controls.StackPanel
    $footerStackPanel.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Right
    $footerStackPanel.Children.Add($closeBtn) | Out-Null
    $footerBorder.Child = $footerStackPanel
    
    [System.Windows.Controls.Grid]::SetRow($footerBorder, 2)
    $mainGrid.Children.Add($footerBorder) | Out-Null
    
    $infoWindow = New-Object System.Windows.Window
    $infoWindow.Title = "Registry Uninstall Keys Information"
    $infoWindow.Width = 700
    $infoWindow.Height = 450
    $infoWindow.WindowStartupLocation = "CenterScreen"
    $infoWindow.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString("#F8F9FA"))
    $infoWindow.ResizeMode = "CanResizeWithGrip"
    $infoWindow.Content = $mainGrid
    
    $closeBtn.Add_Click({ $infoWindow.Close() })
    $infoWindow.ShowDialog() | Out-Null
  })
  
  # Load current paths
  $pathListBox.Items.Clear()
  foreach ($path in $RootPaths) {
    [void]$pathListBox.Items.Add($path)
  }
  
  # Load current scan options into checkboxes
  $servicesCheckBox.IsChecked = $ScanOptions.services
  $tasksCheckBox.IsChecked = $ScanOptions.scheduledTasks
  $runKeysCheckBox.IsChecked = $ScanOptions.runKeys
  $uninstallKeysCheckBox.IsChecked = $ScanOptions.uninstallKeys
  $shortcutsCheckBox.IsChecked = $ScanOptions.startMenuShortcuts
  
  # Load version check options
  $checkVersionsCheckBox = $configWindow.FindName("CheckVersionsCheckBox")
  $checkVersionsCheckBox.IsChecked = $CheckVersions
  
  # Add path button
  $addPathBtn.Add_Click({
    $newPath = $newPathTextBox.Text.Trim()
    
    if ([string]::IsNullOrWhiteSpace($newPath)) {
      [System.Windows.MessageBox]::Show("Please enter a path.", "Empty Path", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
      return
    }
    
    if ($pathListBox.Items -contains $newPath) {
      [System.Windows.MessageBox]::Show("This path already exists.", "Duplicate Path", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
      return
    }
    
    if (-not (Test-Path $newPath)) {
      $msgResult = [System.Windows.MessageBox]::Show("Path does not exist.`n`nAdd anyway?", "Path Not Found", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
      if ($msgResult -ne 'Yes') { return }
    }
    
    [void]$pathListBox.Items.Add($newPath)
    $newPathTextBox.Clear()
  })
  
  # Remove path button
  $removePathBtn.Add_Click({
    if ($pathListBox.SelectedIndex -lt 0) {
      [System.Windows.MessageBox]::Show("Please select a path to remove.", "No Selection", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
      return
    }
    
    $selectedIndex = $pathListBox.SelectedIndex
    $pathListBox.Items.RemoveAt($selectedIndex)
  })
  
  # Save button
  $saveBtn.Add_Click({
    if ($pathListBox.Items.Count -eq 0) {
      [System.Windows.MessageBox]::Show("You must have at least one path.", "Empty Paths", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
      return
    }
    
    # Create config object and save
    $newRootPaths = @()
    foreach ($item in $pathListBox.Items) {
      $newRootPaths += $item
    }
    
    $newScanOptions = @{
      services = $servicesCheckBox.IsChecked
      scheduledTasks = $tasksCheckBox.IsChecked
      runKeys = $runKeysCheckBox.IsChecked
      uninstallKeys = $uninstallKeysCheckBox.IsChecked
      startMenuShortcuts = $shortcutsCheckBox.IsChecked
    }
    
    $configObject = @{
      rootPaths = $newRootPaths
      scanOptions = $newScanOptions
      checkVersions = $checkVersionsCheckBox.IsChecked
      gitHubRepository = "bergerpascal/InstallTracker"
      description = "Folders to scan for changes in InstallTracker (managed via GUI)"
    }
    
    try {
      $configObject | ConvertTo-Json | Set-Content -Path $configFile -Encoding UTF8
      [System.Windows.MessageBox]::Show("Configuration saved successfully!", "Success", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null
      
      # Update global variables with newly saved settings
      $script:RootPaths = $newRootPaths
      $script:ScanOptions = $newScanOptions
      $script:CheckVersions = $checkVersionsCheckBox.IsChecked
      
      $configWindow.Close()
    } catch {
      [System.Windows.MessageBox]::Show("Error saving config: $($_.Exception.Message)", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
    }
  })
  
  # Cancel button
  $cancelBtn.Add_Click({
    $configWindow.Close()
  })
  
  # Reset to defaults button
  $resetBtn.Add_Click({
    $msgResult = [System.Windows.MessageBox]::Show("Reset all settings to defaults?`n`nThis cannot be undone.", "Reset to Defaults", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
    if ($msgResult -ne 'Yes') { return }
    
    # Clear path list and add default paths from Fallback
    $pathListBox.Items.Clear()
    $defaultPaths = @(
      "%USERPROFILE%",
      "%APPDATA%\Microsoft\Windows\Start Menu",
      "C:\Program Files",
      "C:\ProgramData",
      "C:\Users\Public\Desktop",
      "C:\Program Files (x86)"
    )
    foreach ($defaultPath in $defaultPaths) {
      [void]$pathListBox.Items.Add($defaultPath)
    }
    
    # Reset checkboxes to defaults (all enabled except not defined)
    $servicesCheckBox.IsChecked = $true
    $tasksCheckBox.IsChecked = $true
    $runKeysCheckBox.IsChecked = $true
    $uninstallKeysCheckBox.IsChecked = $true
    $shortcutsCheckBox.IsChecked = $true
    
    # Reset version check settings to defaults
    $checkVersionsCheckBox.IsChecked = $true
    $gitHubRepoTextBox.Text = "bergerpascal/InstallTracker"
    
    $newPathTextBox.Clear()
    
    [System.Windows.MessageBox]::Show("Settings reset to defaults.", "Reset Complete", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null
  })
  
  $configWindow.Owner = $window
  $configWindow.ShowDialog() | Out-Null
  
  $preBtn.IsEnabled = $true
  $postBtn.IsEnabled = $true
  $configBtn.IsEnabled = $true
  $exitBtn.IsEnabled = $true
})

$exitBtn.Add_Click({
  $window.Close()
})

# Check if script is running as Administrator
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
$initialStatus = if ($isAdmin) {
  "Ready. Select a mode to begin."
} else {
  "WARNING: Not running as Administrator. For best results, please restart this tool with Administrator privileges.`n`nReady. Select a mode to begin."
}

$statusBox.Text = $initialStatus

# Check for version updates if enabled
if ($CheckVersions -and -not [string]::IsNullOrWhiteSpace($GitHubRepository)) {
  $statusBox.Text = "$initialStatus`n`nChecking for new versions..."
  $window.Dispatcher.Invoke([System.Action]{
    $window.UpdateLayout()
  }, "Normal")
  
  $versionInfo = Test-GitHubVersionCheck -Repository $GitHubRepository -CurrentVersion $scriptVersion
  
  if ($versionInfo) {
    $statusBox.Text = $initialStatus
    
    # Show version check dialog
    if (Show-VersionCheckDialog -UpdateInfo $versionInfo) {
      try {
        # Open the release page in default browser
        Start-Process $updateInfo.DownloadUrl
      } catch {
        Write-Host "Could not open browser: $_"
      }
    }
  } else {
    $statusBox.Text = $initialStatus
  }
}

# Display the GUI window
$window.ShowDialog() | Out-Null
