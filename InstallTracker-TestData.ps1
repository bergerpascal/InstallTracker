#
# InstallTracker Test Data Generator
# Creates and removes test data for all monitored system components
# 
#

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore

$scriptVersion = "1.0.11"

# --- Version Check and Update Logic ---
$script:updateAvailable = $false
$script:updateInfo = $null
$CheckVersions = $true
$GitHubRepository = "bergerpascal/InstallTracker"

# If this is a restart after update, skip the update check
$isUpdateRestart = $env:INSTALLTRACKER_TESTDATA_UPDATE_RESTART -eq "1"

if ($CheckVersions -eq $true -and $GitHubRepository -and -not $isUpdateRestart) {
  $oldEA = $ErrorActionPreference
  $ErrorActionPreference = 'Continue'
  
  try {
    $scriptPath = $MyInvocation.MyCommand.Path
    if (-not $scriptPath) { $scriptPath = $PSCommandPath }
    
    # Fast version check using Range Request - only download first 5KB (version is at the top)
    # Add timestamp to avoid CDN caching
    $cacheParam = [int64](Get-Date -UFormat %s)
    $uri = "https://raw.githubusercontent.com/$GitHubRepository/refs/heads/main/InstallTracker-TestData.ps1?t=$cacheParam"
    
    $latestVersion = $null
    $content = $null
    
    try {
      # Use Range header to get only first 5KB - much faster!
      $webRequest = [System.Net.HttpWebRequest]::Create($uri)
      $webRequest.Timeout = 2000
      $webRequest.AddRange(0, 5120)
      
      try {
        $webResponse = $webRequest.GetResponse()
        $stream = $webResponse.GetResponseStream()
        $reader = New-Object System.IO.StreamReader($stream)
        $content = $reader.ReadToEnd()
        $reader.Close()
        $webResponse.Close()
      } catch {
        # If Range not supported, fallback to regular download
        $webClient = New-Object System.Net.WebClient
        $webClient.Timeout = 2000
        $content = $webClient.DownloadString($uri)
      }
    } catch {
      # Try with Invoke-WebRequest as fallback
      try {
        $response = Invoke-WebRequest -Uri $uri -TimeoutSec 2 -UseBasicParsing -Headers @{Range="bytes=0-5120"}
        $content = $response.Content
      } catch {
        # Network error - skip version check
      }
    }
    
    # Extract version from downloaded content if successful
    if ($content) {
      # Try multiple regex patterns to extract version
      if ($content -match '\$scriptVersion\s*=\s*"([0-9.]+)"') {
        $latestVersion = $matches[1]
      } elseif ($content -match 'scriptVersion\s*=\s*"([0-9.]+)"') {
        $latestVersion = $matches[1]
      }
      
      # Compare versions if extracted successfully
      if ($latestVersion) {
        try {
          $currentVer = [version]$scriptVersion
          $remoteVer = [version]$latestVersion
          
          if ($remoteVer -gt $currentVer) {
            $script:updateAvailable = $true
            $script:updateInfo = @{
              LatestVersion = $latestVersion
              DownloadUrl = $uri
              ScriptPath = $scriptPath
            }
          }
        } catch {
          # Version comparison failed - skip
        }
      }
    }
  } catch {
    # Silently ignore all errors
  } finally {
    $ErrorActionPreference = $oldEA
  }
}

# Test data prefixes and markers
$testMarker = "SystemDiffTest_"
$testGuid = [guid]::NewGuid().ToString().Substring(0, 8)

# Test data paths
$testFolderBase = "C:\ProgramData\${testMarker}Folders"
$testFileBase = "C:\ProgramData\${testMarker}Files"
$testShortcutBase = "C:\ProgramData\${testMarker}Shortcuts"
$testTaskPath = "\${testMarker}Tasks\"
$testServiceName = "${testMarker}Service_${testGuid}"
$testUserDesktop = Join-Path $env:USERPROFILE "Desktop\${testMarker}UserDesktop"
$testAllUserDesktop = Join-Path "C:\Users\Public" "Desktop\${testMarker}AllUserDesktop"
$testProgramFiles = Join-Path "C:\Program Files" "${testMarker}TestApp"
$testProgramFilesX86 = Join-Path "C:\Program Files (x86)" "${testMarker}TestApp"
$testStartMenu = Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs\${testMarker}StartMenu"
$testLocalAppData = Join-Path $env:LOCALAPPDATA "${testMarker}AppData"

# --- GUI Definition ---
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" 
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="InstallTracker Test Data Generator (v$scriptVersion)" 
        WindowStartupLocation="CenterScreen"
        Height="650" Width="800"
        Background="White"
        FontFamily="Segoe UI"
        FontSize="11">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="70"/>
            <RowDefinition Height="*"/>
        </Grid.RowDefinitions>
        
        <!-- Modern Header -->
        <Border Grid.Row="0" Background="#1F2937" Padding="20,0,20,0">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                
                <!-- Title and Version -->
                <StackPanel Grid.Column="0" VerticalAlignment="Center">
                    <TextBlock Text="InstallTracker Test Data Generator" 
                               FontSize="18" FontWeight="Bold" Foreground="White" Margin="0,0,0,5"/>
                    <TextBlock Text="Create or remove test data for system change detection" 
                               FontSize="10" Foreground="#D1D5DB"/>
                </StackPanel>
                
                <!-- Version Badge -->
                <Border Grid.Column="1" VerticalAlignment="Center" Background="#10B981" 
                        CornerRadius="4" Height="36" Margin="0,0,15,0" Padding="12,0">
                    <TextBlock Text="v$scriptVersion" FontSize="10" Foreground="White" FontWeight="Bold"
                               VerticalAlignment="Center" HorizontalAlignment="Center"/>
                </Border>
                
                <!-- Help Button -->
                <Button Name="HelpButton" Grid.Column="2" Content="?" 
                        VerticalAlignment="Center" HorizontalAlignment="Right" Margin="0,0,0,0"
                        Width="36" Height="36" FontSize="16" FontWeight="Bold"
                        Background="#3B82F6" Foreground="White" Cursor="Hand"
                        BorderThickness="0">
                    <Button.Style>
                        <Style TargetType="Button">
                            <Setter Property="Template">
                                <Setter.Value>
                                    <ControlTemplate TargetType="Button">
                                        <Border CornerRadius="4" Background="{TemplateBinding Background}">
                                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                        </Border>
                                        <ControlTemplate.Triggers>
                                            <Trigger Property="IsMouseOver" Value="True">
                                                <Setter Property="Background" Value="#2563EB"/>
                                            </Trigger>
                                        </ControlTemplate.Triggers>
                                    </ControlTemplate>
                                </Setter.Value>
                            </Setter>
                        </Style>
                    </Button.Style>
                </Button>
            </Grid>
        </Border>
        
        <!-- Main Content -->
        <Grid Grid.Row="1" Margin="20,20,20,20">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
            </Grid.RowDefinitions>
            
            <!-- Action Buttons Section -->
            <Border Grid.Row="0" Background="#F3F4F6" CornerRadius="6" Padding="20" Margin="0,0,0,20">
                <StackPanel>
                    <TextBlock Text="Test Data Management" FontSize="14" FontWeight="Bold" Foreground="#1F2937"/>
                    <TextBlock Text="Generate or remove test data across all monitored system components" 
                               FontSize="10" Foreground="#6B7280" TextWrapping="Wrap" Margin="0,0,0,15"/>
                    
                    <!-- Buttons Grid -->
                    <Grid Margin="0,0,0,0">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="10"/>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="10"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            
                            <!-- Create Button -->
                            <Button Name="CreateButton" Grid.Column="0" Content="CREATE TEST DATA" 
                                    Height="50" Foreground="White" FontWeight="Bold" FontSize="12">
                                <Button.Style>
                                    <Style TargetType="Button">
                                        <Setter Property="Background" Value="#10B981"/>
                                        <Setter Property="Template">
                                            <Setter.Value>
                                                <ControlTemplate TargetType="Button">
                                                    <Border Background="{TemplateBinding Background}" CornerRadius="6" Padding="20,10">
                                                        <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
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
                            
                            <!-- Delete Button -->
                            <Button Name="DeleteButton" Grid.Column="2" Content="REMOVE TEST DATA" 
                                    Height="50" Foreground="White" FontWeight="Bold" FontSize="12">
                                <Button.Style>
                                    <Style TargetType="Button">
                                        <Setter Property="Background" Value="#EF4444"/>
                                        <Setter Property="Template">
                                            <Setter.Value>
                                                <ControlTemplate TargetType="Button">
                                                    <Border Background="{TemplateBinding Background}" CornerRadius="6" Padding="20,10">
                                                        <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
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
                            
                            <!-- Exit Button -->
                            <Button Name="ExitButton" Grid.Column="4" Content="Exit" 
                                    Width="80" Height="50" Foreground="White" FontWeight="Bold" FontSize="11">
                                <Button.Style>
                                    <Style TargetType="Button">
                                        <Setter Property="Background" Value="#6B7280"/>
                                        <Setter Property="Template">
                                            <Setter.Value>
                                                <ControlTemplate TargetType="Button">
                                                    <Border Background="{TemplateBinding Background}" CornerRadius="6" Padding="10">
                                                        <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                                    </Border>
                                                    <ControlTemplate.Triggers>
                                                        <Trigger Property="IsMouseOver" Value="True">
                                                            <Setter Property="Background" Value="#4B5563"/>
                                                        </Trigger>
                                                        <Trigger Property="IsPressed" Value="True">
                                                            <Setter Property="Background" Value="#374151"/>
                                                        </Trigger>
                                                    </ControlTemplate.Triggers>
                                                </ControlTemplate>
                                            </Setter.Value>
                                        </Setter>
                                    </Style>
                                </Button.Style>
                            </Button>
                        </Grid>
                    </StackPanel>
                </Border>
            
            <!-- Status Box -->
            <Border Grid.Row="2" Background="White" BorderBrush="#E5E7EB" BorderThickness="1" CornerRadius="6" Padding="0" Margin="0,0,0,0">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*" MinHeight="150"/>
                </Grid.RowDefinitions>
                
                <TextBlock Grid.Row="0" Text="Activity Log" FontSize="12" FontWeight="Bold" Foreground="#1F2937" 
                           Padding="15,15,15,10"/>
                
                <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto" 
                              HorizontalScrollBarVisibility="Disabled"
                              x:Name="StatusScrollViewer">
                    <TextBox Name="StatusBox" 
                             IsReadOnly="True" 
                             Text="Ready. Click CREATE TEST DATA to generate test data."
                             Foreground="#374151" FontFamily="Consolas" 
                             FontSize="10" Padding="15,0,15,15" TextWrapping="Wrap" 
                             Background="White" BorderThickness="0" AcceptsReturn="True"/>
                </ScrollViewer>
            </Grid>
        </Border>
        </Grid>
    </Grid>
</Window>
"@

[xml]$xmlReader = $xaml
$reader = New-Object System.Xml.XmlNodeReader $xmlReader
$window = [System.Windows.Markup.XamlReader]::Load($reader)

$createBtn = $window.FindName("CreateButton")
$deleteBtn = $window.FindName("DeleteButton")
$exitBtn = $window.FindName("ExitButton")
$helpBtn = $window.FindName("HelpButton")
$statusBox = $window.FindName("StatusBox")
$statusScrollViewer = $window.FindName("StatusScrollViewer")

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
  if ($statusScrollViewer) {
    $statusScrollViewer.ScrollToEnd()
  } else {
    $statusBox.ScrollToEnd()
  }
  $window.Dispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Background, [action]{})
}

function New-TestFolders {
  param()
  try {
    Update-Status "Creating test folders..." -Append
    New-Item -ItemType Directory -Path $testFolderBase -Force -ErrorAction SilentlyContinue | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $testFolderBase "SubFolder1") -Force -ErrorAction SilentlyContinue | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $testFolderBase "SubFolder2") -Force -ErrorAction SilentlyContinue | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $testFolderBase "Nested\Level2\Level3") -Force -ErrorAction SilentlyContinue | Out-Null
    Update-Status "[OK] Test folders created at: $testFolderBase" -Append
    return $true
  }
  catch {
    Update-Status "[ERROR] Folder error: $($_.Exception.Message)" -Append
    return $false
  }
}

function New-TestDesktopFiles {
  param()
  try {
    Update-Status "Creating test files on User Desktop..." -Append
    New-Item -ItemType Directory -Path $testUserDesktop -Force -ErrorAction SilentlyContinue | Out-Null
    Set-Content -Path (Join-Path $testUserDesktop "UserDesktopFile.txt") -Value "Test file on User Desktop" -Encoding UTF8 -ErrorAction SilentlyContinue
    Update-Status "[OK] User Desktop files created" -Append
    
    Update-Status "Creating test files on All Users Desktop..." -Append
    New-Item -ItemType Directory -Path $testAllUserDesktop -Force -ErrorAction SilentlyContinue | Out-Null
    Set-Content -Path (Join-Path $testAllUserDesktop "AllUsersDesktopFile.txt") -Value "Test file on All Users Desktop" -Encoding UTF8 -ErrorAction SilentlyContinue
    Update-Status "[OK] All Users Desktop files created" -Append
    
    return $true
  }
  catch {
    Update-Status "[ERROR] Desktop file error: $($_.Exception.Message)" -Append
    return $false
  }
}

function New-TestProgramFilesData {
  param()
  try {
    Update-Status "Creating test data in Program Files..." -Append
    New-Item -ItemType Directory -Path (Join-Path $testProgramFiles "bin") -Force -ErrorAction SilentlyContinue | Out-Null
    Set-Content -Path (Join-Path $testProgramFiles "TestConfig.xml") -Value "<config><setting>test</setting></config>" -Encoding UTF8 -ErrorAction SilentlyContinue
    Set-Content -Path (Join-Path $testProgramFiles "bin\app.exe") -Value "TEST EXECUTABLE" -Encoding ASCII -ErrorAction SilentlyContinue
    Update-Status "[OK] Program Files test data created" -Append
    
    Update-Status "Creating test data in Program Files (x86)..." -Append
    New-Item -ItemType Directory -Path (Join-Path $testProgramFilesX86 "bin") -Force -ErrorAction SilentlyContinue | Out-Null
    Set-Content -Path (Join-Path $testProgramFilesX86 "TestConfig.xml") -Value "<config><setting>test x86</setting></config>" -Encoding UTF8 -ErrorAction SilentlyContinue
    Set-Content -Path (Join-Path $testProgramFilesX86 "bin\app32.exe") -Value "TEST 32-BIT EXECUTABLE" -Encoding ASCII -ErrorAction SilentlyContinue
    Update-Status "[OK] Program Files (x86) test data created" -Append
    
    return $true
  }
  catch {
    Update-Status "[ERROR] Program Files error: $($_.Exception.Message)" -Append
    return $false
  }
}

function New-TestStartMenuShortcuts {
  param()
  try {
    Update-Status "Creating shortcuts in all locations..." -Append
    $createdCount = 0
    
    # All shortcut locations
    $shortcutLocations = @(
      @{ Path = $testUserDesktop; Label = "User Desktop" },
      @{ Path = $testAllUserDesktop; Label = "All Users Desktop" },
      @{ Path = $testStartMenu; Label = "User Start Menu" },
      @{ Path = Join-Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs" "${testMarker}StartMenu"; Label = "All Users Start Menu" }
    )
    
    $shell = New-Object -ComObject WScript.Shell
    
    foreach ($location in $shortcutLocations) {
      try {
        $locationPath = $location.Path
        $locationLabel = $location.Label
        
        New-Item -ItemType Directory -Path $locationPath -Force -ErrorAction SilentlyContinue | Out-Null
        
        # Create 3 shortcuts in each location
        $targets = @(
          @{ Name = "Notepad"; Path = "C:\Windows\System32\notepad.exe" },
          @{ Name = "Calculator"; Path = "C:\Windows\System32\calc.exe" },
          @{ Name = "Explorer"; Path = "C:\Windows\explorer.exe" }
        )
        
        foreach ($target in $targets) {
          $shortcutPath = Join-Path $locationPath "${testMarker}$($target.Name).lnk"
          $link = $shell.CreateShortCut($shortcutPath)
          $link.TargetPath = $target.Path
          $link.Description = "InstallTracker Test - $($target.Name)"
          $link.Save()
          $createdCount++
        }
        
        Update-Status "[OK] Created shortcuts in $locationLabel (3 shortcuts)" -Append
      }
      catch {
        Update-Status "[WARNING] Could not create shortcuts in $($location.Label): $($_.Exception.Message)" -Append
      }
    }
    
    Update-Status "[OK] All shortcuts created ($createdCount total)" -Append
    return $true
  }
  catch {
    Update-Status "[ERROR] Shortcut error: $($_.Exception.Message)" -Append
    return $false
  }
}

function New-TestLocalAppDataFiles {
  param()
  try {
    Update-Status "Creating test files in LocalAppData..." -Append
    New-Item -ItemType Directory -Path $testLocalAppData -Force -ErrorAction SilentlyContinue | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $testLocalAppData "Cache") -Force -ErrorAction SilentlyContinue | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $testLocalAppData "Data") -Force -ErrorAction SilentlyContinue | Out-Null
    
    Set-Content -Path (Join-Path $testLocalAppData "AppConfig.ini") -Value "[Settings]`nTestValue=Enabled" -Encoding UTF8 -ErrorAction SilentlyContinue
    Set-Content -Path (Join-Path $testLocalAppData "Cache\CacheFile.tmp") -Value "Temporary cache data" -Encoding UTF8 -ErrorAction SilentlyContinue
    Set-Content -Path (Join-Path $testLocalAppData "Data\UserData.json") -Value '{"test":"data","version":"1.0"}' -Encoding UTF8 -ErrorAction SilentlyContinue
    
    Update-Status "[OK] LocalAppData test files created at: $testLocalAppData" -Append
    return $true
  }
  catch {
    Update-Status "[ERROR] LocalAppData error: $($_.Exception.Message)" -Append
    return $false
  }
}

function New-TestFiles {
  param()
  try {
    Update-Status "Creating test files..." -Append
    New-Item -ItemType Directory -Path $testFileBase -Force -ErrorAction SilentlyContinue | Out-Null
    
    Set-Content -Path (Join-Path $testFileBase "TestFile1.txt") -Value "This is test file 1`nCreated by InstallTracker Test Generator" -Encoding UTF8 -ErrorAction SilentlyContinue
    Set-Content -Path (Join-Path $testFileBase "TestFile2.log") -Value "Log file for testing`n[INFO] Test log entry" -Encoding UTF8 -ErrorAction SilentlyContinue
    Set-Content -Path (Join-Path $testFileBase "TestConfig.ini") -Value "[Settings]`nTestKey=TestValue" -Encoding UTF8 -ErrorAction SilentlyContinue
    
    New-Item -ItemType Directory -Path (Join-Path $testFileBase "SubFolder") -Force -ErrorAction SilentlyContinue | Out-Null
    Set-Content -Path (Join-Path $testFileBase "SubFolder\TestFile3.txt") -Value "Nested test file" -Encoding UTF8 -ErrorAction SilentlyContinue
    
    Update-Status "[OK] Test files created at: $testFileBase" -Append
    return $true
  }
  catch {
    Update-Status "[ERROR] File error: $($_.Exception.Message)" -Append
    return $false
  }
}

function New-TestShortcuts {
  param()
  try {
    Update-Status "Creating test shortcuts..." -Append
    New-Item -ItemType Directory -Path $testShortcutBase -Force -ErrorAction SilentlyContinue | Out-Null
    
    $shell = New-Object -ComObject WScript.Shell
    $shortcut1Path = Join-Path $testShortcutBase "TestApp1.lnk"
    $shortcut2Path = Join-Path $testShortcutBase "TestApp2.lnk"
    
    $link = $shell.CreateShortCut($shortcut1Path)
    $link.TargetPath = "C:\Windows\System32\notepad.exe"
    $link.Description = "Test Notepad Shortcut"
    $link.Save()
    
    $link = $shell.CreateShortCut($shortcut2Path)
    $link.TargetPath = "C:\Windows\System32\calc.exe"
    $link.Description = "Test Calculator Shortcut"
    $link.Save()
    
    Update-Status "[OK] Test shortcuts created at: $testShortcutBase" -Append
    return $true
  }
  catch {
    Update-Status "[ERROR] Shortcut error: $($_.Exception.Message)" -Append
    return $false
  }
}

function New-TestRegistryEntries {
  param()
  try {
    Update-Status "Creating registry test entries..." -Append
    
    $testValue1 = "C:\Windows\System32\notepad.exe"
    $testValue2 = "C:\Windows\System32\calc.exe"
    
    $oldErrorPref = $ErrorActionPreference
    $ErrorActionPreference = 'SilentlyContinue'
    
    New-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run" -Name "${testMarker}TestApp1" -Value $testValue1 -PropertyType String -Force | Out-Null
    New-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run" -Name "${testMarker}TestApp2" -Value $testValue2 -PropertyType String -Force | Out-Null
    
    Update-Status "[OK] Registry Run entries created in HKCU" -Append
    
    $ErrorActionPreference = $oldErrorPref
    return $true
  }
  catch {
    Update-Status "[ERROR] Registry error: $($_.Exception.Message)" -Append
    return $false
  }
}

function New-TestUninstallEntries {
  param()
  try {
    Update-Status "Creating test Uninstall entries..." -Append
    $createdCount = 0
    
    # HKCU User-level Uninstall entries
    $uninstallPaths = @(
      @{ Path = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall"; Label = "HKCU" },
      @{ Path = "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall"; Label = "HKLM (x64)" },
      @{ Path = "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"; Label = "HKLM (x86)" }
    )
    
    foreach ($uninstallInfo in $uninstallPaths) {
      $uninstallPath = $uninstallInfo.Path
      $label = $uninstallInfo.Label
      
      try {
        # Create 5 test app entries
        for ($i = 1; $i -le 5; $i++) {
          $pathEntry = Join-Path $uninstallPath "${testMarker}TestApp$i"
          New-Item -Path $pathEntry -Force -ErrorAction SilentlyContinue | Out-Null
          New-ItemProperty -Path $pathEntry -Name "DisplayName" -Value "InstallTracker Test Application $i" -PropertyType String -Force -ErrorAction SilentlyContinue | Out-Null
          New-ItemProperty -Path $pathEntry -Name "DisplayVersion" -Value "$i.0.0" -PropertyType String -Force -ErrorAction SilentlyContinue | Out-Null
          $createdCount++
        }
        
        Update-Status "[OK] Uninstall entries created in $label (5 entries)" -Append
      }
      catch {
        Update-Status "[WARNING] Could not write to $label (may need admin privileges): $($_.Exception.Message)" -Append
      }
    }
    
    Update-Status "[OK] Test Uninstall entries created ($createdCount total)" -Append
    return $true
  }
  catch {
    Update-Status "[ERROR] Uninstall entry error: $($_.Exception.Message)" -Append
    return $false
  }
}

function New-TestService {
  param()
  try {
    Update-Status "Creating test service: $testServiceName" -Append
    
    $batchFile = Join-Path $env:TEMP "${testServiceName}.bat"
    @"
@echo off
REM Test Service Batch File
timeout /t 3600 /nobreak
"@ | Set-Content -Path $batchFile -Encoding ASCII -ErrorAction SilentlyContinue
    
    $newService = New-Service -Name $testServiceName -BinaryPathName $batchFile -DisplayName "SystemDiffTest Service" -StartupType Manual -ErrorAction SilentlyContinue
    
    if ($newService) {
      Update-Status "[OK] Service created: $testServiceName" -Append
    } else {
      Update-Status "[WARNING] Service creation skipped (may need admin privileges)" -Append
    }
    return $true
  }
  catch {
    Update-Status "[ERROR] Service error: $($_.Exception.Message)" -Append
    return $false
  }
}

function New-TestScheduledTask {
  param()
  try {
    Update-Status "Creating scheduled task..." -Append
    
    $action = New-ScheduledTaskAction -Execute "cmd.exe" -Argument "/c echo Test" -ErrorAction SilentlyContinue
    $trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddHours(1) -ErrorAction SilentlyContinue
    $settings = New-ScheduledTaskSettingsSet -RunOnlyIfNetworkAvailable -DontStopOnIdleEnd -ErrorAction SilentlyContinue
    
    if ($action -and $trigger -and $settings) {
      Register-ScheduledTask -TaskName "${testMarker}Task1" -Action $action -Trigger $trigger -Settings $settings -TaskPath "\${testMarker}Tasks\" -Force -ErrorAction SilentlyContinue | Out-Null
      Update-Status "[OK] Scheduled task created" -Append
    } else {
      Update-Status "[WARNING] Could not create scheduled task" -Append
    }
    return $true
  }
  catch {
    Update-Status "[ERROR] Task error: $($_.Exception.Message)" -Append
    return $false
  }
}

function Remove-TestService {
  param()
  try {
    Update-Status "Removing test services..." -Append
    $services = Get-Service -Name "SystemDiffTest_Service_*" -ErrorAction SilentlyContinue
    $removed = 0
    
    if ($services) {
      foreach ($service in $services) {
        try {
          Stop-Service -Name $service.Name -Force -ErrorAction SilentlyContinue
          
          # Use WMI for deletion (compatible with PowerShell 5.1+)
          $wmiService = Get-WmiObject Win32_Service -Filter "Name='$($service.Name)'" -ErrorAction SilentlyContinue
          if ($wmiService) {
            $wmiService.Delete() | Out-Null
            $removed++
            Update-Status "[OK] Service removed: $($service.Name)" -Append
          }
        }
        catch {
          Update-Status "[WARNING] Could not remove service $($service.Name): $($_.Exception.Message)" -Append
        }
      }
    } else {
      Update-Status "[INFO] No test services found" -Append
    }
    
    if ($removed -gt 0) {
      Update-Status "[OK] Total services removed: $removed" -Append
    }
  }
  catch {
    Update-Status "[ERROR] Service removal error: $($_.Exception.Message)" -Append
  }
}

function Remove-TestScheduledTasks {
  param()
  try {
    Update-Status "Removing scheduled tasks..." -Append
    $tasks = Get-ScheduledTask -TaskPath "\${testMarker}Tasks\" -ErrorAction SilentlyContinue
    if ($tasks) {
      foreach ($task in $tasks) {
        Unregister-ScheduledTask -TaskName $task.TaskName -TaskPath "\${testMarker}Tasks\" -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
      }
      Update-Status "[OK] Scheduled tasks removed" -Append
    } else {
      Update-Status "[INFO] No scheduled tasks found" -Append
    }
  }
  catch {
    Update-Status "[ERROR] Task removal error: $($_.Exception.Message)" -Append
  }
}

function Remove-TestRegistryEntries {
  param()
  try {
    Update-Status "Removing registry test entries..." -Append
    $regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"
    $properties = Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue
    $removed = 0
    
    if ($properties) {
      foreach ($prop in $properties.PSObject.Properties) {
        if ($prop.Name -like "${testMarker}*") {
          Remove-ItemProperty -Path $regPath -Name $prop.Name -Force -ErrorAction SilentlyContinue
          $removed++
        }
      }
    }
    
    Update-Status "[OK] Registry Run entries removed ($removed total)" -Append
    return $true
  }
  catch {
    Update-Status "[ERROR] Registry removal error: $($_.Exception.Message)" -Append
    return $false
  }
}

function Remove-TestUninstallEntries {
  param()
  try {
    Update-Status "Removing test Uninstall entries..." -Append
    $removed = 0
    
    # Uninstall paths from all hives
    $uninstallPaths = @(
      @{ Path = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall"; Label = "HKCU" },
      @{ Path = "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall"; Label = "HKLM (x64)" },
      @{ Path = "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"; Label = "HKLM (x86)" }
    )
    
    foreach ($uninstallInfo in $uninstallPaths) {
      $uninstallPath = $uninstallInfo.Path
      $label = $uninstallInfo.Label
      
      try {
        $items = Get-ChildItem -Path $uninstallPath -ErrorAction SilentlyContinue
        if ($items) {
          foreach ($item in $items) {
            if ($item.PSChildName -like "${testMarker}*") {
              Remove-Item -Path $item.PSPath -Force -ErrorAction SilentlyContinue
              $removed++
            }
          }
        }
        Update-Status "[OK] Checked $label" -Append
      }
      catch {
        Update-Status "[WARNING] Could not access $label (may need admin privileges)" -Append
      }
    }
    
    Update-Status "[OK] Uninstall entries removed ($removed total)" -Append
    return $true
  }
  catch {
    Update-Status "[ERROR] Uninstall removal error: $($_.Exception.Message)" -Append
    return $false
  }
}

function Remove-TestFolders {
  param()
  try {
    Update-Status "Removing test folders..." -Append
    if (Test-Path $testFolderBase) {
      Remove-Item -Path $testFolderBase -Recurse -Force -ErrorAction SilentlyContinue
      Update-Status "[OK] Test folders removed" -Append
    } else {
      Update-Status "[INFO] Test folders not found" -Append
    }
    return $true
  }
  catch {
    Update-Status "[ERROR] Folder removal error: $($_.Exception.Message)" -Append
    return $false
  }
}

function Remove-TestFiles {
  param()
  try {
    Update-Status "Removing test files..." -Append
    if (Test-Path $testFileBase) {
      Remove-Item -Path $testFileBase -Recurse -Force -ErrorAction SilentlyContinue
      Update-Status "[OK] Test files removed" -Append
    } else {
      Update-Status "[INFO] Test files not found" -Append
    }
    return $true
  }
  catch {
    Update-Status "[ERROR] File removal error: $($_.Exception.Message)" -Append
    return $false
  }
}

function Remove-TestShortcuts {
  param()
  try {
    Update-Status "Removing test shortcuts..." -Append
    if (Test-Path $testShortcutBase) {
      Remove-Item -Path $testShortcutBase -Recurse -Force -ErrorAction SilentlyContinue
      Update-Status "[OK] Test shortcuts removed" -Append
    } else {
      Update-Status "[INFO] Test shortcuts not found" -Append
    }
    return $true
  }
  catch {
    Update-Status "[ERROR] Shortcut removal error: $($_.Exception.Message)" -Append
    return $false
  }
}

function Remove-TestDesktopFiles {
  param()
  try {
    Update-Status "Removing User Desktop test files..." -Append
    if (Test-Path $testUserDesktop) {
      Remove-Item -Path $testUserDesktop -Recurse -Force -ErrorAction SilentlyContinue
      Update-Status "[OK] User Desktop files removed" -Append
    } else {
      Update-Status "[INFO] User Desktop test files not found" -Append
    }
    
    Update-Status "Removing All Users Desktop test files..." -Append
    if (Test-Path $testAllUserDesktop) {
      Remove-Item -Path $testAllUserDesktop -Recurse -Force -ErrorAction SilentlyContinue
      Update-Status "[OK] All Users Desktop files removed" -Append
    } else {
      Update-Status "[INFO] All Users Desktop test files not found" -Append
    }
    
    return $true
  }
  catch {
    Update-Status "[ERROR] Desktop file removal error: $($_.Exception.Message)" -Append
    return $false
  }
}

function Remove-TestLocalAppDataFiles {
  param()
  try {
    Update-Status "Removing LocalAppData test files..." -Append
    if (Test-Path $testLocalAppData) {
      Remove-Item -Path $testLocalAppData -Recurse -Force -ErrorAction SilentlyContinue
      Update-Status "[OK] LocalAppData test files removed" -Append
    } else {
      Update-Status "[INFO] LocalAppData test files not found" -Append
    }
    return $true
  }
  catch {
    Update-Status "[ERROR] LocalAppData removal error: $($_.Exception.Message)" -Append
    return $false
  }
}

function Remove-TestProgramFilesData {
  param()
  try {
    Update-Status "Removing Program Files test data..." -Append
    if (Test-Path $testProgramFiles) {
      Remove-Item -Path $testProgramFiles -Recurse -Force -ErrorAction SilentlyContinue
      Update-Status "[OK] Program Files test data removed" -Append
    } else {
      Update-Status "[INFO] Program Files test data not found" -Append
    }
    
    Update-Status "Removing Program Files (x86) test data..." -Append
    if (Test-Path $testProgramFilesX86) {
      Remove-Item -Path $testProgramFilesX86 -Recurse -Force -ErrorAction SilentlyContinue
      Update-Status "[OK] Program Files (x86) test data removed" -Append
    } else {
      Update-Status "[INFO] Program Files (x86) test data not found" -Append
    }
    
    return $true
  }
  catch {
    Update-Status "[ERROR] Program Files removal error: $($_.Exception.Message)" -Append
    return $false
  }
}

function Remove-TestStartMenuShortcuts {
  param()
  try {
    Update-Status "Removing shortcuts from all locations..." -Append
    $removedCount = 0
    
    # All shortcut locations
    $shortcutLocations = @(
      @{ Path = $testUserDesktop; Label = "User Desktop" },
      @{ Path = $testAllUserDesktop; Label = "All Users Desktop" },
      @{ Path = $testStartMenu; Label = "User Start Menu" },
      @{ Path = Join-Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs" "${testMarker}StartMenu"; Label = "All Users Start Menu" }
    )
    
    foreach ($location in $shortcutLocations) {
      try {
        $locationPath = $location.Path
        $locationLabel = $location.Label
        
        if (Test-Path $locationPath) {
          Remove-Item -Path $locationPath -Recurse -Force -ErrorAction SilentlyContinue
          Update-Status "[OK] Removed shortcuts from $locationLabel" -Append
          $removedCount++
        } else {
          Update-Status "[INFO] No shortcuts found in $locationLabel" -Append
        }
      }
      catch {
        Update-Status "[WARNING] Could not remove $($location.Label): $($_.Exception.Message)" -Append
      }
    }
    
    if ($removedCount -gt 0) {
      Update-Status "[OK] All shortcuts removed from $removedCount locations" -Append
    }
    return $true
  }
  catch {
    Update-Status "[ERROR] Shortcut removal error: $($_.Exception.Message)" -Append
    return $false
  }
}

$createBtn.Add_Click({
  $startTime = Get-Date
  $createBtn.IsEnabled = $false
  $deleteBtn.IsEnabled = $false
  
  Update-Status "Starting test data creation..." -Append
  Update-Status "Timestamp: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -Append
  Update-Status "" -Append
  
  $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
  if (-not $isAdmin) {
    Update-Status "[WARNING] Not running as Administrator!" -Append
    Update-Status "Some operations (Services, Program Files, HKLM Registry) may fail." -Append
    Update-Status "" -Append
  }
  
  New-TestFolders
  Update-Status "" -Append
  New-TestFiles
  Update-Status "" -Append
  New-TestShortcuts
  Update-Status "" -Append
  New-TestDesktopFiles
  Update-Status "" -Append
  New-TestProgramFilesData
  Update-Status "" -Append
  New-TestStartMenuShortcuts
  Update-Status "" -Append
  New-TestLocalAppDataFiles
  Update-Status "" -Append
  New-TestRegistryEntries
  Update-Status "" -Append
  New-TestUninstallEntries
  Update-Status "" -Append
  New-TestScheduledTask
  Update-Status "" -Append
  New-TestService
  Update-Status "" -Append
  
  $endTime = Get-Date
  $duration = $endTime - $startTime
  $durationStr = "{0:D2}:{1:D2}:{2:D2}" -f $duration.Hours, $duration.Minutes, $duration.Seconds
  Update-Status "[SUCCESS] Test data creation completed! (Duration: $durationStr)" -Append
  Update-Status "You can now run InstallTracker.ps1 to capture PRE/POST snapshots" -Append
  
  $createBtn.IsEnabled = $true
  $deleteBtn.IsEnabled = $true
})

$deleteBtn.Add_Click({
  $startTime = Get-Date
  $createBtn.IsEnabled = $false
  $deleteBtn.IsEnabled = $false
  
  $msgBoxInput = [System.Windows.MessageBox]::Show(
    "Are you sure you want to remove all test data?",
    "Confirm Deletion",
    [System.Windows.MessageBoxButton]::YesNo,
    [System.Windows.MessageBoxImage]::Warning
  )
  
  if ($msgBoxInput -eq 'Yes') {
    Update-Status "Starting test data removal..." -Append
    Update-Status "Timestamp: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -Append
    Update-Status "" -Append
    
    Remove-TestService
    Update-Status "" -Append
    Remove-TestScheduledTasks
    Update-Status "" -Append
    Remove-TestRegistryEntries
    Update-Status "" -Append
    Remove-TestUninstallEntries
    Update-Status "" -Append
    Remove-TestShortcuts
    Update-Status "" -Append
    Remove-TestFiles
    Update-Status "" -Append
    Remove-TestFolders
    Update-Status "" -Append
    Remove-TestDesktopFiles
    Update-Status "" -Append
    Remove-TestProgramFilesData
    Update-Status "" -Append
    Remove-TestLocalAppDataFiles
    Update-Status "" -Append
    Remove-TestStartMenuShortcuts
    Update-Status "" -Append
    
    $endTime = Get-Date
    $duration = $endTime - $startTime
    $durationStr = "{0:D2}:{1:D2}:{2:D2}" -f $duration.Hours, $duration.Minutes, $duration.Seconds
    Update-Status "[SUCCESS] Test data removal completed! (Duration: $durationStr)" -Append
  } else {
    Update-Status "Deletion cancelled by user." -Append
  }
  
  $createBtn.IsEnabled = $true
  $deleteBtn.IsEnabled = $true
})

$exitBtn.Add_Click({
  $window.Close()
})

$helpBtn.Add_Click({
  $readmeUrl = "https://github.com/$GitHubRepository/blob/main/README.md"
  try {
    Start-Process $readmeUrl
  } catch {
    [System.Windows.MessageBox]::Show("Could not open browser to README.`n`nURL: $readmeUrl", "Help", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null
  }
})

Update-Status "Ready. Click CREATE TEST DATA to generate test data for InstallTracker testing."

# Show update notification at startup if available
if ($script:updateAvailable -eq $true -and $script:updateInfo) {
  $oldEA = $ErrorActionPreference
  $ErrorActionPreference = 'SilentlyContinue'
  
  try {
    $result = [System.Windows.MessageBox]::Show(
      "A new version is available!`n`nInstalled: v$scriptVersion`nAvailable: v$($script:updateInfo.LatestVersion)`n`nDo you want to download and install the update?",
      "InstallTracker TestData Update",
      [System.Windows.MessageBoxButton]::YesNo,
      [System.Windows.MessageBoxImage]::Information
    )
    
    if ($result -eq "Yes") {
      try {
        $tempDownloadPath = [System.IO.Path]::GetTempFileName() -replace '\.tmp$', '.ps1'
        $wc = New-Object System.Net.WebClient
        $wc.DownloadFile($script:updateInfo.DownloadUrl, $tempDownloadPath)
        
        # Create backup of current version
        $backupPath = $script:updateInfo.ScriptPath -replace '\.ps1$', '_backup.ps1'
        Copy-Item -Path $script:updateInfo.ScriptPath -Destination $backupPath -Force
        
        # Wait for download to complete
        Start-Sleep -Milliseconds 500
        
        # Delete the original file first to clear any file locks/caches
        Remove-Item -Path $script:updateInfo.ScriptPath -Force
        
        # Wait a moment
        Start-Sleep -Milliseconds 300
        
        # Copy the new version to the original location
        Copy-Item -Path $tempDownloadPath -Destination $script:updateInfo.ScriptPath -Force
        
        # Clean up temp file
        Remove-Item -Path $tempDownloadPath -Force
        
        [System.Windows.MessageBox]::Show(
          "Update installed successfully!`n`nThe script will now restart.",
          "Update Complete",
          [System.Windows.MessageBoxButton]::OK,
          [System.Windows.MessageBoxImage]::Information
        ) | Out-Null
        
        # Wait to ensure everything is written to disk
        Start-Sleep -Milliseconds 1000
        
        # Start the new script in a NEW process using Start-Process
        $scriptPath = $script:updateInfo.ScriptPath
        Start-Process -FilePath "powershell.exe" -ArgumentList "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", $scriptPath -WindowStyle Normal
        
        # Exit the old process completely
        exit 0
      } catch {
        # Silently ignore download errors
      }
    }
  } catch {
    # Silently ignore all errors
  }
  
  $ErrorActionPreference = $oldEA
}

# If this is an update restart, skip the GUI and exit - the new version will show it
if ($isUpdateRestart) {
  exit 0
}

$window.ShowDialog() | Out-Null
