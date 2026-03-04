#Requires -RunAsAdministrator
<#
.SYNOPSIS
    WinToPalette Installer & Uninstaller
    Downloads and installs WinToPalette from GitHub releases

.DESCRIPTION
    This script downloads the latest WinToPalette release from GitHub,
    installs the Interception driver, and sets up auto-startup.
    Also handles uninstallation.

.PARAMETER Uninstall
    Uninstall WinToPalette instead of installing it

.EXAMPLE
    PS> .\Install-WinToPalette.ps1
    PS> .\Install-WinToPalette.ps1 -Uninstall
#>

param(
    [switch]$Uninstall = $false
)

# ============================================================================
# CONFIGURATION
# ============================================================================
$GitHubRepo = "zehevi/WinToPalette"
$InstallPath = "$env:ProgramFiles\WinToPalette"
$AppName = "WinToPalette"

# ============================================================================
# UTILITY FUNCTIONS
# ============================================================================

function Write-Success {
    param([string]$Message)
    Write-Host "[OK] $Message" -ForegroundColor Green
}

function Write-Error_ {
    param([string]$Message)
    Write-Host "[ERR] $Message" -ForegroundColor Red
}

function Write-Warning_ {
    param([string]$Message)
    Write-Host "[!] $Message" -ForegroundColor Yellow
}

function Write-Info {
    param([string]$Message)
    Write-Host "[*] $Message" -ForegroundColor Cyan
}

function Show-Header {
    $header = @"
================================================================
                  WinToPalette - Installation
                Windows Key > PowerToys Command Palette
================================================================
"@
    Write-Host $header -ForegroundColor Cyan
}

function Test-Administrator {
    $currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object System.Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Get-LatestReleaseUrl {
    Write-Info "Fetching latest release information..."
    $releases = Invoke-RestMethod -Uri "https://api.github.com/repos/$GitHubRepo/releases" -ErrorAction SilentlyContinue
    
    if (-not $releases -or $releases.Count -eq 0) {
        Write-Error_ "No releases found"
        return $null
    }
    
    $latestRelease = $releases[0]
    $zipAsset = $latestRelease.assets | Where-Object { $_.name -like "*.zip" } | Select-Object -First 1
    
    if (-not $zipAsset) {
        Write-Error_ "No .zip asset found in latest release"
        return $null
    }
    
    return @{
        Url     = $zipAsset.browser_download_url
        Version = $latestRelease.tag_name
        Name    = $zipAsset.name
    }
}

function Download-Release {
    param([string]$Url, [string]$Version)
    
    $downloadPath = "$env:TEMP\WinToPalette-$Version.zip"
    
    Write-Info "Downloading WinToPalette $Version..."
    $progressPreference = 'SilentlyContinue'
    $result = Invoke-WebRequest -Uri $Url -OutFile $downloadPath -ErrorAction SilentlyContinue
    $progressPreference = 'Continue'
    
    if (Test-Path $downloadPath) {
        Write-Success "Downloaded successfully"
        return $downloadPath
    }
    else {
        Write-Error_ "Download failed"
        return $null
    }
}

function Extract-Release {
    param([string]$ZipPath)
    
    $extractPath = "$env:TEMP\WinToPalette-extracted"
    
    if (Test-Path $extractPath) {
        Remove-Item $extractPath -Recurse -Force
    }
    
    Write-Info "Extracting files..."
    $result = Expand-Archive -Path $ZipPath -DestinationPath $extractPath -Force -ErrorAction SilentlyContinue
    
    if (Test-Path $extractPath) {
        Write-Success "Extracted successfully"
        return $extractPath
    }
    else {
        Write-Error_ "Extraction failed"
        return $null
    }
}

function Install-Interception {
    Write-Info "Setting up Interception driver..."
    
    $interceptionExe = Get-ChildItem -Path "$env:TEMP\WinToPalette-extracted" -Recurse -Filter "install-interception.exe" -ErrorAction SilentlyContinue | Select-Object -First 1
    
    if ($interceptionExe) {
        & $interceptionExe.FullName | Out-Null
        Write-Success "Interception driver installed"
    }
    else {
        Write-Warning_ "Interception installer not found. You may need to install it manually."
    }
}

function Install-Application {
    param([string]$ExtractPath)
    
    Write-Info "Installing application files..."
    
    # Create installation directory
    if (-not (Test-Path $InstallPath)) {
        New-Item -ItemType Directory -Path $InstallPath -Force | Out-Null
    }
    
    # Copy files
    $appFolder = Get-ChildItem -Path $ExtractPath -Directory -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($appFolder) {
        Copy-Item -Path "$($appFolder.FullName)\*" -Destination $InstallPath -Recurse -Force
        Write-Success "Application installed to $InstallPath"
        return $true
    }
    else {
        Write-Error_ "Could not find application folder in release"
        return $false
    }
}

function Register-Startup {
    Write-Info "Registering for auto-startup..."
    
    $exePath = Get-ChildItem -Path $InstallPath -Filter "WinToPalette.exe" -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
    
    if (-not $exePath) {
        Write-Error_ "WinToPalette.exe not found"
        return $false
    }
    
    $TaskName = "WinToPalette"
    $TaskPath = "\WinToPalette\"
    $CurrentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    
    $TaskAction = New-ScheduledTaskAction -Execute $exePath.FullName
    $TaskTrigger = New-ScheduledTaskTrigger -AtLogOn -User $CurrentUser
    $TaskPrincipal = New-ScheduledTaskPrincipal -UserId $CurrentUser -LogonType Interactive -RunLevel Highest
    $TaskSettings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable
    
    $existingTask = Get-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath -ErrorAction SilentlyContinue
    
    if ($existingTask) {
        Set-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath -Action $TaskAction -Principal $TaskPrincipal -Trigger $TaskTrigger -Settings $TaskSettings | Out-Null
    }
    else {
        Register-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath -Action $TaskAction -Principal $TaskPrincipal -Trigger $TaskTrigger -Settings $TaskSettings | Out-Null
    }
    
    Write-Success "Auto-startup registered"
    return $true
}

function Uninstall-Application {
    Show-Header
    Write-Warning_ "Uninstalling WinToPalette..."
    Write-Host ""
    
    # Confirmation prompt
    Write-Host "This will:" -ForegroundColor Yellow
    Write-Host "  • Stop any running WinToPalette process"
    Write-Host "  • Remove the startup task"
    Write-Host "  • Delete application files from $InstallPath"
    Write-Host ""
    $response = Read-Host "Proceed with uninstallation? (y/N)"
    if ($response -ne 'y' -and $response -ne 'Y') {
        Write-Info "Uninstallation cancelled"
        return
    }
    
    Write-Host ""
    
    # Stop running process
    Write-Info "Stopping running instance..."
    Get-Process -Name "WinToPalette" -ErrorAction SilentlyContinue | Stop-Process -Force
    Start-Sleep -Milliseconds 500
    Write-Success "Process stopped"
    
    # Remove scheduled task
    Write-Info "Removing startup task..."
    $TaskName = "WinToPalette"
    $TaskPath = "\WinToPalette\"
    $existingTask = Get-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath -ErrorAction SilentlyContinue
    if ($existingTask) {
        Unregister-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
    }
    Write-Success "Startup task removed"
    
    # Remove application directory
    Write-Info "Removing application files..."
    if (Test-Path $InstallPath) {
        Remove-Item $InstallPath -Recurse -Force -ErrorAction SilentlyContinue
    }
    Write-Success "Application files removed"
    
    # Prompt about logs
    Write-Info "Log files location: %APPDATA%\WinToPalette"
    $response = Read-Host "Remove log files? (y/N)"
    if ($response -eq 'y' -or $response -eq 'Y') {
        $logPath = "$env:APPDATA\WinToPalette"
        if (Test-Path $logPath) {
            Remove-Item $logPath -Recurse -Force -ErrorAction SilentlyContinue
            Write-Success "Log files removed"
        }
    }
    
    Write-Host ""
    Write-Success "WinToPalette uninstalled successfully"
    Write-Warning_ "Reboot recommended for complete Interception driver removal"
}

function Install-Application-Main {
    Show-Header
    
    # Verify admin
    if (-not (Test-Administrator)) {
        Write-Error_ "This script requires administrator privileges"
        exit 1
    }
    
    Write-Success "Administrator privileges confirmed"
    Write-Host ""
    
    # Get latest release
    $releaseInfo = Get-LatestReleaseUrl
    if (-not $releaseInfo) {
        exit 1
    }
    
    Write-Info "Latest version: $($releaseInfo.Version)"
    Write-Host ""
    
    # Confirmation prompt
    Write-Host "This will:" -ForegroundColor Yellow
    Write-Host "  • Download WinToPalette from GitHub"
    Write-Host "  • Install to $InstallPath"
    Write-Host "  • Install Interception driver"
    Write-Host "  • Register for auto-startup"
    Write-Host ""
    $response = Read-Host "Proceed with installation? (y/N)"
    if ($response -ne 'y' -and $response -ne 'Y') {
        Write-Info "Installation cancelled"
        exit 0
    }
    
    Write-Host ""
    
    # Download
    $zipPath = Download-Release -Url $releaseInfo.Url -Version $releaseInfo.Version
    if (-not $zipPath) {
        exit 1
    }
    
    # Extract
    $extractPath = Extract-Release -ZipPath $zipPath
    if (-not $extractPath) {
        exit 1
    }
    
    Write-Host ""
    
    # Install Interception
    Install-Interception
    
    # Install application
    Write-Host ""
    if (-not (Install-Application -ExtractPath $extractPath)) {
        exit 1
    }
    
    # Register startup
    Write-Host ""
    if (-not (Register-Startup)) {
        exit 1
    }
    
    # Cleanup
    Write-Host ""
    Write-Info "Cleaning up temporary files..."
    Remove-Item $zipPath -Force -ErrorAction SilentlyContinue
    Remove-Item $extractPath -Recurse -Force -ErrorAction SilentlyContinue
    
    Write-Host ""
    Write-Success "Installation complete!"
    Write-Info "WinToPalette is ready. Press Windows key to launch PowerToys Command Palette."
    Write-Warning_ "If this is your first time: reboot for Interception driver to become active."
}

# ============================================================================
# MAIN
# ============================================================================

if ($Uninstall) {
    Uninstall-Application
}
else {
    Install-Application-Main
}
