# WinToPalette Setup Script
# Requires: Administrator privileges

param(
    [switch]$SkipDriver = $false,
    [switch]$SkipCompile = $false,
    [switch]$Uninstall = $false,
    [string]$InterceptionInstallerPath = ""
)

# Colors for output
function Write-Success { Write-Host $args -ForegroundColor Green }
function Write-Error_ { Write-Host $args -ForegroundColor Red }
function Write-Warning_ { Write-Host $args -ForegroundColor Yellow }
function Write-Info { Write-Host $args -ForegroundColor Cyan }

$InterceptionZipUrl = "https://github.com/oblitum/Interception/releases/download/v1.0.1/Interception.zip"

function Test-DotnetSdkAvailable {
    $dotnet = Get-Command dotnet -ErrorAction SilentlyContinue
    if (-not $dotnet) {
        return $false
    }

    try {
        $sdks = & dotnet --list-sdks 2>$null
        if (-not $sdks) {
            return $false
        }

        return ($sdks | Measure-Object).Count -gt 0
    }
    catch {
        return $false
    }
}

function Install-DotnetSdkIfPossible {
    Write-Info "No .NET SDK detected. Attempting installation via winget..."
    $winget = Get-Command winget -ErrorAction SilentlyContinue
    if (-not $winget) {
        Write-Warning_ "winget is not available. Cannot auto-install .NET SDK."
        return $false
    }

    try {
        & winget install --id Microsoft.DotNet.SDK.8 -e --accept-package-agreements --accept-source-agreements
        if ($LASTEXITCODE -ne 0) {
            Write-Warning_ "winget failed to install .NET SDK 8 (exit code: $LASTEXITCODE)."
            return $false
        }

        Write-Success ".NET SDK 8 installation command completed"
        Start-Sleep -Seconds 2
        return (Test-DotnetSdkAvailable)
    }
    catch {
        Write-Warning_ "Automatic .NET SDK installation failed: $_"
        return $false
    }
}

function Test-InterceptionInstalled {
    $serviceKeyboard = Get-Service -Name "keyboard" -ErrorAction Ignore
    $serviceMouse = Get-Service -Name "mouse" -ErrorAction Ignore
    $driverKeyboard = Join-Path $env:WINDIR "System32\drivers\keyboard.sys"
    $driverMouse = Join-Path $env:WINDIR "System32\drivers\mouse.sys"
    $dllSystem32 = Join-Path $env:WINDIR "System32\interception.dll"
    $dllSysWow64 = Join-Path $env:WINDIR "SysWOW64\interception.dll"

    $filtersInstalled = $false
    try {
        $kbdClassKey = "HKLM:\SYSTEM\CurrentControlSet\Control\Class\{4D36E96B-E325-11CE-BFC1-08002BE10318}"
        $mouClassKey = "HKLM:\SYSTEM\CurrentControlSet\Control\Class\{4D36E96F-E325-11CE-BFC1-08002BE10318}"

        $kbdUpper = (Get-ItemProperty -Path $kbdClassKey -ErrorAction SilentlyContinue).UpperFilters
        $mouUpper = (Get-ItemProperty -Path $mouClassKey -ErrorAction SilentlyContinue).UpperFilters

        if ($kbdUpper -and ($kbdUpper -contains "keyboard")) {
            $filtersInstalled = $true
        }

        if ($mouUpper -and ($mouUpper -contains "mouse")) {
            $filtersInstalled = $true
        }
    } catch {
        # Best-effort detection only.
    }

    $isInstalled =
        ($null -ne $serviceKeyboard) -or
        ($null -ne $serviceMouse) -or
        (Test-Path $driverKeyboard) -or
        (Test-Path $driverMouse) -or
        (Test-Path $dllSystem32) -or
        (Test-Path $dllSysWow64) -or
        $filtersInstalled

    return $isInstalled
}

function Get-InterceptionInstaller {
    param(
        [string]$InstallerPath,
        [string]$DestinationDirectory
    )

    if (-not [string]::IsNullOrWhiteSpace($InstallerPath)) {
        if (Test-Path $InstallerPath) {
            return (Resolve-Path $InstallerPath).Path
        }

        Write-Error_ "Provided -InterceptionInstallerPath does not exist: $InstallerPath"
        return $null
    }

    return $null
}

function Get-InstallerCandidateFromExtractedPath {
    param(
        [string]$SearchRoot
    )

    if (-not (Test-Path $SearchRoot)) {
        return $null
    }

    $candidateNames = @(
        "install-interception.exe",
        "interception.exe",
        "installer.exe",
        "install.bat"
    )

    foreach ($name in $candidateNames) {
        $candidate = Get-ChildItem -Path $SearchRoot -Recurse -File -Filter $name -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($candidate) {
            return $candidate.FullName
        }
    }

    return $null
}

function Stage-InterceptionDllFromExtractedPath {
    param(
        [string]$ExtractedRoot,
        [string]$DestinationDirectory
    )

    if (-not (Test-Path $ExtractedRoot)) {
        return
    }

    $preferredCandidates = @(
        (Join-Path $ExtractedRoot "Interception\library\x64\interception.dll"),
        (Join-Path $ExtractedRoot "Interception\library\x86\interception.dll")
    )

    $dllSource = $preferredCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1
    if (-not $dllSource) {
        $dllSource = (Get-ChildItem -Path $ExtractedRoot -Recurse -File -Filter "interception.dll" -ErrorAction SilentlyContinue | Select-Object -First 1).FullName
    }

    if ($dllSource) {
        $dllDest = Join-Path $DestinationDirectory "interception.dll"
        Copy-Item -Path $dllSource -Destination $dllDest -Force
        Write-Success "Staged interception.dll for project build: $dllDest"
    }
}

function Invoke-InterceptionInstaller {
    param(
        [string]$InstallerPath,
        [ValidateSet("install", "uninstall")]
        [string]$Action = "install"
    )

    if (-not (Test-Path $InstallerPath)) {
        Write-Error_ "Installer path not found: $InstallerPath"
        return 1
    }

    $arg = if ($Action -eq "install") { "/install" } else { "/uninstall" }
    $extension = [System.IO.Path]::GetExtension($InstallerPath).ToLowerInvariant()

    if ($extension -eq ".exe") {
        $proc = Start-Process -FilePath $InstallerPath -ArgumentList $arg -Wait -PassThru
        return [int]$proc.ExitCode
    }

    if ($extension -eq ".bat" -or $extension -eq ".cmd") {
        $proc = Start-Process -FilePath "cmd.exe" -ArgumentList "/c `"$InstallerPath`" $arg" -Wait -PassThru
        return [int]$proc.ExitCode
    }

    $proc = Start-Process -FilePath $InstallerPath -Wait -PassThru
    return [int]$proc.ExitCode
}

function Install-InterceptionFromZip {
    param(
        [string]$DestinationDirectory,
        [string]$ZipUrl
    )

    if (-not (Test-Path $DestinationDirectory)) {
        New-Item -ItemType Directory -Path $DestinationDirectory -ErrorAction SilentlyContinue | Out-Null
    }

    $zipPath = Join-Path $DestinationDirectory "Interception.zip"
    $extractPath = Join-Path $DestinationDirectory "Interception_extracted"

    try {
        Write-Info "Downloading Interception zip from GitHub release..."
        Invoke-WebRequest -Uri $ZipUrl -OutFile $zipPath -ErrorAction Stop
        Write-Success "Interception zip downloaded"
    } catch {
        Write-Warning_ "Failed to download Interception zip: $_"
        return -1
    }

    try {
        if (Test-Path $extractPath) {
            Remove-Item -Path $extractPath -Recurse -Force -ErrorAction SilentlyContinue
        }

        Expand-Archive -Path $zipPath -DestinationPath $extractPath -Force
        Write-Success "Interception zip extracted"
    } catch {
        Write-Warning_ "Failed to extract Interception zip: $_"
        return -2
    }

    $installerPath = Get-InstallerCandidateFromExtractedPath -SearchRoot $extractPath
    if (-not $installerPath) {
        Write-Warning_ "Could not find an installer inside extracted Interception package."
        return -3
    }

    Stage-InterceptionDllFromExtractedPath -ExtractedRoot $extractPath -DestinationDirectory $DestinationDirectory

    Write-Info "Installing Interception driver from extracted package: $installerPath"
    $exitCode = Invoke-InterceptionInstaller -InstallerPath $installerPath -Action install
    if ($exitCode -ne 0) {
        Write-Warning_ "Extracted Interception installer exited with code $exitCode."
    }
    else {
        Write-Success "Extracted Interception installer completed"
    }

    return $exitCode
}

# Check for administrator privileges
function Test-Administrator {
    $testPath = "TestAdminRights"
    $adminTest = New-Item -Path $env:TEMP -Name $testPath -ItemType Directory -ErrorAction SilentlyContinue
    if ($adminTest) {
        Remove-Item $adminTest -ErrorAction SilentlyContinue
        return $true
    }
    return $false
}

Write-Info "=== WinToPalette Setup ==="
Write-Info ""

# Verify admin privileges
if (-not (Test-Administrator)) {
    Write-Error_ "This script requires administrator privileges."
    Write-Info "Please run PowerShell as Administrator and try again."
    exit 1
}

Write-Success "Administrator privileges confirmed"
Write-Info ""

# Set paths
$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjectPath = $ScriptPath
$InterceptionPath = Join-Path $ProjectPath "interception"
$BuiltAppPaths = @(
    (Join-Path $ProjectPath "WinToPalette\bin\x64\Release\net8.0-windows\WinToPalette.exe"),
    (Join-Path $ProjectPath "WinToPalette\bin\Release\net8.0-windows\x64\WinToPalette.exe"),
    (Join-Path $ProjectPath "WinToPalette\bin\Release\net8.0-windows\WinToPalette.exe")
)

Write-Info "Project path: $ProjectPath"
Write-Info ""

# Handle uninstall
if ($Uninstall) {
    Write-Warning_ "Uninstalling WinToPalette..."
    Write-Info ""
    
    # Stop running process
    Write-Info "Stopping running instance..."
    $process = Get-Process -Name "WinToPalette" -ErrorAction SilentlyContinue
    if ($process) {
        Stop-Process -Name "WinToPalette" -Force -ErrorAction SilentlyContinue
        Start-Sleep -Milliseconds 500  # Give it time to stop
        Write-Success "Stopped running process"
    } else {
        Write-Info "No running instance found"
    }
    Write-Info ""
    
    # Remove from startup (scheduled task)
    $TaskName = "WinToPalette"
    $TaskPath = "\WinToPalette\"
    
    Write-Info "Removing scheduled task..."
    $existingTask = Get-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath -ErrorAction SilentlyContinue
    if ($existingTask) {
        Unregister-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath -Confirm:$false -ErrorAction Stop
        Write-Success "Removed scheduled task"
    } else {
        Write-Info "Scheduled task not found"
    }
    Write-Info ""
    
    # Uninstall Interception driver
    $localInstaller = Get-ChildItem -Path $InterceptionPath -Recurse -File -Filter "install-interception.exe" -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($localInstaller) {
        Write-Info "Attempting Interception uninstall using: $($localInstaller.FullName)"
        $uninstallCode = Invoke-InterceptionInstaller -InstallerPath $localInstaller.FullName -Action uninstall
        if ($uninstallCode -eq 0) {
            Write-Success "Interception uninstall command completed"
        } else {
            Write-Warning_ "Interception uninstall command exited with code $uninstallCode"
            Write-Warning_ "You can run manually: `"$($localInstaller.FullName)`" /uninstall"
        }
    } else {
        Write-Warning_ "Interception installer not found locally. Run install-interception.exe /uninstall manually if needed."
    }
    Write-Info ""
    
    # Optional: Clean up logs
    $logPath = "$env:APPDATA\WinToPalette"
    if (Test-Path $logPath) {
        Write-Info "Log files found at: $logPath"
        $response = Read-Host "Remove log files? (y/N)"
        if ($response -eq 'y' -or $response -eq 'Y') {
            Remove-Item -Path $logPath -Recurse -Force -ErrorAction SilentlyContinue
            Write-Success "Removed log files"
        } else {
            Write-Info "Log files kept for reference"
        }
    }
    Write-Info ""
    
    Write-Success "WinToPalette uninstalled successfully"
    Write-Warning_ "IMPORTANT: Reboot recommended for complete Interception driver removal"
    exit 0
}

# Download and install Interception driver if not skipped
if (-not $SkipDriver) {
    Write-Info "Setting up Interception driver..."
    $driverInstallExitCode = $null
    
    if (-not (Test-Path $InterceptionPath)) {
        New-Item -ItemType Directory -Path $InterceptionPath -ErrorAction SilentlyContinue | Out-Null
    }
    
    if (Test-InterceptionInstalled) {
        Write-Success "Interception appears to be installed already."
    }
    else {
        $InterceptionExePath = Get-InterceptionInstaller -InstallerPath $InterceptionInstallerPath -DestinationDirectory $InterceptionPath

        if (-not $InterceptionExePath) {
            $winget = Get-Command winget -ErrorAction SilentlyContinue
            if ($winget) {
                Write-Info "Interception installer not provided. Trying winget package install..."
                try {
                    & winget install --id oblitum.Interception -e --accept-package-agreements --accept-source-agreements
                    if ($LASTEXITCODE -eq 0) {
                        Write-Success "Interception installed using winget"
                        $driverInstallExitCode = 0
                    }
                    else {
                        Write-Warning_ "winget could not install oblitum.Interception (exit code: $LASTEXITCODE)."
                        $driverInstallExitCode = $LASTEXITCODE
                    }
                } catch {
                    Write-Warning_ "winget install failed: $_"
                }
            }
            else {
                Write-Warning_ "winget not found. Skipping automatic Interception installation."
            }

            if (-not (Test-InterceptionInstalled)) {
                Write-Info "Trying GitHub release fallback (Interception.zip v1.0.1)..."
                $zipInstallExitCode = Install-InterceptionFromZip -DestinationDirectory $InterceptionPath -ZipUrl $InterceptionZipUrl
                if ($zipInstallExitCode -ge 0) {
                    $driverInstallExitCode = $zipInstallExitCode
                }
                if ($zipInstallExitCode -lt 0) {
                    Write-Warning_ "GitHub release fallback did not complete successfully."
                }
            }
        }

        if ($InterceptionExePath) {
            Write-Info "Installing Interception driver from: $InterceptionExePath"
            $exitCode = Invoke-InterceptionInstaller -InstallerPath $InterceptionExePath -Action install
            $driverInstallExitCode = $exitCode
            if ($exitCode -ne 0) {
                Write-Warning_ "Interception installer exited with code $exitCode."
            }
            else {
                Write-Success "Interception installer completed"
            }
        }

        Start-Sleep -Seconds 1

        if (-not (Test-InterceptionInstalled)) {
            if ($driverInstallExitCode -eq 0) {
                Write-Warning_ "Interception installer reported success, but the driver is not yet detectable."
                Write-Warning_ "A reboot is usually required for this to take effect."
                Write-Info "Reboot Windows, then re-run: .\setup.ps1 -SkipDriver"
            }
            else {
                Write-Error_ "Interception driver is still not detected."
                Write-Info "Manual fix:"
                Write-Info "  1) Download installer from https://github.com/oblitum/Interception/releases"
                Write-Info "  2) Run it as Administrator"
                Write-Info "  3) Re-run this script with -SkipDriver, or provide -InterceptionInstallerPath <path-to-interception.exe>"
                exit 1
            }
        }

        if (Test-InterceptionInstalled) {
            Write-Success "Interception driver installation verified"
        } else {
            Write-Warning_ "Interception installation pending reboot verification"
        }
    }
    
    Write-Info ""
}

# Compile the application if not skipped
if (-not $SkipCompile) {
    Write-Info "Compiling WinToPalette..."

    if (-not (Test-DotnetSdkAvailable)) {
        $sdkInstalled = Install-DotnetSdkIfPossible
        if (-not $sdkInstalled) {
            Write-Error_ "No .NET SDK found. Building requires .NET SDK 8.x or later."
            Write-Info "Install manually from: https://aka.ms/dotnet/download"
            Write-Info "After installation, re-run: .\setup.ps1 -SkipDriver"
            exit 1
        }

        Write-Success ".NET SDK detected"
    }
    
    $ProjectFile = Join-Path $ProjectPath "WinToPalette\WinToPalette.csproj"
    if (-not (Test-Path $ProjectFile)) {
        Write-Error_ "Project file not found: $ProjectFile"
        exit 1
    }
    
    Push-Location $ProjectPath
    try {
        $buildOutput = & dotnet build WinToPalette\WinToPalette.csproj -c Release -p:Platform=x64 2>&1
        $buildExitCode = $LASTEXITCODE
        $buildOutput | ForEach-Object { Write-Host $_ }

        if ($buildExitCode -ne 0) {
            Write-Error_ "Failed to compile application"
            Write-Info "Build diagnostics:"
            $buildOutput | ForEach-Object { Write-Info $_ }
            exit 1
        }
        Write-Success "Application compiled successfully"
    } catch {
        Write-Error_ "Build error: $_"
        exit 1
    } finally {
        Pop-Location
    }
    
    Write-Info ""
}

# Verify the built application
$BuiltAppPath = $BuiltAppPaths | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $BuiltAppPath) {
    Write-Error_ "Built application not found: $BuiltAppPath"
    Write-Info "Make sure the build succeeded and the path is correct"
    exit 1
}

Write-Success "Built application found at: $BuiltAppPath"
Write-Info ""

# Create scheduled task for startup
Write-Info "Creating scheduled task for startup..."
$TaskName = "WinToPalette"
$TaskPath = "\WinToPalette\"
$CurrentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
$TaskActionParams = @{
    Execute  = $BuiltAppPath
}
$TaskAction = New-ScheduledTaskAction @TaskActionParams -ErrorAction Stop

# Run in the interactive user session with highest privileges
$TaskPrincipal = New-ScheduledTaskPrincipal `
    -UserId $CurrentUser `
    -LogonType Interactive `
    -RunLevel Highest `
    -ErrorAction Stop

# Trigger at user logon (interactive desktop session)
$TaskTrigger = New-ScheduledTaskTrigger `
    -AtLogOn `
    -User $CurrentUser `
    -ErrorAction Stop

# Settings for the task
$TaskSettings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -StartWhenAvailable `
    -RunOnlyIfNetworkAvailable:$false `
    -ErrorAction Stop

# Create or update the task
$existingTask = Get-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath -ErrorAction SilentlyContinue
if ($existingTask) {
    Write-Info "Updating existing scheduled task..."
    Set-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath `
        -Action $TaskAction `
        -Principal $TaskPrincipal `
        -Trigger $TaskTrigger `
        -Settings $TaskSettings `
        -ErrorAction Stop | Out-Null
} else {
    Write-Info "Creating new scheduled task..."
    Register-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath `
        -Action $TaskAction `
        -Principal $TaskPrincipal `
        -Trigger $TaskTrigger `
        -Settings $TaskSettings `
        -ErrorAction Stop | Out-Null
}

Write-Success "Scheduled task created for startup (will run at user logon with highest privileges)"
Write-Info ""

# Launch the application
Write-Info "Launching WinToPalette..."
Start-Process -FilePath $BuiltAppPath -Verb RunAs
Write-Info ""

Write-Success "Setup script completed and WinToPalette is now running"
