<#
.SYNOPSIS
    Downloads and installs the latest CMV Local Printing Agent

.DESCRIPTION
    This script downloads the latest release from GitHub, extracts it, and runs the installer.
    Designed for IT team deployment on Windows machines.

.PARAMETER Version
    Specific version to download (e.g., "v1.2.3", "v1.0.0-alpha"). Default is "latest".

.PARAMETER InstallPath
    Directory where the agent will be extracted. Default is ".\cmv-agent"

.PARAMETER SkipInstaller
    If specified, downloads and extracts but does not run the installer script.

.EXAMPLE
    .\download-and-install.ps1
    Downloads and installs the latest stable version

.EXAMPLE
    .\download-and-install.ps1 -Version "v1.2.3"
    Downloads and installs specific version v1.2.3

.EXAMPLE
    .\download-and-install.ps1 -Version "v1.0.0-alpha" -InstallPath "C:\CMV-Agent"
    Downloads alpha version to custom location

.EXAMPLE
    .\download-and-install.ps1 -SkipInstaller
    Downloads and extracts but does not run installer (for testing)

.NOTES
    Requires: Internet connectivity, PowerShell 5.1+
    GitHub API rate limit: 60 requests/hour (unauthenticated)
#>

param(
    [string]$Version = "latest",
    [string]$InstallPath = ".\cmv-agent",
    [switch]$SkipInstaller
)

# Configuration
$RepoOwner = "franchise-factory"
$RepoName = "cmv-local-printing-agent-releases"
$TempZip = "$env:TEMP\cmv-agent-$([guid]::NewGuid()).zip"

# Color output helpers
function Write-Header($message) {
    Write-Host "`n$message" -ForegroundColor Cyan
    Write-Host ("=" * $message.Length) -ForegroundColor Cyan
}

function Write-Success($message) {
    Write-Host "  [OK] $message" -ForegroundColor Green
}

function Write-Info($message) {
    Write-Host "  >> $message" -ForegroundColor Yellow
}

function Write-Detail($message) {
    Write-Host "     $message" -ForegroundColor Gray
}

function Write-Failure($message) {
    Write-Host "  [X] $message" -ForegroundColor Red
}

# Main execution
Write-Header "CMV Local Printing Agent Installer"

try {
    # Step 1: Get release information
    Write-Info "Fetching release information..."

    if ($Version -eq "latest") {
        $apiUrl = "https://api.github.com/repos/$RepoOwner/$RepoName/releases/latest"
        Write-Detail "Requesting latest release from GitHub API"
    } else {
        $apiUrl = "https://api.github.com/repos/$RepoOwner/$RepoName/releases/tags/$Version"
        Write-Detail "Requesting version $Version from GitHub API"
    }

    try {
        $release = Invoke-RestMethod -Uri $apiUrl -ErrorAction Stop
    }
    catch {
        if ($_.Exception.Response.StatusCode -eq 404) {
            throw "Version '$Version' not found. Please check available versions at: https://github.com/$RepoOwner/$RepoName/releases"
        }
        elseif ($_.Exception.Response.StatusCode -eq 403) {
            throw "GitHub API rate limit exceeded. Please try again in an hour or contact IT support."
        }
        else {
            throw "Failed to fetch release information: $($_.Exception.Message)"
        }
    }

    $versionTag = $release.tag_name
    Write-Success "Found version: $versionTag"

    if ($release.prerelease) {
        Write-Detail "Note: This is a pre-release version (alpha/beta)"
    }

    # Step 2: Find Windows zip asset
    Write-Info "Locating Windows package..."

    $asset = $release.assets | Where-Object { $_.name -like "*windows.zip" } | Select-Object -First 1

    if (-not $asset) {
        throw "No Windows package found in release $versionTag. Available assets: $($release.assets.name -join ', ')"
    }

    $downloadUrl = $asset.browser_download_url
    $assetSize = [math]::Round($asset.size / 1MB, 2)

    Write-Success "Package: $($asset.name) ($assetSize MB)"
    Write-Detail "Download count: $($asset.download_count)"

    # Step 3: Download
    Write-Info "Downloading package..."
    Write-Detail "From: $downloadUrl"
    Write-Detail "To: $TempZip"

    try {
        # Download with progress
        $ProgressPreference = 'SilentlyContinue'  # Faster download
        Invoke-WebRequest -Uri $downloadUrl -OutFile $TempZip -ErrorAction Stop
        $ProgressPreference = 'Continue'
    }
    catch {
        throw "Download failed: $($_.Exception.Message)"
    }

    $downloadedSize = [math]::Round((Get-Item $TempZip).Length / 1MB, 2)
    Write-Success "Download complete ($downloadedSize MB)"

    # Step 4: Verify download
    if ($downloadedSize -ne $assetSize) {
        Write-Warning "Downloaded size ($downloadedSize MB) differs from expected ($assetSize MB)"
    }

    # Step 5: Extract
    Write-Info "Extracting files..."

    if (Test-Path $InstallPath) {
        Write-Detail "Install path exists, removing old files..."
        try {
            Remove-Item -Path $InstallPath -Recurse -Force -ErrorAction Stop
        }
        catch {
            throw "Failed to remove existing installation at $InstallPath. Please close any running agent processes and try again."
        }
    }

    try {
        New-Item -ItemType Directory -Path $InstallPath -Force | Out-Null
        Expand-Archive -Path $TempZip -DestinationPath $InstallPath -Force -ErrorAction Stop
    }
    catch {
        throw "Extraction failed: $($_.Exception.Message)"
    }

    Write-Success "Extraction complete"
    Write-Detail "Location: $InstallPath"

    # Step 6: Verify files
    Write-Info "Verifying installation files..."

    $exePath = Join-Path $InstallPath "cmv-local-printing-agent.exe"
    $dllPath = Join-Path $InstallPath "libusb-1.0.dll"
    $installerPath = Join-Path $InstallPath "windows-service-install.ps1"

    $missingFiles = @()

    if (-not (Test-Path $exePath)) { $missingFiles += "cmv-local-printing-agent.exe" }
    if (-not (Test-Path $dllPath)) { $missingFiles += "libusb-1.0.dll" }
    if (-not (Test-Path $installerPath)) { $missingFiles += "windows-service-install.ps1" }

    if ($missingFiles.Count -gt 0) {
        throw "Missing required files: $($missingFiles -join ', ')"
    }

    $exeSize = [math]::Round((Get-Item $exePath).Length / 1MB, 2)
    Write-Success "All files verified"
    Write-Detail "cmv-local-printing-agent.exe ($exeSize MB)"
    Write-Detail "libusb-1.0.dll"
    Write-Detail "windows-service-install.ps1"

    # Step 7: Run installer
    if ($SkipInstaller) {
        Write-Header "Download Complete (Installer Skipped)"
        Write-Success "Files extracted to: $InstallPath"
        Write-Info "To install manually, run as Administrator:"
        Write-Detail "cd $InstallPath"
        Write-Detail ".\windows-service-install.ps1"
    }
    else {
        Write-Info "Running installer..."
        Write-Detail "This may require Administrator privileges"
        Write-Detail "You may be prompted for elevation"
        Write-Host ""

        try {
            # Change to install directory and run installer
            Push-Location $InstallPath

            # Run installer script
            & .\windows-service-install.ps1

            $installerExitCode = $LASTEXITCODE
            Pop-Location

            if ($installerExitCode -eq 0) {
                Write-Header "Installation Complete!"
                Write-Success "Version: $versionTag"
                Write-Success "Location: $InstallPath"
                Write-Info "The CMV Local Printing Agent service should now be running"
                Write-Detail "Check service status: Get-Service CMVLocalPrintingAgent"
            }
            else {
                Write-Failure "Installer reported errors (exit code: $installerExitCode)"
                Write-Info "Please check the installer output above for details"
                exit $installerExitCode
            }
        }
        catch {
            Pop-Location
            throw "Installer execution failed: $($_.Exception.Message)"
        }
    }

    # Success summary
    Write-Host ""
    Write-Host "===============================================" -ForegroundColor Cyan
    Write-Host "  Deployment Summary" -ForegroundColor Cyan
    Write-Host "===============================================" -ForegroundColor Cyan
    Write-Host "  Version:    $versionTag" -ForegroundColor White
    Write-Host "  Location:   $InstallPath" -ForegroundColor White
    Write-Host "  Downloaded: $downloadedSize MB" -ForegroundColor White
    Write-Host "===============================================" -ForegroundColor Cyan
    Write-Host ""

} catch {
    Write-Host ""
    Write-Failure "Installation failed!"
    Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    Write-Host "Troubleshooting:" -ForegroundColor Yellow
    Write-Host "  - Check internet connectivity" -ForegroundColor Gray
    Write-Host "  - Verify version exists: https://github.com/$RepoOwner/$RepoName/releases" -ForegroundColor Gray
    Write-Host "  - Check GitHub API rate limits (60/hour)" -ForegroundColor Gray
    Write-Host "  - Ensure sufficient disk space" -ForegroundColor Gray
    Write-Host "  - Try running PowerShell as Administrator" -ForegroundColor Gray
    Write-Host ""

    exit 1
} finally {
    # Cleanup
    if (Test-Path $TempZip) {
        Remove-Item $TempZip -Force -ErrorAction SilentlyContinue
    }
}
