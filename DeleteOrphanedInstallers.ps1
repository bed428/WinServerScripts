# Define the path to the Windows Installer directory
$installersPath = "C:\windows\installer"
# Define the path to the directory where orphaned files will be moved
$destinationPath = "C:\OrphanedInstallers"

# Ensure the directories exist
if (-Not (Test-Path -Path $installersPath)) {
    Write-Host "The path $installersPath does not exist." -ForegroundColor Red
    exit
}

if (-Not (Test-Path -Path $destinationPath)) {
    try {
        New-Item -ItemType Directory -Path $destinationPath -Force
        Write-Host "Created directory: $destinationPath" -ForegroundColor Green
    } catch {
        Write-Host "Failed to create directory: $destinationPath. Error: $_" -ForegroundColor Red
        exit
    }
}

# Get all files in the directory and filter by .msi and .msp extensions
$installerFiles = Get-ChildItem -Path $installersPath | Where-Object { $_.Extension -eq ".msi" -or $_.Extension -eq ".msp" }

# Get all installed products using WMI
$installedProducts = Get-WmiObject -Query "SELECT LocalPackage FROM Win32_Product WHERE LocalPackage IS NOT NULL"

# Get all installed patches from the registry
$installedPatches = Get-ChildItem -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Patches" -Recurse |
    Get-ItemProperty -Name LocalPackage -ErrorAction SilentlyContinue |
    Where-Object { $_.LocalPackage -ne $null }

# Loop through each installer file and check if it's orphaned
foreach ($file in $installerFiles) {
    $inUse = $false

    # Check if the file is in use by installed products
    foreach ($product in $installedProducts) {
        if ($product.LocalPackage -eq $file.FullName) {
            $inUse = $true
            break
        }
    }

    # Check if the file is in use by installed patches if not already marked as in use
    if (-Not $inUse) {
        foreach ($patch in $installedPatches) {
            if ($patch.LocalPackage -eq $file.FullName) {
                $inUse = $true
                break
            }
        }
    }

    # Move the file if it's not in use
    if (-Not $inUse) {
        try {
            Move-Item -Path $file.FullName -Destination $destinationPath -Force
            Write-Host "Moved orphaned installer file: $($file.Name) to $destinationPath" -ForegroundColor Green
        } catch {
            Write-Host "Failed to move file: $($file.Name). Error: $_" -ForegroundColor Red
        }
    } else {
        Write-Host "File $($file.Name) is still in use. Skipping." -ForegroundColor Yellow
    }
}

Write-Host "Cleanup complete." -ForegroundColor Cyan
