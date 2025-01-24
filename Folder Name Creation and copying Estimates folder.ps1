# Define paths
$sourcePath = [System.IO.Path]::Combine([System.Environment]::GetFolderPath('Desktop'), 'station information')
$destinationPath = [System.IO.Path]::Combine([System.Environment]::GetFolderPath('Desktop'), 'Station Information 2')

# Create destination folder if it doesn't exist
if (-not (Test-Path $destinationPath)) {
    New-Item -Path $destinationPath -ItemType Directory
}

# Get each subfolder within the source path
Get-ChildItem -Path $sourcePath -Directory | ForEach-Object {
    $subFolder = $_.FullName
    $destinationSubFolder = Join-Path -Path $destinationPath -ChildPath $_.Name

    # Create the corresponding subfolder in the destination if it doesn't exist
    if (-not (Test-Path $destinationSubFolder)) {
        New-Item -Path $destinationSubFolder -ItemType Directory
    }

    # Create 'MTD Depository' folder in the destination subfolder
    $mtdDepositoryFolder = Join-Path -Path $destinationSubFolder -ChildPath "MTD Depository"
    if (-not (Test-Path $mtdDepositoryFolder)) {
        New-Item -Path $mtdDepositoryFolder -ItemType Directory
    }

    # Check if 'estimates' folder exists in the subfolder
    $estimatesFolder = Join-Path -Path $subFolder -ChildPath "Estimates"
    if (Test-Path $estimatesFolder) {
        $destinationEstimatesFolder = Join-Path -Path $destinationSubFolder -ChildPath "Estimates"
        
        # Copy the 'estimates' folder to the destination
        Write-Host "Copying $estimatesFolder to $destinationEstimatesFolder"
        Copy-Item -Path $estimatesFolder -Destination $destinationEstimatesFolder -Recurse -Force
    } else {
        Write-Host "'estimates' folder not found in $subFolder"
    }
}

