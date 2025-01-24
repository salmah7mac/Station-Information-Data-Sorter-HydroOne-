# Define paths
$desktopPath = [System.IO.Path]::Combine([System.Environment]::GetFolderPath('Desktop'), 'station information 2')
$constancePath = [System.IO.Path]::Combine([System.Environment]::GetFolderPath('Desktop'), 'station information')
$excelFilePath = [System.IO.Path]::Combine([System.Environment]::GetFolderPath('Desktop'), 'file 7.csv')

# Define folder names
$folders = @('CAP Seals', 'Misor Change Forms', 'IT Spotcheck', 'Cross Phase', 'Eur', 'Annual Metering Inspection', 'Audit', 'Drawings', 'CT Commissioning' , 'VT Commissioning' , "Registration", 'SLD ' , 'other')

# Create the 'stations 2' folder and subfolders
if (-not (Test-Path $desktopPath)) {
    New-Item -Path $desktopPath -ItemType Directory -Force
    Write-Output "Created folder: $desktopPath"
}

foreach ($folder in $folders) {
    $folderPath = [System.IO.Path]::Combine($desktopPath, $folder)
    if (-not (Test-Path $folderPath)) {
        New-Item -Path $folderPath -ItemType Directory -Force
        Write-Output "Created folder: $folderPath"
    }
}

# Create a COM object to access Excel
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false  # Excel will run in the background

try {
    $workbook = $excel.Workbooks.Open($excelFilePath)
    $worksheet = $workbook.Sheets.Item(1)  # Assuming data is in the first sheet

    # Get the last row with data in column A
    $lastRow = $worksheet.Cells.Find("*", $worksheet.Cells.Item(1,1), [Type]::Missing, [Type]::Missing, [Microsoft.Office.Interop.Excel.XlSearchOrder]::xlByRows, [Microsoft.Office.Interop.Excel.XlSearchDirection]::xlPrevious, $false).Row
    Write-Output "Last row in Excel file: $lastRow"

    # Create a dictionary to hold files from the Constance directory
    $fileDictionary = @{}
    Get-ChildItem -Path $constancePath -Recurse -File | ForEach-Object {
        $fileDictionary[$_.Name] = $_.FullName
    }

    # Process each row in the Excel sheet
    for ($row = 2; $row -le $lastRow; $row++) {
        $fileName = $worksheet.Cells.Item($row, 1).Text  # Column A
        $category = $worksheet.Cells.Item($row, 5).Text  # Column E

        # Determine the year from columns D, G, or F
        $year = $worksheet.Cells.Item($row, 4).Text  # Column D
        if ([string]::IsNullOrWhiteSpace($year)) {
            $year = $worksheet.Cells.Item($row, 7).Text  # Column G
            if ([string]::IsNullOrWhiteSpace($year)) {
                $year = $worksheet.Cells.Item($row, 6).Text  # Column F
                if ([string]::IsNullOrWhiteSpace($year)) {
                    $year = "Unknown"  # Default year if none found
                }
            }
        }

        # Get the station name from column H
        $stationName = $worksheet.Cells.Item($row, 8).Text  # Column H

        # Determine the high-level folder based on the station name
        $stationFolder = [System.IO.Path]::Combine($desktopPath, $stationName)
        if (-not (Test-Path $stationFolder)) {
            New-Item -Path $stationFolder -ItemType Directory -Force
            Write-Output "Created folder: $stationFolder"
        }

        # Determine the target folder based on category
        $targetFolder = [System.IO.Path]::Combine($stationFolder, $category)
        if (-not (Test-Path $targetFolder)) {
            New-Item -Path $targetFolder -ItemType Directory -Force
            Write-Output "Created folder: $targetFolder"
        }

        # Determine the year-based subfolder
        $yearFolder = [System.IO.Path]::Combine($targetFolder, $year)
        if (-not (Test-Path $yearFolder)) {
            New-Item -Path $yearFolder -ItemType Directory -Force
            Write-Output "Created folder: $yearFolder"
        }

        # Check if the file exists in the dictionary
        if ($fileDictionary.ContainsKey($fileName)) {
            $filePath = $fileDictionary[$fileName]

            # Move the file to the year-based folder
            try {
                Move-Item -Path $filePath -Destination $yearFolder -Force
                Write-Output "Moved file: $fileName to $yearFolder"
            } catch {
                Write-Output "Failed to move file: $fileName. Error: $_"
            }
        }
    }
} catch {
    Write-Output "Error processing Excel file: $_"
} finally {
    $workbook.Close($false)
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
