function Extract-Year($filename) {
    $currentYear = (Get-Date).Year
    $matches = [regex]::Matches($filename, '\d{4}')

    foreach ($match in $matches) {
        $year = [int]$match.Value
        if ($year -ge 1990 -and $year -le $currentYear) {
            return $year
        }
    }

}

function Extract-YearFromlastmodified($filename) {
    # Get the last modified date of the file
    $lastModified = (Get-Item $filename).LastWriteTime
    $year = $lastModified.Year

    return $year
}

function Extract-YearFromPath($filePath) {
    $currentYear = (Get-Date).Year
    $matches = [regex]::Matches($filePath, '\d{4}')

    foreach ($match in $matches) {
        $year = [int]$match.Value
        if ($year -ge 1990 -and $year -le $currentYear) {
            return $year
        }
    }

    return $null  # If no valid year found in the path
}

function Categorize-File($name) {
    $nameLower = $name.ToLower()

    if ($nameLower -like '*.jpg' -or $nameLower -like '*.png') {
        return 'Photos'
    }
     elseif ($nameLower -like '*.dwg') {
        return 'Drawings'
    }
    elseif ($nameLower.Contains('photos') -or $nameLower.Contains('pictures')) {
        return 'Photos'
    }
    elseif ($nameLower.Contains('change') -or $nameLower.Contains('change form') -or $nameLower.Contains('static data sheet') -or $nameLower.Contains('communication') -or $nameLower -like '*ChangeForm*') {
        return 'MISOR Change Forms'
    }
     elseif ($nameLower.Contains('ami') -or $nameLower.Contains('sp0245') -or $nameLower.Contains('inspection') -or $nameLower.Contains('it inspection') -or $nameLower.Contains('annual meter inspection') -or $nameLower -like '*Annual Metering Inspection*' ) {
        return "Visual Inspections"
    }
    elseif ($nameLower.Contains('audit')) {
        return 'Audit'
    } 
   
    elseif ($nameLower.Contains('ct commissioning') -or $nameLower.Contains('vt commissioning')) {
        if ($nameLower.Contains('ct')) {
            return 'CT Commissioning'
        } else {
            return 'VT Commissioning'
        }
    }
    
    elseif ($nameLower.Contains('1411') -or $nameLower.Contains('itspotcheck') -or $nameLower.Contains('it spot check')) {
        return "IT Spotcheck"
    }
     elseif ($nameLower.Contains('ssla') -or $nameLower.Contains('mec') -or $nameLower.Contains('eitrp') -or $nameLower.Contains('tt')  -or $nameLower.Contains('totalization') -or $nameLower.Contains('ssl')) {
        return "Registration"
    }
    elseif ($nameLower.Contains('cross phase') -or $nameLower -like '*CrossPhase*' ) {
        return "Cross Phase"
    }
    elseif ($nameLower.Contains('tru') -or $nameLower.Contains('inj') -or $nameLower -like '*EUR*') {
        return "EUR"
    }
    elseif ($nameLower.Contains('pt commissioning')) {
        return 'PT Commissioning'
    } 
    elseif ($nameLower.Contains('sld')-or $nameLower -like '*SLD*') {
        return "SLD"
    }
      elseif ($nameLower.Contains('sp203') -or $nameLower.Contains('sp0203')) {
        return "Visual Inspections"
    }
      elseif ($nameLower.Contains('sp204') -or $nameLower.Contains('sp0204')) {
        return "Visual Inspections"
    }
    elseif ($nameLower -match 'cap|capseal|seal') {
        return 'CAP Seals'
    }
    elseif ($nameLower -match 'cap|capseal|seal') {
        return 'CAP Seals'
    }
    else {
        return 'Other'
    }
}



# Function to extract folder name from the path after 'Station Information'
# Function to extract folder name from the path after 'Station Information'
function Get-FolderNameAfterStations {
    param (
        [string]$Path
    )
    # Split the path by backslashes
    $pathParts = $Path -split '\\'
    # Find the index of 'Station Information'
    $stationInfoIndex = [array]::IndexOf($pathParts, 'station information')
    # Return the folder name immediately after 'Station Information'
    if ($stationInfoIndex -ne -1 -and $stationInfoIndex + 1 -lt $pathParts.Length) {
        return $pathParts[$stationInfoIndex + 1]
    }
    return "Unknown"  # Return 'Unknown' if 'Station Information' is not found
}

try {
    Get-ChildItem -Path "C:\Users\214580\desktop\station information" -Recurse -File | ForEach-Object {
        try {
            # Skip files located in 'Estimates' folder
            if ($_.FullName -notmatch "\\Estimates\\") {
                Write-Host "Processing file: $($_.Name)"
                $length = $_.Length
                $type = $_.Extension
                $yearFromModified = Extract-YearFromLastModified($_.FullName)
                $folderName = Get-FolderNameAfterStations($_.FullName)

                $row = [PSCustomObject]@{
                    FileName = $_.Name
                    FileLength = $length
                    FileType = $_.Extension
                    Year = Extract-Year($_.Name)
                    Category = Categorize-File($_.Name)
                    YearFromLastModified = $yearFromModified
                    YearFromPath = Extract-YearFromPath($_.FullName)
                    FolderNameAfterStations = $folderName
                }
                $row | Export-Csv -Path "C:\Users\214580\Desktop\file 8.csv" -NoTypeInformation -Append
            } else {
                Write-Host "Skipping file in 'Estimates' folder: $($_.Name)"
            }
        } catch {
            Write-Host "Error processing file $($_.Name): $_"
        }
    }
} catch {
    Write-Host "General error: $_"
}


