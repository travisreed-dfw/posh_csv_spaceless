# Define column names
$columnName1 = "ControlHeader"
$columnName2 = "ControlHeader2"

# Get all CSV files in the current directory
$csvFiles = Get-ChildItem -Filter *.csv

# Process each CSV file
foreach ($csvPath in $csvFiles) {

    # Check if the file name matches the date pattern and skip if it does
    if ($csvPath.Name -match "_\d{8}_\d{4}\.csv$") {
        Write-Output "Skipping file: $($csvPath.Name) as it seems to have been processed before."
        continue
    }

    # Read the CSV file, replace spaces first, then replace question marks
    $content = Get-Content -Path $csvPath.FullName -Raw -Encoding Default
    $noSpacesContent = $content -replace ' ', ''
    
    # Convert the string content back to CSV object for structured manipulation
    $csvData = $noSpacesContent | ConvertFrom-Csv

    # Update the specified columns
    foreach ($row in $csvData) {
        if ($row.PSObject.Properties.Name -contains $columnName1) {
            $row.$columnName1 = $row.$columnName1 -replace ',', ', '
        }
        if ($row.PSObject.Properties.Name -contains $columnName2) {
            $row.$columnName2 = $row.$columnName2 -replace ',', ', '
        }
    }
    
    # Remove all question marks from the entire CSV data
    $csvData = $csvData | ForEach-Object {
        $_.PSObject.Properties | ForEach-Object {
            $_.Value = $_.Value -replace '\?', ''
        }
        return $_
    }

    # Generate the new filename with date and time appended
    $currentDate = Get-Date -Format "yyyyMMdd_HHmm"
    $newFileName = "$($csvPath.BaseName)_$currentDate.csv"
    $newFilePath = Join-Path -Path $csvPath.DirectoryName -ChildPath $newFileName

    # Save the modified content to the new CSV file with Default encoding
    $csvData | Export-Csv -Path $newFilePath -NoTypeInformation -Encoding Default
}
