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

    # Import the CSV
    $data = Import-Csv $csvPath.FullName

    # Process each row and cell to remove trailing spaces and update specified columns
    $processedData = $data | ForEach-Object {
        $row = $_

        # Update the specified columns
        if ($row.PSObject.Properties.Name -contains $columnName1) {
            $row.$columnName1 = ($row.$columnName1 -replace ',', ', ').TrimEnd()
        }
        if ($row.PSObject.Properties.Name -contains $columnName2) {
            $row.$columnName2 = ($row.$columnName2 -replace ',', ', ').TrimEnd()
        }

        # Trim trailing spaces from all columns
        $row.PsObject.Properties | ForEach-Object {
            $row.$($_.Name) = $row.$($_.Name).TrimEnd()
        }

        return $row
    }

    # Generate the new filename with date and time appended
    $currentDate = Get-Date -Format "yyyyMMdd_HHmm"
    $newFileName = "$($csvPath.BaseName)_$currentDate.csv"
    $newFilePath = Join-Path -Path $csvPath.DirectoryName -ChildPath $newFileName

    # Export the processed data to the new file
    $processedData | Export-Csv $newFilePath -NoTypeInformation -Encoding Default

    Write-Output "File saved to: $newFilePath"
}
