# Define column names
$columnName1 = "ControlHeader"
$columnName2 = "ControlHeader2"

# Get all CSV files in the current directory excluding already processed ones
$csvFiles = Get-ChildItem -Filter *.csv | Where-Object { $_.Name -notmatch "_\d{8}_\d{4}\.csv$" }

# Process each CSV file
foreach ($csvPath in $csvFiles) {
    try {
        # Import the CSV
        $data = Import-Csv $csvPath.FullName -Encoding Default

        # Skip if the file is empty
        if (-not $data) {
            Write-Host "Skipping empty file: $($csvPath.Name)"
            continue
        }

        # Process each row to trim spaces and then update specified columns
        foreach ($row in $data) {
            # Trim spaces from all columns
            foreach ($property in $row.PSObject.Properties) {
                $row.$property.Name = $row.$property.Name.Trim()
            }
        
            # Update the specified columns
            if ($columnName1 -in $row.PSObject.Properties.Name) {
                # Replace comma only if not followed by a space
                $row.$columnName1 = $row.$columnName1 -replace ',(?! )', ', '
            }
            if ($columnName2 -in $row.PSObject.Properties.Name) {
                # Replace comma only if not followed by a space
                $row.$columnName2 = $row.$columnName2 -replace ',(?! )', ', '
            }
        }

        # Backup original file
        #Copy-Item -Path $csvPath.FullName -Destination "$($csvPath.FullName).backup"

        # Generate the new filename with date and time appended
        $currentDate = Get-Date -Format "yyyyMMdd_HHmm"
        $newFileName = "$($csvPath.BaseName)_$currentDate.csv"
        $newFilePath = Join-Path -Path $csvPath.DirectoryName -ChildPath $newFileName

        # Export the processed data to the new file
        $data | Export-Csv $newFilePath -NoTypeInformation -Encoding UTF8

        Write-Host "File saved to: $newFilePath"
    } catch {
        $errorMsg = "An error occurred processing $($csvPath.Name): $_"
        Write-Error $errorMsg
        Add-Content -Path "errorlog.txt" -Value $errorMsg
    }
}
