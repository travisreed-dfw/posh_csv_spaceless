# Define column names
$columnName1 = "ControlHeader"
$columnName2 = "ControlHeader2"

# Get all CSV and Excel files in the current directory
$filePaths = Get-ChildItem -Filter '{*.csv,*.xlsx}'

# Process each file
foreach ($filePath in $filePaths) {

    # Check if the file name matches the date pattern and skip if it does
    if ($filePath.Name -match "_\d{8}_\d{4}\.") {
        Write-Output "Skipping file: $($filePath.Name) as it seems to have been processed before."
        continue
    }

    # Variable to hold data
    $data = @()

    # Import the file based on its type
    if ($filePath.Extension -eq '.csv') {
        $data = Import-Csv -Path $filePath.FullName
    } elseif ($filePath.Extension -eq '.xlsx') {
        # Create a new Excel application object
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $workbook = $excel.Workbooks.Open($filePath.FullName)
        $worksheet = $workbook.Worksheets.Item(1)
        $range = $worksheet.UsedRange
        $data = $range.Value2 | ConvertTo-Csv -Delimiter "," -NoTypeInformation | ConvertFrom-Csv

        # Close Excel without saving and release COM objects
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    } else {
        Write-Warning "Unsupported file type for $($filePath.Name)"
        continue
    }

    # Remove spaces and update columns with comma replacements
    foreach ($row in $data) {
        foreach ($property in $row.PSObject.Properties) {
            # Remove spaces
            $property.Value = $property.Value -replace ' ', ''

            # Replace commas for the specified columns
            if ($property.Name -eq $columnName1 -or $property.Name -eq $columnName2) {
                $property.Value = $property.Value -replace ',', ', '
            }
        }
    }

    # Generate the new filename with date and time appended
    $currentDate = Get-Date -Format "yyyyMMdd_HHmm"
    $newFileName = "$($filePath.BaseName)_$currentDate$($filePath.Extension)"
    $newFilePath = Join-Path -Path $filePath.DirectoryName -ChildPath $newFileName

    # Save the modified content back based on its type
    if ($filePath.Extension -eq '.csv') {
        $data | Export-Csv -Path $newFilePath -NoTypeInformation -Delimiter ","
    } elseif ($filePath.Extension -eq '.xlsx') {
        # Open Excel again for saving data
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $workbook = $excel.Workbooks.Add()
        $worksheet = $workbook.Worksheets.Item(1)
        $worksheet.Name = "Processed Data"
        
        # Convert data to a format suitable for Excel and populate the worksheet
        $dataArray = $data | ConvertTo-Csv -Delimiter "," -NoTypeInformation | ConvertFrom-Json
        $worksheet.Range("A1").Resize($dataArray.Length, $dataArray[0].Length).Value = $dataArray

        # Save the workbook
        $workbook.SaveAs($newFilePath)
        $workbook.Close($true)
        $excel.Quit()
        
        # Release COM objects again
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}
