# Define column names
$columnName1 = "ControlHeader"
$columnName2 = "ControlHeader2"

# Get all CSV files in the current directory
$csvFiles = Get-ChildItem -Filter *.csv

# Function to detect file encoding
function Get-FileEncoding {
    param([string]$path)
    
    [byte[]]$byte = get-content -Encoding byte -ReadCount 4 -TotalCount 4 -Path $path
    
    if ($byte[0] -eq 0x2b -and $byte[1] -eq 0x2f -and $byte[2] -eq 0x76) { return New-Object System.Text.UTF7Encoding }
    if ($byte[0] -eq 0xef -and $byte[1] -eq 0xbb -and $byte[2] -eq 0xbf) { return New-Object System.Text.UTF8Encoding }
    if ($byte[0] -eq 0xff -and $byte[1] -eq 0xfe) { return New-Object System.Text.UnicodeEncoding }
    if ($byte[0] -eq 0xfe -and $byte[1] -eq 0xff) { return New-Object System.Text.BigEndianUnicodeEncoding }
    if ($byte[0] -eq 0 -and $byte[1] -eq 0 -and $byte[2] -eq 0xfe -and $byte[3] -eq 0xff) { return New-Object System.Text.UTF32Encoding }
    
    return New-Object System.Text.ASCIIEncoding
}

# Process each CSV file
foreach ($csvPath in $csvFiles) {

    # Check if the file name matches the date pattern and skip if it does
    if ($csvPath.Name -match "_\d{8}_\d{4}\.csv$") {
        Write-Output "Skipping file: $($csvPath.Name) as it seems to have been processed before."
        continue
    }

    # Detect the encoding of the file
    $encoding = Get-FileEncoding -Path $csvPath.FullName
    
    # Read the CSV file and replace spaces using the detected encoding
    $content = Get-Content -Path $csvPath.FullName -Raw -Encoding $encoding.EncodingName
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

    # Generate the new filename with date and time appended
    $currentDate = Get-Date -Format "yyyyMMdd_HHmm"
    $newFileName = "$($csvPath.BaseName)_$currentDate.csv"
    $newFilePath = Join-Path -Path $csvPath.DirectoryName -ChildPath $newFileName

    # Save the modified content to the new CSV file in UTF-8
    $csvData | Export-Csv -Path $newFilePath -NoTypeInformation -Encoding UTF8
}
