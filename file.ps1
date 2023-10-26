# Check if the ImportExcel module is available
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    # If not, install it
    Write-Output "Installing the ImportExcel module..."
    Install-Module -Name ImportExcel -Scope CurrentUser -Force
    Import-Module ImportExcel
} else {
    # If already installed, just import it
    Import-Module ImportExcel
}

# Define column names
$columnName1 = "ControlHeader"
$columnName2 = "ControlHeader2"

# Get all CSV and Excel files in the current directory
$filePaths = Get-ChildItem -Filter '*.csv,*.xlsx'

# Process each file
foreach ($filePath in $filePaths) {

    # Check if the file name matches the date pattern and skip if it does
    if ($filePath.Name -match "_\d{8}_\d{4}\.") {
        Write-Output "Skipping file: $($filePath.Name) as it seems to have been processed before."
        continue
    }

    # Import the file based on its type
    if ($filePath.Extension -eq '.csv') {
        $data = Import-Csv -Path $filePath.FullName
    } elseif ($filePath.Extension -eq '.xlsx') {
        $data = Import-Excel -Path $filePath.FullName
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
        $data | Export-Excel -Path $newFilePath
    }
}
