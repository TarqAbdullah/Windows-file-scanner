# Define the list of known "normal" file extensions
$normalExtensions = @(
    ".txt"
)

# Directory to scan
$scanDirectory = "DETERMINE THE SCANNING DIRECTORY HERE"

# Output Excel file path
$outputExcelPath = "DETERMINE THE OUTPUT LOCATION HERE"

# Empty hash table to hold extensions and file paths
$abnormalFiles = @{}

# Function to calculate the hash of a file
function Get-FileHashValue {
    param (
        [string]$filePath
    )
    try {
        # Calculate the hash using SHA256 by default
        $hash = Get-FileHash -Path $filePath -Algorithm SHA256
        return $hash.Hash
    } catch {
        Write-Host "Failed to calculate hash for $filePath. Error: $_"
        return "N/A"
    }
}

# Get all files recursively in the specified directory
$files = Get-ChildItem -Path $scanDirectory -File -Recurse

# Loop through each file and check its extension
foreach ($file in $files) {
    $extension = [System.IO.Path]::GetExtension($file.FullName)

    # Check if the extension is not in the normal extensions list
    if ($normalExtensions -notcontains $extension) {
        # If the extension is not already in the hash table, initialize an array
        if (-not $abnormalFiles.ContainsKey($extension)) {
            $abnormalFiles[$extension] = @()
        }

        # Add the file details (path and hash) to the corresponding extension's array
        $fileHash = Get-FileHashValue -filePath $file.FullName
        $abnormalFiles[$extension] += [PSCustomObject]@{
            FilePath = $file.FullName
            FileHash = $fileHash
        }
    }
}

# If abnormal files were found
if ($abnormalFiles.Count -gt 0) {
    # Create a list to hold data for Excel export
    $excelData = @()

    foreach ($extension in $abnormalFiles.Keys) {
        foreach ($fileDetails in $abnormalFiles[$extension]) {
            # Create a row with the extension, file path, and hash
            $excelData += [PSCustomObject]@{
                "Extension"   = $extension
                "File Path"   = $fileDetails.FilePath
                "File Hashes"   = $fileDetails.FileHash
            }
        }
    }

    # Export the data to an Excel sheet using ImportExcel module
    $excelData | Export-Excel -Path $outputExcelPath -AutoSize -WorksheetName "Abnormal Files"
	
	# Close Excel application after export (**NEW LINE**)
	$excelApp = New-Object -ComObject Excel.Application
	$excelApp.Quit()
	
	# Path to the Python script
	$scriptPath = "PYTHON SCRIPT LOCATION"

	# Run the Python script and wait for it to complete
	Start-Process python -ArgumentList $scriptPath -Wait
	
    Write-Host "Abnormal file extensions report with hashes saved to $outputExcelPath"
} else {
    Write-Host "No abnormal file extensions found."
}