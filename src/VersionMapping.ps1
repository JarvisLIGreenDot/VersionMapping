# Define the paths of the two folders to compare
$sourceFolderPath = "C:\MyFolder\VersionMapping\versionmap\source"
$targetFolderPath = "C:\MyFolder\VersionMapping\versionmap\targer"

# Define the output Excel file path
$componentName ="CoreProject1"
$outputExcelPath = ".\version_comparison_"+ $componentName +".xlsx"

# Define the file extensions to compare
$fileExtensionsToCompare = @(".dll", ".exe")

# Initialize an empty array to store file information
$fileComparisonList = @()

# Get all files in the source folder, including subdirectories, and filter by extension
$sourceFiles = Get-ChildItem -Path $sourceFolderPath -File -Recurse | Where-Object { $fileExtensionsToCompare -contains $_.Extension }

# Get all files in the target folder, including subdirectories, and filter by extension
$targetFiles = Get-ChildItem -Path $targetFolderPath -File -Recurse | Where-Object { $fileExtensionsToCompare -contains $_.Extension }

# Create a hashtable to store files from the target folder for quick lookup
$targetFilesHashTable = @{}
foreach ($file in $targetFiles) {
    $relativePath = $file.FullName.Substring($targetFolderPath.Length).TrimStart("\")
    $targetFilesHashTable[$relativePath] = $file
}

# Compare files from the source folder
foreach ($sourceFile in $sourceFiles) {
    # Get the relative path of the source file
    $relativePath = $sourceFile.FullName.Substring($sourceFolderPath.Length).TrimStart("\")
    
    # Get the version information of the source file
    $sourceVersionInfo = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($sourceFile.FullName)

    # Check if there is a file with the same relative path in the target folder
    if ($targetFilesHashTable.ContainsKey($relativePath)) {
        $targetFile = $targetFilesHashTable[$relativePath]

        # Get the version information of the target file
        $targetVersionInfo = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($targetFile.FullName)

        # Compare file versions
        $comparisonResult = if ($sourceVersionInfo.FileVersion -eq $targetVersionInfo.FileVersion) { "Same" } else { "Different" }

        # Create an object to store file comparison information
        $fileComparison = [PSCustomObject]@{
            FileName = $sourceFile.Name
            SourceFilePath = "Source/"+$sourceFile.Name
            TargetFilePath = "Target/"+$targetFile.Name
            SourceFileVersion = $sourceVersionInfo.FileVersion
            TargetFileVersion = $targetVersionInfo.FileVersion
            ComparisonResult = $comparisonResult
        }
    } else {
        # If there is no file with the same relative path in the target folder
        $fileComparison = [PSCustomObject]@{
            FileName = $sourceFile.Name
            SourceFilePath =  "Source/"+$sourceFile.Name
            TargetFilePath =  "Target Not Found"
            SourceFileVersion = $sourceVersionInfo.FileVersion
            TargetFileVersion = "Not Found"
            ComparisonResult = "File Missing"
        }
    }

    # Add the file comparison information to the array
    $fileComparisonList += $fileComparison
}

# Create a hashtable to store files from the source folder for quick lookup
$sourceFilesHashTable = @{}
foreach ($file in $sourceFiles) {
    $relativePath = $file.FullName.Substring($sourceFolderPath.Length).TrimStart("\")
    $sourceFilesHashTable[$relativePath] = $file
}

# Compare files from the target folder that are not in the source folder
foreach ($targetFile in $targetFiles) {
    $relativePath = $targetFile.FullName.Substring($targetFolderPath.Length).TrimStart("\")
    if (-not $sourceFilesHashTable.ContainsKey($relativePath)) {
        $targetVersionInfo = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($targetFile.FullName)
        $fileComparison = [PSCustomObject]@{
            FileName          = "Target/" + $targetFile.Name
            SourceFilePath    = "Source Not Found"
            TargetFilePath    = "Target/"+$targetFile.Name
            SourceFileVersion = "Not Found"
            TargetFileVersion = $targetVersionInfo.FileVersion
            ComparisonResult  = "File Missing"
        }
        $fileComparisonList += $fileComparison
    }
}

$fileComparisonPath = [PSCustomObject]@{
    FileName          = "----"
        SourceFilePath    = "Source Path=" + $sourceFolderPath
        TargetFilePath    = "Target Path=" + $targetFolderPath
        SourceFileVersion = "----"
        TargetFileVersion = "----"
        ComparisonResult  = "----"
}
$fileComparisonList += $fileComparisonPath

# Export the file comparison information to an Excel file
$fileComparisonList | Export-Excel -Path $outputExcelPath -AutoSize -WorksheetName "Comparison"


Write-Output "Version comparison has been exported to $outputExcelPath"
Write-Output "File path:"
Write-Output "Source----$sourceFolderPath"
Write-Output "Target----$targetFolderPath"

