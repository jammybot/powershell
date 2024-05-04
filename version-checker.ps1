# Install-Module -Name ImportExcel -Scope CurrentUser
# winget install winget -a
# Install-Module -Name Microsoft.PowerShell.Winget -Scope CurrentUser
#
# Import required modules
Import-Module ImportExcel

# Define log file path
$UptoDateFilePath = "./UptoDateLog.txt"
$NoMatchFilePath = "./NoMatchLog.txt"
$NewUpdateFilePath ="./NewUpdateLog.txt"

# Step 2: Read software names from Excel file
$excelFilePath = "E:\Uni\scripts\Book2.xlsx"
$softwareList = Import-Excel -Path $excelFilePath -WorksheetName "Sheet1" -HeaderRow 1

# Step 3-6: Iterate through each software entry, search in Winget, compare versions, and update Excel
foreach ($softwareEntry in $softwareList) {
    $softwareName = $softwareEntry.Software
    $currentVersion = $softwareEntry.Version
    # Skip processing if either software name is null
    if ([string]::IsNullOrWhiteSpace($softwareName)) {
        Write-Host "Skipping empty row."
        continue
    }
    # Search for the software in Winget assumes that the first one is the one we want. 
    $wingetSearchResult = Find-WingetPackage -Name $softwareName | Select-Object -First 1
    # Basic if statement to catch if there is any software called that
    if ($wingetSearchResult) {
        # For loop to compare
        foreach ($result in $wingetSearchResult) {
            $timestamp = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
            $latestVersion = $result.Version
            # Compare versions
            if ($latestVersion -eq $currentVersion) {
                Write-Output "$timestamp - $softwareName is up to date." | Out-File -FilePath $UptoDateFilePath -Append
            } else {
                Write-Output "$timestamp - $softwareName is not up to date. Current version: $softwareVersion, Latest version: $latestVersion" | Out-File -FilePath $NewUpdateFilePath -Append
                # Update Excel with latest version
                $softwareEntry.LatestVersion = $latestVersion
                Export-Excel -Path $excelFilePath -WorksheetName "Sheet1" -AutoSize -InputObject $softwareList -ClearSheet 
            }
        }
    } else {
        Write-Output "$timestamp - No matching software found for '$softwareName'." | Out-File -FilePath $NoMatchFilePath -Append
    }
}
