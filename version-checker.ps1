# Install-Module -Name ImportExcel -Scope CurrentUser
# winget install winget -a
# Install-Module -Name Microsoft.PowerShell.Winget -Scope CurrentUser

# Import required modules
Import-Module ImportExcel

# Step 2: Read software names from Excel file
$excelFilePath = "E:\Uni\scripts\Book1.xlsx"
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
            $latestVersion = $result.Version
            # Compare versions
            if ($latestVersion -eq $currentVersion) {
                Write-Output "$softwareName is up to date."
            } else {
                Write-Output "$softwareName is not up to date. Current version: $softwareVersion, Latest version: $latestVersion"
                # Update Excel with latest version
                $softwareEntry.LatestVersion = $latestVersion
                Export-Excel -Path $excelFilePath -WorksheetName "Sheet1" -AutoSize -InputObject $softwareList -ClearSheet 
            }
        }
    } else {
        Write-Host "No matching software found for '$softwareName'."
    }
}
