# Parameters
$InputCsv = Read-Host "Enter the path to the input CSV file (e.g., groups.csv)"
$OutputCsv = "GroupSharePointSites.csv"

# Ensure you are connected to Exchange Online
# Connect-ExchangeOnline

# Import the list of groups from CSV
$groupsList = Import-Csv -Path $InputCsv

# Prepare an array to hold results
$outputCollection = @()

foreach ($entry in $groupsList) {
    $GroupIdentity = $entry.GroupEmail
    Write-Host "Processing: $GroupIdentity"

    # Find the group using Get-UnifiedGroup
    $group = Get-UnifiedGroup -Identity $GroupIdentity -ErrorAction SilentlyContinue
    if (-not $group) {
        Write-Host "Group not found: $GroupIdentity" -ForegroundColor Red
        continue
    }

    # Prepare output object
    $output = [PSCustomObject]@{
        GroupName     = $group.DisplayName
        GroupEmail    = $group.PrimarySmtpAddress
        SharePointUrl = $group.SharePointSiteUrl
    }

    $outputCollection += $output
}

# Export all results to CSV
$outputCollection | Export-Csv -Path $OutputCsv -NoTypeInformation

Write-Host "Exported SharePoint site URLs to $OutputCsv"
