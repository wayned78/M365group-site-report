# Parameters
$InputCsv = Read-Host "Enter the path to the input CSV file (e.g., groups.csv)"
$OutputCsv = "GroupSharePointSites.csv"

# Connect to Microsoft Graph (uncomment if not already connected)
# Connect-MgGraph -Scopes "Group.Read.All","Sites.Read.All"

# Import the list of groups from CSV
$groupsList = Import-Csv -Path $InputCsv

# Prepare an array to hold results
$outputCollection = @()

foreach ($entry in $groupsList) {
    $GroupIdentity = $entry.GroupEmail
    Write-Host "Processing: $GroupIdentity"

    # Find the group
    $group = Get-MgGroup -Filter "mail eq '$GroupIdentity'"
    if (-not $group) {
        Write-Host "Group not found: $GroupIdentity" -ForegroundColor Red
        continue
    }

    # Get the SharePoint site for the group
    $site = Get-MgGroupSite -GroupId $group.Id

    # Prepare output object
    $output = [PSCustomObject]@{
        GroupName     = $group.DisplayName
        GroupEmail    = $group.Mail
        SharePointUrl = $site.WebUrl
    }

    $outputCollection += $output
}

# Export all results to CSV
$outputCollection | Export-Csv -Path $OutputCsv -NoTypeInformation

Write-Host "Exported SharePoint site URLs to $OutputCsv"
