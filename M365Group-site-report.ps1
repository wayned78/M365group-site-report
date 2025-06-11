# Parameters
$GroupIdentity = Read-Host "Enter the M365 group email or display name"
$OutputCsv = "GroupSharePointSites.csv"

# Connect to Microsoft Graph (uncomment if not already connected)
# Connect-MgGraph -Scopes "Group.Read.All","Sites.Read.All"

# Find the group
$group = Get-MgGroup -Filter "displayName eq '$GroupIdentity' or mail eq '$GroupIdentity'"
if (-not $group) {
    Write-Host "Group not found!" -ForegroundColor Red
    exit
}

# Get the SharePoint site for the group
$site = Get-MgGroupSite -GroupId $group.Id

# Prepare output object
$output = [PSCustomObject]@{
    GroupName = $group.DisplayName
    GroupEmail = $group.Mail
    SharePointUrl = $site.WebUrl
}

# Export to CSV
$output | Export-Csv -Path $OutputCsv -NoTypeInformation

Write-Host "Exported SharePoint site URL to $OutputCsv"
