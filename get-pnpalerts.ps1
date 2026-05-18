# ============================================================
# USFinTech Tenant - Export All SharePoint Alert Subscribers
# Requires: PnP.PowerShell module, SharePoint Admin role
# Output: CSV with all alert subscribers across all sites
# ============================================================

$tenantAdminUrl = "https://usfintech-admin.sharepoint.com"  # adjust to actual admin URL
$outputPath     = "C:\Reports\USFinTech_SharePointAlerts_$(Get-Date -Format 'yyyyMMdd').csv"

# Connect interactively (MFA-safe)
Connect-PnPOnline -Url $tenantAdminUrl -Interactive

# Get all site collections
$allSites = Get-PnPTenantSite -IncludeOneDriveSites:$false | Select-Object -ExpandProperty Url

$alertResults = @()
$errors       = @()

foreach ($siteUrl in $allSites) {
    Write-Host "Processing: $siteUrl" -ForegroundColor Cyan
    try {
        Connect-PnPOnline -Url $siteUrl -Interactive

        $alerts = Get-PnPAlert -AllUsers

        foreach ($alert in $alerts) {
            $alertResults += [PSCustomObject]@{
                SiteUrl       = $siteUrl
                AlertTitle    = $alert.Title
                UserEmail     = $alert.User.Email
                UserName      = $alert.User.Title
                ListName      = $alert.List.Title
                AlertType     = $alert.AlertType        # AllChanges, ItemAdded, etc.
                Frequency     = $alert.AlertFrequency   # Immediate, Daily, Weekly
                Status        = $alert.Status
                AlertId       = $alert.Id
            }
        }
    }
    catch {
        $errors += [PSCustomObject]@{
            SiteUrl = $siteUrl
            Error   = $_.Exception.Message
        }
        Write-Host "  ERROR on $siteUrl`: $_" -ForegroundColor Red
    }
}

# Export results
$alertResults | Export-Csv -Path $outputPath -NoTypeInformation -Encoding UTF8
Write-Host "`nDone. $($alertResults.Count) alerts found across $($allSites.Count) sites." -ForegroundColor Green
Write-Host "Report saved to: $outputPath" -ForegroundColor Green

# Optional: show errors
if ($errors.Count -gt 0) {
    Write-Host "`n$($errors.Count) sites had errors — review manually." -ForegroundColor Yellow
    $errors | Format-Table -AutoSize
}
