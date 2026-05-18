# ============================================================
# USFinTech Tenant - Export All SharePoint Alert Subscribers
# Requires: PnP.PowerShell module, SharePoint Admin role
# Output: CSV with all alert subscribers across all sites
# ============================================================

$tenantAdminUrl = "https://usfintech-admin.sharepoint.com"
$outputPath     = "C:\Reports\USFinTech_SharePointAlerts_$(Get-Date -Format 'yyyyMMdd').csv"

# Connect to tenant admin
Connect-PnPOnline -Url $tenantAdminUrl -Interactive

# Get all site collections — exclude Archive and TMS sites, keep Url + Title
$allSites = Get-PnPTenantSite -IncludeOneDriveSites:$false |
    Where-Object {
        $_.Url   -notmatch "archive" -and
        $_.Url   -notmatch "TMS" -and
        $_.Title -notmatch "archive" -and
        $_.Title -notmatch "TMS"
    } |
    Select-Object Url, Title

$totalSites   = $allSites.Count
$alertResults = @()
$errors       = @()
$siteCounter  = 0

Write-Host "`n Total sites to process (after exclusions): $totalSites" -ForegroundColor White
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkGray

foreach ($site in $allSites) {

    $siteCounter++
    $percent = [math]::Round(($siteCounter / $totalSites) * 100, 1)

    # Progress line — Site Name prominently in yellow
    Write-Host "`n[$siteCounter / $totalSites | $percent%] " -ForegroundColor DarkGray -NoNewline
    Write-Host "$($site.Title)" -ForegroundColor Yellow -NoNewline
    Write-Host " — $($site.Url)" -ForegroundColor DarkGray

    Write-Progress -Activity "Scanning USFinTech Sites for Alerts" `
                   -Status "$siteCounter of $totalSites — $($site.Title)" `
                   -PercentComplete $percent

    try {
        Connect-PnPOnline -Url $site.Url -ClientId "GUID" -Interactive

        $alerts     = Get-PnPAlert -AllUsers
        $alertCount = $alerts.Count

        Write-Host "   → $alertCount alert(s) found" -ForegroundColor $(if ($alertCount -gt 0) { "Green" } else { "DarkGray" })

        foreach ($alert in $alerts) {
            $alertResults += [PSCustomObject]@{
                SiteName   = $site.Title
                SiteUrl    = $site.Url
                AlertTitle = $alert.Title
                UserEmail  = $alert.User.Email
                UserName   = $alert.User.Title
                ListName   = $alert.List.Title
                AlertType  = $alert.AlertType
                Frequency  = $alert.AlertFrequency
                Status     = $alert.Status
                AlertId    = $alert.Id
            }
        }
    }
    catch {
        $errors += [PSCustomObject]@{
            SiteName = $site.Title
            SiteUrl  = $site.Url
            Error    = $_.Exception.Message
        }
        Write-Host "   → ERROR: $_" -ForegroundColor Red
    }
}

Write-Progress -Activity "Scanning USFinTech Sites for Alerts" -Completed

# Export results
$alertResults | Export-Csv -Path $outputPath -NoTypeInformation -Encoding UTF8

Write-Host "`n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkGray
Write-Host " COMPLETE" -ForegroundColor Green
Write-Host " Sites scanned    : $totalSites" -ForegroundColor White
Write-Host " Total alerts     : $($alertResults.Count)" -ForegroundColor White
Write-Host " Sites with errors: $($errors.Count)" -ForegroundColor $(if ($errors.Count -gt 0) { "Yellow" } else { "Green" })
Write-Host " Report saved     : $outputPath" -ForegroundColor White
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkGray

if ($errors.Count -gt 0) {
    Write-Host "`n Sites with errors:" -ForegroundColor Yellow
    $errors | Format-Table -AutoSize
}
