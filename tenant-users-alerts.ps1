# ============================================================
# USFinTech Tenant - Export All SharePoint Alert Subscribers
# Hybrid: Tenant sweep + per-user alert detail extraction
# ============================================================

$tenantAdminUrl = "https://usfintech-admin.sharepoint.com"
$outputPath     = "C:\data\SwS\PP\USFinTech_SharePointAlerts_$(Get-Date -Format 'yyyyMMdd').csv"
$logPath        = "C:\data\SwS\PP\USFinTech_SharePointAlerts_Errors_$(Get-Date -Format 'yyyyMMdd').log"
$clientId       = "83b4e9ee-5d0a-4310-a55c-a5a83e6366e3"

# Connect to tenant admin
Connect-PnPOnline -Url $tenantAdminUrl -ClientId $clientId -Interactive

# Get all site collections — exclude Archive, TMS, OneDrive
$allSites = Get-PnPTenantSite -IncludeOneDriveSites:$false |
    Where-Object {
        $_.Url   -notmatch "archive" -and
        $_.Url   -notmatch "TMS" -and
        $_.Url   -notmatch "sharepoint\.com/personal" -and
        $_.Url   -notmatch "-my\." -and
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

    Write-Host "`n[$siteCounter / $totalSites | $percent%] " -ForegroundColor DarkGray -NoNewline
    Write-Host "$($site.Title)" -ForegroundColor Yellow -NoNewline
    Write-Host " — $($site.Url)" -ForegroundColor DarkGray

    Write-Progress -Activity "Scanning USFinTech Sites for Alerts" `
                   -Status "$siteCounter of $totalSites — $($site.Title)" `
                   -PercentComplete $percent

    try {
        Connect-PnPOnline -Url $site.Url -ClientId $clientId -Interactive

        # Get members users of the site — same approach as your original script
        $users = Get-PnPUser | Where-Object {
            $_.LoginName -like "*membership*" -and $_.Email
        }

        $siteAlertCount = 0

        foreach ($user in $users) {
            try {
                $alerts = Get-PnPAlert -User $user

                foreach ($alert in $alerts) {

                    # Expand list and list URL — same as your original script
                    $alertList = $null
                    $listUrl   = $null
                    try {
                        $alertList = Get-PnPProperty -ClientObject $alert -Property List
                        $listUrl   = Get-PnPProperty -ClientObject $alertList -Property DefaultViewURL
                    }
                    catch { <#Do nothing#> }

                    $alertResults += [PSCustomObject]@{
                        SiteName     = $site.Title
                        SiteUrl      = $site.Url
                        UserDisplay  = $user.Title
                        UserLogin    = $user.LoginName
                        UserEmail    = $user.Email
                        AlertTitle   = $alert.Title
                        List         = $alertList.Title
                        URL          = $listUrl
                        Frequency    = $alert.AlertFrequency
                        AlertType    = $alert.AlertType
                        EventType    = $alert.EventType
                        Status       = $alert.Status
                        Delivery     = $alert.DeliveryChannels
                    }
                    $siteAlertCount++
                }
            }
            catch {
                # Silent — user may have no alerts
            }
        }

        Write-Host "   → $siteAlertCount alert(s) found across $($users.Count) users" `
            -ForegroundColor $(if ($siteAlertCount -gt 0) { "Green" } else { "DarkGray" })

    }
    catch {
        $errors += [PSCustomObject]@{
            SiteName = $site.Title
            SiteUrl  = $site.Url
            Error    = $_.Exception.Message
        }
        $details  = @()
        $details += "Time: $(Get-Date)"
        $details += "Message: $($_.Exception.Message)"
        $details += "Category: $($_.CategoryInfo.CategoryInfoName)"
        $details += "FullyQualifiedErrorId: $($_.FullyQualifiedErrorId)"
        $details += "Invocation: $($_.InvocationInfo.PositionMessage)"
        $details += "------------------------------------"
        $details -join "`r`n" | Out-File -FilePath $logPath -Append -Encoding UTF8

        Write-Host "   → ERROR: $($_.Exception.Message)" -ForegroundColor Red
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
Write-Host " Error log        : $logPath" -ForegroundColor White
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkGray

if ($errors.Count -gt 0) {
    Write-Host "`n Sites with errors:" -ForegroundColor Yellow
    $errors | Format-Table -AutoSize
}
