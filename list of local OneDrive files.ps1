#How to make list of local OneDrive files that haven't finished sync'ing

$ErrorActionPreference = 'Inquire'
function Get-OneDriveRoots {
    # Use a HashSet to ensure unique paths
    $roots = New-Object System.Collections.Generic.HashSet[string]
    
    # 1. Check common environment variables (fastest method)
    foreach ($name in 'OneDrive', 'OneDriveCommercial', 'OneDriveConsumer') {
        $p = [Environment]::GetEnvironmentVariable($name, 'User')
        if ($p -and (Test-Path $p)) {
            $roots.Add((Get-Item $p).FullName) | Out-Null
        }
    }

    # 2. Dynamically scan all account subkeys in the registry
    $regPath = 'HKCU:\Software\Microsoft\OneDrive\Accounts'
    if (Test-Path $regPath) {
        Get-ChildItem $regPath | ForEach-Object {
            # Get the 'UserFolder' property without failing if it's missing
            $val = Get-ItemProperty -Path $_.PsPath -Name 'UserFolder' -ErrorAction SilentlyContinue
            if ($val -and $val.UserFolder -and (Test-Path $val.UserFolder)) {
                $roots.Add((Get-Item $val.UserFolder).FullName) | Out-Null
            }
        }
    }

    # FIX: Cast to array or return the collection explicitly
    #return ,$roots.ToArray() # Added , to avoid unpacking, or just use: [string[]]$roots
    # Corrected return:
    return [string[]]$roots
}

# Example usage:
$myOneDrivePaths = Get-OneDriveRoots
$myOneDrivePaths


$shell = New-Object -ComObject Shell.Application
$StatusHeader = 'Status'   # If your Windows is not in English, change this to the localized column name

$indexCache = @{}

function Get-StatusIndex([string]$folder) {
  if ($indexCache.ContainsKey($folder)) { return $indexCache[$folder] }
  $ns = $shell.NameSpace($folder)
  if (!$ns) { $indexCache[$folder] = -1 ; return -1 }
  $idx = -1
  for ($i = 0; $i -lt 500; $i++) {
    $h = $ns.GetDetailsOf($null, $i)
    if ([string]::IsNullOrWhiteSpace($h)) { continue }
    if ($h -eq $StatusHeader -or $h -eq 'Sync Status') { $idx = $i ; break }
  }
  $indexCache[$folder] = $idx
  return $idx
}

function Get-ItemStatus([string]$path) {
  $folder = Split-Path -LiteralPath $path -Parent
  $name = Split-Path -LiteralPath $path -Leaf
  $ns = $shell.NameSpace($folder)
  if (!$ns) { return $null }
  $idx = Get-StatusIndex $folder
  if ($idx -lt 0) { return $null }
  $item = $ns.ParseName($name)
  if (!$item) { return $null }
  return $ns.GetDetailsOf($item, $idx)
}

$roots = Get-OneDriveRoots
if (-not $roots -or $roots.Count -eq 0) { Write-Error 'No OneDrive folders found' ; return }

# Match any status text that indicates not fully synced. Tweak if you want to be stricter.
$pattern = '(sync|pending|problem|error|conflict|paused|processing)'

$results = New-Object System.Collections.Generic.List[object]

foreach ($root in $roots) {
  Get-ChildItem -LiteralPath $root -File -Recurse -Force | ForEach-Object {
    $st = Get-ItemStatus $_.FullName
    if ($st -and ($st -match $pattern -and $st -notmatch 'up to date')) {
      $results.Add([pscustomobject]@{ Path = $_.FullName ; Status = $st })
    }
  }
}

$csv = "C:\data\OneDrive_NotSynced.csv"
$results | Sort-Object Path | Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8

Write-Host "Items not fully synced: $($results.Count)"
Write-Host "Saved to: $csv"
