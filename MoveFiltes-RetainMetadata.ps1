# Import PnP Module
Import-Module SharePointPnPPowerShellOnline

# Variables
$sourceSiteUrl = "https://yourtenant.sharepoint.com/sites/source"
$destinationSiteUrl = "https://yourtenant.sharepoint.com/sites/destination"
$sourceLibrary = "SourceDocumentLibrary"
$destinationLibrary = "DestinationDocumentLibrary"
$logFile = "MigrationLog.txt"

# Function to write log
function Write-Log {
    param([string]$message)
    Add-Content -Path $logFile -Value $message
}

# Function to connect to SharePoint site
function Connect-ToSite {
    param([string]$siteUrl)
    try {
        Connect-PnPOnline -Url $siteUrl -UseWebLogin -ErrorAction Stop
    }
    catch {
        Write-Log "Error connecting to site: $siteUrl. Error: $_"
        throw
    }
}

# Connect to source site
Connect-ToSite -siteUrl $sourceSiteUrl

try {
    # Get files from the source library
    $files = Get-PnPListItem -List $sourceLibrary -ErrorAction Stop
}
catch {
    Write-Log "Error retrieving files from source library: $sourceLibrary. Error: $_"
    throw
}

# Connect to destination site
Connect-ToSite -siteUrl $destinationSiteUrl

# Loop through each file and copy
foreach ($file in $files) {
    try {
        $fileStream = Get-PnPFile -Url $file["FileRef"] -AsFilestream -ErrorAction Stop
        $fileName = [System.IO.Path]::GetFileName($file["FileRef"])
        
        # Upload file to destination library
        Add-PnPFile -Path $fileStream -Folder $destinationLibrary -NewFileName $fileName -ErrorAction Stop

        # Get the uploaded file and set metadata
        $uploadedFile = Get-PnPListItem -List $destinationLibrary -Query "<View><Query><Where><Eq><FieldRef Name='FileLeafRef'/><Value Type='File'>$fileName</Value></Eq></Where></Query></View>"
        Set-PnPListItem -List $destinationLibrary -Identity $uploadedFile.Id -Values @{
            "Editor" = $file["Editor"].LookupId; 
            "Author" = $file["Author"].LookupId;
            "Created" = $file["Created"];
            "Modified" = $file["Modified"]
        } -ErrorAction Stop
    }
    catch {
        Write-Log "Error processing file: $fileName. Error: $_"
    }
}

# Disconnect
Disconnect-PnPOnline
