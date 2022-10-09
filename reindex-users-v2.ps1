
# Re-index SPO user profiles script
# Author: Mikael Svenson - @mikaelsvenson
# Blog: http://techmikael.com

<#
.SYNOPSIS
Script to trigger re-indexing of all user profiles

.Description
If you perform search schema mappings after profiles exist you have to update the last modified time on a profile for it to be re-indexed.
This script ensures all profiles are updated with a new time stamp. Once the import job completes allow 4-24h for profiles to be updated in search.

If used in automation replace Connect-PnPOnline with somethine which works for you.

.Parameter url
The site you will use to host the import file. Can be any site you have write access to. DO NOT use the admin site.

.Parameter tempPath
The absolute path on your local disk where we store the import file

.Example 
.\reindex-users-v2.ps1 -url https://contoso.sharepoint.com -tempPath D:\repos\SPO-Trigger-Reindex
Specify temp path

.Example 
.\reindex-users-v2.ps1 -url https://contoso.sharepoint.com
Use script path as the temp path

#>

param(
    [Parameter(Mandatory = $true, ValueFromPipeline = $true)][string]$url,
    [Parameter(Mandatory = $false, ValueFromPipeline = $true)][string]$tempPath = $PSScriptRoot
)

Write-Output "An updates version exists at https://pnp.github.io/script-samples/spo-request-pnp-reindex-user-profile/README.html which is worth taking a look at."

$hasPnP = (Get-Module PnP.PowerShell -ListAvailable).Length
if ($hasPnP -eq 0) {
    Write-Output "This script requires PnP PowerShell, trying to install"
    Find-Module PnP.PowerShell
    Install-Module PnP.PowerShell
}
Import-Module PnP.PowerShell


# Replace connection method as needed
Connect-PnPOnline -Url $url -UseWebLogin

Write-Output "Retrieving all user profiles"
$profiles = Submit-PnPSearchQuery -Query '-AccountName:spofrm -AccountName:spoapp -AccountName:app@sharepoint -AccountName:spocrawler -AccountName:spocrwl -PreferredName:"Foreign Principal"' `
    -SourceId "b09a7990-05ea-4af9-81ef-edfab16c4e31" -SelectProperties "aadobjectid", "department", "write" ` -All -TrimDuplicates:$false -RelevantResults

$fragmentTemplate = "{{""IdName"": ""{0}"",""Department"": ""{1}""}}";
$accountFragments = @();

$profiles |% {
    $aadId =  $_.aadobjectid + ""
    $dept = $_.department + ""
    if(-not [string]::IsNullOrWhiteSpace($aadId) -and $aadId -ne "00000000-0000-0000-0000-000000000000") {
        $accountFragments += [string]::Format($fragmentTemplate,$aadId,$dept)
    }

}
Write-Output "Found $($accountFragments.Count) profiles"
$json = "{""value"":[" + ($accountFragment -join ',') + "]}"

$propertyMap = @{}
$propertyMap.Add("Department", "Department")

$filename = "upa-batch-trigger";
$web = Get-PnPWeb
$folder = $web.GetFolderByServerRelativeUrl("/");

# Cleanup
$files = $folder.Files
$folders = $folder.Folders
Get-PnPProperty -ClientObject $folder -Property Files,Folders

$files |% {
    if($_.Name -like "*$filename*") {
        Write-Output "Remove old import file"
        $_.DeleteObject()
    }
}

$folders |% {
    if($_.Name -like "*$filename*") {
        Write-Output "Remove old import status folder"
        $_.DeleteObject()
    }
}
Invoke-PnPQuery
# End cleanup

$json > "$tempPath/$filename.txt"

Write-Output "Kicking off import job - Please be patient and allow for 4-24h before profiles are updates in search.`n`nDo NOT re-run because you are impatient!"


$job = New-PnPUPABulkImportJob -UserProfilePropertyMapping $propertyMap -IdType CloudId -IdProperty "IdName" -Folder "/" -Path "$tempPath/$filename.txt"
Remove-Item -Path "$tempPath/$filename.txt"

Write-Output "You can check the status of your job with: Get-PnPUPABulkImportStatus -JobId $($job.JobId)"
$job



