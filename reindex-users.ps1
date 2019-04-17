param(
    [Parameter(Mandatory = $true, ValueFromPipeline = $true)]$url,
    [ValidateSet('SPS-Birthday', 'Department')][System.String]$changeProperty = "Department"
)
# Re-index SPO user profiles script
# Author: Mikael Svenson - @mikaelsvenson
# Blog: http://techmikael.com

$hasPnP = (Get-Module SharePointPnPPowerShellOnline -ListAvailable).Length
if ($hasPnP -eq 0) {
    Write-Host "This script requires PnP PowerShell, trying to install"
    Find-Module SharePointPnPPowerShellOnline
    Install-Module SharePointPnPPowerShellOnline
}
Import-Module SharePointPnPPowerShellOnline

function Reset-UserProfiles( $siteUrl ) {
    Write-Host "Retrieving all user profiles" -ForegroundColor Green
    $profiles = Submit-PnPSearchQuery -Query '-AccountName:spofrm -AccountName:spoapp -AccountName:app@sharepoint -AccountName:spocrawler -AccountName:spocrwl -PreferredName:"Foreign Principal"' -SourceId "b09a7990-05ea-4af9-81ef-edfab16c4e31" -SelectProperties "AccountName", "LastModifiedTime", "PreferredName" `
        -All -TrimDuplicates:$false -RelevantResults

    $count = $profiles.Count
    Write-Host "Iterating $count profiles" -ForegroundColor Green
    foreach ($p in $profiles) {
        Write-Host $p.AccountName "Last saved:" $p.LastModifiedTime -ForegroundColor Cyan
        $props = Get-PnPUserProfileProperty -Account $p.AccountName

        if ( $changeProperty -eq "SPS-Birthday" ) {
            $birthday = $props.UserProfileProperties["SPS-Birthday"]
            if ( $null -eq $birthday) {
                Write-Host "`tSkipping as user doesn't have the SPS-Birthday field" -ForegroundColor Yellow
                continue
            }

            # Force save by setting a random birthday value
            Set-PnPUserProfileProperty -Account $p.AccountName -PropertyName "SPS-Birthday" -Value [DateTime]::Now.ToString("yyyyMMddHHmmss.0Z")
            if ( $birthday -eq "" ) {
                Write-Host "`tKeeping birthday as not defined" -ForegroundColor Green
                Set-PnPUserProfileProperty -Account $p.AccountName -PropertyName "SPS-Birthday" -Value [String]::Empty
            }
            else {
                $oldDate = [DateTime]::Parse($birthday)
                Write-Host "`tRe-setting birthday to" $oldDate -ForegroundColor Green
                Set-PnPUserProfileProperty -Account $p.AccountName -PropertyName "SPS-Birthday" -Value $oldDate
            }
        }
        if ( $changeProperty -eq "Department" ) {
            $oldDepartment = $props.UserProfileProperties["Department"]
            if ( $null -eq $oldDepartment) {
                Write-Host "`tSkipping as user doesn't have the Department field" -ForegroundColor Yellow
                continue
            }
            Set-PnPUserProfileProperty -Account $p.AccountName -PropertyName "Department" -Value "mAdcOW reindex placeholder"
            Write-Host "`tRe-setting Department to" $oldDepartment -ForegroundColor Green
            Set-PnPUserProfileProperty -Account $p.AccountName -PropertyName "Department" -Value $oldDepartment
        }
        $count--
        Write-Host "`t$count profiles to go" -ForegroundColor Green
    }
}

if ( $url.tolower() -notlike '*-admin*') {
    Write-Host "This script has to be executed against the admin site of SPO. Eg. https://tenant-admin.sharepoint.com" -ForegroundColor Yellow
    return
}
Connect-PnPOnline -Url $url -SPOManagementShell -SkipTenantAdminCheck
Reset-UserProfiles -siteUrl $url