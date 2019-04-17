param(
    [Parameter(Mandatory = $true, ValueFromPipeline = $true)]$url,
    [ValidateSet('skip', 'on', 'off')][System.String]$enableAllManagedProperties = "skip"  
)
# Re-index SPO tenant script, and enable ManagedProperties managed property
# Author: Mikael Svenson - @mikaelsvenson
# Blog: http://techmikael.com

$hasPnP = (Get-Module SharePointPnPPowerShellOnline -ListAvailable).Length
if ($hasPnP -eq 0) {
    Write-Host "This script requires PnP PowerShell, trying to install"
    Find-Module SharePointPnPPowerShellOnline
    Install-Module SharePointPnPPowerShellOnline
}
Import-Module SharePointPnPPowerShellOnline

function Reset-Webs( $siteUrl ) {
    Connect-PnPOnline -Url $siteUrl -SPOManagementShell -SkipTenantAdminCheck
    Write-Host "Processing $siteUrl" -ForegroundColor White
    $web = Get-PnPWeb
    if ( $enableAllManagedProperties -ne "skip" ) {
        Set-AllManagedProperties -web $web -enableAllManagedProps $enableAllManagedProperties
    }
    
    Write-Host "`tSite marked for re-indexing" -ForegroundColor Green
    Request-PnPReIndexWeb

    if ($enableAllManagedProperties -ne "skip") {
        $subWebs = Get-PnPSubWebs -Recurse
        if ($subWebs.Count -gt 0) {
            foreach ($subWeb in $subWebs) {
                Reset-Webs($subWeb.Url)
            }
        }
    }
}

function Set-AllManagedProperties( $web, $enableAllManagedProps ) {
    $clientContext = $web.Context
    $lists = Get-PnPList

    foreach ($list in $lists) {
        Write-Host "`t$($list.Title)"

        if ( $list.NoCrawl ) {
            Write-Host "`t`tSkipping list due to not being crawled" -ForegroundColor Yellow
            continue
        }

        $skip = $false;
        $eventReceivers = Get-PnPEventReceiver -List $list

        foreach ( $eventReceiver in $eventReceivers ) {
            if ( $eventReceiver.ReceiverClass -eq "Microsoft.SharePoint.Publishing.CatalogEventReceiver" ) {
                $skip = $true
                Write-Host "`t`tSkipping list as it's published as a catalog" -ForegroundColor Yellow
                break
            }
        }
        if ( $skip ) { continue }

        $folder = $list.RootFolder
        $props = $folder.Properties
        $clientContext.Load($folder)
        $clientContext.Load($props)
        $clientContext.ExecuteQuery()

        if ( $enableAllManagedProps -eq "on" ) {
            Write-Host "`t`tEnabling all managed properties" -ForegroundColor Green
            $props["vti_indexedpropertykeys"] = "UAB1AGIAbABpAHMAaABpAG4AZwBDAGEAdABhAGwAbwBnAFMAZQB0AHQAaQBuAGcAcwA=|SQBzAFAAdQBiAGwAaQBzAGgAaQBuAGcAQwBhAHQAYQBsAG8AZwA=|"
            $props["IsPublishingCatalog"] = "True"
        }
        if ( $enableAllManagedProps -eq "off" ) {
            Write-Host "`t`tDisabling all managed properties" -ForegroundColor Green
            $props["vti_indexedpropertykeys"] = $null
            $props["IsPublishingCatalog"] = $null
        }
        $folder.Update()
        $clientContext.ExecuteQuery()
    }
}

Connect-PnPOnline -Url $url -SPOManagementShell -SkipTenantAdminCheck
Reset-Webs -siteUrl $url
