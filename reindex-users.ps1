param([Parameter(Mandatory=$true,ValueFromPipeline=$true)]$url, [Parameter(Mandatory=$true,ValueFromPipeline=$true)][string]$username, [Parameter(Mandatory=$true,ValueFromPipeline=$true)][string]$password, [ValidateSet('skip','on','off')][System.String]$enableAllManagedProperties="skip"  )
# Re-index SPO tenant script, and enable ManagedProperties managed property
# Author: Mikael Svenson - @mikaelsvenson
# Blog: http://techmikael.blogspot.com

# Modified by Eric Skaggs on 10/21/2014 - had trouble running this script as it was; functionality has not been changed

function Reset-UserProfiles( $siteUrl )
{
	#$clientContext = [mAdcOW.Hack.UPA]::GetContext($siteUrl)
	$clientContext = New-Object Microsoft.SharePoint.Client.ClientContext($siteUrl)
	$clientContext.Credentials = $credentials 
	 
	if (!$clientContext.ServerObjectIsNull.Value) 
	{ 
		Write-Host "Connected to SharePoint Online site: '$siteUrl'" -ForegroundColor Green 
	} 
	$start = 0
	$rowLimit = 100
	do
	{
		$query = new-object Microsoft.SharePoint.Client.Search.Query.KeywordQuery($clientContext)
		$query.QueryText="*"
		$query.SourceId= [Guid]"b09a7990-05ea-4af9-81ef-edfab16c4e31"
		$query.StartRow=$start
		$query.RowLimit=$rowLimit 
		$query.SelectProperties.Add("accountname")
		$query.SelectProperties.Add("write")
		$query.SelectProperties.Add("crawltime")
		$query.SelectProperties.Add("PreferredName")
		$query.TrimDuplicates = $false
		$executor = new-object Microsoft.SharePoint.Client.Search.Query.SearchExecutor($clientContext)
		$result = $executor.ExecuteQuery($query)
		$clientContext.ExecuteQuery()

		$currentCount = 0
		if ($result.Value -ne $null)
		{			
			$currentCount = $result.Value.ResultRows.Length
			Write-Host "Iterating $currentCount profiles" -ForegroundColor Green
			$start = ($start + $rowLimit)
			foreach ($dictionary in $result.Value.ResultRows)
			{
				Write-Host $dictionary["accountname"] "Saved:" $dictionary["write"] "Indexed:" $dictionary["crawltime"] -ForegroundColor Cyan
				$pm = New-Object Microsoft.SharePoint.Client.UserProfiles.PeopleManager($clientContext)
				$props = $pm.GetPropertiesFor($dictionary["accountname"]);
				$clientContext.Load($props)
				$clientContext.ExecuteQuery()
					
				$birthday = $props.UserProfileProperties["SPS-Birthday"]
				if( $birthday -eq $null) {
					Write-Host "`tSkipping as user doesn't have the SPS-Birthday field" -ForegroundColor Yellow
					continue
				}

				# Force save by setting a random birthday value
				$pm.SetSingleValueProfileProperty($props.AccountName, "SPS-Birthday",  [DateTime]::Now.ToString("yyyyMMddHHmmss.0Z"));
				$clientContext.ExecuteQuery()

				if( $birthday -eq "" ) {
					Write-Host "`tRe-setting birthday to blank" -ForegroundColor Green
					$pm.SetSingleValueProfileProperty($props.AccountName, "SPS-Birthday",  [String]::Empty);
				} else {
					$oldDate = [DateTime]::Parse($birthday)
					Write-Host "`tRe-setting birthday to" $oldDate -ForegroundColor Green	
					$pm.SetSingleValueProfileProperty($props.AccountName, "SPS-Birthday",  $oldDate);
				}
				$clientContext.ExecuteQuery()
			}
		}
	}
	while ($currentCount -eq $rowLimit)	
}


# change to the path of your CSOM dlls and add their types
$csomPath = "c:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI"
Add-Type -Path "$csomPath\Microsoft.SharePoint.Client.dll" 
Add-Type -Path "$csomPath\Microsoft.SharePoint.Client.Runtime.dll" 
Add-Type -Path "$csomPath\Microsoft.SharePoint.Client.Search.dll" 
Add-Type -Path "$csomPath\Microsoft.SharePoint.Client.UserProfiles.dll" 

#convert input password to a secure password 
$securePassword = ConvertTo-SecureString $password -AsPlainText -Force 
$credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($username, $securePassword) 
$spoCredentials = New-Object System.Management.Automation.PSCredential($username, $securePassword)
Reset-UserProfiles -siteUrl $url