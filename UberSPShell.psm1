# Author: Fırat Akcan
# E-mail: akcan.firat@gmail.com
# 2019

Import-Module Easy-Peasy
Add-PSSnapinIfNotYetAdded Microsoft.SharePoint.PowerShell

function Add-SharepointShellAdmin {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory = $true,HelpMessage = "Please provide the name of the new Farm Administrator in the form of DOMAIN\Username")]
		[ValidateNotNullOrEmpty()]
		[string]$newFarmAdministrator
	)
	SPLogMe
	
	$caWebApp = Get-SPWebApplication -IncludeCentralAdministration | Where-Object { $_.DisplayName -like "*Central Administration*" }
	$caSite = $caWebApp.Sites[0]
	$caWeb = $caSite.RootWeb

	$farmAdministrators = $caWeb.SiteGroups["Farm Administrators"]
	$farmAdministrators.AddUser($newFarmAdministrator, "", $newFarmAdministrator, "Configured via UberSPShell")

	$caWeb.Dispose()
	$caSite.Dispose()

	$caDB = Get-SPContentDatabase -WebApplication $caWebApp
	Add-SPShellAdmin -Database $caDB -UserName $newFarmAdministrator
}

function Enable-SPDeveloperDashboard {
	SPLogMe
	
	$svc = [Microsoft.SharePoint.Administration.SPWebService]::ContentService
	$dds = $svc.DeveloperDashboardSettings
	$dds.DisplayLevel = "On"
	$dds.Update()
	Write-Host "Sharepoint Developer dashboard: On" -ForegroundColor green
}

function Disable-SPDeveloperDashboard {
	SPLogMe
	
	$svc = [Microsoft.SharePoint.Administration.SPWebService]::ContentService
	$dds = $svc.DeveloperDashboardSettings
	$dds.DisplayLevel = "Off"
	$dds.Update()
	Write-Host "Sharepoint Developer dashboard: Off" -ForegroundColor green
}

function Switch-SPDeveloperDashboard {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory = $false,HelpMessage = "On, Off")]
		[ValidateSet('On','Off')]
		[string]$option
	)
	SPLogMe
	
	$svc = [Microsoft.SharePoint.Administration.SPWebService]::ContentService
	$dds = $svc.DeveloperDashboardSettings
	if($option -ne $null)
	{
		$dds.DisplayLevel = $option
	}
	else {
		if($dds.DisplayLevel -eq "Off") {
			$dds.DisplayLevel  = "On"
		}
		else {
			$dds.DisplayLevel  = "Off"
		}
	}
	$dds.Update()
	Write-Host "Sharepoint Developer dashboard: $($dds.DisplayLevel)" -ForegroundColor green
}

function Backup-WSPSolutions {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory = $false,HelpMessage = "Please provide the path of the backup folder, otherwise default path will be used.")]
		[ValidateNotNullOrEmpty()]
		[string]$backupFolder = $env:CustomWspBackupPath
	)
	SPLogMe
	
	$now = (Get-Date).ToString("yyyyMMdd_HHmm")
	$currentBackupFolder = Join-Path -Path $backupFolder -ChildPath $now
	
	$Command = {
		$currentBackupFolder = {0}
		New-Item -ItemType Directory -Force -Path $currentBackupFolder -Force
		Get-SPSolution | ForEach-Object {
			$_.SolutionFile.SaveAs("$currentBackupFolder\$($_.Name)")
		}
	}
	$cmd = Replace-ScriptBlockWithArguments -Command $Command -Arguments $currentBackupFolder	
	DoParallel-OnSPServers -Command $cmd
	
	return $currentBackupFolder
}

function Backup-SPWebConfig {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory = $false,HelpMessage = "Please provide the path of the backup folder, otherwise default path will be used.")]
		[ValidateNotNullOrEmpty()]
		[string]$backupFolder = $env:CustomWebconfigBackupPath
	)
	SPLogMe
	
	DoParallel-OnSPServers -Command {
		Get-SPWebApplication | % {
			$zone = $_.AlternateUrls[0].UrlZone
			$iisSettings = $_.IISSettings[$zone]
			$webconfig = Join-Path -Path $iisSettings.Path.FullName -ChildPath "web.config"			
			$wss = $iisSettings.Path.Name
			$wssBackupPath = Join-Path -Path $backupFolder -ChildPath $wss
			if(-not (Test-Path -Path $wssBackupPath)){
				New-Item -ItemType directory -Path $wssBackupPath
			}
			$wDate = (Get-Date).ToString("_hh_MM_yyyy_HH_mm_")
			$destination = Join-Path -Path $wssBackupPath -ChildPath "web$wDate.config"
			Copy-Item $webconfig -Destination $destination
			Write-Host "Backup for " $_ "on" $server -foreground Green
		}
	}
}

function Rebuild-DistributedCacheServeInstance {
	SPLogMe
	
	$SPFarm = Get-SPFarm
	$cacheClusterName = "SPDistributedCacheCluster_" + $SPFarm.Id.ToString()
	$cacheClusterManager = [Microsoft.SharePoint.DistributedCaching.Utilities.SPDistributedCacheClusterInfoManager]::Local
	$cacheClusterInfo = $cacheClusterManager.GetSPDistributedCacheClusterInfo($cacheClusterName);
	$instanceName = "SPDistributedCacheService Name=AppFabricCachingService"
	$serviceInstance = Get-SPServiceInstance | Where-Object { ($_.Service.ToString()) -eq $instanceName -and ($_.Server.Name) -eq $env:COMPUTERNAME }
	$cacheClusterInfo.CacheHostsInfoCollection
	Remove-SPDistributedCacheServiceInstance
	$serviceInstance.Delete()
	Add-SPDistributedCacheServiceInstance
	$cacheClusterInfo.CacheHostsInfoCollection
}

function Get-SPSolutionLastDeploymentSucceeded {
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true,HelpMessage = "Solution name needed")]
		[ValidateNotNullOrEmpty()]
		[string]$SolutionName
	)
	SPLogMe
	
	# Get the solution again to make sure all deployment info is up-to-date
	$solution = Get-SPSolution -Identity $SolutionName -ErrorAction SilentlyContinue
	if (-not ($solution)) {
		Write-Host "Unable to find solution '$SolutionName'." -ForegroundColor Red
		return $false
	}

	# Check the solution deployment's last operation result was successful
	$lastOperationResult = $solution.LastOperationResult
	if ($lastOperationResult -eq [Microsoft.SharePoint.Administration.SPSolutionOperationResult]::DeploymentSucceeded) {
		Write-Host "Solution '$SolutionName' last deployment succeeded."
		return $true
	}

	$lastOperationDetails = $solution.LastOperationDetails
	Write-Host "Solution '$SolutionName' last operation result is '$lastOperationResult'." -ForegroundColor Red
	Write-Host "Details: $lastOperationDetails" -ForegroundColor Red

	return $false
}

function Create-ManagedPropertiesForSearchService {
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string] $CsvFile = [string]::Empty
	)
	SPLogMe
	
	$LogTime = Get-Date -Format yyyy-MM-dd_hh-mm
	
	if(-not ((Test-Path -Path $CsvFile) -and $Path.EndsWith('.csv'))) {
		Write-Host "There is no file like $CsvFile or it doesn't end with .csv" -Foreground Red
		return $false
	}
	$path = (Get-Item $CsvFile).Directory.FullName
		
	#Get the Search Service Application
	$searchapp = Get-SPEnterpriseSearchServiceApplication
	#Iterate through the CSV file
	Import-Csv $csvfile | ? {
		#Get the Search SErvice Application
		$category = Get-SPEnterpriseSearchMetadataCategory -SearchApplication $searchapp -Identity $_.Category
		#Get the Crawled Property First
		$crawledProperty = Get-SPEnterpriseSearchMetadataCrawledProperty -SearchApplication $searchapp -Name $_.CrawledPropertyName -Category $category -EA silentlycontinue
		#If the Crawled Property is not null, then go inside
		if ($crawledProperty) {
			# Check whether Managed Property already exists
			$property = Get-SPEnterpriseSearchMetadataManagedProperty -SearchApplication $searchapp -Identity $_.ManagedPropertyName -EA silentlycontinue
			if ($property) {
				Write-Host -f yellow "Cannot create managed property" $_.ManagedPropertyName "because it already exists"
				$ExistingManagedProp = "Cannot create managed property " + $_.ManagedPropertyName + " because it already exists" 
				$ExistingManagedProp | Out-File "$path\ExistingManagedProp_$LogTime.txt" -Append
			}
			else {
				# If already not there, then create Managed Property
				New-SPEnterpriseSearchMetadataManagedProperty -Name $_.ManagedPropertyName -SearchApplication $searchapp -Type $_.Type -Description $_.Description -Queryable $true -Retrievable $true
				#Get the managed Property which Just now, we created
				$mp = Get-SPEnterpriseSearchMetadataManagedProperty -SearchApplication $searchapp -Identity $_.ManagedPropertyName
				$mp.Sortable = [System.Convert]::ToBoolean($_.Sortable)
				$mp.Refinable = [System.Convert]::ToBoolean($_.Refinable)
				#Map the Managed Property with the Corresponding Crawled Property
				New-SPEnterpriseSearchMetadataMapping -SearchApplication $searchapp -ManagedProperty $mp -CrawledProperty $crawledProperty
			}
		}
		else {
			Write-Host -foreground Yellow "The specified crawled property " $_.CrawledPropertyName " does not exists... Please check whether you have given valid crawled property name"
			$CrawledPropErrorLog = "The specified crawled property " + $_.CrawledPropertyName + " does not exists... Please check whether you have given valid crawled property name" 
			$CrawledPropErrorLog | Out-File "$path\CrawledPropErrorLogs_$LogTime.txt" -Append
		}
	}
}

function Change-SPDistributedCacheServiceAccount {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory = $true)]
		[ValidateScript(
			{
				$_ -in (Get-SPManagedAccount | % UserName)
			}
		)]
		[string]$Account
	)
	SPLogMe
	
	$farm = Get-SPFarm
	$cacheService = $farm.Services | ? { $_.Name -eq "AppFabricCachingService" }
	$cacheService.ProcessIdentity.CurrentIdentityType = "SpecificUser"
	Write-Host ("Current service account for distributed cache is {0}" -f $cacheService.ProcessIdentity.ManagedAccount.UserName) -foreground yellow
	$cacheService.ProcessIdentity.ManagedAccount = (Get-SPManagedAccount -Identity $Account)
	$cacheService.ProcessIdentity.Update()
	$cacheService.ProcessIdentity.Deploy()
	Write-Host ("Updated service account for distributed cache is {0}" -f $cacheService.ProcessIdentity.ManagedAccount.UserName) -foreground green
}

function Import-SPPropertyBag {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory = $true)]
		[ValidateSet("SPFarm","SPWebApplication","SPSite","SPWeb")]
		[string]$Level,
		[Parameter(Mandatory = $false)]
		[ValidateNotNullOrEmpty()]
		[string]$Url,
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$CsvFile
	)
	SPLogMe
	
	$parent = $null
	if ($Level -eq "SPFarm") {
		$parent = Get-SPFarm
		$properties = $parent.Properties
	}
	else {
		if($Url -eq $null){
			Write-Host "You have to pass a valid url as an argument if you select SPWebApplication, SPSite, SPWeb as a level." -Foreground Red
			return $false
		}
		if ($Level -eq "SPWebApplication") {
			$parent = Get-SPWebApplication -Identity $Url
			$properties = $parent.Properties
		}
		elseif ($Level -eq "SPSite") {
			$parent = Get-SPsite -Identity $Url
			$properties = $parent.RootWeb.Properties
		}
		elseif ($Level -eq "SPWeb") {
			$parent = Get-SPWeb -Identity $Url
			$properties = $parent.Properties
		}
		else {
			Write-Host "Congratulations!`n \|/ You have done the impossible \|/" -Foreground Red
			return $false
		}
	}
	
	if(-not (Test-Path -Path $CsvFile)) {
		Write-Host "There is no file like $CsvFile" -Foreground Red
		return $false
	}
	$importProperties = Import-Csv -Delimiter ";" -Path $CsvFile
	if ($importProperties -ne $null) {
		$importProperties.GetEnumerator() | % {
			if ($properties.ContainsKey($_.Key)) { 
				$properties[$_.Key] = $_.Value 
			} 
			else { 
				$properties.Add($_.Key,$_.Value) 
			}
		}
		$parent.Properties = $properties
		$parent.Update()
		Write-Host "$Level Property Bag is imported." -Foreground Green
		return $true
	}
	else {
		Write-Host "There is no data on the properties file: $CsvFile" -Foreground Red		
		return $false
	}
}

function Export-SPPropertyBag {
	[CmdletBinding()]
	param(		
		[Parameter(Mandatory = $true)]
		[ValidateSet("SPFarm","SPWebApplication","SPSite","SPWeb")]
		[string]$Level,
		[Parameter(Mandatory = $false)]
		[ValidateNotNullOrEmpty()]
		[string]$Url,
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$CsvFile
	)
	SPLogMe
	
	if ($Level -eq "SPFarm") {
		$parent = Get-SPFarm
		$properties = $parent.Properties
	}
	else {
		if($Url -eq $null){
			Write-Host "You have to pass a valid url as an argument if you select SPWebApplication, SPSite, SPWeb as a level." -Foreground Red
			return $false
		}
		if ($Level -eq "SPWebApplication") {
			$parent = Get-SPWebApplication -Identity $Url
			$properties = $parent.Properties
		}
		elseif ($Level -eq "SPSite") {
			$parent = Get-SPsite -Identity $Url
			$properties = $parent.RootWeb.Properties
		}
		elseif ($Level -eq "SPWeb") {
			$parent = Get-SPWeb -Identity $Url
			$properties = $parent.Properties
		}
		else {
			Write-Host "Congratulations!`n \|/ You have done the impossible \|/" -Foreground Red
			return $false
		}
	}
	
	$properties.GetEnumerator() | % { 
		New-Object PSObject -Property @{ Key = $_.Name; Value = $_.Value } 
	} | Export-Csv $CsvFile -notype -Delimiter ";"
	Write-Host "$Level Property Bag is exported." -Foreground Green
	return $true
}

function Get-SPServersOn {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string] $Server,
		[Parameter(Mandatory = $false)]
		[ValidateNotNullOrEmpty()]
		[PSCredential] $Credential
	)
	SPLogMe
	
	if($Credential -eq $null) {
		return (Invoke-Command -ComputerName $Server -ScriptBlock { 
			Get-SPServer | ? { $_.Role -eq "Application" } | % { $_.Address } 
		} )
	}
	else {
		return (Invoke-Command -Credential $Credential -ComputerName $Server -ScriptBlock { 
			Get-SPServer | ? { $_.Role -eq "Application" } | % { $_.Address } 
		} )
	}
}

function Get-SPServers {
	SPLogMe
	
	return (Get-SPServer | ? { $_.Role -eq "Application" } | % { $_.Address })
}

function Get-SPDistributedCacheServersStatus {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory = $false)]
		[ValidateNotNullOrEmpty()]
		[PSCredential] $Credential,
		[switch] $NodeByNode
	)
	SPLogMe
	
	[array]$servers = Get-SPServer | ? {($_.ServiceInstances | % TypeName) -contains 'Distributed Cache'} | % { $_.Address }
	$scriptBlock = {
		Add-PSSnapin Microsoft.SharePoint.PowerShell
		Get-SPServiceInstance | Where-Object { ($_.Service.ToString()) -eq "SPDistributedCacheService Name=AppFabricCachingService" } | Select-Object Server,Status
		Use-CacheCluster
		Get-CacheHost
	}
	switch ($NodeByNode.IsPresent) {
		$false {
			if($Credential -eq $null) {
				return (DoParallel-OnSPServers -Servers $servers -Command $scriptBlock)
			}
			else {
				return (DoParallel-OnSPServers -Servers $servers -Credential $Credential -Command $scriptBlock)
			}
		}
		$true {			
			[array]$result = @()
			$servers | % {
				if($Credential -ne $null) {
					$res = (Invoke-Command -ComputerName $_ -Credential $Credential -ScriptBlock $scriptBlock)
				}
				else {
					$res = (Invoke-Command -ComputerName $_ -ScriptBlock $scriptBlock)
				}
				$result += $res
			}
			return $result
		}
	}
}

function Add-SPSites2Localhost {
	SPLogMe
	
	$hostsfile = Join-Path -Path $env:SystemRoot -ChildPath "System32\drivers\etc\hosts"
	$date = Get-Date -UFormat "%y%m%d%H%M%S"
	$filecopy = $hostsfile + '.' + $env:USERNAME + '.' + $date + '.copy'
	Copy-Item $hostsfile -Destination $filecopy

	# Get a list of the AAMs and weed out the duplicates
	$hosts = Get-SPAlternateURL | ForEach-Object { $_.incomingurl.Replace("https://","").Replace("http://","") } | Where-Object { $_.ToString() -notlike "*:*" } | Select-Object -Unique

	# Get the contents of the Hosts file
	$file = Get-Content $hostsfile
	$file = $file | Out-String

	# write the AAMs to the hosts file, unless they already exist.
	$hosts | ForEach-Object {
		if ($file.Contains($_)) {
			Write-Host "Entry for $_ already exists. Skipping"
		}
		else {
			Write-Host "Adding entry for $_";
			Add-Content -Path $hostsfile -Value "127.0.0.1 `t $_ "
		}
	}
	# Disable the loopback check, since everything we just did will fail if it's enabled
	$regPath = HKLM:\System\CurrentControlSet\Control\Lsa
	$regName = DisableLoopbackCheck
	if(-not (Test-RegistryValue -Path $regPath -Name $regName)) {
		New-ItemProperty $regPath -Name $regName -Value 1 -PropertyType dword
	}
	else {
		$val = (Get-RegistryValue -Path $regPath -Name $regName)
		if($val -ne 1) {
			Set-ItemProperty $regPath -Name $regName -Value 1 -PropertyType dword
		}
	}
}

function Get-SPServersNeedsUpgrade {
	SPLogMe
	
	[array]$servers = Get-SPServer | ? { $_.Role -eq "Application" -and $_.NeedsUpgrade -eq $True } | % Name	
	[array]$serversToBeUpgrade = @()
	
	foreach($server in $servers) {
		$result = (Invoke-Command -ComputerName $server -ScriptBlock { 
			param($pServer) 
			Get-SPServer $pServer | % NeedsUpgrade 
		} -ArgumentList $server )
		
		if($result -eq $True){
			$serversToBeUpgrade += $server
		}
	}
	return $serversToBeUpgrade
}

function Start-SPWindowsServices {
	SPLogMe
	
	$SPservices = @("SPAdminV4", "SPTimerV4", "SPTraceV4", "SPUserCodV4", "SPWriterV4", "OSearch15", "W3SVC")

	foreach ($service in $SPservices)
	{
		Write-Host -ForegroundColor green "Starting $STservices ..."
		Start-Service -Name $service
	}
	iisreset /start
}

function Stop-SPWindowsServices {
	SPLogMe
	
	$SPservices = @("W3SVC", "SPTimerV4", "SPTraceV4", "SPUserCodV4", "SPWriterV4", "OSearch15", "SPAdminV4")

	iisreset /stop
	foreach ($service in $SPservices)
	{
		Write-Host -ForegroundColor Red "Stopping $STservices ..."
		Stop-Service -Name $service
	}	
}

function Upgrade-SPContentDB {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory = $true,HelpMessage = "Content database name needed")]
		[ValidateNotNullOrEmpty()]
		[string]$Name
	)
	SPLogMe
	
	Upgrade-SPContentDatabase -id (Get-SPContentDatabase -Identity $Name).Id
}

function Copy-SPList {
	################################################################    
	# Rahul Rashu | https://gallery.technet.microsoft.com/Powershell-script-to-copy-3890a86f
	################################################################
	#.Synopsis            
	#  Copies a list or document library from one web in a site collection to another web in the same site collection            
	#.DESCRIPTION            
	# Use this function to copy a list or document library and all of its items to a new list in the same web or a different web. You can only copy list and document libraries between webs in same site collection.            
	#.Parameter SourceWebUrl            
	#  The full url to the web that hosts the list that will be copied            
	#.Parameter SourceListName             
	#  The list title of the list that will be copied            
	#.Parameter DestinationWebUrl            
	#  The full url to the web where the list will be copied to            
	#.Parameter DestinationListName            
	#  The name given to the list created at the destination web. If this is omitted, the source list name will be used.            
	#.EXAMPLE            
	#  Copy-SPList -SourceWebUrl http://corporate -SourceListName "SecretDocuments" -DestinationWebUrl http://corporate/moresecureweb            
	#  Copy the SecretDocuments document library from the http://corporate web to the http://corporate/moresecure web, keeping the same list name.            
	#.EXAMPLE            
	#  Copy-SPList -SourceWebUrl http://corporate -SourceListName "SecretDocuments" -DestinationWebUrl http://corporate/lesssecureweb -DestinationListName "NotSoSecretDocuments"            
	#  Copy the SecretDocuments document library from the http://corporate web to the http://corporate/lesssecure web, changing the name of the list to "NotSoSecretDocuments".            
	#.EXAMPLE            
	#  Copy-SPList -SourceWebUrl http://corporate -SourceListName "SecretDocuments" -DestinationWebUrl http://corporate -DestinationListName "ACopyOfTheSecretDocuments"            
	#  Create a copy the SecretDocuments document library in the same web, http://corporate, with the new name "ACopyOfTheSecretDocuments".                
	################################################################            

	[CmdletBinding()]
	param(
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$SourceWebUrl,
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$SourceListName,
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$DestinationWebUrl,
		[Parameter(Mandatory = $false)]
		[ValidateNotNullOrEmpty()]
		[string]$DestinationListName
	)
	SPLogMe
	
	$numberOfActions = 6
	$progressCounter = 0
	Write-Progress -Id 1 -ParentId 0 -Activity "Copying list from site $SourceWebUrl to $DestinationWebUrl" -PercentComplete ($progressCounter++ / $numberOfActions * 100) -Status "Connecting to the sourcen web, $SourceWebUrl."
	
	$site = New-Object Microsoft.SharePoint.SPSite ($SourceWebUrl)
	$web = $site.OpenWeb()
	Write-Progress -Id 1 -ParentId 0 -Activity "Copying list from site $SourceWebUrl to $DestinationWebUrl" -PercentComplete ($progressCounter++ / $numberOfActions * 100) -Status "Connecting to the destination web, $DestinationWebUrl.";
	$destinationSite = New-Object Microsoft.SharePoint.SPSite ($DestinationWebUrl);
	$destinationWeb = $destinationSite.OpenWeb()
	try
	{
		Write-Progress -Id 1 -ParentId 0 -Activity "Copying list from site $SourceWebUrl to $DestinationWebUrl" -PercentComplete ($progressCounter++ / $numberOfActions * 100) -Status "Getting the source list, $SourceListName."
		$sourceList = $web.Lists[$SourceListName]
		$id = [guid]::NewGuid()
		$templateName = [string]::Format("{0}-{1}",$sourceList.Title,$id.ToString())
		$templateFileName = $templateName
		$destinationListDescription = $sourceList.Description
		Write-Progress -Id 1 -ParentId 0 -Activity "Copying list from site $SourceWebUrl to $DestinationWebUrl" -PercentComplete ($progressCounter++ / $numberOfActions * 100) -Status "Saving the source list as a temmplate."
		$sourceList.SaveAsTemplate($templateFileName,$templateName,$sourceList.Description,$true)
		if ([string]::IsNullOrEmpty($DestinationListName)) { 
			$DestinationListName = $SourceListName
		}
		$listTemplate = $site.GetCustomListTemplates($web)[$templateName]
		Write-Progress -Id 1 -ParentId 0 -Activity "Copying list from site $SourceWebUrl to $DestinationWebUrl" -PercentComplete ($progressCounter++ / $numberOfActions * 100) -Status "Creating a new list ($DestinationListName) in the $DestinationWebUrl site."
		$destinationWeb.Lists.Add($destinationListName,$destinationListDescription,$listTemplate) | Out-Null
		$destinationWeb.Update()
		$listTemplates = $site.RootWeb.Lists["List Template Gallery"]
		$lt = $listTemplates.Items | Where-Object { $_.Title -eq $templateName }
		if ($lt -ne $null) { 
			$lt.Delete()
		}
		Write-Host "The list $SourceListName has been copied to $DestinationWebUrl" -ForegroundColor Green
		Write-Progress -Id 1 -ParentId 0 -Activity "Copying list from site $SourceWebUrl to $DestinationWebUrl" -PercentComplete ($progressCounter++ / $numberOfActions * 100) -Status "The list $SourceListName has been copied to $DestinationWebUrl"
		sleep 3
	}
	catch
	{
		Write-Progress -Id 1 -ParentId 0 -Activity "Copying list from site $SourceWebUrl to $DestinationWebUrl" -PercentComplete (100) -Status "Failed to copy the list $SourceListName"
		sleep 3
		Write-Host "An error occurred: $_"
	}
	finally
	{
		if ($web -ne $null) { 
			$web.Dispose() 
		}
		if ($site -ne $null) { 
			$site.Dispose() 
		}
		if ($destinationWeb -ne $null) { 
			$destinationWeb.Dispose() 
		}
		if ($destinationSite -ne $null) { 
			$destinationSite.Dispose() 
		}
	}
}

function Reset-AllSPIIS {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory = $false)]
		[ValidateNotNullOrEmpty()]
		[PSCredential]$Credential,
		[switch]$NodeByNode
	)
	SPLogMe
	
	switch ($NodeByNode.IsPresent) {
		$false {
			if($Credential -eq $null) {
				return (DoParallel-OnSPServers -Command { Start-Process -FilePath "iisreset.exe" -NoNewWindow -Wait })
			}
			else {
				return (DoParallel-OnSPServers -Credential $Credential -Command { Start-Process -FilePath "iisreset.exe" -NoNewWindow -Wait })
			}
		}
		$true {
			[array]$servers = Get-SPServer | ? { $_.Role -eq "Application" } | % { $_.Address }
			[array]$result = @()
			$servers | % {
				if($Credential -ne $null) {
					$res = Invoke-Command -ComputerName $_ -Credential $Credential -ScriptBlock { Start-Process -FilePath "iisreset.exe" -NoNewWindow -Wait }
				}
				else {
					$res = Invoke-Command -ComputerName $_ -ScriptBlock { Start-Process -FilePath "iisreset.exe" -NoNewWindow -Wait }
				}
				$result += $res
			}
			return $result
		}
	}	
}

function DoParallel-OnSPServers {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory = $false)]
		[ValidateNotNullOrEmpty()]
		[array] $Servers = (Get-SPServer | ? { $_.Role -eq "Application" } | % { $_.Address }),
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string] $Command,
		[Parameter(Mandatory = $false)]
		[PSCredential] $Credential,
		[Parameter(Mandatory = $false)]
		[int] $Timeout = 60
	)
	SPLogMe
	
	[array]$result = @()
	if($Command -notlike "*Add-PSSnapin Microsoft.SharePoint.PowerShell*"){
		$Command = "Add-PSSnapin Microsoft.SharePoint.PowerShell`r`n" + $Command
	}
	$Servers | % {
		if($Credential -ne $null) {
			$res = Invoke-Command -ComputerName $_ -Credential $Credential -ScriptBlock $Command -AsJob
		}
		else {
			$res = Invoke-Command -ComputerName $_ -ScriptBlock $Command -AsJob
		}
		$result += $res
	}
	$timeoutCounter = 1
	$rTimeout = 0
	while( $result.Count -ne ($result | ? { $_.State -eq "Completed"}).Count ){
		$waitingfor = $result | ? {$_.State -ne "Completed"} | % { $_.Location }
		Write-Host " Waiting for $waitingfor"
		Start-Sleep -Seconds $timeoutCounter
		$rTimeout += $timeoutCounter
		if($rTimeout -gt $Timeout)
		{
			$str = ""
			$result | ? { $_.State -ne "Completed"} | % { $str += (", {0}" -f $_.Location) }
			if($str.Length -gt 2){
				$str = $str.Substring(2)
			}
			Write-Host -Foreground Red "Timeout Exception for $str"
			break
		}
		$timeoutCounter++
	}
	[array]$report = @()
	$arrow = "{0}{1}{2}" -f [char]9584, [char]9830, [char]9588
	$result | % { 
		$location = $_.Location
		if($_.State -eq "Completed"){ 
			$report += ("{0}:`n{1}`t{2}" -f $location, $arrow, ($_ | Receive-Job) )
		}
		else {
			$report += ("{0}:`n{1}`t{2}" -f $location, $arrow, $_.Jobstateinfo.State)
		}
	}
	$result | Remove-Job
	return $report
}

function SPLogMe {
	if(![System.Diagnostics.EventLog]::SourceExists("UberSPShell")) {
		New-EventLog -LogName "Windows PowerShell" -Source "UberSPShell"
	}
	
	$CallStack = (Get-PSCallStack).Command
	$Args = (Get-PSCallStack).Arguments
	if($CallStack.Count -ge 1) {
		$CallerFunc = $CallStack[1]
		$Arg = $Args[1]
		$fqdnHostname = $env:COMPUTERNAME + "." + $env:USERDNSDOMAIN
		$User = $env:USERNAME
		$elevated = if(([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator")) {"Yes"} else {"No"}
		$Message = "Function: $CallerFunc`r`nArguments: $Arg`r`nUser: $User`r`nElevated Powershell Console: $elevated`r`nHost: $fqdnHostname`r`n`r`n$User called $CallerFunc within Easy-Peasy module at $fqdnHostname"
		Write-EventLog -LogName "Windows PowerShell" -Source "UberSPShell" -EntryType Information -EventID 0 -Message $Message
	}
}

function Get-SPProductInformation {
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $false)]
		[string] $Path = [string]::Empty
	)
	SPLogMe

	$patchList = @()
	$products = Get-SPProduct

	if ($products.Count -lt 1) {
		Write-Host -ForegroundColor Red "No Products found."
		break
	}

	foreach ($product in $products.PatchableUnitDisplayNames) {
		$unit = $products.GetPatchableUnitInfoByDisplayName($product)
		$i = 0

		foreach ($patch in $unit.Patches) {
			$obj = [pscustomobject]@{
				DisplayName = ''
				IsLatest = ''
				Patch = ''
				Version = ''
				SupportUrl = ''
				MissingFrom = ''
			}

			$obj.DisplayName = $unit.DisplayName

			if ($unit.LatestPatch.Version.ToString() -eq $unit.Patches[$i].Version.ToString()) {
				$obj.IsLatest = "Yes"
			}
			else {
				$obj.IsLatest = "No"
			}

			if (($unit.Patches[$i].PatchName) -ne [string]::Empty) {
				if ($unit.Patches[$i].ServersMissingThis.Count -ge 1) {
					$missing = [System.String]::Join(',',$unit.Patches[$i].ServersMissingThis.ServerName)
				}
				else {
					$missing = ''
				}

				$obj.Patch = $unit.Patches[$i].PatchName
				$obj.Version = $unit.Patches[$i].Version.ToString()
				$obj.SupportUrl = $unit.Patches[$i].Link.AbsoluteUri
				$obj.MissingFrom = $missing
				$missing = $null
			}
			else {
				$obj.Patch = "N/A"
				$obj.Version = "N/A"
				$obj.SupportUrl = "N/A"
				$obj.MissingFrom = "N/A"
			}

			$patchList += $obj
			$obj = $null
			++ $i
		}
	}

	if ($Path -ne '') {
		try {
			Test-Path $Path | Out-Null
		}
		catch {
			Write-Host -ForegroundColor Red "Invalid path."
			break
		}
		$date = Get-Date -Format MM-dd
		$farm = Get-SPFarm

		if ($Path.EndsWith('.csv')) {
			$patchList | Export-Csv $Path -NoTypeInformation
			Write-Host -ForegroundColor Green "Build information exported to $Path"
		}
		else {
			$patchList | Export-Csv "$Path\$date-$($farm.EncodedFarmId).csv" -NoTypeInformation
			Write-Host -ForegroundColor Green "Build information exported to $Path\$date-$($farm.EncodedFarmId).csv"
		}
	}
	else {
		return $patchList
	}
}

function Start-AuditReportInterface {
	<#
	.Synopsis
	   This module will extract the audit log events for the specified time period, at the indicated location.
	.DESCRIPTION
	   This module will extract the audit log events for the specified time period, at the indicated location.
	   There are 3 options: 
		- Extract just the <new file uploaded> events from the document libraries;
		- Extract the standard Audit log report with a custom date interval;
		- Extract a the same report as in option two, but with a more user friendly set of data
	#>
	function Get-DocumentLibStatistics {
		[CmdletBinding()]
		param()
		$site = Get-SPWeb (Read-Host "Please provide the site URL")
		Write-Host "Retrieved Site object" -ForegroundColor Green
		#minimum <created date> value
		$format = (Get-Date).ToShortDateString()
		[string]$SaveFormat = "Doc_Stats_" + ((Get-Date).Day).ToString() + "_" + (Get-Date).Month.ToString() + "_" + (Get-Date).Year.ToString() + ".csv"
		$StartD = (Get-Date -Date "$(Read-Host "Please enter the oldest item date you want to start with in the format: $($format)")").ToShortDateString()
		do {
			[string]$location = (Read-Host "Type the path, where you want to save the results")
		}
		until (Test-Path $location)
		$location = $location + $saveformat
		Write-Host "Checking the site lists" -ForegroundColor Green
		$lists = $site.Lists
		foreach ($list in $lists) {
			$count = $list.Items.Count - 1
			if ($list.basetemplate -eq "DocumentLibrary" -and $count -ge "1" -and $list.Title -ne "Style Library" -and $list.Title -ne "Site assets") {
				Write-Host "Found document library: $($list.Title)" -ForegroundColor Green
				foreach ($doc in $list.Items) {
					if ($doc.file.TimeCreated -ne $null -and $doc.file.TimeLastModified -ne $null) {
						[datetime]$x = $doc.file.TimeCreated
						if ($x -ge $StartD -and $x -ne $null) {
							Write-Host "Date created: $($x)"
							Write-Host "File: $($doc.Name)" -ForegroundColor DarkCyan
							$createdD = $x.ToShortDateString()
							$createdT = $x.ToLongTimeString()
							$properties = @{
								'App Id' = "";
								'Event Data' = "";
								'Event Source' = "SharePoint";
								'Custom Event Name' = " ";
								'Event' = "Created";
								'Occurred (GMT)' = "$($createdD) $($createdT)";
								'Document Location' = "$($site.URL)/$($doc.URL)";
								'User ID' = "$($doc.file.author.UserLogin)";
								'Item Type' = "Document";
								'Item Id' = "$($doc.file.UniqueId)";
								'Site Id' = "$($site.ID)";

							} #properties
							$obj = New-Object –TypeName PSObject –Property $properties
							Write-Output $obj | Select-Object 'Site Id','Item Id','Item Type','User ID','Document Location','Occurred (GMT)','Event','Custom Event Name','Event Source','Event Data','App Id' | Export-Csv -Path "$($location)" -Delimiter "," -Encoding UTF8 -Append -NoTypeInformation -Force
						} #time validation
					} #date validation
				} #file loop
			} #template validation
		} #list for loop ends
		Write-Host "Report generated here - $($location)" -ForegroundColor Green
	} # function ends

	function Get-AuditReport {
		#values for the audit log
		$format = (Get-Date).ToShortDateString()
		do {
			[string]$location = Read-Host "Type the path, where you want to save the results" }
		until (Test-Path $location)
		[string]$SaveFormat = "Custom_AuditQuery_" + ((Get-Date).Day).ToString() + "_" + (Get-Date).Month.ToString() + "_" + (Get-Date).Year.ToString() + ".csv"
		$location = $location + $SaveFormat
		$StartD = Get-Date -Date "$(Read-Host "Please enter the start date in the format: $($format)")"
		$EndD = Get-Date -Date "$(Read-Host "Please enter the end date in the format: $($format)")"
		$s1 = Get-SPsite (Read-Host "Please enter the site, to run the query against")
		$q1 = New-Object Microsoft.SharePoint.SPAuditQuery ($s1)
		$q1.SetRangeStart($StartD)
		$q1.SetRangeEnd($EndD)
		$s1.Audit.GetEntries($q1) | Select-Object @{ label = 'Site Id'; e = { "$($_.siteid)" } },@{ label = 'Item Id'; e = { "$($_.itemid)" } },@{ label = 'Item Type'; e = { $_.itemtype } },@{ label = 'User Id'; e = { $_.userid } },@{ label = 'Document location'; e = { $_.doclocation } },@{ label = 'Occurred (GMT)'; e = { $_.occurred } },@{ label = "Event"; e = { $_.eventname } },@{ label = 'Custom Event Name'; e = { $_.eventsource } },@{ label = 'Event Data'; e = { $_.eventdata } },@{ label = 'App Id'; e = { $_.appprincipalid } } | Export-Csv -Path $location -Delimiter "," -Encoding UTF8 -Append -NoTypeInformation -Force
		Write-Host "Report generated here - $($location)" -ForegroundColor Green
	}

	function Get-CustomAuditReport {
		$format = (Get-Date).ToShortDateString()
		do {
			[string]$location = Read-Host "Type the path, where you want to save the results" }
		until (Test-Path $location)
		[string]$SaveFormat = "CustomFormat_Audit_" + ((Get-Date).Day).ToString() + "_" + (Get-Date).Month.ToString() + "_" + (Get-Date).Year.ToString() + ".csv"
		$location = $location + $SaveFormat
		$StartD = Get-Date -Date "$(Read-Host "Please enter the start date in the format: $($format)")"
		$EndD = Get-Date -Date "$(Read-Host "Please enter the end date in the format: $($format)")"
		$s1 = Get-SPsite (Read-Host "Please enter the site, to run the query against")
		$q1 = New-Object Microsoft.SharePoint.SPAuditQuery ($s1)
		$q1.SetRangeStart($StartD)
		$q1.SetRangeEnd($EndD)
		$entries = $s1.Audit.GetEntries($q1)
		$w = Get-SPWeb $s1.Url
		Write-Host "Looping through the Audit entries to customize the output. This will take a while."
		foreach ($entry in $entries) {
			Write-Host "*" -NoNewline -ForegroundColor Red
			Write-Host "*" -NoNewline -ForegroundColor Green
			Write-Host "*" -NoNewline -ForegroundColor White
			Write-Host "*" -NoNewline -ForegroundColor Black
			$occurredD = $entry.occurred.ToShortDateString()
			$occurredT = $entry.occurred.ToShortTimeString()
			$occurred = $occurredD + $occurredT
			$uid = $entry.userid
			if ($uid -gt 0) {
				$userID = $W.allusers.getbyid($uid).userlogin + " " + "<" + $W.allusers.getbyid($uid).Name + ">"
				$DocLoc = $entry.doclocation
				$properties = @{
					'App Id' = "$($entry.appprincipalid)";
					'Event Data' = "$($entry.eventdata)";
					'Event Source' = "$($entry.eventsource)";
					'Custom Event Name' = "$($entry.eventsource)";
					'Event' = "$($entry.eventname)";
					'Occurred (GMT)' = "$($occurred)";
					'Document Location' = "$($DocLoc)";
					'User ID' = "$($userID)";
					'Item Type' = "$($entry.itemtype)";
					'Item Id' = "$($entry.itemId)";
					'Site Id' = "$($entry.siteID)";
					'Site URL' = "$($w.Url)";

				} #properties
				$obj1 = New-Object –TypeName PSObject –Property $properties
				Write-Output $obj1 | Select-Object 'Site URL','Site Id','Item Id','Item Type','User ID','Document Location','Occurred (GMT)','Event','Custom Event Name','Event Source','Event Data','App Id' | Export-Csv -Path $location -Delimiter "," -Encoding UTF8 -Append -NoTypeInformation -Force
			} #if end
		} #for loop end
	} #function end
	
	do {
		[int]$userMenuChoice = 0
		while ($userMenuChoice -lt 1 -or $userMenuChoice -gt 4) {
			Write-Host " "
			Write-Host "Audit Menu" -BackgroundColor Yellow -ForegroundColor Black
			Write-Host "1. File statistics" -ForegroundColor Yellow
			Write-Host "2. Custom Audit Report" -ForegroundColor Yellow
			Write-Host "3. Custom Format Audit Report" -ForegroundColor Yellow
			Write-Host "4. Quit and Exit" -ForegroundColor Yellow

			[int]$userMenuChoice = Read-Host "Please choose an option"
			switch ($userMenuChoice) {
				1 { 
					Write-Host "Preparing..." -ForegroundColor DarkGreen
					Start-Sleep -Seconds 3
					Get-DocumentLibStatistics
				}
				2 { 
					Write-Host "Preparing..." -ForegroundColor DarkGreen
					Start-Sleep -Seconds 3
					Get-AuditReport
				}
				3 { 
					Write-Host "Preparing..." -ForegroundColor DarkGreen
					Start-Sleep -Seconds 3
					Get-CustomAuditReport
				}
				4 { Write-Host "Thank you for using the module. Have a nice day!" -ForegroundColor DarkGreen
					Start-Sleep -Seconds 1
				}
				default {
					Write-Host "Invalid Choice" 
				}
			}
		}
	} while ($userMenuChoice -ne 4)
}

function Upgrade-SPContentDatabasesIfNeeded {
	SPLogMe	
	
	$spDbs = Get-SPContentDatabase | ? { $_.NeedsUpgrade } | % Name
	if($spDbs.Count -gt 0) {		
		Write-Host "There is $($spsDbs.Count) content DB(s) to upgrade." -Foreground Green
		$spDbs | % { Upgrade-SPContentDatabase $_ }
	}
	else {
		Write-Host "There is no content DB to upgrade." -Foreground Red
	}
	
}

function Get-SPContentDatabasesNeedUpgrade {
	SPLogMe	
	
	return (Get-SPContentDatabase | ? { $_.NeedsUpgrade } | % Name)
}

function Upgrade-SPServersIfNeeded {
	SPLogMe
	
	$servers = Get-SPServersNeedsUpgrade
	if($servers.Count -gt 0) {		
		Write-Host "There is $($servers.Count) server to upgrade." -Foreground Green
		$arrow = "{0}{1}{2}" -f [char]9584, [char]9830, [char]9588
		
		foreach($server in $servers) { 
			$result = (Invoke-Command -ComputerName $server -ScriptBlock { psconfig.exe -cmd upgrade b2b -force })
			Write-Host ("{0}:`n{1}`t{2}" -f $server, $arrow, $result)
		}
	}
	else {
		Write-Host "There is no server to upgrade." -Foreground Red
	}
}

function Set-EmailOptionForUserProfiles {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory = $true)]
		[ValidateSet("On","Off")]
		[string]$Option,
		[Parameter(Mandatory = $true)]
		[ValidateScript(
			{
				(Get-SPsite $_) -ne $null
			}
		)]
		[string]$Site
	)
	SPLogMe	
	
	[Reflection.Assembly]::LoadWithPartialName("Microsoft.Office.Server") | Out-Null
	[Reflection.Assembly]::LoadWithPartialName("Microsoft.Sharepoint") | Out-Null

	# 118 email On
	# 126 email Off
	$iOption = 126
	if($Option -eq "On") {
		$iOption = 118
	}
	else {
		$iOption = 126
	}
	
	$serviceContext = Get-SPServiceContext ($Site)
	$profileManager = New-Object Microsoft.Office.Server.UserProfiles.UserProfileManager ($serviceContext)
	$profiles = $profileManager.GetEnumerator()
	$index = 0

	foreach ($profile in $profiles)
	{
		$AccountName = $profile[[Microsoft.Office.Server.UserProfiles.PropertyConstants]::AccountName].Value
		$logstr = "User: " + $AccountName + " | "

		try
		{
			$logstr = $logstr + "EMailOptionsSet:BEGIN | "
			$updated = $false
			$sOption = if($Profile["SPS-EmailOptin"].Value -eq 118) { "On" } else { "Off" }
			$logstr = $logstr + " Option Value: " + $sOption + " | "
			if ($Profile["SPS-EmailOptin"].Value -ne $iOption)
			{
				$Profile["SPS-EmailOptin"].Value = $iOption
				$updated = $true
			}
			if ($updated)
			{
				$Profile.Commit()
				$sOption = if($Profile["SPS-EmailOptin"].Value -eq 118) { "On" } else { "Off" }
				$logstr = $logstr + "OK, new value: " + $sOption + " | "
			}
			else
			{
				$logstr = $logstr + "NotNeccessary | "
			}
		}
		catch
		{
			$logstr = $logstr + "ERROR | " + $_.Exception.Message
		}
		$index = $index + 1
		$remaining = $profileManager.Count - $index
		$logstr = [string]$index + "/" + [string]$profileManager.Count + ":" + [string]$remaining + " | " + $logstr + "PROCESSEND!!! | "
		Write-Host $logstr
	}
	$site.Dispose()
}

function FineTune-DistributedCaches {
	SPLogMe
	
	#DistributedLogonTokenCache
	$DLTC = Get-SPDistributedCacheClientSetting -ContainerType DistributedLogonTokenCache
	$DLTC.MaxConnectionsToServer = 1
	$DLTC.requestTimeout = "3000"
	$DLTC.channelOpenTimeOut = "3000"
	Set-SPDistributedCacheClientSetting -ContainerType DistributedLogonTokenCache $DLTC

	#DistributedViewStateCache
	$DVSC = Get-SPDistributedCacheClientSetting -ContainerType DistributedViewStateCache
	$DVSC.MaxConnectionsToServer = 1
	$DVSC.requestTimeout = "3000"
	$DVSC.channelOpenTimeOut = "3000"
	Set-SPDistributedCacheClientSetting -ContainerType DistributedViewStateCache $DVSC

	#DistributedAccessCache
	$DAC = Get-SPDistributedCacheClientSetting -ContainerType DistributedAccessCache
	$DAC.MaxConnectionsToServer = 1
	$DAC.requestTimeout = "3000"
	$DAC.channelOpenTimeOut = "3000"
	Set-SPDistributedCacheClientSetting -ContainerType DistributedAccessCache $DAC

	#DistributedActivityFeedCache
	$DAF = Get-SPDistributedCacheClientSetting -ContainerType DistributedActivityFeedCache
	$DAF.MaxConnectionsToServer = 1
	$DAF.requestTimeout = "3000"
	$DAF.channelOpenTimeOut = "3000"
	Set-SPDistributedCacheClientSetting -ContainerType DistributedActivityFeedCache $DAF

	#DistributedActivityFeedLMTCache
	$DAFC = Get-SPDistributedCacheClientSetting -ContainerType DistributedActivityFeedLMTCache
	$DAFC.MaxConnectionsToServer = 1
	$DAFC.requestTimeout = "3000"
	$DAFC.channelOpenTimeOut = "3000"
	Set-SPDistributedCacheClientSetting -ContainerType DistributedActivityFeedLMTCache $DAFC

	#DistributedBouncerCache
	$DBC = Get-SPDistributedCacheClientSetting -ContainerType DistributedBouncerCache
	$DBC.MaxConnectionsToServer = 1
	$DBC.requestTimeout = "3000"
	$DBC.channelOpenTimeOut = "3000"
	Set-SPDistributedCacheClientSetting -ContainerType DistributedBouncerCache $DBC

	#DistributedDefaultCache
	$DDC = Get-SPDistributedCacheClientSetting -ContainerType DistributedDefaultCache
	$DDC.MaxConnectionsToServer = 1
	$DDC.requestTimeout = "3000"
	$DDC.channelOpenTimeOut = "3000"
	Set-SPDistributedCacheClientSetting -ContainerType DistributedDefaultCache $DDC

	#DistributedSearchCache
	$DSC = Get-SPDistributedCacheClientSetting -ContainerType DistributedSearchCache
	$DSC.MaxConnectionsToServer = 1
	$DSC.requestTimeout = "3000"
	$DSC.channelOpenTimeOut = "3000"
	Set-SPDistributedCacheClientSetting -ContainerType DistributedSearchCache $DSC

	#DistributedSecurityTrimmingCache
	$DTC = Get-SPDistributedCacheClientSetting -ContainerType DistributedSecurityTrimmingCache
	$DTC.MaxConnectionsToServer = 1
	$DTC.requestTimeout = "3000"
	$DTC.channelOpenTimeOut = "3000"
	Set-SPDistributedCacheClientSetting -ContainerType DistributedSecurityTrimmingCache $DTC

	#DistributedServerToAppServerAccessTokenCache
	$DSTAC = Get-SPDistributedCacheClientSetting -ContainerType DistributedServerToAppServerAccessTokenCache
	$DSTAC.MaxConnectionsToServer = 1
	$DSTAC.requestTimeout = "3000"
	$DSTAC.channelOpenTimeOut = "3000"
	Set-SPDistributedCacheClientSetting -ContainerType DistributedServerToAppServerAccessTokenCache $DSTAC
	
}

function Get-SPManagedAccountsPassword {
	SPLogMe
	
	function _Bindings
	{
		return [System.Reflection.BindingFlags]::CreateInstance -bor
		[System.Reflection.BindingFlags]::GetField -bor
		[System.Reflection.BindingFlags]::Instance -bor
		[System.Reflection.BindingFlags]::NonPublic
	}
	
	function _GetFieldValue ([object]$o,[string]$fieldName)
	{
		$bindings = _Bindings
		return $o.GetType().GetField($fieldName,$bindings).GetValue($o);
	}
	
	function _ConvertTo-UnsecureString ([System.Security.SecureString]$string)
	{
		$intptr = [System.IntPtr]::Zero
		$unmanagedString = [System.Runtime.InteropServices.Marshal]::SecureStringToGlobalAllocUnicode($string)
		$unsecureString = [System.Runtime.InteropServices.Marshal]::PtrToStringUni($unmanagedString)
		[System.Runtime.InteropServices.Marshal]::ZeroFreeGlobalAllocUnicode($unmanagedString)
		return $unsecureString
	}

	return Get-SPManagedAccount | Select-Object UserName,@{ Name = "Password"; Expression = { _ConvertTo-UnsecureString (_GetFieldValue $_ "m_Password").SecureStringValue } }
}

function Get-SPDatabasesBackupSize {
	SPLogMe
	
	return Get-SPDatabase | Sort-Object DiskSizeRequired -desc | ForEach-Object { $db = 0; $cArray = @{} } { $db += $_.DiskSizeRequired; $cArray.Add($_.Name,$_.DiskSizeRequired / 1GB) } { $cArray | Format-Table -AutoSize @{ label = 'SP Database'; Expression = { $_.Key } },@{ label = 'Size (GB)'; Expression = { $_.Value } }; Write-Host "`nTotal Storage = " ("{0:n0} GB" -f ($db / 1GB)) }
}

function Do-GracefulShutdownDistributedCacheServices {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory = $true)]
		[ValidateScript(
			{
				$_ -in (Get-SPServer | ? {($_.ServiceInstances | % TypeName) -contains 'Distributed Cache'} | % Address )
			}
		)]
		[string]$DistributedCacheHostName = $env:COMPUTERNAME
	)	
	SPLogMe
	
	$arrow = "{0}{1}{2}" -f [char]9584, [char]9830, [char]9588
	if($DistributedCacheHostName -eq $env:COMPUTERNAME){
		Write-Host ("{0}:`n{1}`t" -f $DistributedCacheHostName, $arrow)
		try {
			Use-CacheCluster
			Get-AFCacheClusterHealth			
			$startTime = Get-Date
			$currentTime = $startTime
			$elapsedTime = $currentTime - $startTime
			$timeOut = 900			
			Write-Host "`tShutting down distributed cache host."
			$hostInfo = Stop-CacheHost -Graceful -CachePort 22233 -HostName $dcServer
			while ($elapsedTime.TotalSeconds -le $timeOut -and $hostInfo.Status -ne 'Down')
			{
				Write-Host "`tHost Status : [$($hostInfo.Status)]"
				Start-Sleep (5)
				$currentTime = Get-Date
				$elapsedTime = $currentTime - $startTime
				#Get-AFCacheClusterHealth
				$hostInfo = Get-CacheHost -HostName $dcServer -CachePort 22233
			}
			Write-Host "`tStopping distributed cache host was successful. Updating Service status in SharePoint."
			Stop-SPDistributedCacheServiceInstance -Graceful
			Write-Host "`tTo start service, please use Central Administration site."
		}
		catch [System.Exception] {
			Write-Host "`tUnable to stop cache host within 15 minutes."
		}			
	}
	else {
		$result = (Invoke-Command -ComputerName $DistributedCacheHostName -ScriptBlock { 
			param($dcServer)		
			Add-PSSnapin Microsoft.SharePoint.PowerShell		
			try {
				Use-CacheCluster
				Get-AFCacheClusterHealth				
				$startTime = Get-Date
				$currentTime = $startTime
				$elapsedTime = $currentTime - $startTime
				$timeOut = 900				
				Write-Host "Shutting down distributed cache host."
				$hostInfo = Stop-CacheHost -Graceful -CachePort 22233 -HostName $dcServer
				while ($elapsedTime.TotalSeconds -le $timeOut -and $hostInfo.Status -ne 'Down')
				{
					Write-Host "Host Status : [$($hostInfo.Status)]"
					Start-Sleep (5)
					$currentTime = Get-Date
					$elapsedTime = $currentTime - $startTime
					#Get-AFCacheClusterHealth
					$hostInfo = Get-CacheHost -HostName $dcServer -CachePort 22233
				}
				Write-Host "Stopping distributed cache host was successful. Updating Service status in SharePoint."
				Stop-SPDistributedCacheServiceInstance -Graceful
				Write-Host "To start service, please use Central Administration site."
			}
			catch [System.Exception] {
				Write-Host "Unable to stop cache host within 15 minutes."
			}	
		} -ArgumentList $DistributedCacheHostName )		
		Write-Host ("{0}:`n{1}`t{2}" -f $DistributedCacheHostName, $arrow, $result)
	}
}

function Rename-SPSite {
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true,HelpMessage = "Enter the old site collection URL")]
		[ValidateNotNullOrEmpty()]
		[string]$OldUrl,
		[Parameter(Mandatory = $true,HelpMessage = "Enter the new site collection URL")]
		[ValidateNotNullOrEmpty()]
		[string]$NewUrl
	)
	SPLogMe
	
	$site = Get-SPsite $OldUrl
	$site.Rename($NewUrl)
	Write-Host (Get-SPsite $NewUrl)
}

function Release-SPFileLock {
	[CmdletBinding()]
	param(
		[Parameter(Position = 0,Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[System.String]$WebUrl,
		[Parameter(Position = 1,Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[System.String]$FileURL
	)
	SPLogMe

	$web = Get-SPWeb $WebUrl
	$file = $web.GetFile($FileURL)

	if ($file.LockId -ne $null)	{
		$userId = $file.LockedByUser.Id
		Write-Host "The file has locked by:" $file.LockedByUser.LoginName ",Lock Expires on: " $file.LockExpires
		#impersonation to release the lock
		$user = $web.allusers.GetById($userId)
		$impersonateSite = New-Object Microsoft.SharePoint.SPSite ($web.Url,$user.UserToken);
		$impersonateWeb = $impersonateSite.OpenWeb();
		$impersonateItem = $impersonateWeb.GetFile($FileURL);
		$impersonateItem.ReleaseLock($impersonateItem.LockId)
		Write-Host "The file lock has been released!"
		$file
	}
	else {
		$file
		Write-Host "The file is not locked " $file.Name
	}
	$web.Dispose()
}

function RestartAll-SPTimerJobServices {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory = $false)]
		[PSCredential] $Credential,
		[switch] $NodeByNode
	)
	SPLogMe
	
	$scriptBlock = { 
		Restart-Service -Name SPTimerV4 
	}
	
	switch ($NodeByNode.IsPresent) {
		$false {
			if($Credential -eq $null) {
				return (DoParallel-OnSPServers -Command $scriptBlock)
			}
			else {
				return (DoParallel-OnSPServers -Credential $Credential -Command $scriptBlock)
			}
		}
		$true {			
			[array]$servers = Get-SPServers
			[array]$result = @()
			$servers | % {
				if($Credential -ne $null) {
					$res = (Invoke-Command -ComputerName $_ -Credential $Credential -ScriptBlock $scriptBlock)
				}
				else {
					$res = (Invoke-Command -ComputerName $_ -ScriptBlock $scriptBlock)
				}
				$result += $res
			}
			return $result
		}
	}
}

function RecycleAll-SPWebApplicatonPools {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory = $false)]
		[PSCredential] $Credential,
		[switch] $NodeByNode
	)
	SPLogMe
	
	$scriptBlock = {  
		Import-Module Easy-Peasy
		Add-PSSnapinIfNotYetAdded Microsoft.SharePoint.PowerShell
		$pUrls = Get-SPWebApplication | % Url
		Recycle-AppPoolsByURLorName -WithoutInteraction -Url $pUrls
	}
	
	switch ($NodeByNode.IsPresent) {
		$false {
			if($Credential -eq $null) {
				return (DoParallel-OnSPServers -Command $scriptBlock)
			}
			else {
				return (DoParallel-OnSPServers -Credential $Credential -Command $scriptBlock)
			}
		}
		$true {			
			[array]$servers = Get-SPServers
			[array]$result = @()
			$servers | % {
				if($Credential -ne $null) {
					$res = (Invoke-Command -ComputerName $_ -Credential $Credential -ScriptBlock $scriptBlock)
				}
				else {
					$res = (Invoke-Command -ComputerName $_ -ScriptBlock $scriptBlock)
				}
				$result += $res
			}
			return $result
		}
	}
}

function Get-SPDatabaseDiskSize {
	SPLogMe
	
	return (Get-SPDatabase | Sort-Object DiskSizeRequired -desc | Format-Table Name,@{ Label = "Size in MB"; Expression = { $_.DiskSizeRequired / 1024 / 1024 } })
}

function Set-SPUsageLogRetentionDay {
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true,HelpMessage = "Enter the log retention day")]
		[ValidateNotNullOrEmpty()]
		[string]$RetentionDay
	)
	SPLogMe
	
	Get-SPUsageDefinition | ? { $_.Enabled } | Set-SPUsageDefinition -DaysRetained $RetentionDay
	return Get-SPUsageDefinition
}

function Set-SPUserAsSiteCollectionAdminOnWebApplication {
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true,HelpMessage = 'username in format DOMAIN\username')]
		[ValidateNotNullOrEmpty()]
		[string]$Username = "",
		[Parameter(Mandatory = $true,HelpMessage = 'url for web application e.g. http://collab')]
		[ValidateNotNullOrEmpty()]
		[string]$WebApplicationUrl = ""
	)
	SPLogMe

	Write-Host "Setting up user $Username as site collection admin on all sitecollections in Web Application $WebApplicationUrl"
	$webApplication = Get-SPWebApplication $WebApplicationUrl
	if ($webApplication -ne $null) {
		foreach ($siteCollection in $webApplication.Sites) {
			Write-Host "Setting up user $Username as site collection admin for $siteCollection"
			$userToBeMadeSiteCollectionAdmin = $siteCollection.RootWeb.EnsureUser($Username)
			if ($userToBeMadeSiteCollectionAdmin.IsSiteAdmin -ne $true) {
				$userToBeMadeSiteCollectionAdmin.IsSiteAdmin = $true
				$userToBeMadeSiteCollectionAdmin.Update()
				Write-Host "User is now site collection admin for $siteCollection"
			}
			else {
				Write-Host "User is already site collection admin for $siteCollection"
			}
			Write-Host "Current Site Collection Admins for site: " $siteCollection.Url " " $siteCollection.RootWeb.SiteAdministrators
		}
	}
	else {
		Write-Host "Could not find Web Application $WebApplicationUrl" -foreground Red
	}	
}

function Update-SPProfilePictureThumbnails {
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true,HelpMessage = 'Enter the URL for my site host location')]
		[ValidateNotNullOrEmpty()]
		[string]$MySiteURL = ""
	)
	SPLogMe
	
	Update-SPProfilePhotoStore -CreateThumbnailsForImportedPhotos $true -MySiteHostLocation $MySiteURL
}

function Add-SPCodeDomAuthorizedType {
	<#
    .Synopsis
       Adds the necessary authorizedType elements to all web.config files for all non-central admin web applications
    .DESCRIPTION
       Adds the necessary authorizedType elements to all web.config files for all non-central admin web applications
     .EXAMPLE
       Add-SPCodeDomAuthorizedType
    #>
	[CmdletBinding()]
	param ()
	SPLogMe

	begin {
		$farmMajorVersion = (Get-SPFarm -Verbose:$false).buildversion.major
		$contentService = [Microsoft.SharePoint.Administration.SPWebService]::ContentService
		$typeNames = @("CodeBinaryOperatorExpression","CodePrimitiveExpression","CodeMethodInvokeExpression","CodeMethodReferenceExpression","CodeFieldReferenceExpression","CodeThisReferenceExpression","CodePropertyReferenceExpression")
	}

	process {
		if (@($contentService.WebConfigModifications | Where-Object { $_.Name -eq "NetFrameworkAuthorizedTypeUpdate" }).Count -gt 0)
		{
			Write-Host "Existing NetFrameworkAuthorizedTypeUpdate entries found, this script only need to be run once per farm."
			return
		}

		if ($farmMajorVersion -le 14) # 2010, 2007
		{
			foreach ($typeName in $typeNames) {
				# System, Version=2.0.0.0
				$netFrameworkConfig = New-Object Microsoft.SharePoint.Administration.SPWebConfigModification
				$netFrameworkConfig.Path = "configuration/System.Workflow.ComponentModel.WorkflowCompiler/authorizedTypes"
				$netFrameworkConfig.Name = "authorizedType[@Assembly='System, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089'][@Namespace='System.CodeDom'][@TypeName='{0}'][@Authorized='True']" -f $typeName
				$netFrameworkConfig.Owner = "NetFrameworkAuthorizedTypeUpdate"
				$netFrameworkConfig.Sequence = 0
				$netFrameworkConfig.Type = [Microsoft.SharePoint.Administration.SPWebConfigModification+SPWebConfigModificationType]::EnsureChildNode
				$netFrameworkConfig.Value = '<authorizedType Assembly="System, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089" Namespace="System.CodeDom" TypeName="{0}" Authorized="True"/>' -f $typeName
				$contentService.WebConfigModifications.Add($netFrameworkConfig);
			}
		}
		else # 2013+
		{
			foreach ($typeName in $typeNames) {
				# System, Version=4.0.0.0
				$netFrameworkConfig = New-Object Microsoft.SharePoint.Administration.SPWebConfigModification
				$netFrameworkConfig.Path = "configuration/System.Workflow.ComponentModel.WorkflowCompiler/authorizedTypes/targetFx"
				$netFrameworkConfig.Name = "authorizedType[@Assembly='System, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089'][@Namespace='System.CodeDom'][@TypeName='{0}'][@Authorized='True']" -f $typeName
				$netFrameworkConfig.Owner = "NetFrameworkAuthorizedTypeUpdate"
				$netFrameworkConfig.Sequence = 0
				$netFrameworkConfig.Type = [Microsoft.SharePoint.Administration.SPWebConfigModification+SPWebConfigModificationType]::EnsureChildNode
				$netFrameworkConfig.Value = '<authorizedType Assembly="System, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089" Namespace="System.CodeDom" TypeName="{0}" Authorized="True"/>' -f $typeName
				$contentService.WebConfigModifications.Add($netFrameworkConfig);
			}
		}

		Write-Host "Updating web.configs"
		$contentService.Update()
		$contentService.ApplyWebConfigModifications();
	}
	end {
	}
}

function Remove-SPCodeDomAuthorizedType {
	<#
    .Synopsis
       Removes any web configuration entires owned by "NetFrameworkAuthorizedTypeUpdate" 
    .DESCRIPTION
       Removes any web configuration entires owned by "NetFrameworkAuthorizedTypeUpdate"
    .EXAMPLE
        Remove-SPCodeDomAuthorizedType
    #>
	[CmdletBinding()]
	param()
	SPLogMe
	
	begin {
		$contentService = [Microsoft.SharePoint.Administration.SPWebService]::ContentService
	}
	process {
		$webConfigModifications = @($contentService.WebConfigModifications | Where-Object { $_.Owner -eq "NetFrameworkAuthorizedTypeUpdate" })
		foreach ($webConfigModification in $webConfigModifications)
		{
			Write-Host "Found instance owned by NetFrameworkAuthorizedTypeUpdate"
			$contentService.WebConfigModifications.remove($webConfigModification) | Out-Null
		}
		if ($webConfigModifications.Count -gt 0)
		{
			$contentService.Update()
			$contentService.ApplyWebConfigModifications()
		}
	}
	end {
	}
}

function Audit-SPUserProfile {
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true,HelpMessage = 'Enter the URL for my site host location')]
		[ValidateNotNullOrEmpty()]
		[string]$MySiteURL = "",
		[Parameter(Mandatory = $true,HelpMessage = 'Enter an account to audit')]
		[ValidateNotNullOrEmpty()]
		[string]$Account = ""
	)
	SPLogMe
	
	$site = Get-SPsite $MySiteURL
	$context = Get-SPServiceContext $site
	$profileManager = New-Object Microsoft.Office.Server.UserProfiles.UserProfileManager ($context)
	$report = @()	
	if ($profileManager.UserExists($Account)) {
		$userProfile = $profileManager.GetUserProfile($Account)
		$changes = $userProfile.GetChanges()
		foreach ($change in $changes) {
			$val = "AccountName: " + $change.AccountName + " ChangeType: " + $change.ChangeType + " EventTime: " + $change.EventTime + " NewValue: " + $change.NewValue + " PropertyDescription: " + $change.ProfileProperty.Description
			$report += $val
		}
	}
	else {
		Write-Host "Profile for user" $Account "cannot be found"
	}
	$site.Dispose()
	return $report
}

function Update-SPUserProfileNewsfeedPrivacy {
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true,HelpMessage = 'Enter a URL')]
		[ValidateNotNullOrEmpty()]
		[string]$URL = ""
	)
	SPLogMe
	
	$Site = Get-SPSite $URL
	$ServiceContext = Get-SPServiceContext $Site
	$ProfileManager = New-Object Microsoft.Office.Server.UserProfiles.UserProfileManager $ServiceContext
	$Profiles = $ProfileManager.GetEnumerator()

	$progressCounter = 1
	foreach ($Profile in $Profiles) {
		Write-Progress -Activity "Changing privacy settings for $($Profile.DisplayName)" -PercentComplete ($progressCounter++ / $ProfileManager.Count * 100) -Status "Changing privacy options..."
		
		# 4095 is decimal value of 111111111111, meaning all options are checked. 
		# Binary value describes what options are checked
		# so 000000000001 --> 1 (1 being the value you want to set as the field value
		# would mean only "Participation in communities" would be checked
		$Profile["SPS-PrivacyActivity"].Value = 0
		$Profile["SPS-EmailOptin"].Value = 126
		$Profile.Commit()
	}
}

function Deploy-WSPSolutions {
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true,HelpMessage = 'Enter a hostname')]
		[ValidateNotNullOrEmpty()]
		[string]$Hostname = "",
		[Parameter(Mandatory = $false,HelpMessage = 'Enter a URL to deploy the solution')]
		[ValidateScript(
			{
				if($env:COMPUTERNAME -eq $Hostname) {
					$sites = (Get-SPWebApplication | % Url )					
				}
				else {
					$sites = (Invoke-Command -ComputerName $Hostname -ScriptBlock { return (Get-SPWebApplication | % Url) } )
				}
				foreach($site in $_){
					$site -in $sites
				}
			}
		)]
		[array]$SiteCollectionUrl,
		[Parameter(Mandatory = $true,HelpMessage = 'Enter a valid path which contains WSP packages')]
		[ValidateNotNullOrEmpty()]
		[string]$DestinationUrl,
		[Parameter(Mandatory = $false,HelpMessage = 'Redeploy solutions to already deployed web applications')]
		[switch]$RedeployWebApps
	)
	SPLogMe
	
	try {
		$reDeploySolution = $false
		switch ($RedeployWebApps.IsPresent) {
			$false {
				$reDeploySolution = $false
			}
			$true {
				$reDeploySolution = $true
			}
		}
		
		$result = (Invoke-Command -ComputerName $Hostname -ScriptBlock {			
			Import-Module UberSPShell			
			function Deploy-WSP ($webApp, $wspURL, $redeploy) {
				try {
					$wspFiles = Get-ChildItem $wspURL | ? { !($_.psiscontainer) } | ? { $_.Name -like "*.wsp" } 
					foreach ($file in $wspFiles) {
						$solution = Get-SPSolution -Identity $file.Name
						$urls = $solution.DeployedWebApplications.Url
						if ($solution.Deployed -eq $true) {
							Write-Host -ForegroundColor Green "Uninstalling the solution $($solution.Name) ... "
							if ($solution.ContainsWebApplicationResource) {
								foreach($url in $urls) {
									Write-Host -NoNewline " $url"
									Uninstall-SPSolution -Identity $solution.Name -WebApplication $url -ErrorAction Stop -Confirm:$false
								}
							}
							else {
								Uninstall-SPSolution -Identity $solution.Name -Confirm:$false
							}
							Write-Host -NoNewline "Waiting to finish "
							while ((Get-SPSolution -Identity $solution.Name).JobExists) {
								Write-Host -NoNewline .
								Start-Sleep -Seconds 1
							}
							Remove-SPSolution -Identity $solution.Name -Force -Confirm:$false
							Write-Host -NoNewline " done!"
						}
						Write-Host -ForegroundColor Green "Installing your solution to ..."
						Add-SPSolution -LiteralPath $file.FullName
						if($redeploy -eq $true) {
							foreach($url in $urls) {
								Write-Host -NoNewline " $url"
								Install-SPSolution -WebApplication $url -GACDeployment -FullTrustBinDeployment -Identity $file.Name -CompatibilityLevel All -ErrorAction Stop							
							}
						}
						if($webApp.Count -gt 0) {
							foreach($url in $webApp) {
								Write-Host -NoNewline " $url"
								Install-SPSolution -WebApplication $url -GACDeployment -FullTrustBinDeployment -Identity $file.Name -CompatibilityLevel All -ErrorAction Stop								
							}
						}
						else {
							Write-Host -NoNewline " Global"
							Install-SPSolution -GACDeployment -FullTrustBinDeployment -Identity $file.Name -CompatibilityLevel All -ErrorAction Stop
						}						
					}
				}
				catch {
					$ErrorMessage = $_.Exception.Message
					throw $ErrorMessage
				}
			}
			
			$backupPath = Backup-WSPSolutions
			try {
				DoParallel-OnSPServers -Command { Stop-Service -Name "SPTimerV4" }
				Deploy-WSP ($args[1], $args[0], $args[2])
				DoParallel-OnSPServers -Command { Start-Service -Name "SPTimerV4" }
			}
			catch {
				Write-Host -ForegroundColor Red "Failed to deploy, rollback solution..."
				try {
					DoParallel-OnSPServers -Command { Stop-Service -Name "SPTimerV4" }
					Deploy-WSP ($args[1], $backupPath, $args[2])
					DoParallel-OnSPServers -Command { Start-Service -Name "SPTimerV4" }
				}
				catch {
					Write-Host -ForegroundColor Red "FATAL ERROR! Unable to deploy the backed up packages!"
					throw "FATAL ERROR! Unable to deploy the backed up packages!"
				}
				throw "Failed to deploy, rollback solution..."
			}
		} -Args $DestinationUrl, $SiteCollectionUrl, $reDeploySolution )		
		Write-Host $result
	}
	catch {
		$ErrorMessage = $_.Exception.Message
		throw $ErrorMessage
	}
}

function Copy-SPPackages {
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true,HelpMessage = "Enter source server's hostname")]
		[ValidateNotNullOrEmpty()]
		[string]$SourceServer = "",
		[Parameter(Mandatory = $true,HelpMessage = "Enter destination server's hostname")]
		[ValidateNotNullOrEmpty()]
		[string]$DestinationServer = "",
		[Parameter(Mandatory = $true,HelpMessage = "Enter a valid path on source server or URL to obtain necessary files")]
		[ValidateNotNullOrEmpty()]
		[string]$SourcePath,
		[Parameter(Mandatory = $true,HelpMessage = "Enter a valid path on destination server or URL to deploy necessary files")]
		[ValidateNotNullOrEmpty()]
		[string]$DestinationPath,
		[Parameter(Mandatory = $false,HelpMessage = "Give a valid credential on source server")]
		[ValidateNotNullOrEmpty()]
		[PSCredential]$SourceServerCredential,
		[Parameter(Mandatory = $false,HelpMessage = "Give a valid username on the service")]
		[ValidateNotNullOrEmpty()]
		[string]$SourceUrlUsername,
		[Parameter(Mandatory = $false,HelpMessage = "Give a valid password on the service")]
		[ValidateNotNullOrEmpty()]
		[string]$SourceUrlPassword,
		[Parameter(Mandatory = $false,HelpMessage = "Give a valid api key on the service")]
		[ValidateNotNullOrEmpty()]
		[string]$SourceApiKey,
		[Parameter(Mandatory = $false,HelpMessage = "Give a valid credential on destination server")]
		[ValidateNotNullOrEmpty()]
		[PSCredential]$DestinationServerCredential
	)
	SPLogMe
	
	$sourceSession = $null
	$destinationSession = $null
	if($SourceServer -like "http*"){
		continue
	}
	else {
		if($SourceServerCredential -ne $null) {
			$sourceSession = New-PSSession -ComputerName $SourceServer -Credential $SourceServerCredential
		}
		else {
			$sourceSession = New-PSSession -ComputerName $SourceServer
		}
	}
	if($DestinationServerCredential -ne $null) {
		$destinationSession = New-PSSession -ComputerName $DestinationServer -Credential $DestinationServerCredential
	}
	else {
		$destinationSession = New-PSSession -ComputerName $DestinationServer
	}	
	
	$now = (Get-Date).ToString("yyyyMMdd_HHmm")
	$dropboxPath = Join-Path -Path $env:CustomTempPath -ChildPath "dropbox_$now"
	New-Item -ItemType Directory -Path $dropboxPath
	
	if($sourceSession -ne $null) {
		Copy-Item -FromSession $sourceSession -Path $SourcePath -Destination $dropboxPath -Recurse
	}
	else {
		#function Get-RedirectedUrl {
		#	Param (
		#		[Parameter(Mandatory=$true)]
		#		[String]$URL
		#	)
		#	$request = [System.Net.WebRequest]::Create($url)
		#	$request.AllowAutoRedirect=$false
		#	$response=$request.GetResponse()
		#	If ($response.StatusCode -eq "Found")
		#	{
		#		$response.GetResponseHeader("Location")
		#	}
		#}
		#$remoteFileName = ([System.IO.Path]::GetFileName((Get-RedirectedUrl $webPath))).replace("%20"," ")
		
		$webPath = $SourceServer + "/artifactory/api/storage/$SourcePath"		
		$remoteFileName = $SourcePath.Split("/")[-1]		
		$tmpFile = Join-Path -Path $dropboxPath -ChildPath $remoteFileName
		
		[SecureString]$Pass = $SourceUrlPassword | ConvertTo-SecureString -AsPlainText -Force
		$Pass.MakeReadOnly()
		$creds = New-Object System.Management.Automation.PSCredential($SourceUrlUsername,$Pass)
		$Pass = $null
		
		if(!(String-IsNullOrEmpty $SourceApiKey)) {
			$params = @{
				Uri = $webPath
				Method = "Get"
				Headers = @{ "X-JFrog-Art-Api" = $SourceApiKey }
				Verbose = $false
			}
			if ($creds -ne $null) {
				$params.Credential = $creds
			} 
			else {
				$params.UseDefaultCredentials = $true
			}
			Invoke-RestMethod @params -OutFile $tmpFile
		}
		else {
			$pair = "$($SourceUrlUsername):$($SourceUrlPassword)"
			$bytes = [System.Text.Encoding]::ASCII.GetBytes($pair)
			$base64 = [System.Convert]::ToBase64String($bytes)
			$basicAuthValue = "Basic $base64"
			$headers = @{ 			
				"Authorization" = $basicAuthValue 
				"Accept" = "application/json"			
			}
			Invoke-WebRequest -Credential $creds -Headers $headers -Method Get -Uri $webPath -OutFile $tmpFile
		}
	}
	Write-Host "Step 1/2 is completed."
	
	Copy-Item -ToSession $destinationSession -Path $dropboxPath -Destination $DestinationPath -Recurse
	Write-Host "Step 2/2 is completed. Done!"
	Remove-Item -Path $dropboxPath -Recurse
}

function Add-SPAllowedInlineDownloadedMimeTypes {
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $false, HelpMessage = 'Enter a hostname')]
		[ValidateNotNullOrEmpty()]
		[string]$Hostname = $env:COMPUTERNAME,
		[Parameter(Mandatory = $true, HelpMessage = 'Enter a URL to deploy the solution')]
		[ValidateScript(
			{
				if($env:COMPUTERNAME -eq $Hostname) {
					$sites = (Get-SPWebApplication | % Url )					
				}
				else {
					$sites = (Invoke-Command -ComputerName $Hostname -ScriptBlock { return (Get-SPWebApplication | % Url) } )
				}
				foreach($site in $_){
					$site -in $sites
				}
			}
		)]
		[array]$Url,
		[Parameter(Mandatory = $true,HelpMessage = 'Enter a valid path which contains WSP packages')]
		[ValidateNotNullOrEmpty()]
		[string]$MimeType
	)
	SPLogMe	
	
	if($env:COMPUTERNAME -eq $Hostname) {
		foreach($site in $Url){
			$app = Get-SPWebApplication -Identity $site
			if ($app.AllowedInlineDownloadedMimeTypes.Contains($MimeType)) {
				Write-Host "Mime-type" $MimeType "already added to AllowedInlineDownloadedMimeTypes."
			} 
			else {
				$app.AllowedInlineDownloadedMimeTypes.Add($MimeType)
				Write-Host "Mime-type" $MimeType "added to AllowedInlineDownloadedMimeTypes."
			}
			$app.Update()
		}	
	}
	else {
		$result = (Invoke-Command -ComputerName $Hostname -ScriptBlock {
			Import-Module UberSPShell
			SPLogMe
		
			$Url = $args[0]
			$MimeType = $args[1]
			foreach($site in $Url){
				$app = Get-SPWebApplication -Identity $site
				if ($app.AllowedInlineDownloadedMimeTypes.Contains($MimeType)) {
					Write-Host "Mime-type" $MimeType "already added to AllowedInlineDownloadedMimeTypes."
				} 
				else {
					$app.AllowedInlineDownloadedMimeTypes.Add($MimeType)
					Write-Host "Mime-type" $MimeType "added to AllowedInlineDownloadedMimeTypes."
				}
				$app.Update()
			}	
		} -Args $Url, $MimeType )
		Write-Host $result
	}
}

function Flush-SPBlobCache {
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $false, HelpMessage = 'Enter a hostname')]
		[ValidateNotNullOrEmpty()]
		[string]$Hostname = $env:COMPUTERNAME,
		[Parameter(Mandatory = $true, HelpMessage = 'Enter a URL to deploy the solution')]
		[ValidateScript(
			{
				if($env:COMPUTERNAME -eq $Hostname) {
					$sites = (Get-SPWebApplication | % Url )					
				}
				else {
					$sites = (Invoke-Command -ComputerName $Hostname -ScriptBlock { return (Get-SPWebApplication | % Url) } )
				}
				foreach($site in $_){
					$site -in $sites
				}
			}
		)]
		[array]$Url
	)
	SPLogMe	
	
	foreach($site in $Url){
		Write-Host "Flushing BLOB Cache for:" $site
		$webApplication = Get-SPWebApplication -Identity $site
		[Microsoft.SharePoint.Publishing.PublishingCache]::FlushBlobCache($webApplication)
		Write-Host "Flushed the BLOB cache for:" $webApplication
	}	
}

function AddOrUpdate-SPFarmProperty {
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$Key,
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$Value
	)
	SPLogMe	
	
	Write-Host "Updating farm property " $Key " -> " $Value
	$farm = Get-SPFarm
	$properties = $farm.Properties
	$properties[$Key] = $Value
	$farm.Update()
}

function Get-SPFarmProperty {
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$Key
	)
	SPLogMe	
	
	$farm = Get-SPFarm
	$value = $farm.Properties[$Key]
	if ($value -eq $null) {
		throw $Key + " is not set in Farm Property Bag."
	}
	return $value
}

function Get-SPUserProfiles {
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$MySiteUrl
	)
	SPLogMe
	
	#Get site objects and connect to User Profile Manager service 
	$site = Get-SPsite $MySiteUrl
	$context = Get-SPServiceContext $site
	$profileManager = New-Object Microsoft.Office.Server.UserProfiles.UserProfileManager ($context)
	$enumer = $profileManager.GetEnumerator()
	return $enumer
}

function Get-SPSearchServiceManagedProperties {
	[CmdletBinding()]
	param ()
	SPLogMe
	
	$svcapp = Get-SPServiceApplication | Where-Object { $_.Name -like "Search Service Application*" }
	return Get-SPEnterpriseSearchMetadataManagedProperty -SearchApplication $svcapp | Format-Table -Property Name,ManagedType,Searchable,FullTextQueriable,Queryable,Retrievable,Refinable,Sortable,SafeForAnonymous,EnabledForScoping,EqualityMatchOnly -AutoSize
}

function Disable-SPFeatureInAllSites {
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$WebAppUrl,
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$FeatureName
	)
	SPLogMe
	
	try {
		$feature = Get-SPFeature $FeatureName
		$webApp = Get-SPWebApplication -Identity $WebAppUrl
		$webApp | Get-SPsite -Limit all | ForEach-Object {
			try {
				if ($_.Features[$feature.Id]) {
					Write-Host "Found feature " $FeatureName " in site collection " $_.Url
					Disable-SPFeature $feature -Url $_.Url -Force -Confirm:$false
					Write-Host "Disabled feature " $FeatureName " in site collection " $_.Url
				}
				$_.Dispose()
			} 
			catch {
				Write-Host "Exception in enumerating sites or disabling feature " $FeatureName "." $_.Message
			}
		}
	} 
	catch {
		Write-Host "Exception in retrieving Feature or Web Application:" $_.Message
		break
	}
}

function Enable-SPFeatureInAllSites {
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$WebAppUrl,
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$FeatureName
	)
	SPLogMe
	
	try {
		$feature = Get-SPFeature $FeatureName
		$webApp = Get-SPWebApplication -Identity $WebAppUrl
		$webApp | Get-SPsite -Limit all | ForEach-Object {
			try {
				if ($_.Features[$feature.Id] -eq $null) {
					Write-Host "Feature " $FeatureName " in site collection " $_.Url " is not enabled"
					Disable-SPFeature $feature -Url $_.Url -Force -Confirm:$false
					Write-Host "Enabled feature " $FeatureName " in site collection " $_.Url
				}
				$_.Dispose()
			} 
			catch {
				Write-Host "Exception in enumerating sites or enabling feature " $FeatureName "." $_.Message
			}
		}
	} 
	catch {
		Write-Host "Exception in retrieving Feature or Web Application:" $_.Message
		break
	}
}

function Enable-SPFeatureInAllWebs {
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$WebAppUrl,
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$FeatureName
	)
	SPLogMe
	
	try {
		$feature = Get-SPFeature $FeatureName
		$webApp = Get-SPWebApplication -Identity $WebAppUrl
		$webApp | Get-SPsite -Limit all | ForEach-Object {
			try {
				$_ | Get-SPWeb -Limit all | ForEach-Object {
					if ($_.Features[$feature.Id] -eq $null) {
						Write-Host "Feature " $FeatureName " in site/subsite " $_.Url " is not enabled"
						Enable-SPFeature $feature -Url $_.Url -Force -Confirm:$false
						Write-Host "Enabled feature " $FeatureName " in site/subsite " $_.Url
					}
					$_.Dispose()
				}
				$_.Dispose()
			} 
			catch {
				Write-Host "Exception in enumerating sites/webs or enabling feature " $FeatureName "." $_.Message
			}
		}
	} 
	catch {
		Write-Host "Exception in retrieving Feature or Web Application:" $_.Message
		break
	}
}

function Disable-SPFeatureInAllWebs {
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$WebAppUrl,
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$FeatureName
	)
	SPLogMe
	
	try {
		$feature = Get-SPFeature $FeatureName
		$webApp = Get-SPWebApplication -Identity $WebAppUrl
		$webApp | Get-SPsite -Limit all | ForEach-Object {
			try {
				$_ | Get-SPWeb -Limit all | ForEach-Object {
					if ($_.Features[$feature.Id]) {
						Write-Host "Found feature " $FeatureName " in site/subsite " $_.Url
						Disable-SPFeature $feature -Url $_.Url -Force -Confirm:$false
						Write-Host "Disabled feature " $FeatureName " in site/subsite " $_.Url
					}
					$_.Dispose()
				}
				$_.Dispose()
			} 
			catch {
				Write-Host "Exception in enumerating sites/webs or disabling feature " $FeatureName "." $_.Message
			}
		}
	} 
	catch {
		Write-Host "Exception in retrieving Feature or Web Application:" $_.Message
		break
	}
}

function Configure-SPSocialFeedCache {
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[int]$TTLHours = 2166, # 90 days
		[Parameter(Mandatory = $false)]
		[ValidateNotNullOrEmpty()]
		[int]$ObjectCountLimit = 5000,
		[Parameter(Mandatory = $false)]
		[ValidateNotNullOrEmpty()]
		[int]$RoomForGrowth = 200
	)
	SPLogMe
	
	$upa = Get-SPServiceApplication | Where-Object { $_.TypeName -eq "User Profile Service Application" }
	$upa.FeedCacheTTLHours = $TTLHours
	$upa.FeedCacheLastModifiedTimeTtlDeltaHours = $TTLHours
	$upa.FeedCacheObjectCountLimit = $ObjectCountLimit
	$upa.FeedCacheRoomForGrowth = $RoomForGrowth
	$upa.Update()
}

function Get-SPEventReceivers {
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$WebAppUrl
	)
	SPLogMe
	
	$webApp = Get-SPWebApplication -Identity $WebAppUrl
	$sites = $webApp | Get-SPsite -Limit all
	
	$result = @()
	foreach ($site in $sites) {
		$web = $site.RootWeb		
		foreach($list in $web.Lists) {
			$list.EventReceivers | % { $result += [PSCustomObject]@{
					List = $list.Title
					EventReceieverType = $_.Type
					EventReceieverClass = $_.Class 					
				} 
			}
		}
	}
	return $result
}

function Export-SPRootCertificate {
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$Path = (Join-Path -Path $env:CertificatesFolder -ChildPath "SPRootAuthority.cer")
	)
	SPLogMe
	
	return ((Get-SPCertificateAuthority).RootCertificate.Export("Cert") | Set-Content $Path -Encoding byte)
}

function Refresh-SPMissingRootCertificate {
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$Guid,
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$Path = (Join-Path -Path $env:CertificatesFolder -ChildPath "SPRootAuthority.cer")
	)
	SPLogMe
	
	$farm = Get-SPFarm
	$certObject = $farm.GetObject($Guid)

	if ($certObject -and $certObject.SecureStringValue -eq $null) {
		Write-Host "SecureStringValue is null, please run this script as an administrator." -ForegroundColor Red
		return
	}

	if ($certObject) {
		# convert from secure store to plaintext string
		$bstr = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($certObject.SecureStringValue)) 
		# convert the plaintext string to a byte[]
		$exportedCertificate = [Convert]::FromBase64String($bstr) 
		# create the cert from the byte array
		$certificate = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 ($exportedCertificate,$null,[System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::MachineKeySet) 
		# if the certifiate exists, save it locally
		if ($certificate) {
			# save the cert locally
			$certificate.Export("Cert") | Set-Content $Path -Encoding byte -Force 
			# make sure the save worked
			if (Test-Path $Path -PathType Leaf) {
				# get the cert in a format we can push back into SharePoint
				$certificate = Get-PfxCertificate -FilePath $Path
				$exportedCert = $certificate.Export("Pkcs12")
				$secureStrCert = ConvertTo-SecureString -String $([Convert]::ToBase64String($exportedCert)) -Force -AsPlainText
				# get the SharePoint object and updated it with the serialized cert
				$certObjectId = [guid]::Parse($Guid)
				$certObject = (Get-SPFarm).GetObject($certObjectId)
				$certObject.UpdateSecureStringValue($secureStrCert)
			}
			else {
				Write-Host "File not found: $Path" -ForegroundColor Red
			}
		}
	}
	else {
		Write-Host "Could not get the EncryptedString object from the config database" -ForegroundColor Red
	}
}

function Get-SPMissingWebPartDetails {
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$DatabaseServer,
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$Path
	)
	SPLogMe
		
	function RunSQLQuery ($SqlServer, $SqlDatabase, $SqlQuery) {
		$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
		$SqlConnection.ConnectionString = "Server =" + $SqlServer + "; Database =" + $SqlDatabase + "; Integrated Security = True"
		$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
		$SqlCmd.CommandText = $SqlQuery
		$SqlCmd.Connection = $SqlConnection
		$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
		$SqlAdapter.SelectCommand = $SqlCmd
		$DataSet = New-Object System.Data.DataSet
		$SqlAdapter.Fill($DataSet)
		$SqlConnection.Close()
		$DataSet.Tables[0]
	}
	
	function GetWebPartDetails ($DatabaseServer, $wpid, $DBname) {
		#Define SQL Query and set in Variable
		$Query = "SELECT * from AllDocs with (nolock) inner join AllWebParts with (nolock) on AllDocs.Id = AllWebParts.tp_PageUrlID where AllWebParts.tp_WebPartTypeID = '" + $wpid + "'"

		#Runing SQL Query to get information about Assembly (looking in EventReceiver Table) and store it in a Table
		$QueryReturn = @(RunSQLQuery -SqlServer $DatabaseServer -SqlDatabase $DBname -SqlQuery $Query | Select-Object Id, SiteId, DirName, LeafName, WebId, ListId, tp_ZoneID, tp_DisplayName)

		#Actions for each element in the table returned
		foreach ($event in $QueryReturn) {
			if ($event.id -ne $null) {
				#Get Site URL
				$site = Get-SPsite -Limit all | Where-Object { $_.Id -eq $event.SiteId }
				#Log information to Host
				Write-Host $wpid -NoNewline -ForegroundColor yellow
				Write-Host ";" -NoNewline
				Write-Host $site.Url -NoNewline -ForegroundColor green
				Write-Host "/" -NoNewline -ForegroundColor green
				Write-Host $event.LeafName -ForegroundColor green -NoNewline
				Write-Host ";" -NoNewline
				Write-Host $site.Url -NoNewline -ForegroundColor gray
				Write-Host "/" -NoNewline -ForegroundColor gray
				Write-Host $event.DirName -ForegroundColor gray -NoNewline
				Write-Host "/" -NoNewline -ForegroundColor gray
				Write-Host $event.LeafName -ForegroundColor gray -NoNewline
				Write-Host "?contents=1" -ForegroundColor gray -NoNewline
				Write-Host ";" -NoNewline
				Write-Host $event.tp_ZoneID -ForegroundColor cyan
			}
		}
	}
	
	$input = @(Get-Content $Path)
	
	#Log the CVS Column Title Line
	Write-Host "WebPartID; PageUrl; MaintenanceUrl; WpZoneID" -ForegroundColor Red

	foreach ($event in $input) {
		$wpid = $event.split(";")[0]
		$DBname = $event.split(";")[1]
		GetWebPartDetails $DatabaseServer $wpid $dbname
	}
}

function Get-SPMissingAssemblyDetails {
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$DatabaseServer,
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$Path
	)
	SPLogMe
		
	function RunSQLQuery ($SqlServer,$SqlDatabase,$SqlQuery) {
		$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
		$SqlConnection.ConnectionString = "Server =" + $SqlServer + "; Database =" + $SqlDatabase + "; Integrated Security = True"
		$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
		$SqlCmd.CommandText = $SqlQuery
		$SqlCmd.Connection = $SqlConnection
		$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
		$SqlAdapter.SelectCommand = $SqlCmd
		$DataSet = New-Object System.Data.DataSet
		$SqlAdapter.Fill($DataSet)
		$SqlConnection.Close()
		$DataSet.Tables[0]
	}
	
	function GetAssemblyDetails ($DatabaseServer, $assembly, $DBname) {
		#Define SQL Query and set in Variable
		$Query = "SELECT * from EventReceivers with (nolock) where Assembly = '" + $assembly + "'"
		#$Query = "SELECT * from EventReceivers where Assembly = 'Microsoft.Office.InfoPath.Server, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c'" 
		#Runing SQL Query to get information about Assembly (looking in EventReceiver Table) and store it in a Table
		$QueryReturn = @(RunSQLQuery -SqlServer $DatabaseServer -SqlDatabase $DBname -SqlQuery $Query | Select-Object Id, Name, SiteId, WebId, HostId, HostType)
		#Actions for each element in the table returned
		foreach ($event in $QueryReturn) {
			#HostID (check http://msdn.microsoft.com/en-us/library/ee394866(v=prot.13).aspx for HostID Type reference)
			if ($event.HostType -eq 0) {
				$site = Get-SPsite -Limit all | Where-Object { $_.Id -eq $event.SiteId }
				#Get the EventReceiver Site Object
				$er = $site.EventReceivers | Where-Object { $_.Id -eq $event.Id }
				Write-Host $assembly -NoNewline -ForegroundColor yellow
				Write-Host ";" -NoNewline
				Write-Host $site.Url -NoNewline -ForegroundColor gray
				Write-Host ";" -NoNewline
				Write-Host $er.Name -ForegroundColor green -NoNewline
				Write-Host ";" -NoNewline
				Write-Host $er.Class -ForegroundColor cyan
				#$er.Delete()
			}
			if ($event.HostType -eq 1) {
				$site = Get-SPsite -Limit all | Where-Object { $_.Id -eq $event.SiteId }
				$web = $site | Get-SPWeb -Limit all | Where-Object { $_.Id -eq $event.WebId }
				#Get the EventReceiver Site Object
				$er = $web.EventReceivers | Where-Object { $_.Id -eq $event.Id }
				$er.Name
				Write-Host $assembly -NoNewline -ForegroundColor yellow
				Write-Host ";" -NoNewline
				Write-Host $web.Url -NoNewline -ForegroundColor gray
				Write-Host ";" -NoNewline
				Write-Host $er.Name -ForegroundColor green -NoNewline
				Write-Host ";" -NoNewline
				Write-Host $er.Class -ForegroundColor cyan
				#$er.Delete()
			}

			if ($event.HostType -eq 2) {
				$site = Get-SPsite -Limit all | Where-Object { $_.Id -eq $event.SiteId }
				$web = $site | Get-SPWeb -Limit all | Where-Object { $_.Id -eq $event.WebId }
				$list = $web.Lists | Where-Object { $_.Id -eq $event.HostId }
				#Get the EventReceiver List Object
				$er = $list.EventReceivers | Where-Object { $_.Id -eq $event.Id }
				Write-Host $assembly -NoNewline -ForegroundColor yellow
				Write-Host ";" -NoNewline
				Write-Host $web.Url -NoNewline -ForegroundColor gray
				Write-Host "/" -NoNewline -ForegroundColor gray
				Write-Host $list.RootFolder -NoNewline -ForegroundColor gray
				Write-Host ";" -NoNewline
				Write-Host $er.Name -ForegroundColor green -NoNewline
				Write-Host ";" -NoNewline
				Write-Host $er.Class -ForegroundColor cyan
				#$er.Delete()
			}
		}
	}
	
	$input = @(Get-Content $Path)
	
	#Log the CVS Column Title Line
	Write-Host "Assembly; Url; EventReceiverName; EventReceiverClass" -ForegroundColor Red

	foreach ($event in $input) {
		$assembly = $event.split(";")[0]
		$DBname = $event.split(";")[1]
		GetAssemblyDetails $DatabaseServer $assembly $dbname
	}
}

function Get-SPMissingSetupFileDetails {
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$DatabaseServer,
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$Path
	)
	SPLogMe
		
	function RunSQLQuery ($SqlServer, $SqlDatabase, $SqlQuery) {
		$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
		$SqlConnection.ConnectionString = "Server =" + $SqlServer + "; Database =" + $SqlDatabase + "; Integrated Security = True"
		$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
		$SqlCmd.CommandText = $SqlQuery
		$SqlCmd.Connection = $SqlConnection
		$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
		$SqlAdapter.SelectCommand = $SqlCmd
		$DataSet = New-Object System.Data.DataSet
		$SqlAdapter.Fill($DataSet)
		$SqlConnection.Close()
		$DataSet.Tables[0]
	}
	
	function GetFileUrl ($DatabaseServer, $filepath, $DBname) {
		#Define SQL Query and set in Variable  
		$Query = "SELECT * from AllDocs with (nolock) where SetupPath = '" + $filepath + "'"
		#Runing SQL Query to get information about the MissingFiles and store it in a Table  
		$QueryReturn = @(Run-SQLQuery -SqlServer $DatabaseServer -SqlDatabase $DBname -SqlQuery $Query | Select-Object Id,SiteId,DirName,LeafName,WebId,ListId)
		foreach ($event in $QueryReturn) {
			if ($event.id -ne $null) {
				$site = Get-SPsite -Limit all | Where-Object { $_.Id -eq $event.SiteId }
				#get the URL of the Web:  
				$web = $site | Get-SPWeb -Limit all | Where-Object { $_.Id -eq $event.WebId }
				#Write the SPWeb URL to host  
				Write-Host $filepath -NoNewline -ForegroundColor yellow
				Write-Host ";" -NoNewline
				Write-Host $web.Url -NoNewline -ForegroundColor green
				#get the URL of the actual file:  
				$file = $web.GetFile([guid]$event.Id)
				#Write the relative URL to host  
				Write-Host "/" -NoNewline -ForegroundColor green
				Write-Host $file.Url -ForegroundColor green
				$file.Delete()
			}
		}
	}
	
	$input = @(Get-Content $Path)
	
	#Log the CVS Column Title Line
	Write-Host "MissingSetupFile; Url" -ForegroundColor Red 
	
	foreach ($event in $input) {
		$filepath = $event.split(";")[0]
		$DBname = $event.split(";")[1]
		GetFileUrl $DatabaseServer $filepath $dbname
	}
}

function Repopulate-SPSiteCollectionFeeds {
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$SiteUrl,
		[Parameter(Mandatory = $false)]
		[ValidateNotNullOrEmpty()]
		[string]$ProxyName = "User Profile Service Application"
	)
	SPLogMe
	
	# Allocate process memory
	Start-SPAssignment -Global
	# Get the UPS Proxy for use with the cache commands
	$proxy = Get-SPServiceApplicationProxy | Where-Object { $_.Name -eq $ProxyName }
	# Initialize the distributed cache repopulation
	Update-SPRepopulateMicroblogLMTCache -ProfileServiceApplicationProxy $proxy
	# Obtain Service Context based on URL
	$siteCollections = Get-SPsite -Identity "$SiteUrl/.*" -Regex -Limit ALL #Get-SPSite $SiteUrl
	foreach ($siteCollection in $siteCollections) {
		$context = Get-SPServiceContext $siteCollection
		# Access the user profiles through the context
		$UserProfileManager = New-Object Microsoft.Office.Server.UserProfiles.UserProfileManager($context)
		$profiles = $UserProfileManager.GetEnumerator()
		# Perform the cache repopulation for each user
		foreach ($profile in $profiles) {
			if ($profile.item("SPS-PersonalSiteCapabilities").Value -eq 14) {
				$AccountName = $profile.item("AccountName").Value
				Write-Host "Updating the Newsfeed for $AccountName"
				Update-SPRepopulateMicroblogFeedCache -ProfileServiceApplicationProxy $proxy -AccountName $AccountName
			}
		}
	}
	# Dispose of process memory
	Stop-SPAssignment -Global
}

function Add-SPPeoplePickerADProvider {
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$WebAppUrl,
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$Domain,
		[Parameter(Mandatory = $false)]
		[ValidateNotNullOrEmpty()]
		[string]$ShortDomain,
		[Parameter(Mandatory = $false)]
		[ValidateNotNullOrEmpty()]
		[string]$LoginName,
		[Parameter(Mandatory = $false)]
		[ValidateNotNullOrEmpty()]
		[SecureString]$Password,
		[Parameter(Mandatory = $false)]
		[ValidateNotNullOrEmpty()]
		[bool]$IsForest
	)
	SPLogMe
	
	$farm = Get-SPFarm
	$farm.Properties["disable-netbios-dc-resolve"] = $true
	$farm.Update()

	$wa = Get-SPWebApplication $WebAppUrl
	
	$temp = New-Object Microsoft.SharePoint.Administration.SPPeoplePickerSearchActiveDirectoryDomain
	$temp.DomainName = $Domain
	$temp.ShortDomainName = $ShortDomain
	$temp.LoginName = $LoginName
	$temp.IsForest = $IsForest
	$temp.setpassword($Password)
	
	$wa.PeoplePickerSettings.SearchActiveDirectoryDomains.Add($temp)
	$wa.PeoplePickerSettings.ShowUserInfoListSuggestionsInClaimsMode = $true
	$wa.PeoplePickerSettings.ActiveDirectoryRestrictIsolatedNameLevel = $false
	$wa.Update()
	return $wa.PeoplePickerSettings.SearchActiveDirectoryDomains
}

function Clear-SPPeoplePickerADProvider {
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$WebAppUrl
	)
	SPLogMe
	
	$farm = Get-SPFarm
	$farm.Properties["disable-netbios-dc-resolve"] = $true
	$farm.Update()

	$wa = Get-SPWebApplication $WebAppUrl
	$wa.PeoplePickerSettings.SearchActiveDirectoryDomains.Clear()	
	$wa.Update()
	
	return $wa.PeoplePickerSettings.SearchActiveDirectoryDomains
}

function Remove-SPPeoplePickerADProvider {
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$WebAppUrl,
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$Domain
	)
	SPLogMe
	
	$farm = Get-SPFarm
	$farm.Properties["disable-netbios-dc-resolve"] = $true
	$farm.Update()

	$wa = Get-SPWebApplication $WebAppUrl
	
	$others = $wa.PeoplePickerSettings.SearchActiveDirectoryDomains | ? { $_.DomainName -ne $Domain }
	
	$wa.PeoplePickerSettings.SearchActiveDirectoryDomains.Clear()
	$others | % { $wa.PeoplePickerSettings.SearchActiveDirectoryDomains.Add($_) }
	$wa.PeoplePickerSettings.ShowUserInfoListSuggestionsInClaimsMode = $true
	$wa.PeoplePickerSettings.ActiveDirectoryRestrictIsolatedNameLevel = $false
	$wa.Update()
	return $wa.PeoplePickerSettings.SearchActiveDirectoryDomains
}

function Get-SPSiteMetrics {
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$Save
	)
	SPLogMe

	Write-Output "Extracting structure information to $Save..."
	"<Metrics>" | Out-File -FilePath $Save -Append:$false
	$spWebApps = Get-SPWebApplication
	$spWAcount = 1
	foreach ($spWebApp in $spWebApps) {
		$percentComplete = [math]::round((($spWAcount * 100) / $spWebApps.Count), 2)
		Write-Progress -Activity "Reading SharePoint farm structure..." -Status "Enumerating Web Applications, $percentComplete% completed..." -Id 0 -PercentComplete $percentComplete -CurrentOperation "Web Application: $($spWebApp.DisplayName) [Url: $($spWebApp.Url)]"
		
		"`t<WebApplication DisplayName='$($spWebApp.DisplayName)' Url='$($spWebApp.Url)' SiteCount='$($spWebApp.Sites.Count)'>" | Out-File -FilePath $Save -Append:$true
		# export $spWebApp.DisplayName
		# export $spWebApp.Url
		# export $spWebApp.Sites.Count, check if bigger than ?
		$spSiteCount = 1
		foreach ($spSite in $spWebApp.Sites) {
			$sitePercentComplete = [math]::round((($spSiteCount * 100) / $spWebApp.Sites.Count), 2)
			Write-Progress -Activity "Enumerating site collections, $sitePercentComplete% completed..." -Id 1 -PercentComplete $sitePercentComplete -CurrentOperation "Site collection: $($spSite.Url)" -ParentId 0
			# export $spSite.Url
			# export $spSite.Usage.Storage, check if bigger than 100GB
			# export $spSite.AllWebs.Count, check if bigger than 250.000	
			"`t`t<SiteCollection Url='$($spSite.Url)' Database='$($spSite.ContentDatabase.Name)' Storage='$($spSite.Usage.Storage)' WebCount='$($spSite.AllWebs.Count)'/>" | Out-File -FilePath $Save -Append:$true
			if ($spSite.Usage.Storage -gt 100GB) {
				Write-Output "Warning: site collection $($spSite.Url) is larger than 100GB. Site collection size: $([int]($spSite.Usage.Storage/1GB))GB" -ForegroundColor Yellow
			}
			if ($spSite.AllWebs.Count -gt 250000) {
				Write-Output "Warning: site collection $($spSite.Url) has more than 250.000 sites. Number of sites: $($spSite.AllWebs.Count)" -ForegroundColor Yellow
			}
			$spSiteCount++
		}
		"`t</WebApplication>" | Out-File -FilePath $Save -Append:$true
		$spWAcount++
	}
	"</Metrics>" | Out-File -FilePath $Save -Append:$true
	Write-Output "Completed!"
}

function Get-SPSitesSize {
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[ValidateScript(
			{
				$_ -in (Get-SPWebApplication | % Url )			
			}
		)]
		[string]$WebApplicationUrl
	)
	SPLogMe
		
	$WebApp = Get-SPWebApplication -Identity $WebApplicationUrl
	$result = @()
	$Sites = Get-SPsite -WebApplication $WebApp -Limit All
	foreach ($Site in $Sites) {
		$SizeInKB = $Site.Usage.Storage /1024
		$SizeInMB = $SizeInKB / 1024
		$SizeInGB = $SizeInMB / 1024
		$result += New-Object PsObject @{ Site = $Site.RootWeb.Title; URL = $Site.Url; ContentDatabase = $Site.ContentDatabase.Name;
		Size = [math]::Round($SizeInMB, 2) }
		$Site.Dispose()
	}
	return $result
}

function Set-SPQuotaToSiteCollection {
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[ValidateScript(
			{
				$_ -in ([Microsoft.SharePoint.Administration.SPWebService]::ContentService.QuotaTemplates | % Name )			
			}
		)]
		[string]$TemplateName,
		[Parameter(Mandatory = $true)]
		[ValidateScript(
			{
				[Microsoft.SharePoint.SPSite]::Exists([System.Uri]$_, $true)
			}
		)]
		[string]$SiteCollectionUrl
		
	)
	SPLogMe

	$Site = Get-SPsite $SiteCollectionUrl
	Set-SPSite -Identity $Site -QuotaTemplate $TemplateName
}

#region Function Exports
Export-ModuleMember -Function Add-SharepointShellAdmin 
Export-ModuleMember -Function Enable-SPDeveloperDashboard 
Export-ModuleMember -Function Disable-SPDeveloperDashboard 
Export-ModuleMember -Function Switch-SPDeveloperDashboard 
Export-ModuleMember -Function Backup-WSPSolutions 
Export-ModuleMember -Function Backup-SPWebConfig
Export-ModuleMember -Function Rebuild-DistributedCacheServeInstance
Export-ModuleMember -Function Get-SPSolutionLastDeploymentSucceeded 
Export-ModuleMember -Function Create-ManagedPropertiesForSearchService
Export-ModuleMember -Function Change-SPDistributedCacheServiceAccount
Export-ModuleMember -Function Import-SPPropertyBag
Export-ModuleMember -Function Export-SPPropertyBag
Export-ModuleMember -Function Get-SPServersOn
Export-ModuleMember -Function Get-SPServers 
Export-ModuleMember -Function Get-SPDistributedCacheServersStatus
Export-ModuleMember -Function Add-SPSites2Localhost 
Export-ModuleMember -Function Get-SPServersNeedsUpgrade 
Export-ModuleMember -Function Start-SPWindowsServices
Export-ModuleMember -Function Stop-SPWindowsServices
Export-ModuleMember -Function Upgrade-SPContentDB 
Export-ModuleMember -Function Copy-SPList
Export-ModuleMember -Function Reset-AllSPIIS 
Export-ModuleMember -Function DoParallel-OnSPServers
Export-ModuleMember -Function SPLogMe 
Export-ModuleMember -Function Get-SPProductInformation
Export-ModuleMember -Function Start-AuditReportInterface
Export-ModuleMember -Function Upgrade-SPContentDatabasesIfNeeded
Export-ModuleMember -Function Get-SPContentDatabasesNeedUpgrade
Export-ModuleMember -Function Upgrade-SPServersIfNeeded
Export-ModuleMember -Function Set-EmailOptionForUserProfiles
Export-ModuleMember -Function FineTune-DistributedCaches
Export-ModuleMember -Function Get-SPManagedAccountsPassword
Export-ModuleMember -Function Get-SPDatabasesBackupSize
Export-ModuleMember -Function Do-GracefulShutdownDistributedCacheServices 
Export-ModuleMember -Function Rename-SPSite
Export-ModuleMember -Function Release-SPFileLock 
Export-ModuleMember -Function RestartAll-SPTimerJobServices
Export-ModuleMember -Function RecycleAll-SPWebApplicatonPools
Export-ModuleMember -Function Get-SPDatabaseDiskSize
Export-ModuleMember -Function Set-SPUsageLogRetentionDay
Export-ModuleMember -Function Set-SPUserAsSiteCollectionAdminOnWebApplication
Export-ModuleMember -Function Update-SPProfilePictureThumbnails
Export-ModuleMember -Function Add-SPCodeDomAuthorizedType
Export-ModuleMember -Function Remove-SPCodeDomAuthorizedType
Export-ModuleMember -Function Audit-SPUserProfile
Export-ModuleMember -Function Update-SPUserProfileNewsfeedPrivacy
Export-ModuleMember -Function Deploy-WSPSolutions 
Export-ModuleMember -Function Copy-SPPackages
Export-ModuleMember -Function Add-SPAllowedInlineDownloadedMimeTypes
Export-ModuleMember -Function Flush-SPBlobCache
Export-ModuleMember -Function AddOrUpdate-SPFarmProperty
Export-ModuleMember -Function Get-SPFarmProperty
Export-ModuleMember -Function Get-SPUserProfiles
Export-ModuleMember -Function Get-SPSearchServiceManagedProperties
Export-ModuleMember -Function Disable-SPFeatureInAllSites
Export-ModuleMember -Function Enable-SPFeatureInAllSites
Export-ModuleMember -Function Enable-SPFeatureInAllWebs
Export-ModuleMember -Function Disable-SPFeatureInAllWebs
Export-ModuleMember -Function Configure-SPSocialFeedCache
Export-ModuleMember -Function Export-SPRootCertificate
Export-ModuleMember -Function Refresh-SPMissingRootCertificate
Export-ModuleMember -Function Get-SPMissingWebPartDetails
Export-ModuleMember -Function Get-SPMissingAssemblyDetails
Export-ModuleMember -Function Get-SPMissingSetupFileDetails
Export-ModuleMember -Function Repopulate-SPSiteCollectionFeeds
Export-ModuleMember -Function Add-SPPeoplePickerADProvider
Export-ModuleMember -Function Clear-SPPeoplePickerADProvider
Export-ModuleMember -Function Remove-SPPeoplePickerADProvider
Export-ModuleMember -Function Get-SPSiteMetrics
Export-ModuleMember -Function Get-SPSitesSize
#endregion
