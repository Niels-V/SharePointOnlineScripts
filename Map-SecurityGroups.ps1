Import-Module .\SharePointOnlineScripts.psm1
#set credentials for source system
Set-SpoCredential -UserName "example@orgA.onmicrosoft.com" -Password "xyz"
#set credentials for target system 
$cred = New-Object -TypeName System.Management.Automation.PSCredential -argumentlist "example@orgB.onmicrosoft.com", $(convertto-securestring "123" -asplaintext -force)
Connect-SPOService -Url "https://orgB-admin.sharepoint.com" -Credential $cred

function Get-Groups ([string]$url) {
	$groups = @()
	
	$x = Invoke-SpoRest -Url $url
	foreach ($e in $x.feed.entry) {
		$group = New-Object PSObject                                       
		$group | add-member Noteproperty Title       $e.content.properties.Title                 
		$group | add-member Noteproperty LoginName   $e.content.properties.LoginName
		$groups += $group
	}
	return $groups
}

function Get-SpoGroup([string]$siteUrl) {
	$g = Get-Groups -Url "$siteUrl/_api/web/siteusers?`$filter=PrincipalType%20eq%204"
	
	$hash = @{            
        Site       = $siteUrl                 
        Groups             = $g              
	}                           
                                    
    $Object = New-Object PSObject -Property $hash
	return $Object
}

function Get-SiteMapping([string]$oldSite)
{
	switch ($oldSite) 
    { 
		"https://orgA.sharepoint.com/sites/site1/"         	{ return "https://orgB.sharepoint.com/sites/site1"}
		"https://orgA-2.sharepoint.microsoftonline.com" 	{ return "https://orgB.sharepoint.com/sites/site2"}
	}
	return ""
}

#Get all the source sites in a array
$oldSites = @("https://orgA-2.sharepoint.microsoftonline.com",
			  "https://orgA.sharepoint.com/sites/site1"
)

#retrieve old site security group info by calling Rest Uservice via Get-SpoGroup
$sites = @()
foreach ($oldSite in $oldSites) {
	$s = Get-SpoGroup -siteUrl $oldSite
	$sites += $s
}

$groupMapping = @()

foreach ($site in $sites) {
	#get oldsite-newsite mapping
	$newSite = Get-SiteMapping -oldSite $site.Site
	
	foreach ($group in $site.Groups) {
		#create AD Group into the new site
		$u = Add-SPOUser -Site $newSite -LoginName $group.Title -Group "Visitors"
		#$u.LoginName now contains the sid
		$hash = @{            
			Title   = $group.Title                 
			OldSid  = $group.LoginName
			NewSid  = $u.LoginName
		}                           
        $groupMap = New-Object PSObject -Property $hash
		$groupMapping += $groupMap
	}
}
$groupMapping  | Export-Csv -Path .\GroupMapping.csv -NoTypeInformation