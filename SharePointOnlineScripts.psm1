Add-Type –Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type –Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"

$Credentials = $null

function Invoke-SpoRest(
	[Parameter(Mandatory=$True)][String]$Url,[Parameter(Mandatory=$False)]
	[Microsoft.PowerShell.Commands.WebRequestMethod]$Method = [Microsoft.PowerShell.Commands.WebRequestMethod]::Get
	)
{
	Write-Debug "Calling REST service $Method $Url"
	$request = [System.Net.WebRequest]::Create($Url)
	$request.Credentials = $script:Credentials
	$request.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f")
	$request.Accept = "application/xml"
	$request.Method=$Method
	$response = $request.GetResponse()
	$requestStream = $response.GetResponseStream()
	$readStream = New-Object System.IO.StreamReader $requestStream
	$data=$readStream.ReadToEnd()
	[xml]$results = $data
	return $results
}

function Get-SpoCredential {
	<#
	.Synopsis
		Returns the stored SharePoint online credentials
	.Description
		Returns the stored SharePoint online credentials
	#>
	return $script:Credentials
}

function Set-SpoCredential(
	[Parameter(Mandatory=$True)][String]$UserName,
	[Parameter(Mandatory=$False)][String]$Password
)
{
	if([string]::IsNullOrEmpty($Password)) {
		$SecurePassword = Read-Host -Prompt "Enter the password for user $UserName" -AsSecureString
	}
	else {
		$SecurePassword = $Password | ConvertTo-SecureString -AsPlainText -Force
	}
	$script:Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($UserName, $SecurePassword)
	Write-Verbose "Credentials set for user $UserName"
}

function Switch-SpoFeature(
	[Parameter(Mandatory=$true)][String]$Url,
	[Parameter(Mandatory=$true)][ValidateSet('Web','Site')][String]$Scope,
	[Parameter(Mandatory=$true)][GUID]$FeatureId = [System.Guid]::Empty,
	[switch]$Disable,
	[switch]$Force
) 
{
<#
	.Synopsis
		Enables or disables a SharePoint feature on the specified URL
	.Description
		Enables or disables a SharePoint feature on the specified URL
	#>
	$clientContext = New-Object Microsoft.SharePoint.Client.ClientContext($Url)
	$clientContext.Credentials = $script:Credentials
 
	$features =  $null
	if ($Scope -eq "Web") {
		$features = $clientContext.Web.Features
	}
	elseif ($Scope -eq "Site") {
		$features = $clientContext.Site.Features
	}
	else {
		Write-Error "Wrong scope defined, only site and web are supported!"
		return
	}
	$clientContext.Load($features)
	$clientContext.ExecuteQuery()
	
	if (-Not $Disable)
	{
		$features.Add($FeatureId, $Force, [Microsoft.SharePoint.Client.FeatureDefinitionScope]::None)
	}
	else
	{
		$features.Remove($FeatureId, $Force)
	}
 
	try
	{
		$clientContext.ExecuteQuery()
		if (-Not $Disable)
		{
			Write-Host "Feature '$FeatureId' successfully activated.."
		}
		else
		{
			Write-Host "Feature '$FeatureId' successfully deactivated.."
		}
	}
	catch
	{
		Write-Error "An error occurred whilst activating/deactivating the Feature. Error detail: $($_)"
	}
	finally 
	{
		$clientContext.Dispose()
	}
	
}

function Get-SpoWorkflows(
	[Parameter(Mandatory=$true)][String]$Url
)
{
	$clientContext = New-Object Microsoft.SharePoint.Client.ClientContext($Url)
	$clientContext.Credentials = $script:Credentials
	
	$rootWeb = $clientContext.Site.RootWeb
	
	$clientContext.Load($rootWeb)
	$clientContext.ExecuteQuery()
	
	return Get-WorkflowFromWeb($rootWeb)
}

function Get-WorkflowFromWeb($Web) {
	$workflows = @()
	$clientContext.Load($web.Lists)
	$clientContext.ExecuteQuery()
	foreach($list in $web.Lists)
	{
		$clientContext.Load($list.WorkflowAssociations)
		$clientContext.ExecuteQuery()
		foreach($wf in $list.WorkflowAssociations)
        {
			if ($wf.Name -notlike "*Previous Version*")
            {
				$workflow = new-object PSObject
				$workflow | add-member -membertype NoteProperty -name "WebUrl" -value $web.Url
				$workflow | add-member -membertype NoteProperty -name "ListTitle" -value $list.Title
				$workflow | add-member -membertype NoteProperty -name "WorkflowName" -value $wf.Name
				$workflows += $workflow
			}
		}
	}
	$clientContext.Load($web.Webs)
	$clientContext.ExecuteQuery()
	foreach($subweb in $web.Webs)
    {
		 $subwebFlows = Get-WorkflowFromWeb($subweb)
		 $workflows += $subwebFlows
    }
	return $workflows
}

function Get-SharePointGroupsWithMembers($SiteCollectionUrl) {
	$groups = @()
	Write-Verbose "Calling $SiteCollectionUrl/_api/web/sitegroups"
	$x = Invoke-SpoRest -Url "$SiteCollectionUrl/_api/web/sitegroups"
	foreach ($e in $x.feed.entry) {
		$group = New-Object PSObject                                       
		$group | add-member Noteproperty Title       $e.content.properties.Title                 
		$group | add-member Noteproperty Id   $e.id
		
		$usersEndpoint = $e.id+"/Users"
		Write-Verbose "Calling $usersEndpoint"
		$u = Invoke-SpoRest -Url $usersEndpoint
		$users = @()
		foreach ($y in $u.feed.entry) {
			$user = New-Object PSObject                                       
			$user | add-member Noteproperty Title       $y.content.properties.Title                 
			$user | add-member Noteproperty Id   $y.content.properties.Id.InnerText
			$principalTypeId = [int]$y.content.properties.PrincipalType.InnerText
			$principalType = [Microsoft.SharePoint.Client.Utilities.PrincipalType].GetEnumName($principalTypeId)
			$user | add-member Noteproperty MemberType    $principalType	
			$users += $user
		}
		
		$group | add-member Noteproperty Users $users
		$groups += $group
	}
	return $groups
}

function Get-Webs($SiteUrl) {
	$groups = @()
	Write-Verbose "Calling $SiteUrl/_api/web/webs"
	$x = Invoke-SpoRest -Url "$SiteUrl/_api/web/webs"
	foreach ($e in $x.feed.entry) {
		$group = New-Object PSObject                                       
		$group | add-member Noteproperty Title       $e.content.properties.Title     
		$group | add-member Noteproperty Url       $e.content.properties.Url     		
		$group | add-member Noteproperty Id   $e.id
		
		$groups += $group
		$groups += Get-Webs -SiteUrl $e.content.properties.Url
		
	}
			
	return $groups
}

function Get-SiteCollectionPermissions($SiteCollectionUrl) {
	$subwebs = Get-Webs($SiteCollectionUrl)
	$results = @()
	
	$s = Get-SitePermissions($SiteCollectionUrl)
	$results += $s
	foreach($web in $subwebs) {
		$s = Get-SitePermissions($web.Url)
		$results += $s
	}
	return $results
}

function Get-SitePermissions($SiteUrl) {
	$roleAssignments = @()
	Write-Verbose "Calling $SiteUrl/_api/web/RoleAssignments"
	$x = Invoke-SpoRest -Url "$SiteUrl/_api/web/RoleAssignments"
	foreach ($entry in $x.feed.entry) {
		$principalId = $entry.content.properties.PrincipalId.InnerText
		
		Write-Verbose "Calling $SiteUrl/_api/web/RoleAssignments/GetByPrincipalId($principalId)/Member"
		$memberResponse = Invoke-SpoRest -Url "$SiteUrl/_api/web/RoleAssignments/GetByPrincipalId($principalId)/Member"
		
		$memberTitle = $memberResponse.entry.content.properties.Title
		$principalTypeId = [int]$memberResponse.entry.content.properties.PrincipalType.InnerText
		$principalType = [Microsoft.SharePoint.Client.Utilities.PrincipalType].GetEnumName($principalTypeId)
		Write-Verbose "Resolved $memberTitle with $principalType - $principalTypeId"
		Write-Verbose "$SiteUrl/_api/web/RoleAssignments/GetByPrincipalId($principalId)/RoleDefinitionBindings"
		$roleDefBindingsResponse = Invoke-SpoRest -Url "$SiteUrl/_api/web/RoleAssignments/GetByPrincipalId($principalId)/RoleDefinitionBindings"
		
		$rdbs = ""
		
		foreach	($r in $roleDefBindingsResponse.feed.entry) {
			$rdbs += $r.content.properties.Name + ";"
		}
	
		$roleAssignment = New-Object PSObject                                       
		$roleAssignment | add-member Noteproperty Id        $principalId
		$roleAssignment | add-member Noteproperty Site      $SiteUrl
		$roleAssignment | add-member Noteproperty Member    $memberTitle
		$roleAssignment | add-member Noteproperty MemberType    $principalType		
		$roleAssignment | add-member Noteproperty Roles		$rdbs
		$roleAssignments += $roleAssignment
	}
	return $roleAssignments
}

	
export-modulemember -function Invoke-SpoRest, Set-SpoCredential, Get-SpoCredential, Switch-SpoFeature, Get-SpoWorkflows, Get-SharePointGroupsWithMembers, Get-SiteCollectionPermissions, Get-SitePermissions, Get-Webs