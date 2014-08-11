Add-Type –Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type –Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"

$Credentials = $null

function Invoke-SpoRest(
	[Parameter(Mandatory=$True)][String]$Url,[Parameter(Mandatory=$False)]
	[Microsoft.PowerShell.Commands.WebRequestMethod]$Method = [Microsoft.PowerShell.Commands.WebRequestMethod]::Get
	)
{
	$request = [System.Net.WebRequest]::Create($Url)
	$request.Credentials = $script:Credentials
	$request.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f")
	$request.Accept = "application/xml"
	$request.Method=[Microsoft.PowerShell.Commands.WebRequestMethod]::Get
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
		$SecurePassword = Read-Host -Prompt "Enter the password" -AsSecureString
	}
	else {
		$SecurePassword = $Password | ConvertTo-SecureString -AsPlainText -Force
	}
	$script:Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($UserName, $SecurePassword)
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


	
export-modulemember -function Invoke-SpoRest, Set-SpoCredential, Get-SpoCredential, Switch-SpoFeature, Get-SpoWorkflows