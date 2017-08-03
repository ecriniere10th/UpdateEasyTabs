##############################################################################
##############################################################################
##
## Update EasyTabs for Sharepoint
## Author : Eric Criniere
## Date : 01 August 2017
## Version : 1.00
## 
## 
## 
################################################################################

## Load SharePoint Modules

Add-PSSnapin Microsoft.SharePoint.PowerShell

## Load SharePoint EasyTabs Functions

function Enumerate-WebPartsSites{
<#
.SYNOPSIS
Enumerate through Sharepoint WebParts per each site
.DESCRIPTION
Enumerate through Sharepoint WebParts per each site
.PARAMETER SPFarmURL
SharePoint Farm URL
.EXAMPLE
Enumerate-WebPartsSites -SPFarmURL http://spsite.domain.com
#>
[CmdletBinding()]
Param (
    [Parameter(
        Mandatory=$True,
        ValueFromPipeline=$True,
        HelpMessage="SharePoint Farm URL")]
    [Alias('spuri')]
    [string[]]
    $SPFarmURL
)

$site = new-object Microsoft.SharePoint.SPSite $SPFarmURL    
foreach($web in $site.AllWebs) {
        $pWeb = [Microsoft.SharePoint.Publishing.PublishingWeb]::GetPublishingWeb($web)
        $pages = $web.Lists["Pages"]
        if ($pages) {foreach ($item in $pages.Items) {
                $fileUrl = $webUrl + “/” + $item.File.Url
                $manager = $item.file.GetLimitedWebPartManager([System.Web.UI.WebControls.Webparts.PersonalizationScope]::Shared);
                $wps = $manager.webparts
                $wps | select-object @{Expression={$pWeb.Url};Label=”WebURL”},@{Expression={$fileUrl};Label=”PageURL”},DisplayTitle,ZoneID,PartOrder,ID,Hidden
            }
            }
    }
}

function Enumerate-WebPartsSitePages {
<#
.SYNOPSIS
Enumerate through Sharepoint WebParts per each site
.DESCRIPTION
Enumerate through Sharepoint WebParts per each site
.PARAMETER SPFarmURL
SharePoint Farm URL
.EXAMPLE
Enumerate-WebPartsSitePages -SPFarmURL http://spsite.domain.com
#>
[CmdletBinding()]
Param (
[Parameter(
    Mandatory=$True,
    ValueFromPipeline=$True,
    HelpMessage="SharePoint Farm URL")]
[Alias('spuri')]
[string[]]
$SPFarmURL
)
$site = new-object Microsoft.SharePoint.SPSite $SPFarmURL    
foreach($web in $site.AllWebs) {
        $pages = $null
        $pages = $web.Lists["Site Pages"]
        if ($pages) {foreach ($item in $pages.Items) {
                $fileUrl = $webUrl + “/” + $item.File.Url
                $manager = $item.file.GetLimitedWebPartManager([System.Web.UI.WebControls.Webparts.PersonalizationScope]::Shared);
                $wps = $manager.webparts
                $wps | select-object @{Expression={$pWeb.Url};Label=”WebURL”},@{Expression={$fileUrl};Label=”PageURL”}, DisplayTitle, ZoneID,PartOrder,ID,Hidden
            }
        }                
    Write-Host “… completed processing” $web
}
}

function Enumerate-WebPartsRoot {
<#
.SYNOPSIS
Enumerate through Sharepoint WebParts per each site
.DESCRIPTION
Enumerate through Sharepoint WebParts per each site
.PARAMETER SPFarmURL
SharePoint Farm URL
.EXAMPLE
Enumerate-WebPartsRoot -SPFarmURL http://spsite.domain.com
#>
[CmdletBinding()]
Param (
[Parameter(
    Mandatory=$True,
    ValueFromPipeline=$True,
    HelpMessage="SharePoint Farm URL")]
[Alias('spuri')]
[string[]]
$SPFarmURL
)
    $site = new-object Microsoft.SharePoint.SPSite $SPFarmURL    
    foreach($web in $site.AllWebs) {
            $pages = $web.GetFile("default.aspx")
            if($pages){
            $manager = $web.GetLimitedWebPartManager("default.aspx",[System.Web.UI.WebControls.WebParts.PersonalizationScope]::Shared)
            $wps = $manager.webparts
            $wps | select-object @{Expression={$web.Url};Label=”WebURL”},@{Expression={$pages};Label=”PageURL”}, DisplayTitle, ZoneID,PartOrder,ID,Hidden
        }
    }
}

function Enumerate-WebParts {
<#
.SYNOPSIS
Enumerate through Sharepoint WebParts per each site
.DESCRIPTION
Enumerate through Sharepoint WebParts per each site
.PARAMETER SPFarmURL
SharePoint Farm URL
.EXAMPLE
Enumerate-WebParts -SPFarmURL http://spsite.domain.com
#>
[CmdletBinding()]
Param (
[Parameter(
    Mandatory=$True,
    ValueFromPipeline=$True,
    HelpMessage="SharePoint Farm URL")]
[Alias('spuri')]
[string[]]
$SPFarmURL
)
    $site = new-object Microsoft.SharePoint.SPSite $SPFarmURL    
    foreach($web in $site.AllWebs) {
        if ([Microsoft.SharePoint.Publishing.PublishingWeb]::IsPublishingWeb($web)) {
            $pWeb = [Microsoft.SharePoint.Publishing.PublishingWeb]::GetPublishingWeb($web)
            $pages = $pWeb.PagesList
            foreach ($item in $pages.Items) {
                $fileUrl = $webUrl + “/” + $item.File.Url
                $manager = $item.file.GetLimitedWebPartManager([System.Web.UI.WebControls.Webparts.PersonalizationScope]::Shared);
                $wps = $manager.webparts
                $wps | select-object @{Expression={$pWeb.Url};Label=”WebURL”},@{Expression={$fileUrl};Label=”PageURL”},DisplayTitle,IsVisible, @{Expression={$_.GetType().ToString()};Label=”Type”}
            }
        }
        else {
            $pages = $null
            $pages = $web.Lists["Site Pages"]
            if ($pages) {foreach ($item in $pages.Items) {
                    $fileUrl = $webUrl + “/” + $item.File.Url
                    $manager = $item.file.GetLimitedWebPartManager([System.Web.UI.WebControls.Webparts.PersonalizationScope]::Shared);
                    $wps = $manager.webparts
                    $wps | select-object @{Expression={$pWeb.Url};Label=”WebURL”},@{Expression={$fileUrl};Label=”PageURL”},DisplayTitle,IsVisible, @{Expression={$_.GetType().ToString()};Label=”Type”}
                }
            }
            else {
            $pages = $null
            $pages = $web.GetFile("default.aspx")
            $manager = $web.GetLimitedWebPartManager("default.aspx",[System.Web.UI.WebControls.WebParts.PersonalizationScope]::Shared)
            $wps = $manager.webparts
            $wps | select-object @{Expression={$pWeb.Url};Label=”WebURL”},@{Expression={$fileUrl};Label=”PageURL”},DisplayTitle,IsVisible, @{Expression={$_.GetType().ToString()};Label=”Type”}
            }
        }        Write-Host “… completed processing” $web
    }
}

function Get-SPEasyTabs{
<#
.SYNOPSIS
Retreive Sharepoint EasyTabs WebParts per each site
.DESCRIPTION
Retreive Sharepoint EasyTabs WebParts per each site
.PARAMETER SiteArray
Array of SharePoint sites with EasyTabs Webparts
.EXAMPLE
Get-SPEasyTabs -SPSiteArray $SPEasyTabsWP
#>
[CmdletBinding()]
Param (
    [Parameter(
        Mandatory=$True,
        ValueFromPipeline=$True
        )]
    $SPSiteArray
)
$fgimport = $null
$fgimport = @()

foreach($fg in $SPSiteArray)
{
	if(($fg.displaytitle -eq "Easy Tabs 2010 - Orange") -or ($fg.displaytitle -eq "Easy Tabs 2007") -or ($fg.displaytitle -eq "Easy Tabs 2010 - Gray") -or ($fg.displaytitle -eq "EasyTabs2013"))
	{
		$fgimport += $fg
	}	
}

$fgimport
}

function Add-SPEasyTabs{
<#
.SYNOPSIS
Adds Sharepoint EasyTabs WebParts per each site
.DESCRIPTION
Adds Sharepoint EasyTabs WebParts per each site
.PARAMETER SiteArray
Array of SharePoint sites with EasyTabs Webparts
.EXAMPLE
Add-SPEasyTabs -SPSiteArray $SPEasyTabsWP
#>
[CmdletBinding()]
Param (
    [Parameter(
        Mandatory=$True,
        ValueFromPipeline=$True
        )]
    $SPSiteArray
)
foreach($fgwp in $SPSiteArray)
	{
		$WebPartFileName = "EasyTabs2013.dwp"
		$WebPartZoneIndex = 1
		$WebPartFilePath = "D:\EasyTabs\$WebPartFileName"
		$WebPartZoneID = $fgwp.ZoneID
		$SiteURL =  $fgwp.WebURL + "/"
		$PageURL = $fgwp.PageURL.Substring(1)
		$web = Get-SPWeb $SiteUrl
		$pweb = [Microsoft.SharePoint.Publishing.PublishingWeb]::GetPublishingWeb($web)
		$wpm = $web.GetLimitedWebPartManager($PageURL, [System.Web.UI.WebControls.WebParts.PersonalizationScope]::Shared)
		write-host $($fgwp.WebURL + $fgwp.PageURL) -foregroundcolor Yellow
		[xml]$WebPartXml = get-content $WebPartFilePath
		$SR = New-Object System.IO.StringReader($WebPartXml.OuterXml)
		$XTR = New-Object System.Xml.XmlTextReader($SR)
		$Err = $null
		$WP = $wpm.ImportWebPart($XTR, [ref] $Err)
		$wpm.AddWebPart($WP, $WebPartZoneID, $WebPartZoneIndex)
	}

}

function Remove-SPEasyTabs{
<#
.SYNOPSIS
Removes Sharepoint EasyTabs WebParts per each site
.DESCRIPTION
Removes Sharepoint EasyTabs WebParts per each site
.PARAMETER SiteArray
Array of SharePoint sites with EasyTabs Webparts
.EXAMPLE
Remove-SPEasyTabs -SPSiteArray $SPEasyTabsWP
#>
[CmdletBinding()]
Param (
    [Parameter(
        Mandatory=$True,
        ValueFromPipeline=$True
        )]
    $SPSiteArray
)

foreach($fgwp in $SPSiteArray)
	{
		$WebPartZoneID = $fgwp.ZoneID
		$SiteURL =  $fgwp.WebURL + "/"
		$PageURL = $fgwp.PageURL
		$pageUrlv2 = $SiteURL + $PageURL
		$webpartID = $fgwp.id
		$web = Get-SPWeb $SiteUrl
		$pweb = [Microsoft.SharePoint.Publishing.PublishingWeb]::GetPublishingWeb($web)
		$wpm = $web.GetLimitedWebPartManager($PageURL, [System.Web.UI.WebControls.WebParts.PersonalizationScope]::Shared)
		write-host $pageUrlv2 -foregroundcolor Yellow
		$wpm.DeleteWebPart($wpm.Webparts[$webpartId])
	}
}

function CheckIn-SPEasyTabs{
<#
.SYNOPSIS
Checks in Sharepoint EasyTabs WebParts per each site
.DESCRIPTION
Checks in Sharepoint EasyTabs WebParts per each site
.PARAMETER SiteArray
Array of SharePoint sites with EasyTabs Webparts
.EXAMPLE
CheckIn-SPEasyTabs -SPSiteArray $SPEasyTabsWP
#>
[CmdletBinding()]
Param (
    [Parameter(
        Mandatory=$True,
        ValueFromPipeline=$True
        )]    
    $SPSiteArray
)
	foreach($fgwp in $SPSiteArray)
	{
		$SiteURL =  $fgwp.WebURL + "/"
		$PageURL = $fgwp.PageURL
		$web = Get-SPWeb $SiteUrl
		$pweb = [Microsoft.SharePoint.Publishing.PublishingWeb]::GetPublishingWeb($web)
		$pageUrlv2 = $fgwp.WebURL + $fgwp.PageURL
		$page = $web.GetFile($pageUrlv2)
		$wpm = $web.GetLimitedWebPartManager($PageURL, [System.Web.UI.WebControls.WebParts.PersonalizationScope]::Shared)
		Write-Host $pageUrlv2 -foregroundcolor Yellow
		Write-host "Checked Out Status is : "$page.CheckOutStatus -foregroundcolor Yellow
		if(!($page.checkoutstatus -eq "None"))
		{
			$page.checkin("Updated Webpart")
		}
	}
}

function CheckOut-SPEasyTabs{
<#
.SYNOPSIS
Checks out Sharepoint EasyTabs WebParts per each site
.DESCRIPTION
Checks out Sharepoint EasyTabs WebParts per each site
.PARAMETER SiteArray
Array of SharePoint sites with EasyTabs Webparts
.EXAMPLE
CheckIn-SPEasyTabs -SPSiteArray $SPEasyTabsWP
#>
[CmdletBinding()]
Param (
    [Parameter(
        Mandatory=$True,
        ValueFromPipeline=$True
        )]    
    $SPSiteArray
)
	foreach($fgwp in $SPSiteArray)
	{
		$SiteURL =  $fgwp.WebURL + "/"
		$PageURL = $fgwp.PageURL
		$web = Get-SPWeb $SiteUrl
		$pweb = [Microsoft.SharePoint.Publishing.PublishingWeb]::GetPublishingWeb($web)
		$pageUrlv2 = $fgwp.WebURL + $fgwp.PageURL
		$page = $web.GetFile($pageUrlv2)
		$wpm = $web.GetLimitedWebPartManager($PageURL, [System.Web.UI.WebControls.WebParts.PersonalizationScope]::Shared)
		Write-Host $pageUrlv2 -foregroundcolor Yellow
		Write-host "Checked Out Status is : "$page.CheckOutStatus -foregroundcolor Yellow
		if($page.checkoutstatus -eq "None")
		{
			$page.checkout()
		}
	}
}

function Check-SPEasyTabs{
<#
.SYNOPSIS
Checks to see if Sharepoint EasyTabs WebParts are checked in or out per each site
.DESCRIPTION
Checks to see if Sharepoint EasyTabs WebParts are checked in or out per each site
.PARAMETER SiteArray
Array of SharePoint sites with EasyTabs Webparts
.EXAMPLE
Check-SPEasyTabs -SPSiteArray $SPEasyTabsWP
#>
[CmdletBinding()]
Param (
    [Parameter(
        Mandatory=$True,
        ValueFromPipeline=$True
        )]    
    $SPSiteArray
)
foreach($fgwp in $SPSiteArray)
	{
		$SiteURL =  $fgwp.WebURL + "/"
		$PageURL = $fgwp.PageURL
		$web = Get-SPWeb $SiteUrl
		$pweb = [Microsoft.SharePoint.Publishing.PublishingWeb]::GetPublishingWeb($web)
		$pageUrlv2 = $fgwp.WebURL + $fgwp.PageURL
		$page = $web.GetFile($pageUrlv2)
		$wpm = $web.GetLimitedWebPartManager($PageURL, [System.Web.UI.WebControls.WebParts.PersonalizationScope]::Shared)
		Write-Host $pageUrlv2 -foregroundcolor Yellow
		Write-host "Checked Out Status is : "$page.CheckOutStatus -foregroundcolor Yellow
	}
}

function Publish-SPEasyTabs{
<#
.SYNOPSIS
Publishes Sharepoint EasyTabs WebParts per each site
.DESCRIPTION
Publishes Sharepoint EasyTabs WebParts per each site
.PARAMETER SiteArray
Array of SharePoint sites with EasyTabs Webparts
.EXAMPLE
Publish-SPEasyTabs -SPSiteArray $SPEasyTabsWP
#>
[CmdletBinding()]
Param (
    [Parameter(
        Mandatory=$True,
        ValueFromPipeline=$True
        )]    
    $SPSiteArray
)
	foreach($fgwp in $SPSiteArray)
	{
		$SiteURL =  $fgwp.WebURL + "/"
		$PageURL = $fgwp.PageURL
		$web = Get-SPWeb $SiteUrl
		$pweb = [Microsoft.SharePoint.Publishing.PublishingWeb]::GetPublishingWeb($web)
		$pageUrlv2 = $fgwp.WebURL + $fgwp.PageURL
		$page = $web.GetFile($pageUrlv2)
		$wpm = $web.GetLimitedWebPartManager($PageURL, [System.Web.UI.WebControls.WebParts.PersonalizationScope]::Shared)
		Write-Host $pageUrlv2 -foregroundcolor Yellow
		Write-host "Checked Out Status is : "$page.CheckOutStatus -foregroundcolor Yellow
		$page.checkin("Updated Webpart")
		$page.publish("")
		$web.Update()
	}
}

function Verify-SPEasyTabs{
<#
.SYNOPSIS
Verify status of Sharepoint EasyTabs WebParts per each site
.DESCRIPTION
Verify status of Sharepoint EasyTabs WebParts per each site
.PARAMETER SiteArray
Array of SharePoint sites with EasyTabs Webparts
.EXAMPLE
Verify-SPEasyTabs -SPSiteArray $SPEasyTabsWP
#>
[CmdletBinding()]
Param (
    [Parameter(
        Mandatory=$True,
        ValueFromPipeline=$True
        )]    
    $SPSiteArray
)
	foreach($fgwp in $SPSiteArray)
	{
		$SiteURL =  $fgwp.WebURL + "/"
		$PageURL = $fgwp.PageURL
		$web = Get-SPWeb $SiteUrl
		$pweb = [Microsoft.SharePoint.Publishing.PublishingWeb]::GetPublishingWeb($web)
		$wpm = $web.GetLimitedWebPartManager($PageURL, [System.Web.UI.WebControls.WebParts.PersonalizationScope]::Shared)
		write-host $($fgwp.WebURL + $fgwp.PageURL) -foregroundcolor Yellow
		foreach($wp in $wpm.webparts)
		{
			if($wp.zoneid -eq $fgwp.zoneid)
			{
				$wp | Select Title,ZoneID,PartOrder,Hidden
			}
		}
	}
}


## Parameters

## SharePoint Farm URL
$Url = "http://spfarm.domain.com/"

## Begin with null values for variables that will enumerate all webparts for the SharePoint Farm provided

$fgwpet = $null
$fgwproot = $null
$fgwpsites = $null
$fgwpsitepages = $null
$fgimportet = $null
$fgimportroot = $null
$fgimportsites = $null
$fgimportsitepages = $null

## Sequence of events

## Enumerate SharePoint webparts

$fgwpet = Enumerate-WebParts -SPFarmURL $Url
$fgwproot = Enumerate-WebPartsRoot -SPFarmURL $Url
$fgwpsites = Enumerate-WebPartsSites -SPFarmURL $Url
$fgwpsitepages = Enumerate-WebPartsSitePages -SPFarmURL $Url

## Retreive all EasyTab webparts for each SharePoint site

$fgimportet = Get-SPEasyTabs -SPSiteArray $fgwpet
$fgimportroot = Get-SPEasyTabs -SPSiteArray $fgwproot
$fgimportsites = Get-SPEasyTabs -SPSiteArray $fgwpsites
$fgimportsitepages = Get-SPEasyTabs -SPSiteArray $fgwpsitepages

## Verify status of EasyTab webparts for each SharePoint site : Are sites checked in or out?
## Status : None = Checked In
## Status : LongTerm = Checked Out

Check-SPEasyTabs -SPSiteArray $fgimportet
Check-SPEasyTabs -SPSiteArray $fgimportroot
Check-SPEasyTabs -SPSiteArray $fgimportsites
Check-SPEasyTabs -SPSiteArray $fgimportsitepage

## Check out SharePoint sites

CheckOut-SPEasyTabs -SPSiteArray $fgimportet
CheckOut-SPEasyTabs -SPSiteArray $fgimportroot
CheckOut-SPEasyTabs -SPSiteArray $fgimportsites
CheckOut-SPEasyTabs -SPSiteArray $fgimportsitepage

## Verify status of EasyTab webparts for each SharePoint site : Are sites checked in or out?
## Status : None = Checked In
## Status : LongTerm = Checked Out

Check-SPEasyTabs -SPSiteArray $fgimportet
Check-SPEasyTabs -SPSiteArray $fgimportroot
Check-SPEasyTabs -SPSiteArray $fgimportsites
Check-SPEasyTabs -SPSiteArray $fgimportsitepage

## Remove EasyTab webparts for each SharePoint sites

Remove-SPEasyTabs -SPSiteArray $fgimportet
Remove-SPEasyTabs -SPSiteArray $fgimportroot
Remove-SPEasyTabs -SPSiteArray $fgimportsites
Remove-SPEasyTabs -SPSiteArray $fgimportsitepage

## Check in SharePoint sites

CheckIn-SPEasyTabs -SPSiteArray $fgimportet
CheckIn-SPEasyTabs -SPSiteArray $fgimportroot
CheckIn-SPEasyTabs -SPSiteArray $fgimportsites
CheckIn-SPEasyTabs -SPSiteArray $fgimportsitepage

## Verify status of EasyTab webparts for each SharePoint site : Are sites checked in or out?
## Status : None = Checked In
## Status : LongTerm = Checked Out

Check-SPEasyTabs -SPSiteArray $fgimportet
Check-SPEasyTabs -SPSiteArray $fgimportroot
Check-SPEasyTabs -SPSiteArray $fgimportsites
Check-SPEasyTabs -SPSiteArray $fgimportsitepage

## Check out SharePoint sites

CheckOut-SPEasyTabs -SPSiteArray $fgimportet
CheckOut-SPEasyTabs -SPSiteArray $fgimportroot
CheckOut-SPEasyTabs -SPSiteArray $fgimportsites
CheckOut-SPEasyTabs -SPSiteArray $fgimportsitepage

## Add new EasyTab webparts for each SharePoint site listed

Add-SPEasyTabs -SPSiteArray $fgimportet
Add-SPEasyTabs -SPSiteArray $fgimportroot
Add-SPEasyTabs -SPSiteArray $fgimportsites
Add-SPEasyTabs -SPSiteArray $fgimportsitepage

## Check in SharePoint sites

CheckIn-SPEasyTabs -SPSiteArray $fgimportet
CheckIn-SPEasyTabs -SPSiteArray $fgimportroot
CheckIn-SPEasyTabs -SPSiteArray $fgimportsites
CheckIn-SPEasyTabs -SPSiteArray $fgimportsitepage

## Publish SharePoint sites

Publish-SPEasyTabs -SPSiteArray $fgimportet
Publish-SPEasyTabs -SPSiteArray $fgimportroot
Publish-SPEasyTabs -SPSiteArray $fgimportsites
Publish-SPEasyTabs -SPSiteArray $fgimportsitepage

## Verify EasyTab webparts are present for the SharePoint sites listed

Verify-SPEasyTabs -SPSiteArray $fgimportet
Verify-SPEasyTabs -SPSiteArray $fgimportroot
Verify-SPEasyTabs -SPSiteArray $fgimportsites
Verify-SPEasyTabs -SPSiteArray $fgimportsitepage