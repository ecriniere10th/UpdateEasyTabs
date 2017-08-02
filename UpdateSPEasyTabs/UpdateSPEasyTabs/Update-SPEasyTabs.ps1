######################################################################################################################################

##############################################################################
##
## SAP SharePoint Update Reporting
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
    [String[]]
    $SPSiteArray
)
foreach($fg in $SPSiteArray)
{
	if(($fg.displaytitle -eq "Easy Tabs 2010 - Orange") -or ($fg.displaytitle -eq "Easy Tabs 2007") -or ($fg.displaytitle -eq "Easy Tabs 2010 - Gray") -or ($fg.displaytitle -eq "EasyTabs2013"))
	{
		$fgimport += $fg
	}
}



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
    [String[]]
    $SPSiteArray
)
foreach($fgwp in $fgimport){$WebPartFileName = "EasyTabs2013.dwp";$WebPartZoneIndex = 1;$WebPartFilePath = "D:\EasyTabs\$WebPartFileName"; $WebPartZoneID = $fgwp.ZoneID;$SiteURL =  $fgwp.WebURL + "/";$PageURL = $fgwp.PageURL.Substring(1);$web = Get-SPWeb $SiteUrl;$pweb = [Microsoft.SharePoint.Publishing.PublishingWeb]::GetPublishingWeb($web);$wpm = $web.GetLimitedWebPartManager($PageURL, [System.Web.UI.WebControls.WebParts.PersonalizationScope]::Shared);write-host $($fgwp.WebURL + $fgwp.PageURL) -foregroundcolor Yellow;[xml]$WebPartXml = get-content $WebPartFilePath;$SR = New-Object System.IO.StringReader($WebPartXml.OuterXml);$XTR = New-Object System.Xml.XmlTextReader($SR);$Err = $null;$WP = $wpm.ImportWebPart($XTR, [ref] $Err); $wpm.AddWebPart($WP, $WebPartZoneID, $WebPartZoneIndex)}

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
    [String[]]
    $SPSiteArray
)


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
    [String[]]
    $SPSiteArray
)


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
    [String[]]
    $SPSiteArray
)


}






















