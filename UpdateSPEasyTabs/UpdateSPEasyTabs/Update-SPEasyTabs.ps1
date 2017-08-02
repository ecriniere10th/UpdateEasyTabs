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

## Load SharePoint Functions

function Enumerate-WebPartsSites{
<#
.SYNOPSIS
Enumerate through Sharepoint WebParts per each site
.DESCRIPTION
Enumerate through Sharepoint WebParts per each site
.PARAMETER SPFarmURL
The name of the computer to query.  Accepts multiple values and accepts pipeline input.
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
The name of the computer to query.  Accepts multiple values and accepts pipeline input.
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
The name of the computer to query.  Accepts multiple values and accepts pipeline input.
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
The name of the computer to query.  Accepts multiple values and accepts pipeline input.
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
                $wps | select-object @{Expression={$pWeb.Url};Label=”Web URL”},@{Expression={$fileUrl};Label=”Page URL”}, DisplayTitle, IsVisible, @{Expression={$_.GetType().ToString()};Label=”Type”}
            }
        }
        else {
            $pages = $null
            $pages = $web.Lists["Site Pages"]
            if ($pages) {foreach ($item in $pages.Items) {
                    $fileUrl = $webUrl + “/” + $item.File.Url
                    $manager = $item.file.GetLimitedWebPartManager([System.Web.UI.WebControls.Webparts.PersonalizationScope]::Shared);
                    $wps = $manager.webparts
                    $wps | select-object @{Expression={$pWeb.Url};Label=”Web URL”},@{Expression={$fileUrl};Label=”Page URL”}, DisplayTitle, IsVisible, @{Expression={$_.GetType().ToString()};Label=”Type”}
                }
            }
            else {
            $pages = $null
            $pages = $web.GetFile("default.aspx")
            $manager = $web.GetLimitedWebPartManager("default.aspx",[System.Web.UI.WebControls.WebParts.PersonalizationScope]::Shared)
            $wps = $manager.webparts
            $wps | select-object @{Expression={$pWeb.Url};Label=”Web URL”},@{Expression={$fileUrl};Label=”Page URL”}, DisplayTitle, IsVisible, @{Expression={$_.GetType().ToString()};Label=”Type”}
            }
        }        Write-Host “… completed processing” $web
    }
}

































