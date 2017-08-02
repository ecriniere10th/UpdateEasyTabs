function enumerateWebPartsSitePages($Url) {
    $site = new-object Microsoft.SharePoint.SPSite $Url    
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
