function enumerateWebPartsSites($Url) {
    $site = new-object Microsoft.SharePoint.SPSite $Url    
    foreach($web in $site.AllWebs) {
            $pWeb = [Microsoft.SharePoint.Publishing.PublishingWeb]::GetPublishingWeb($web)
            $pages = $web.Lists["Pages"]
            if ($pages) {foreach ($item in $pages.Items) {
                    $fileUrl = $webUrl + “/” + $item.File.Url
                    $manager = $item.file.GetLimitedWebPartManager([System.Web.UI.WebControls.Webparts.PersonalizationScope]::Shared);
                    $wps = $manager.webparts
                    $wps | select-object @{Expression={$pWeb.Url};Label=”WebURL”},@{Expression={$fileUrl};Label=”PageURL”}, DisplayTitle, ZoneID,PartOrder,ID,Hidden
                }
             }
        }
    }