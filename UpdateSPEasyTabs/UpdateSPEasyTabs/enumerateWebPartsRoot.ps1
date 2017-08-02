function enumerateWebPartsRoot($Url) {
    $site = new-object Microsoft.SharePoint.SPSite $Url    
    foreach($web in $site.AllWebs) {
            $pages = $web.GetFile("default.aspx")
            if($pages){
            $manager = $web.GetLimitedWebPartManager("default.aspx",[System.Web.UI.WebControls.WebParts.PersonalizationScope]::Shared)
            $wps = $manager.webparts
            $wps | select-object @{Expression={$web.Url};Label=”WebURL”},@{Expression={$pages};Label=”PageURL”}, DisplayTitle, ZoneID,PartOrder,ID,Hidden
        }
    }
}