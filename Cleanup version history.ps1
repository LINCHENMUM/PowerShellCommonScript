#Add-PSSnapin "Microsoft.SharePoint.PowerShell"
if ((Get-PSSnapin -Name Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue) -eq $null )
{
    Add-PsSnapin Microsoft.SharePoint.PowerShell
}


$w = Get-SPWeb -Identity http://ideltaam.deltaww.com/sites/sales/iabg/fob_price	         
$l = $w.GetList("http://ideltaam.deltaww.com/sites/sales/iabg/fob_price/Lists/FOB_Price")    
$items = $l.Items;            
$f = $l.Fields["Check In Comment"];            
$listType = $l.GetType().Name;            
foreach($item in $items)            
{            
    $itemTitle = $item.Title;            
    if($listType -eq "SPDocumentLibrary")            
    {            
        if($itemTitle -eq ""){$itemTitle = $item["Name"];}            
    }            
    if($item.Versions.Count -gt 0){            
        $vtr = $item.Versions.Count;            
        Write-Host "$itemTitle, has $vtr versions" -foregroundcolor Green;            
                     
        while($vtr -gt 0){                     
            $vtr--;            
            [Microsoft.SharePoint.SPListItemVersion]$iv = $item.Versions[$vtr];            
            $versionNumber = $iv.VersionLabel;            
            if(!$iv.VersionLabel.EndsWith(".0"))            
            {                          
                continue;            
            }                      
            #Write-Host "$itemTitle : Deleted version $versionNumber" -foregroundcolor Yellow;            
            $iv.Delete();            
        }            
    }            
}
echo "Done successfully"