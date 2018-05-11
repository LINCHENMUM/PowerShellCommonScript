#Your Shaeproint Site URL
$web = Get-SPWeb "http://yourweb.com/bu/admin/IT"
$web.AllowUnsafeUpdates = $true

#Your List Name
$list = $web.Lists["Contracts"]
$count = 0

#Loop through all Items in List then loop through all Workflows on each List Items.         
foreach ($listItem in $list.Items) 
{
	foreach ($workflow in $listItem.Workflows) 
	{
		#Disregard Completed Workflows 
		if(($listItem.Workflows | where {$_.InternalState -ne "Completed"}) -ne $null)
		{
			#Cancel Workflows        
			#[Microsoft.SharePoint.Workflow.SPWorkflowManager]::CancelWorkflow($workflow)    
			write-output "Workflow cancelled for : " $listItem.Title
            $count=$count+1 
		}
	}
}
echo $count
$web.Dispose();