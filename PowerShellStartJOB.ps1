##Start a job
$job = Start-Job -Name "CompareSAPandSharePointData" -FilePath "C:\app\Scripts\CompareSAPandSharePointData.ps1"
## View a job
Get-Job
#Get-ScheduledJob
Get-Job | where PSJobTypeName -eq PSScheduledJob
## get a job result
Receive-Job -Job $job
#Receive-Job -Id 7
## Register a scheduled-job
$trigger = New-JobTrigger -Daily -At "09:12"
##$option = New-ScheduledJobOption -WakeToRun -StartIfNotIdle -MultipleInstancesPolicy Queue
Register-ScheduledJob -Name "CompareSAPandSharePointData" -FilePath "C:\app\Scripts\CompareSAPandSharePointData.ps1" -Trigger $trigger -ScheduledJobOption @{WakeToRun=$true; StartIfNotIdle=$true; MultipleInstancePolicy="Queue";RunElevated=$true}
#Unregister job
Unregister-ScheduledJob -Name CompareSAPandSharePointData
##Get ScheduledTask
Get-ScheduledTask
