#Stop

Import-Module WebAdministration
$applicationPools = Get-ChildItem IIS:\AppPools | ? {$_.state -ne "Stopped"}
foreach($applicationPool in $applicationPools)
{ 
Write-host "Stopping application pool ""($($applicationPool.name)""" $applicationPool.Stop() 
}


#Start 
Import-Module WebAdministration 
$applicationPools = Get-ChildItem IIS:\AppPools |? {$_.state -ne "started"} 
foreach($applicationPool in $applicationPools)
{ 
Write-host "Starting application pool ""($($applicationPool.name)""" $applicationPool.Start()
}