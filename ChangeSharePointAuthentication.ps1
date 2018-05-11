# Change to Claims Based Authentication
$setcba = Get-SPWebApplication "http://localhost/"
$setcba.UseClaimsAuthentication = 1;
$setcba.Update()

#To revert back to Classic mode authentication (disabled Claims Based Authentication) just change the 1 to a 0: