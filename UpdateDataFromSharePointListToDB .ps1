# Config the task command line
#powershell.exe -executionPolicy bypass -file C:\app\scripts\UpdateDataFromSharePointListToDB.ps1

if ((Get-PSSnapin -Name Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue) -eq $null )
{
    Add-PsSnapin Microsoft.SharePoint.PowerShell
}

#Database config
$dataSource = "Database server name"
$database = "Database name"
$connectionString = "Server=$dataSource;Database=$database;Trusted_Connection=True;"
#SharePoint config
$spWeb = Get-SPWeb http://yourSharePoint.com/sites/MyWeb
$spList = $web.lists['List name']

function ReadDBTable($sql){
    #echo $sql
    $table = new-object "System.Data.DataTable"
  
	$connection = New-Object System.Data.SqlClient.SqlConnection ($connectionString)
	$connection.Open()
	if($connection.State -eq [System.Data.ConnectionState]::Open)
	{
		$command = New-Object System.Data.SqlClient.SqlCommand ($sql, $connection)
        $recordset = $command.ExecuteReader()
        $table.Load($recordset)

		$recordset.Close()
		$connection.Close();
	}
  return ,$table;
}

function InsertDBTable($sql){
     
	$connection = New-Object System.Data.SqlClient.SqlConnection ($connectionString)
	$connection.Open()
	if($connection.State -eq [System.Data.ConnectionState]::Open)
	{
		$command = New-Object System.Data.SqlClient.SqlCommand ($sql, $connection)
        $command.ExecuteReader()   
		$connection.Close();
	}
}

function UpdateDataTOTaxExemptCertificate($table){	
    $cnt=0;
	foreach($row in $table.Rows){
        $sql="INSERT INTO [dbo].[TaxExemptCertificate]
           ([CustomerName]
           ,[Country]
           ,[CompanyCode]
           ,[SoldToCode]
           ,[ShiptoState]
           ,[TaxExemptStates]
           ,[CustTaxClass]
           ,[MaterialTaxClass]
           ,[Amount]
           ,[ValidFrom]
           ,[ValidTo]
           ,[TaxCode]
           ,[LicenseNo]
           ,[LicenseDate]
           ,[LastModify]
           ,[ModifyBy]
           ,[SAPStatus]
           ,[CreateTime]
           ,[UpdateToSP]
           )
     VALUES
           ('"+$row.CustomerName+"'
           ,'"+$row.Country+"'
           ,'"+$row.CompanyCode+"'
           ,'"+$row.SoldToCode+"'
           ,'"+$row.ShiptoState+"'
           ,'"+$row.TaxExemptStates+"'
           ,'"+$row.CustTaxClass+"'
           ,'"+$row.MaterialTaxClass+"'
           ,"+$row.Amount+"
           ,'"+$row.ValidFrom+"'
           ,'"+$row.ValidTo+"'
           ,'"+$row.TaxCode+"'
           ,'"+$row.LicenseNo+"'
           ,'"+$row.LicenseDate+"'
           ,'"+$row.LastModify+"'
           ,'"+$row.ModifyBy+"'
           ,'"+$row.SAPStatus+"'
           ,getdate()
           ,'No')"
	InsertDBTable $sql
	$cnt+=1
    }
    Write-Host 'Total changes: ' $cnt
}

$tDataRecord = ReadDBTable "SELECT a.[customername],
       a.[country],
       a.[company code]                            AS CompanyCode,
       a.[sold-to (customer code)]                 AS SoldToCode,
       a.[ship-to state]                           AS ShiptoState,
       a.[taxexemptstates],
       a.[cust tax class.]                         AS CustTaxClass,
       a.[material tax class.]                     AS MaterialTaxClass,
       a.[amount (%)]                              AS Amount,
       CONVERT( DATE, a.[valid from] )             AS ValidFrom,
       CONVERT( DATE, a.[valid to] )               AS ValidTo,
       a.[tax code]                                AS TaxCode,
       Isnull( a.[license no. (optional)], NULL )  AS LicenseNo,
       Isnull( a.[license date (optional)], NULL ) AS LicenseDate,
       a.[lastmodify],
       a.[modifyby],
       a.[sapstatus],
       Isnull( a.[sapmessage], NULL )              AS SAPMessage,
       Isnull( a.[sapposttime], NULL )             AS SAPPostTime
FROM   [AMCSQL01].[SharePointViews].[dbo].[VTAXEXEMPTCERTIFICATE_NEW] a
       LEFT JOIN [dbo].[TAXEXEMPTCERTIFICATE] b
              ON a.[company code] = b.[companycode] AND
                 a.[sold-to (customer code)] = b.[soldtocode] AND
                 a.[ship-to state] = b.[shiptostate] AND
                 a.[valid from] = b.[validfrom]
WHERE  a.[sapstatus] = 'ReadyToPost' AND
       b.[companycode] IS NULL";

UpdateDataTOTaxExemptCertificate $tDataRecord