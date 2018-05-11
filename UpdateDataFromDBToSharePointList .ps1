#  Config the task command line
# powershell.exe -executionPolicy bypass -file C:\app\scripts\UpdateDataFromDBToSharePointList.ps1

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

function UpdateDBTable($sql){
     
	$connection = New-Object System.Data.SqlClient.SqlConnection ($connectionString)
	$connection.Open()
	if($connection.State -eq [System.Data.ConnectionState]::Open)
	{
		$command = New-Object System.Data.SqlClient.SqlCommand ($sql, $connection)
        $command.ExecuteReader()   
		$connection.Close();
	}
}

function UpdateDataToResaleCertificateList($table){	
    $cnt=0;
	foreach($row in $table.Rows){
        
		$CompanyCode=$row.CompanyCode
		$SoldToCode=$row.SoldToCode
        $TaxExemptState=$row.TaxExemptStates
        #$TaxExemptState=';#Michigan,MI;#'
        $ValidFrom=$row.ValidFrom

		$SAPStatus=$row.SAPStatus
        if ($row.SAPMessage -eq [DBNull]::Value){
            $SAPMessage=$null;
        }else{
            $SAPMessage=$row.SAPMessage
        }
        if($row.SAPPostTime -eq [DBNull]::Value){
            $SAPPostTime=$null
        }else{
		    $SAPPostTime=$row.SAPPostTime
        }

        #echo $TaxExemptState
        #echo $ValidFrom
        #echo $row.SAPPostTime
        #echo $SAPPostTime
		
		
		 $spQuery = New-Object Microsoft.SharePoint.SPQuery
		 $camlQuery = '<Where>
		   <And>
		    <And>
				<Eq>
				<FieldRef Name="Company_x0020_Code" /><Value Type="Text">' +$CompanyCode + '</Value>
				</Eq>
				<Eq>
				<FieldRef Name="Sold_x0020_To_x0020_Code" /><Value Type="Text">' +$SoldToCode + '</Value>
				</Eq>
			</And>
			<And>
				<Eq>
				<FieldRef Name="Tax_x0020_Exempt_x0020_State" /><Value Type="Text">' +$TaxExemptState + '</Value>
				</Eq>
				<Eq>
				<FieldRef Name="Valid_x0020_From" /><Value Type="Datetime" IncludeTimeValue="FALSE">' +$ValidFrom.ToString("yyyy-MM-ddTHH:mm:ssZ") + '</Value>
				</Eq>
			</And>
		  </And>
		</Where>'
		
		 $spQuery.Query = $camlQuery
		 $spListItems = $spList.GetItems($spQuery)
         foreach($item in $splistItems)
		 {
					#Update SP list
					$item["SAP_x0020_Status"]=$SAPStatus
					$item["SAP_x0020_Message"]=$SAPMessage
					$item["SAP_x0020_Post_x0020_Time"]=$SAPPostTime
					
					#$item.Update()
                    $item.SystemUpdate()

					$UpdateToSP="No"
					#Update DB
                    if($SAPStatus -eq "Success"){
                        $UpdateToSP="Yes"
                    }else{
                        $UpdateToSP="Next"
                    }
					$sql="UPDATE [dbo].[TaxExemptCertificate] set [UpdateToSP]='"+$UpdateToSP+"' 
					WHERE [CompanyCode]='"+$CompanyCode+"' AND [SoldToCode]='"+$SoldToCode+"' AND [TaxExemptStates]='"+$TaxExemptState+"' AND [ValidFrom]='"+$ValidFrom+"'"
					#echo $sql
                    UpdateDBTable $sql
		 }	
           
	$cnt+=1
    }
    Write-Host 'Total changes: ' $cnt  
	
	if($spWeb)
    {
       $spWeb.Dispose()
    }
}

$tDataRecord = ReadDBTable "SELECT x.[companycode],
       x.[soldtocode],
       x.[taxexemptstates],
       x.[validfrom],
       'Fail'                 AS SAPStatus,
       Max( x.[sapposttime] ) AS [SAPPostTime],
       [SAPMessage]=Rtrim( Stuff( ( SELECT ', '
                                           + Concat(z.[shiptostate], ': ', z.[sapmessage]) AS [SAPMessage]
                                    FROM   [dbo].[TAXEXEMPTCERTIFICATE] z
                                    WHERE  x.[companycode] = z.[companycode] AND
                                           x.[soldtocode] = z.[soldtocode] AND
                                           x.[taxexemptstates] = z.[taxexemptstates] AND
                                           x.[validfrom] = z.[validfrom] AND
                                           z.[sapstatus] = 'Fail'
                                    FOR xml path(''), type ).value('.', 'NVARCHAR(MAX)'), 1, 1, '' ) )
FROM   [dbo].[TAXEXEMPTCERTIFICATE] x
WHERE  x.[updatetosp] <> 'Yes' AND
       x.[sapstatus] = 'Fail'
GROUP  BY [companycode],[soldtocode],[taxexemptstates],[validfrom]
UNION
SELECT a.[companycode],
       a.[soldtocode],
       a.[taxexemptstates],
       a.[validfrom],
       'Success' AS SAPStatus,
       a.[sapposttime],
       a.[sapmessage]
FROM   ( SELECT m.[companycode],
                m.[soldtocode],
                m.[taxexemptstates],
                m.[validfrom],
                Max( m.[sapposttime] ) AS [SAPPostTime],
                [SAPMessage]=Rtrim( Stuff( ( SELECT ', ' + ( CASE
                                                               WHEN n.[sapmessage] IS NULL THEN NULL
                                                               ELSE Concat( n.[shiptostate], ': ', n.[sapmessage] )
                                                             END ) AS [SAPMessage]
                                             FROM   [dbo].[TAXEXEMPTCERTIFICATE] n
                                             WHERE  m.[companycode] = n.[companycode] AND
                                                    m.[soldtocode] = n.[soldtocode] AND
                                                    m.[taxexemptstates] = n.[taxexemptstates] AND
                                                    m.[validfrom] = n.[validfrom] AND
                                                    n.[sapstatus] = 'Success'
                                             FOR xml path(''), type ).value('.', 'NVARCHAR(MAX)'), 1, 1, '' ) )
         FROM   [dbo].[TAXEXEMPTCERTIFICATE] m
         WHERE  m.[updatetosp] <> 'Yes' AND
                m.[sapstatus] = 'Success'
         GROUP  BY m.[companycode],m.[soldtocode],m.[taxexemptstates],m.[validfrom] ) a
       LEFT JOIN ( SELECT [companycode],
                          [soldtocode],
                          [taxexemptstates],
                          [validfrom]
                   FROM   [dbo].[TAXEXEMPTCERTIFICATE]
                   WHERE  [updatetosp] <> 'Yes' AND
                          [sapstatus] = 'Fail'
                   GROUP  BY [companycode],[soldtocode],[taxexemptstates],[validfrom] ) b
              ON a.[companycode] = b.[companycode] AND
                 a.[soldtocode] = b.[soldtocode] AND
                 a.[taxexemptstates] = b.[taxexemptstates] AND
                 a.[validfrom] = b.[validfrom]
WHERE  b.[soldtocode] IS NULL";

UpdateDataToResaleCertificateList $tDataRecord