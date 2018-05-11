# Config the task command line
# powershell.exe -executionPolicy bypass -file C:\app\scripts\PowerShellTrigger\CompareSAPandSharePointData.ps1

if ((Get-PSSnapin -Name Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue) -eq $null )
{
    Add-PsSnapin Microsoft.SharePoint.PowerShell
}

#Database config
$dataSource = "Database server name"
$database = "Database name"
$connectionString = "Server=$dataSource;Database=$database;Trusted_Connection=True;"
#SharePoint config
$web = Get-SPWeb http://yourSharePoint.com/sites/MyWeb
$fobList = $web.lists['List name']
$fobUpdatesList = $web.Lists['List name']


function ReadDBTable($sql){
    $res = @()
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

function RetrieveDBRow($title, $table, $mapping){
    $rows = $table.Select("MATNR = '$title'");
    if($rows.Count -eq 0){
        return @{};
    }

    $r0 = $rows[0];
    $res = @{};
    foreach($k in $mapping.keys){
        $dbField = $mapping[$k]
        $dbValue = $r0[$dbField]
        if($dbValue -eq [System.DBNull]::Value){
            $dbValue = $null
        }elseif($dbValue -eq '99991231'){
            $dbValue = '89001231';
        }

        
        if($dbField -eq 'effectiveDate' -or $dbField -eq 'expiredDate'){
            $dbValue = [DateTime]::ParseExact($dbValue,'yyyyMMdd',$null)
        }
        if($dbField -eq 'FRA1' -or $dbField -eq 'ZOA1'){
            $dbValue = $dbValue/100;
        }
        if($dbField -eq 'DEL_FLAG' -and $dbValue -eq 'X'){
            $dbValue = $true
        }elseif($dbField -eq 'DEL_FLAG' -and $dbValue -ne 'X'){
            $dbValue = $false
        }
        
        $res.Add($k, $dbValue);
    }

    return $res;
}

function RetrieveSPRow($item, $mapping){
    $res = @{}
    foreach($k in $mapping.keys){
        $res.Add($k, $item[$k]);
        #Write-Host 'SP items:' $k -ForegroundColor green
        #Write-Host 'SP items:' $item[$k] -ForegroundColor yellow
    }

    return $res;
}

function Map2StringDB($map){
    $str = '';
    foreach($k in $map.Keys){
        if ($k -eq 'FOB Status'){
            $FOBStatus= $k + '  :  ' + $map[$k] + [System.Environment]::NewLine;
        }else{
            $str+= $k + '  :  ' + $map[$k] + [System.Environment]::NewLine;
        }
    }
    if ($FOBStatus -like 'FOB Status*'){
        $str+=$FOBStatus
    }
    return $str;
}

function Map2StringSP($map){
    $str = '';
    foreach($k in $map.Keys){
        if ($k -eq 'FOB Status'){
            $FOBStatus= $k + '  :  ' + $map[$k] + [System.Environment]::NewLine;
        }else{
            $str+= $k + '  :  ' + $map[$k] + [System.Environment]::NewLine;
        }
    }
    return $str;
}

function LogChanges($id, $sp, $db){
    $spStr = Map2StringSP $sp
    $dbStr = Map2StringDB $db
    #Write-Host $id ' --------- '
    #write-host $spStr
    #write-host $dbStr
    #return;

    $item = $fobList.GetItemById($id)
    
    $Duty=$db['Duty%']
    $FOBStatus=$db['FOB Status']
    $EffectDate=$db['Effective Date']
    $DelFlag=$db['DEL FLAG']
    $SapProfiterCenter=$db['sapprofitcenter']
    $DPCSAPStatus=$db['DPCSAPStatus']
    $SAPInfoRecord  =$db['SAPInfoRecord']
    $ExpiredDate =$db['Expired Date']
    $SAPDesc =$db['SAP Desc']
    $Freight =$db['Freight%']

    $item['Duty_x0025_']= $Duty
    $item['Effective_x0020_Date']= $EffectDate
    if ($DelFlag -ne $true){
        $DelFlag=0
    }else{
        $DelFlag=1
    }
    $item['DEL_x0020_FLAG']= $DelFlag
    if($DelFlag -eq 1){
        $item["Life_x0020_Cycles"]="Obsolete"
    }
    $item['sapprofitcenter']= $SapProfiterCenter
    $item['DPCSAPStatus']= $DPCSAPStatus
    $item['SAPInfoRecord']= $SAPInfoRecord
    $item['Expired_x0020_Date']= $ExpiredDate
    $item['SAP_x0020_Desc']= $SAPDesc
    $item['Freight_x0025_']= $Freight
    $item['FOB_x0020_Status']= $FOBStatus
    $item.Update();


    $ni = $fobUpdatesList.AddItem();
    #Write-Host 'SP items:' $spStr -ForegroundColor green
    #Write-Host 'SP items:' $dbStr -ForegroundColor yellow
    $ni['Title'] = $item.Title;
    $ni['ItemID'] = $id;
    $ni['Old Value'] = $spStr;
    $ni['New Value'] = $dbStr;
    $ni.Update();
    
}

function MapEqual($m1, $m2){
    
    #Write-Host 'Total changes: ' $m1.Count   $m2.Count
    if($m1 -eq $null -or $m2 -eq $null){ return $false;}
    if($m1.Count -ne $m2.Count){ return $false;}

    foreach($k in $m1.Keys){
    #Write-Host 'Total changes: ' $m1[$k]  -foregroundcolor green
    #Write-Host 'Total changes: ' $m2[$k]  -ForegroundColor Yellow
        if(-not $m2.ContainsKey($k)){
            return $false;
        }
        if($m1[$k] -ne $m2[$k]){
            return $false;
        }
    }
    
    return $true;
}


function CheckFOBPriceItems($tVInfoRecordLatest,  $tYMARC){
    $mappingVInfoRecordLatest = @{'Expired Date' = 'expiredDate'; 'Effective Date'='effectiveDate';'Freight%'='FRA1'; 'Duty%'='ZOA1';'SAPInfoRecord'='Rate';'SAP Desc'='description'};
    $mappingYMARC = @{'DEL FLAG'='DEL_FLAG';'sapprofitcenter'='PRCTR'};
    
    $cnt = 0;
    foreach($item in $fobList.Items){
        $dbVInforRecordLatest = RetrieveDBRow $item['Title'] $tVInfoRecordLatest $mappingVInfoRecordLatest
        $spVInforRecordLatest = @{};
        if($dbVInforRecordLatest.Count -gt 0){
            $spVInforRecordLatest = RetrieveSPRow $item $mappingVInfoRecordLatest
        }
        
        $dbYMarc = RetrieveDBRow $item['Title'] $tYMARC $mappingYMARC
        $spYMarc = @{}
        if($dbYMarc.Count -gt 0){
            $dbYMarc.Add('DPCSAPStatus', 'IN DPC SAP');
            $dbYMarc.Add('FOB Status', 'Completed');
            
            $spYMarc = RetrieveSPRow $item $mappingYMARC
            $spYMarc.Add('DPCSAPStatus', $item['DPCSAPStatus']);
            #only for comparison to add 'FOB Status' item for old value
            $spYMarc.Add('FOB Status', 'Completed');
        }

        $dbProp = $dbVInforRecordLatest + $dbYMarc
        $spProp = $spVInforRecordLatest + $spYMarc

        if($dbProp.Count -gt 0 -and (-not (MapEqual $dbProp $spProp))){
            LogChanges $item.ID $spProp $dbProp
            $cnt++
        }
    }
    Write-Host 'Total changes: ' $cnt
}

$tVInfoRecordLatest = ReadDBTable "select [MATNR],[expiredDate],[effectiveDate],[FRA1],[ZOA1],[Rate],[description] from dbo.VInfoRecordLatest where Vendor = '849120'";
$tVMARC = ReadDBTable "select [MATNR],[DEL_FLAG],[PRCTR] from dbo.YMARC where WERKS = 'DPC1'";

#$tVInfoRecordLatest = ReadDBTable "select [MATNR],[expiredDate],[effectiveDate],[FRA1],[ZOA1],[Rate],[description] from dbo.VInfoRecordLatest where Vendor = '849120' and MATNR='VFD550CP63A-21'";
#$tVMARC = ReadDBTable "select [MATNR],[DEL_FLAG],[PRCTR] from dbo.YMARC where WERKS = 'DPC1' and MATNR='VFD550CP63A-21'";

CheckFOBPriceItems $tVInfoRecordLatest $tVMARC