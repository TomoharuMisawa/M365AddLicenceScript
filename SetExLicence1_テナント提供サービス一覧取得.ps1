# Office365で利用可能なライセンス情報を出力するモジュール
Install-Module -Name MSOnline
Import-Module MSOnline


Try 
{
    Connect-MsolService -ErrorAction Stop
} 
Catch 
{
    Write-Error -Message "接続エラーです。理由 $_"  -ErrorAction Stop
}


$SKUList = Get-MsolAccountSku

    Write-Host "利用可能なライセンスの一覧表示"
    Write-Host "表示例："
    Write-Host "AccountSKU ID（テナント名:ライセンスの内部名称）"
    Write-Host "ServiceName",  "TargetClass",  "ServiceType" | Format-Wide
    Write-Host "-------------------------------------------------"
    Write-Host "利用可能なサービスの内部名、ターゲット（ユーザーに付与するか、テナントに付与するか）、サービスのタイプ"
    Write-Host ""

$array_outputspled = New-Object System.Collections.ArrayList
$array_outputheader = New-Object System.Collections.ArrayList
$array_outputheader.Add("""[AccountSkuId]"",""[ServiceName]"",""[TargetClass]"",""[ServiceType]""")
$array_outputspled.Add($array_outputheader)
foreach($sku in $SKUList) 
{
    
    Write-Host ""
    Write-Host $sku.AccountSkuId 
    Write-Host "ServiceName",  "TargetClass",  "ServiceType" | Format-Wide
    Write-Host "-------------------------------------------------"

    foreach($serviceObj in $sku.ServiceStatus)
    {
        $array_outputrow = New-Object System.Collections.ArrayList
        Write-Host $serviceObj.ServicePlan.ServiceName,  $serviceObj.ServicePlan.TargetClass,  $serviceObj.ServicePlan.ServiceType | Format-Wide 
        $array_outputrow.Add( """" + $sku.AccountSkuId + """,""" + $serviceObj.ServicePlan.ServiceName + """,""" + $serviceObj.ServicePlan.TargetClass + """,""" + $serviceObj.ServicePlan.ServiceType + """" )
        $array_outputspled.Add($array_outputrow)
    }
}

# ファイルへ保存
$array_outputspled | Out-File C:\Work\list.csv -Encoding UTF8 -Append 
# 終了

