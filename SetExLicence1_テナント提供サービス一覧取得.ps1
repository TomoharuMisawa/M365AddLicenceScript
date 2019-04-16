# Office365で利用可能なライセンス情報を出力するモジュール

## ファイル保存を開く関数
## 利用例
## $LicenceCSVPath =  SaveFileDialog
## Out-File $LicenceCSVPath -Encoding UTF8
function SaveFileDialog()
{
    Add-Type -AssemblyName System.Windows.Forms
    
    $dialog = New-Object System.Windows.Forms.SaveFileDialog
    $dialog.InitialDirectory = "."
    $dialog.Filter = "CSV files (*.csv)| *.csv" 
    $dialog.Title = "保存ファイルを指定してください" 
    $dialog.OverwritePrompt = $true
    
    # ダイアログを表示
    if($dialog.ShowDialog() -eq "OK")
    {      
      #入力されたファイル名を返す
      return $dialog.Filename
    }
}

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

#出力ファイル
Write-Host "＊＊＊　テナントでもっているライセンスを取得します　＊＊＊"
Write-Host "＊＊＊　一覧を保存するライセンス情報一覧ファイル名（CSV）を決定してください　＊＊＊"
$LicenceCSVPath =  SaveFileDialog

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
$array_outputheader.Add("""[AccountSkuId]"",""[ServiceName]"",""[TargetClass]"",""[ServiceType]""") > $null
$array_outputspled.Add($array_outputheader) > $null
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
        $array_outputrow.Add( """" + $sku.AccountSkuId + """,""" + $serviceObj.ServicePlan.ServiceName + """,""" + $serviceObj.ServicePlan.TargetClass + """,""" + $serviceObj.ServicePlan.ServiceType + """" ) > $null
        $array_outputspled.Add($array_outputrow) > $null
    }
}

# ファイルへ保存
$array_outputspled | Out-File $LicenceCSVPath -Encoding UTF8 -Append 
# 終了

