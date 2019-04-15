# Office365で利用可能なライセンスを付加するモジュール
## ファイルを開く関数
## 利用例
## $LicenceCSVPath =  OpenFileDialog
## $Licencearray = Import-CSV $LicenceCSVPath
function OpenFileDialog()
{
    Add-Type -AssemblyName System.Windows.Forms

    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.InitialDirectory = "c:\"
    $OpenFileDialog.filter = "CSV files (*.csv)| *.csv"
    $OpenFileDialog.Title = "CSVファイルを選択してください" 
    $ret = $OpenFileDialog.ShowDialog() 
    if($ret -eq "OK"){ 
        return $OpenFileDialog.FileName
    }
    return ""
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


# SetExLicence1_テナント提供サービス一覧取得のファイルを読み込み

Write-Host "＊＊＊　ユーザーへライセンスを付与または変更します　＊＊＊"
Write-Host "＊＊＊　設定するライセンス情報一覧ファイル（CSV）を選択してください　＊＊＊"

#入力ファイル
$LicenceCSVPath =  OpenFileDialog
$Licencearray = Import-CSV $LicenceCSVPath



Write-Host "＊＊＊　設定するユーザー情報一覧ファイル（CSV）を選択してください　＊＊＊"

#入力ファイル
$CSVPath =  OpenFileDialog

# サービス一覧を突き合わせ、読んだファイルに無いサービスの一覧を作って保持（除外サービスの確定）

$SKUList = Get-MsolAccountSku
$array_outputspled = New-Object System.Collections.ArrayList

foreach($sku in $SKUList) 
{
    foreach($serviceObj in $sku.ServiceStatus)
    {
        $array_outputrow = New-Object System.Collections.ArrayList
        $array_outputrow.AddRange(@($sku.AccountSkuId, $serviceObj.ServicePlan.ServiceName, $serviceObj.ServicePlan.TargetClass, $serviceObj.ServicePlan.ServiceType) )
        $array_outputspled.Add($array_outputrow)
    }
}

for($i = 0; $i -lt $array_outputspled.Count; $i++)
{

    $checkflg = $false
    
    foreach($disablecheck in $Licencearray)
    {

        if(($array_outputspled[$i][0] -eq $disablecheck.AccountSkuId) -and
        ($array_outputspled[$i][1] -eq $disablecheck.ServiceName) -and
        ($array_outputspled[$i][2] -eq $disablecheck.TargetClass) -and
        ($array_outputspled[$i][3] -eq $disablecheck.ServiceType) )
        {

            $checkflg = $true
            break
        }

    }
    if($checkflg)
    {
        $array_outputspled[$i] = $null
    }
}

$newLicencearray = $array_outputspled -ne $null

#ライセンスのカスタマイズ(OFFにするライセンス）
Write-Host $newLicencearray
$disableLicenceHash = New-Object "System.Collections.Generic.Dictionary[string, string]"

foreach($adLicence in $newLicencearray)
{
    if($disableLicenceHash.ContainsKey($adLicence[0]))
    {
        $disableLicenceHash[$adLicence[0]] += ", " + $adLicence[1]
    }
    else 
    {
        $disableLicenceHash.Add($adLicence[0], $adLicence[1])
    }
}



foreach ($key in $disableLicenceHash.Keys){
    $License = $key
    $MyO365Sku ="New-MsolLicenseOptions -AccountSkuId " , $key , " -DisabledPlans ", $disableLicenceHash[$key]

    Write-Host $MyO365Sku

    #固定の設定値
    $UsageLocation = "JP" #ユーザーの地域
    #出力フォルダ
    $date = Get-Date -Format "yyyyMMddHHmm"
    $OutputFolder = [System.IO.Path]::GetDirectoryName($CSVPath)
    $TranscriptPath =  [System.IO.Path]::GetDirectoryName($CSVPath)+"\$date-log-userLicense.txt"

    #######################################
    write "ユーザーへライセンスを付与または変更します"
    #######################################

    Import-CSV $CSVPath | % {

    $UserLicense = Get-MsolUser -UserPrincipalName　$_.UserPrincipalName;

    if($UserLicense.IsLicensed -eq $False){
    #新規ユーザーの場合       
            Write-Host $_.UserPrincipalName;
        
            Write-Host "新規ユーザー";

        #ロケーション設定
            Set-MsolUser `
                -UserPrincipalName $_.UserPrincipalName `
                -UsageLocation $UsageLocation;
    
        #ライセンス付与        
            Set-MsolUserLicense `
                -UserPrincipalName $_.UserPrincipalName `
                -AddLicenses $License `
                -LicenseOptions $MyO365Sku
        }else{
    #既存ユーザーの場合
            Write-Host $_.UserPrincipalName; 
        
            Write-Host "既存ユーザー";

       #ライセンスオプション変更
            Set-MsolUserlicense `
                -UserPrincipalName $_.UserPrincipalName `
                -LicenseOptions $MyO365Sku
        }
    }


    # 待機
    Write-Host "＊＊＊　反映まで60秒お待ちください　＊＊＊"
    Start-Sleep -s 60


    #結果を取得
    Write-Host "＊＊＊　ライセンスのログを出力中　＊＊＊"

    $skuList = @();
    Import-CSV $CSVPath | % {

        Get-MsolUser -UserPrincipalName　$_.UserPrincipalName | % {

            $upn = $_.UserPrincipalName;
            $dpn = $_.DisplayName;

            $_.Licenses | % {
                $sku = $_.AccountSkuId;
                $_.ServiceStatus | ForEach-Object {
                    $skuList += @{
                    UserPrincipalName =$upn;
                    DisplayName = $dpn;
                    AccountSkuId = $sku;
                    ServiceName = $_.ServicePlan.ServiceName;
                    ProvisioningStatus = $_.ProvisioningStatus;
                    }
                } 
            }
        }
    }
  
    $skuList | select `
        @{n="UserPrincipalName"; e={$_.UserPrincipalName}}, `
        @{n="DisplayName"; e={$_.DisplayName}}, `
        @{n="AccountSkuId"; e={$_.AccountSkuId}}, `
        @{n="ServiceName"; e={$_.ServiceName}}, `
        @{n="ProvisioningStatus"; e={$_.ProvisioningStatus}} | `
    Export-Csv -NoTypeInformation -Encoding UTF8 $OutputFolder\$date-set-msoluserlicense.csv -Append


    Write-Host "＊＊＊　ライセンスの割り当てが完了しました　＊＊＊"

}


# 終了

