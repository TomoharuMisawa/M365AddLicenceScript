# Office365で利用可能なライセンスを付加するモジュール

## ファイルを開く関数
## 利用例
## $LicenceCSVPath =  OpenFileDialog
## $Licencearray = Import-CSV $LicenceCSVPath
function OpenFileDialog()
{
    Add-Type -AssemblyName System.Windows.Forms

    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.InitialDirectory = "."
    $OpenFileDialog.filter = "CSV files (*.csv)| *.csv"
    $OpenFileDialog.Title = "CSVファイルを選択してください" 
    $ret = $OpenFileDialog.ShowDialog() 
    if($ret -eq "OK"){ 
        return $OpenFileDialog.FileName
    }
    return ""
}


#Install-Module -Name MSOnline
#Import-Module MSOnline


Try 
{
#    Connect-MsolService -ErrorAction Stop
} 
Catch 
{
    Write-Error -Message "接続エラーです。理由 $_"  -ErrorAction Stop
}


# SetExLicence1_テナント提供サービス一覧取得のファイルを読み込み

Write-Host "＊＊＊　ユーザーへライセンスを付与または変更します　＊＊＊"
Write-Host "＊＊＊　設定するライセンス情報一覧ファイル（CSV）を選択してください　＊＊＊"

#入力ファイル
#$LicenceCSVPath =  OpenFileDialog
$Licencearray = @(Import-CSV $LicenceCSVPath)



Write-Host "＊＊＊　設定するユーザー情報一覧ファイル（CSV）を選択してください　＊＊＊"

#入力ファイル
#$CSVPath =  OpenFileDialog

# サービス一覧を突き合わせ、読んだファイルに無いサービスの一覧を作って保持（除外サービスの確定）
# サービスを全取得
$SKUList = Get-MsolAccountSku
$array_outputspled = New-Object System.Collections.ArrayList

# サービス一覧の整形
foreach($sku in $SKUList) 
{
    foreach($serviceObj in $sku.ServiceStatus)
    {
        $array_outputrow = New-Object System.Collections.ArrayList
        $array_outputrow.AddRange(@($sku.AccountSkuId, $serviceObj.ServicePlan.ServiceName, $serviceObj.ServicePlan.TargetClass, $serviceObj.ServicePlan.ServiceType) ) > $null
        $array_outputspled.Add($array_outputrow) > $null
        $serviceObj = $null
    }
    $sku = $null
}

# サービス一覧から削除するものを抽出
#for($i = 0; $i -lt $array_outputspled.Count; $i++)
foreach($disablecheck in $array_outputspled)
{

    # 5つ目の入れもの(利用か除外か)をつくる
    $disablecheck.Add($false) > $null

    #foreach($disablecheck in $Licencearray)
    for($i = 0; $i -lt $Licencearray.Count; $i++)
    {

        if(($disablecheck[0] -eq $Licencearray[$i].AccountSkuId) -and
        ($disablecheck[1] -eq $Licencearray[$i].ServiceName) -and
        ($disablecheck[2] -eq $Licencearray[$i].TargetClass) -and
        ($disablecheck[3] -eq $Licencearray[$i].ServiceType) )
        {
            $disablecheck[4] = $true
            break
        }
    }
    $disablecheck = $null
}
#再生成
$newLicencearray = $array_outputspled -ne $null
#ライセンスのカスタマイズ(OFFにするライセンス一覧を作る）
Write-Host $newLicencearray
$disableLicenceHash = New-Object "System.Collections.Generic.Dictionary[string, string]"

# 全ライセンス情報をまわして確認
foreach($adLicence in $newLicencearray)
{
    if(-not $disableLicenceHash.ContainsKey($adLicence[0]))
    {
        $disableLicenceHash.Add($adLicence[0], "") > $null
    }
    
    # 使わないライセンス（5番目の項がfalse）のものを抽出
    if($adLicence[4]) 
    { 
        $adLicence = $null
        continue 
    }
    else
    {
        if($disableLicenceHash[$adLicence[0]].Length -eq 0 )
        {
            $disableLicenceHash[$adLicence[0]] += $adLicence[1]
        }
        else
        {
            $disableLicenceHash[$adLicence[0]] += ", " + $adLicence[1]
        }
        $adLicence = $null
    }
}

# ここで全除外ライセンスの元側を抜く。
foreach($al in $SKUList)
{
    ## AccountSku
    $checkstr = $disableLicenceHash[$al.AccountSkuId]
    $num = if($checkstr.Length -eq 0 ) { 0 } else { $checkstr.split(",").Count }

    # 除外サービスと提供サービスの量が一致した場合はライセンス適用を行なわないように調整
    if($num -eq $al.ServiceStatus.Count)
    {
        $disableLicenceHash.Remove($al.AccountSkuId) > $null
    }
    # 適用除外の確認(ユーザー以外のライセンスなど）SKUレベル
    elseif($al.TargetClass -eq "Tenant")
    {
        $disableLicenceHash.Remove($al.AccountSkuId) > $null
    }
    else
    {
        # 適用除外の確認(ユーザー以外のライセンスなど）機能レベル
        foreach($func in $al.ServiceStatus)
        {
            if($func.ServicePlan.TargetClass -eq "Tenant")
            {
                $checkstr = $checkstr.Replace($func.ServicePlan.ServiceName, "")
                $checkstr = $checkstr.Trim(",")
                $checkstr = $checkstr.Trim(" ")
            }
        }
        $disableLicenceHash[$al.AccountSkuId] = $checkstr
    }
    $checkstr = $null
    $num = $null
    $al = $null
}

# ライセンスの適用
foreach ($key in $disableLicenceHash.Keys){
    $LicenseString = $key.ToString()
    $disableOption = New-MsolLicenseOptions -AccountSkuId $LicenseString -DisabledPlans $disableLicenceHash[$key]

    #固定の設定値
    $UsageLocation = "JP" #ユーザーの地域
    #出力フォルダ
    $date = Get-Date -Format "yyyyMMddHHmm"
    $OutputFolder = [System.IO.Path]::GetDirectoryName($CSVPath)
    $TranscriptPath =  [System.IO.Path]::GetDirectoryName($CSVPath)+"\$date-log-userLicense.txt"

    #######################################
    write "ユーザーへライセンスを付与または変更します"
    #######################################

    @(Import-CSV $CSVPath) | % {

    $UserLicense = Get-MsolUser -UserPrincipalName　$_.UserPrincipalName;
    
    # ライセンスがついているか確認無ければ新規ユーザーとする
    $uselicense = $false
    foreach($li in $UserLicense.Licenses)
    {
        if($li.AccountSkuId -eq $LicenseString)
        {
            $uselicense = $true
            break
        }
    }

    if(-not $uselicense)
    {
        
        #新規ユーザーの場合       
        Write-Host $_.UserPrincipalName;
        Write-Host "新規ユーザー";

        #ロケーション設定
        Set-MsolUser `
            -UserPrincipalName $_.UserPrincipalName `
            -UsageLocation $UsageLocation;
    
        #ライセンス付与        
        write-host $LicenseString

        if($disableLicenceHash[$key].Length -eq 0)
        {
            $disableplan = @()
            $x = New-MsolLicenseOptions -AccountSkuId $LicenseString -DisabledPlans $disableplan
            Set-MsolUserLicense `
                -UserPrincipalName $_.UserPrincipalName `
                -AddLicenses $LicenseString `
                -LicenseOptions $x;
        }
        else
        {
            Set-MsolUserLicense `
                -UserPrincipalName $_.UserPrincipalName `
                -AddLicenses $LicenseString `
                -LicenseOptions $disableOption;
        }
    }else{
    #既存ユーザーの場合
        Write-Host $_.UserPrincipalName; 
        
        Write-Host "既存ユーザー";

       #ライセンスオプション変更
        write-host $LicenseString

        if($disableLicenceHash[$key].Length -eq 0)
        {
            $disableplan = @()
            $x = New-MsolLicenseOptions -AccountSkuId $LicenseString -DisabledPlans $disableplan
            Set-MsolUserlicense `
                -UserPrincipalName $_.UserPrincipalName `
                -LicenseOptions $x;
        }
        else
        {
            Set-MsolUserlicense `
                -UserPrincipalName $_.UserPrincipalName `
                -LicenseOptions $disableOption;
        }

    }
}


# 待機
Write-Host "＊＊＊　反映まで60秒お待ちください　＊＊＊"
Start-Sleep -s 1


#結果を取得
Write-Host "＊＊＊　ライセンスのログを出力中　＊＊＊"

    $skuList = @();
    @(Import-CSV $CSVPath) | % {

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
