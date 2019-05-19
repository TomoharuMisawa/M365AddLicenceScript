# Microsoft 管理センターで利用可能なライセンスを付加するモジュール

## ファイルを開く関数
## 利用例
## $LicenceCSVPath =  OpenFileDialog
## $Licencearray = Import-CSV $LicenceCSVPath
## キャンセルした場合は空文字が応答される
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

# モジュールの読み込み
# すでに読み込み済みやインストール済みの場合は消すことも可能
Install-Module -Name MSOnline
Import-Module MSOnline

# 管理センターへ接続
# 接続するライセンスは管理者権限があること
Try 
{
    Connect-MsolService -ErrorAction Stop
} 
Catch 
{
    Write-Error -Message "接続エラーです。理由 $_"  -ErrorAction Stop
}


# ファイルを読み込み
Write-Host "＊＊＊　ユーザーへライセンスを付与または変更します　＊＊＊"
Write-Host "＊＊＊　設定するライセンス情報一覧ファイル（CSV）を選択してください　＊＊＊"
$LicenceCSVPath = OpenFileDialog
# ライセンスファイルはSetExLicence1_テナント提供サービス一覧取得のファイルを読み込む
# ヘッダ行あり。（"[AccountSkuId]"、"[ServiceName]"、"[TargetClass]"、"[ServiceType]"）
$Licencearray = @(Import-CSV $LicenceCSVPath -Header("[AccountSkuId]", "[ServiceName]", "[TargetClass]", "[ServiceType]"))
Write-Host "＊＊＊　設定するユーザー情報一覧ファイル（CSV）を選択してください　＊＊＊"
$CSVPath = OpenFileDialog

# サービス一覧を突き合わせ、読み込んだファイルに無いサービスの一覧を作って保持（除外サービスの確定）
# サービスを全取得
$SKUList = Get-MsolAccountSku
$array_outputspled = New-Object System.Collections.ArrayList

# サービス一覧の整形
# サービスの一覧から機能を取り出し、一覧化する
# $array_outputspledに格納する
foreach($sku in $SKUList) 
{
    foreach($serviceObj in $sku.ServiceStatus)
    {
        $array_outputrow = New-Object System.Collections.ArrayList
        $array_outputrow.AddRange(@($sku.AccountSkuId, $serviceObj.ServicePlan.ServiceName, $serviceObj.ServicePlan.TargetClass, $serviceObj.ServicePlan.ServiceType) ) > $null
        $array_outputspled.Add($array_outputrow) > $null
        $array_outputrow　= $null
        $serviceObj = $null
    }
    $sku = $null
}

# サービス一覧から削除するものを抽出
# すべてのライセンスを抽出したものが$array_outputspled。適用するライセンスを抽出した（ユーザーが選択した）ものが$Licencearray。
# PowerShellでライセンスを賦与する才は、除外したい機能を選択することとなるためここで変換を行なう。
foreach($disablecheck in $array_outputspled)
{

    # 5つ目の入れもの(利用か除外か)をつくる
    # 利用時はtrue、除外時はfalseとなる
    $disablecheck.Add($false) > $null

    # 除外リストと全体リストが一致したときは利用となる。trueにかえる。一致しなかった際は除外としfalseのままとする。
    foreach($lic in $Licencearray)
    {
        if(($disablecheck[0] -eq $Lic."[AccountSkuId]") -and
        ($disablecheck[1] -eq $Lic."[ServiceName]") -and
        ($disablecheck[2] -eq $Lic."[TargetClass]") -and
        ($disablecheck[3] -eq $Lic."[ServiceType]") )
        {
            $disablecheck[4] = $true
            break
        }
    }
    $disablecheck = $null
}

# ライセンスのカスタマイズ(OFFにするライセンス一覧を作る）
# 上記で$array_outputspledには不要機能、必要機能両方がある状態のリストができた。
# そこから、不要機能を除外してPowerShellでライセンス付与時に渡せる形の情報を作成する。
# 不要機能のみを抽出する
$disableLicenceHash = New-Object "System.Collections.Generic.Dictionary[string, string]"
# 全ライセンス情報をまわして確認
foreach($adLicence in $array_outputspled)
{
    if(-not $disableLicenceHash.ContainsKey($adLicence[0]))
    {
        # 初めてライセンスが出てきたときは、ライセンスの入れ物を作る（この時点では除外項目は無い状態となる）
        $disableLicenceHash.Add($adLicence[0], "") > $null
    }
    
    # 使わない機能（5番目の項がfalse）を抽出
    if(-not $adLicence[4]) 
    { 
        if($disableLicenceHash[$adLicence[0]].Length -eq 0 )
        {
            $disableLicenceHash[$adLicence[0]] += $adLicence[1]
        }
        else
        {
            $disableLicenceHash[$adLicence[0]] += ", " + $adLicence[1]
        }
    }
    $adLicence = $null
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
    # 適用除外の確認(ユーザー以外のライセンスなど）ライセンスレベル
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

#固定の設定値
$UsageLocation = "JP" #ユーザーの地域
#出力フォルダ
$date = Get-Date -Format "yyyyMMddHHmm"
$OutputFolder = [System.IO.Path]::GetDirectoryName($CSVPath)
$TranscriptPath =  [System.IO.Path]::GetDirectoryName($CSVPath)+"\$date-log-userLicense.txt"

foreach($sk in $SKUList)
{
    $licenseKey = $null
    $disableplans = $null
    $disableOption = $null
    # 適用除外の確認(ユーザー以外のライセンスなど）SKUレベル
    if($sk.TargetClass -eq "Tenant")
    {
        continue
    }
    if($disableLicenceHash.ContainsKey($sk.AccountSkuId))
    {
        #ライセンス付与の対象の場合、付与する
        $licenseKey = $sk.AccountSkuId
        # 無効化するオプションを決定する
        $disableplans = @()
        if($disableLicenceHash[$licenseKey].Length -gt 0)
        {
            $tempOptionsArray = $disableLicenceHash[$licenseKey].split(",").Trim()
            # ブランクのオプションを消す
            foreach($brankcheck in $tempOptionsArray)
            {
                if(-not($brankcheck -eq ""))
                {
                    $disableplans += $brankcheck
                }
            }
            $tempOptionsArray = $null
        }
        $disableOption = New-MsolLicenseOptions -AccountSkuId $licenseKey -DisabledPlans $disableplans
    
    }
    else
    {
        #ライセンス付与対象ではない場合、ユーザーからライセンスを剥奪する
        $licenseKey = $null
        $disableplans = $null
        $disableOption = $null
    }

    write "#######################################"
    write "ユーザーへ以下ライセンスを付与または変更、削除します"
    write $sk.AccountSkuId
    write "除外機能：" 
    write $disableOption.DisabledServicePlans
    write "#######################################"

    #ここまでに処理するライセンスが確定する
    #ライセンス処理するユーザーを読み込む
    @(Import-CSV $CSVPath) | % {

        $UserLicense = Get-MsolUser -UserPrincipalName　$_.UserPrincipalName;
        Write $_.UserPrincipalName; 

        # ライセンスが付与されているか確認。無ければ新規ユーザーとする。
        $uselicense = $false
        foreach($li in $UserLicense.Licenses)
        {
            if($li.AccountSkuId -eq $sk.AccountSkuId)
            {
                $uselicense = $true
                $li = $null
                break
            }
            $li = $null
        }

        # ライセンス削除のケース
        if($licenseKey -eq $null)
        {
            #　すでにライセンスが付与されていなければ何もしない
            # ライセンスがあれば削除する
            if($uselicense)
            {
                Write "ライセンス削除";
               #ライセンス削除
                Set-MsolUserlicense `
                    -UserPrincipalName $_.UserPrincipalName `
                    -RemoveLicenses　$sk.AccountSkuId;
                
            }
        }
        # ライセンスが付与されていない場合はsageLocationをJPにしてAddLicensesを実施する。
        elseif(-not $uselicense)
        {
        
            #新規ライセンスの場合
            Write "新規ライセンス付与";

            #ロケーション設定
            Set-MsolUser `
                -UserPrincipalName $_.UserPrincipalName `
                -UsageLocation $UsageLocation;
    
            #ライセンス付与        
            Set-MsolUserLicense `
                -UserPrincipalName $_.UserPrincipalName `
                -AddLicenses $licenseKey `
                -LicenseOptions $disableOption;

        }
        else
        {
        #既存ライセンスの場合
            Write "既存ライセンスの変更";

           #ライセンスオプション変更
            Set-MsolUserlicense `
                -UserPrincipalName $_.UserPrincipalName `
                -LicenseOptions $disableOption;

        }
    }
}

# 待機
Write-Host "＊＊＊　反映まで60秒お待ちください　＊＊＊"
Start-Sleep -s 60


# 結果を取得
# 現時点で付与されているライセンス、機能の一覧を出力する
# 出力先はユーザー情報一覧ファイル（CSV）のある場所になる（$CSVPath）
Write-Host "＊＊＊　ライセンスのログを出力中　＊＊＊"

$afterskuList = @();
@(Import-CSV $CSVPath) | % {

    Get-MsolUser -UserPrincipalName　$_.UserPrincipalName | % {

        $upn = $_.UserPrincipalName;
        $dpn = $_.DisplayName;

        $_.Licenses | % {
            $sku = $_.AccountSkuId;
            $_.ServiceStatus | ForEach-Object {
                $afterskuList += @{
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
  
$afterskuList | select `
    @{n="UserPrincipalName"; e={$_.UserPrincipalName}}, `
    @{n="DisplayName"; e={$_.DisplayName}}, `
    @{n="AccountSkuId"; e={$_.AccountSkuId}}, `
    @{n="ServiceName"; e={$_.ServiceName}}, `
    @{n="ProvisioningStatus"; e={$_.ProvisioningStatus}} | `
Export-Csv -NoTypeInformation -Encoding UTF8 $OutputFolder\$date-set-msoluserlicense.csv -Append


Write-Host "＊＊＊　ライセンスの割り当てが完了しました　＊＊＊"

# 終了
