<#
.Synopsis
   Office 365 Exchange 共有メールボックス作成スクリプト

.DESCRIPTION
   Office 365 Exchange に共有メールボックスを作成する。

   下記設定を行う。
   ・共有メールユーザのサインインを無効化
   ・アーカイブを有効化
   ・訴訟ホールドを有効化
   ・SharingPolicy: XXXX Sharing Policy
   ・RetentionPolicy: XXXX MRM Policy
   ・MessageCopyForSentAsEnabled: True
   ・MessageCopyForSendOnBehalfEnabled: True

.LINK

#>

#変数
#******************************************************************************************************
$location = "JP"
$current = Split-Path $myInvocation.MyCommand.path
$nowdate = (Get-Date).ToString("yyyyMMdd")
$outputfile = $current + "\LOG\o365_exch_sharedmailbox_" + $nowdate + ".log"
$pwfile = $current + "\password.enc"
$errflg = 0
#******************************************************************************************************

#情報入力
Write-Host ////////////////////////////////////////////////////////////////////////////////
$dn = Read-Host " 共有メールボックス名を入力してください"
Write-Host ////////////////////////////////////////////////////////////////////////////////
$mail = Read-Host " 共有メールアドレスを入力してください"

function Write-Log($comment,$flag){
    $comment = (Get-Date).ToString("yyyy/MM/dd HH:mm")+"`t"+$comment
    if($flag -eq $null){
        Write-Host $comment
    }
    Write-Output $comment | Out-File -Encoding utf8 -FilePath $outputfile -Append
}

Write-Log "処理開始" 

#Office365 へ接続する
$UserName = "username"
$securePassWord = Get-Content $pwfile | ConvertTo-SecureString
$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePassWord)
$Password = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($BSTR)

#Office365へ接続-----------------------------------------------------------------------------
do{
    Write-Log "Office 365 接続中"
    Start-Sleep -Seconds 2
    $UserCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $UserName,(ConvertTo-SecureString  $Password -AsPlainText -Force)
    Import-Module MSOnline
    $con_ret = Connect-MsolService -Credential $UserCredential 2>&1
}until($con_ret -eq $null)
Write-Log "Office 365 接続成功"

#Exchangeへ接続-----------------------------------------------------------------------------
do{
    Write-Log "Exchange 接続中"
    Start-Sleep -Seconds 2
    $UserCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $UserName,(ConvertTo-SecureString  $Password -AsPlainText -Force)
    $proxyOptions = New-PSSessionOption -ProxyAccessType IEConfig -ProxyAuthentication Negotiate -SkipRevocationCheck
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection -SessionOption $proxyOptions
    $con_ret = Import-PSSession $Session -DisableNameChecking 2>&1
}until($con_ret -ne $null)
Write-Log "Exchange 接続成功"

#共有メールボックス作成-----------------------------------------------------------------------------
New-Mailbox -Shared -Name $dn -Alias $mail.Split("@")[0] -PrimarySmtpAddress $mail

#3分待機
$timer = 180
do{
    Write-Progress "共有メールボックス作成" -SecondsRemaining $timer
    Start-Sleep -Seconds 1
    $timer -= 1
}until($timer -lt 0)

#UPN取得-----------------------------------------------------------------------------
Start-Sleep -Seconds 2
$upn = (Get-MailBox -Identity $mail.Split("@")[0]).UserPrincipalName

#ライセンス生成
$e3AccountSku = Get-MsolAccountSku | ?{$_.AccountSkuId -Match "ENTERPRISEPACK"}
$e3ServiceName = ($e3AccountSku | Select-Object -ExpandProperty ServiceStatus | Select-Object -ExpandProperty ServicePlan).ServiceName
$ServiceName = $e3ServiceName | Select-String -Pattern "EXCHANGE_S_ENTERPRISE" -NotMatch
$LicenseOption =  New-MsolLicenseOptions -AccountSkuId $e3AccountSku.AccountSkuId -DisabledPlans $ServiceName

#ロケーション設定
switch($mail.Split("@")[1]){
    "XXXX.sg"{
        $location = "SG"
    }
    "XXXX.th"{
        $location = "TH"
    }
    default{
        $location = "JP"
    }
}
Start-Sleep -Seconds 2
Set-MsolUser -UserPrincipalName $upn -UsageLocation $location

#ライセンス付与
Start-Sleep -Seconds 2
Set-MsolUserLicense -UserPrincipalName $upn -LicenseOptions $LicenseOption -AddLicenses $e3AccountSku.AccountSkuId

#3分待機
$timer = 180
do{
    Write-Progress "ライセンス割り当て" -SecondsRemaining $timer
    Start-Sleep -Seconds 1
    $timer -= 1
}until($timer -lt 0)

#サインインをブロック
Start-Sleep -Seconds 2
Set-Msoluser -UserPrincipalName $upn -BlockCredential $true

#ポリシー設定
switch($mail.Split("@")[1]){
    "sumitomo-pharma.com.sg"{
        Start-Sleep -Seconds 2
        Set-Mailbox -Identity $mail.Split("@")[0] -SharingPolicy "XXXX Sharing Policy" -RetentionPolicy "XXXX MRM Policy"
    }
    "sumitomo-pharma.co.th"{
        Start-Sleep -Seconds 2
        Set-Mailbox -Identity $mail.Split("@")[0] -SharingPolicy "XXXX Sharing Policy" -RetentionPolicy "XXXX MRM Policy"
    }
    default{
        #Start-Sleep -Seconds 2
        #Set-Mailbox -Identity $mail.Split("@")[0] -SharingPolicy "XXXXXXXX Policy" -RetentionPolicy "XXXXXXXX Policy"
    }
}

#アーカイブ有効
Start-Sleep -Seconds 2
Enable-MailBox -Identity $mail.Split("@")[0] -Archive
Start-Sleep -Seconds 2
Enable-MailBox -Identity $mail.Split("@")[0] -AutoExpandingArchive

#訴訟ホールド/送信済みアイテム有効
Set-Mailbox -Identity $mail.Split("@")[0] -LitigationHoldEnabled $true -MessageCopyForSentAsEnabled $true -MessageCopyForSendOnBehalfEnabled $true

#1分待機
$timer = 60
do{
    Write-Progress "設定反映" -SecondsRemaining $timer
    Start-Sleep -Seconds 1
    $timer -= 1
}until($timer -lt 0)

$msoluser = Get-Msoluser -UserPrincipalName $upn | Select BlockCredential,UsageLocation
$mailbox = Get-MailBox -Identity $mail.Split("@")[0] | Select PrimarySmtpAddress,ArchiveStatus,AutoExpandingArchiveEnabled,LitigationHoldEnabled,SharingPolicy,RetentionPolicy,MessageCopyForSentAsEnabled,MessageCopyForSendOnBehalfEnabled
Write-Log $msoluser False
Write-Log $mailbox False

#結果表示
Write-Host ////////////////////////////////////////////////////////////////////////////////
Write-Host  実行結果:
Write-Host ////////////////////////////////////////////////////////////////////////////////
$msoluser
$mailbox
Write-Host ////////////////////////////////////////////////////////////////////////////////

#O365/ExchangeOnlineに切断
Remove-PSSession $Session

Write-Log "処理終了"

Pause
