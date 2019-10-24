<#
.Synopsis
   Office 365 Exchange ���L���[���{�b�N�X�쐬�X�N���v�g

.DESCRIPTION
   Office 365 Exchange �ɋ��L���[���{�b�N�X���쐬����B

   ���L�ݒ���s���B
   �E���L���[�����[�U�̃T�C���C���𖳌���
   �E�A�[�J�C�u��L����
   �E�i�׃z�[���h��L����
   �ESharingPolicy: XXXX Sharing Policy
   �ERetentionPolicy: XXXX MRM Policy
   �EMessageCopyForSentAsEnabled: True
   �EMessageCopyForSendOnBehalfEnabled: True

.LINK

#>

#�ϐ�
#******************************************************************************************************
$location = "JP"
$current = Split-Path $myInvocation.MyCommand.path
$nowdate = (Get-Date).ToString("yyyyMMdd")
$outputfile = $current + "\LOG\o365_exch_sharedmailbox_" + $nowdate + ".log"
$pwfile = $current + "\password.enc"
$errflg = 0
#******************************************************************************************************

#������
Write-Host ////////////////////////////////////////////////////////////////////////////////
$dn = Read-Host " ���L���[���{�b�N�X������͂��Ă�������"
Write-Host ////////////////////////////////////////////////////////////////////////////////
$mail = Read-Host " ���L���[���A�h���X����͂��Ă�������"

function Write-Log($comment,$flag){
    $comment = (Get-Date).ToString("yyyy/MM/dd HH:mm")+"`t"+$comment
    if($flag -eq $null){
        Write-Host $comment
    }
    Write-Output $comment | Out-File -Encoding utf8 -FilePath $outputfile -Append
}

Write-Log "�����J�n" 

#Office365 �֐ڑ�����
$UserName = "username"
$securePassWord = Get-Content $pwfile | ConvertTo-SecureString
$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePassWord)
$Password = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($BSTR)

#Office365�֐ڑ�-----------------------------------------------------------------------------
do{
    Write-Log "Office 365 �ڑ���"
    Start-Sleep -Seconds 2
    $UserCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $UserName,(ConvertTo-SecureString  $Password -AsPlainText -Force)
    Import-Module MSOnline
    $con_ret = Connect-MsolService -Credential $UserCredential 2>&1
}until($con_ret -eq $null)
Write-Log "Office 365 �ڑ�����"

#Exchange�֐ڑ�-----------------------------------------------------------------------------
do{
    Write-Log "Exchange �ڑ���"
    Start-Sleep -Seconds 2
    $UserCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $UserName,(ConvertTo-SecureString  $Password -AsPlainText -Force)
    $proxyOptions = New-PSSessionOption -ProxyAccessType IEConfig -ProxyAuthentication Negotiate -SkipRevocationCheck
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection -SessionOption $proxyOptions
    $con_ret = Import-PSSession $Session -DisableNameChecking 2>&1
}until($con_ret -ne $null)
Write-Log "Exchange �ڑ�����"

#���L���[���{�b�N�X�쐬-----------------------------------------------------------------------------
New-Mailbox -Shared -Name $dn -Alias $mail.Split("@")[0] -PrimarySmtpAddress $mail

#3���ҋ@
$timer = 180
do{
    Write-Progress "���L���[���{�b�N�X�쐬" -SecondsRemaining $timer
    Start-Sleep -Seconds 1
    $timer -= 1
}until($timer -lt 0)

#UPN�擾-----------------------------------------------------------------------------
Start-Sleep -Seconds 2
$upn = (Get-MailBox -Identity $mail.Split("@")[0]).UserPrincipalName

#���C�Z���X����
$e3AccountSku = Get-MsolAccountSku | ?{$_.AccountSkuId -Match "ENTERPRISEPACK"}
$e3ServiceName = ($e3AccountSku | Select-Object -ExpandProperty ServiceStatus | Select-Object -ExpandProperty ServicePlan).ServiceName
$ServiceName = $e3ServiceName | Select-String -Pattern "EXCHANGE_S_ENTERPRISE" -NotMatch
$LicenseOption =  New-MsolLicenseOptions -AccountSkuId $e3AccountSku.AccountSkuId -DisabledPlans $ServiceName

#���P�[�V�����ݒ�
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

#���C�Z���X�t�^
Start-Sleep -Seconds 2
Set-MsolUserLicense -UserPrincipalName $upn -LicenseOptions $LicenseOption -AddLicenses $e3AccountSku.AccountSkuId

#3���ҋ@
$timer = 180
do{
    Write-Progress "���C�Z���X���蓖��" -SecondsRemaining $timer
    Start-Sleep -Seconds 1
    $timer -= 1
}until($timer -lt 0)

#�T�C���C�����u���b�N
Start-Sleep -Seconds 2
Set-Msoluser -UserPrincipalName $upn -BlockCredential $true

#�|���V�[�ݒ�
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

#�A�[�J�C�u�L��
Start-Sleep -Seconds 2
Enable-MailBox -Identity $mail.Split("@")[0] -Archive
Start-Sleep -Seconds 2
Enable-MailBox -Identity $mail.Split("@")[0] -AutoExpandingArchive

#�i�׃z�[���h/���M�ς݃A�C�e���L��
Set-Mailbox -Identity $mail.Split("@")[0] -LitigationHoldEnabled $true -MessageCopyForSentAsEnabled $true -MessageCopyForSendOnBehalfEnabled $true

#1���ҋ@
$timer = 60
do{
    Write-Progress "�ݒ蔽�f" -SecondsRemaining $timer
    Start-Sleep -Seconds 1
    $timer -= 1
}until($timer -lt 0)

$msoluser = Get-Msoluser -UserPrincipalName $upn | Select BlockCredential,UsageLocation
$mailbox = Get-MailBox -Identity $mail.Split("@")[0] | Select PrimarySmtpAddress,ArchiveStatus,AutoExpandingArchiveEnabled,LitigationHoldEnabled,SharingPolicy,RetentionPolicy,MessageCopyForSentAsEnabled,MessageCopyForSendOnBehalfEnabled
Write-Log $msoluser False
Write-Log $mailbox False

#���ʕ\��
Write-Host ////////////////////////////////////////////////////////////////////////////////
Write-Host  ���s����:
Write-Host ////////////////////////////////////////////////////////////////////////////////
$msoluser
$mailbox
Write-Host ////////////////////////////////////////////////////////////////////////////////

#O365/ExchangeOnline�ɐؒf
Remove-PSSession $Session

Write-Log "�����I��"

Pause
