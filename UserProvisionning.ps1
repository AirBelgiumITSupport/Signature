####################################
########### PARAMETERS #############
####################################

$TEMPLATE_FOLDER = "C:\PowerShell\Signature\"  #Do not forget \ at the end
$TEMPLATE_WITH_MOBILE = $TEMPLATE_FOLDER + "AB-signature-email.html"
$USERS_DATA = "C:\PowerShell\Signature\users_infos.csv"
$USERS_DATA_CONVERTED = "C:\PowerShell\Signature\users_infos_UNICODE.csv"

####################################
#######LOGIN TO AZURE CRED##########
####################################

$usernameAzure = Read-host "Enter AD Azure admin login "
$passwordAzure = Read-host "Enter AD Azure admin password " -AsSecureString
#$usernameAzure = "admin.delhaye@airbelgium.com"
#$passwordAzure = "MYPASSWORD"|ConvertTo-SecureString -AsPlainText -Force

$username = Read-host "Enter AD local admin login ( airbelgium\firstname.lastname ) "
$password =  Read-host "Enter local AD admin password " -AsSecureString 
#$username = "airbelgium\firstname.lastname"
#$password = "MYPASSWORD"|ConvertTo-SecureString -AsPlainText -Force

####################################
#######END LOGIN TO AZURE ##########
####################################


#######CONNECT TO AZURE ##########
$AzureCredential = new-object -typename System.Management.Automation.PSCredential ` -argumentlist $usernameAzure, $passwordAzure
$credential = new-object -typename System.Management.Automation.PSCredential ` -argumentlist $username, $password

write-Host 'Connecting to AD Azure and AD Local using credentials' -ForegroundColor Yellow

Set-ExecutionPolicy RemoteSigned -force
Import-Module ActiveDirectory
#Connect to exchange Online
$exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $AzureCredential -Authentication "Basic" -AllowRedirection
Import-PSSession  $exchangeSession -DisableNameChecking -AllowClobber 

Import-Module SkypeOnlineConnector
#Import-Module LyncOnlineConnector
$sfboSession = New-CsOnlineSession -Credential $AzureCredential –OverrideAdminDomain "airbelgium.onmicrosoft.com"
#$sfbSession = New-CsOnlineSession -Credential $AzureCredential 
#Import-PSSession  $sfbSession
Import-PSSession  $sfboSession -DisableNameChecking -AllowClobber


write-Host 'Starting to update AD Users Attributes and signatures.......' -ForegroundColor Yellow

#Set encoding! é è à ë ï ç
####### CONVERS CSV TO UNICODE ##########
Get-Content $USERS_DATA | Set-Content $USERS_DATA_CONVERTED -encoding UNICODE
####### LOOP ON CSV USERS ##########
$users = Import-Csv -Path $USERS_DATA_CONVERTED -Encoding UNICODE
foreach ($user in $users) {
    $firstNameWeb = $($user."First Name").Trim().replace("é", "&eacute;").replace("è", "&egrave;").replace("ê", "&ecirc;").replace("ë", "&euml;").replace("ï", "&iuml;").replace("ç", "&ccedil;").replace("œ", "&oelig;").replace("â", "&acirc;")
    $lastnameWeb = $($user."Last Name").Trim().replace("é", "&eacute;").replace("è", "&egrave;").replace("ê", "&ecirc;").replace("ë", "&euml;").replace("ï", "&iuml;").replace("ç", "&ccedil;").replace("œ", "&oelig;").replace("â", "&acirc;")
    $firstName = $($user."First Name").Trim()
    $lastname = $($user."Last Name").Trim()
    $displayName = $firstName + " " + $lastname
    $JobTitle = $($user."Job Title").Trim()
    $Manager = $($user."Manager (N+1)").Trim()
    $MobilePhone = $($user."AB Mobile Phone").Trim()
    $BusinessPhone = $($user."Business Phone").Trim()
    $identityName = $($user."E-Mail").Trim()
    $SamAccountName = $($user."E-Mail").replace("@airbelgium.com", "")

    #Configure Hosted VoiceMail .... 
    Set-CsUser -HostedVoiceMail $True -Identity $identityName

    
    if ($SamAccountName -eq "olivier.corvilain"){
        $destinationFile = $TEMPLATE_FOLDER + "AB-Signature_olivi.html"
    }else{
        $destinationFile = $TEMPLATE_FOLDER + "AB-Signature_" + $SamAccountName + ".html"
    }

    $sourceFile = $TEMPLATE_WITH_MOBILE
    (Get-Content $sourceFile) | Set-Content $destinationFile -encoding UTF8

    $eol="`r`n"

    # Replace CR+LF with LF
    $text = [IO.File]::ReadAllText($destinationFile) -replace "`r`n", "`n"
    [IO.File]::WriteAllText($destinationFile, $text)

    # Replace CR with LF
    $text = [IO.File]::ReadAllText($destinationFile) -replace "`r", "`n"
    [IO.File]::WriteAllText($destinationFile, $text)

    #  At this point all line-endings should be LF.

    # Replace LF with intended EOL char
    if ($eol -ne "`n") {
      $text = [IO.File]::ReadAllText($destinationFile) -replace "`n", $eol
      [IO.File]::WriteAllText($destinationFile, $text)
    }




    if ($MobilePhone -eq '' -or $MobilePhone -eq '0'){

        if ($BusinessPhone -eq '' -or $BusinessPhone -eq '0'){
          (Get-Content $destinationFile) | ForEach-Object {
                $_.replace('<p style="font-weight:normal; font-size:9pt; font-family:Verdana, sans-serif!important; margin:0; text-decoration:none!important; color:#000000!important; position:relative; margin-top:5px">M. [[MOBILEPHONE]] &nbsp &nbsp T. [[OFFICEPHONE]]<br>', '')
            } | Set-Content $destinationFile
        }else{
          (Get-Content $destinationFile) | ForEach-Object {
                $_.replace('M. [[MOBILEPHONE]] &nbsp &nbsp ', '')
            } | Set-Content $destinationFile
        }
    }else{
       if ($BusinessPhone -eq '' -or $BusinessPhone -eq '0'){
          (Get-Content $destinationFile) | ForEach-Object {
                $_.replace(' &nbsp &nbsp T. [[OFFICEPHONE]]', '')
            } | Set-Content $destinationFile
        }
        else{
            $sourceFile = $destinationFile
        }
    }



    if ($SamAccountName -eq 'niky.terzakis'){
        $JobTitle = $JobTitle + '*'
        (Get-Content $destinationFile) | ForEach-Object {
                $_.replace('<p style="font-weight:normal; font-size:7pt; font-family:Verdana, sans-serif!important; margin:0; color:#727272;" >This message (including any attachments) contains confidential information intended for a specific individual and purpose, and is protected by law. If you are not the intended recipient, you should delete this message and are hereby notified that any disclosure, duplication, or distribution of this message, or the taking of any action based on it, is strictly prohibited.</p>', '<p style="font-weight:normal; font-size:7pt; font-family:Verdana, sans-serif!important; margin:0; color:#727272;" >This message (including any attachments) contains confidential information intended for a specific individual and purpose, and is protected by law. If you are not the intended recipient, you should delete this message and are hereby notified that any disclosure, duplication, or distribution of this message, or the taking of any action based on it, is strictly prohibited.<br /><br /></p><p style="font-weight:normal; font-size:7pt; font-family:Verdana, sans-serif!important; margin:0; color:#727272;" >(*) 3T Management & Associates SPRL appointed as Chief Executive Officer of Air Belgium SA, through Mr Niky Terzakis, its permanent representative<br /><br /></p><p style="font-weight:normal; font-size:7pt; font-family:Verdana, sans-serif!important; margin:0; color:#727272;" >(*) Air Belgium SA repr&eacute;sent&eacute;e par 3T Management & Associates SPRL comme Administrateur D&eacute;l&eacute;gu&eacute;, elle-m&ecirc;me repr&eacute;sent&eacute;e par Niky Terzakis g&eacute;rant et repr&eacute;sentant permanent</p>')
            } | Set-Content $destinationFile

    }

    if ($MobilePhone -eq '' -or $MobilePhone -eq '0'){
        $MobilePhone = '<not set>'
    }
#UPDATE FIELDS ON AD LOCAL
    Get-ADUser -Filter "SamAccountName -eq '$SamAccountName'" -Properties * -SearchBase "DC=airbelgium,DC=com" -Server 192.168.40.5 -Credential $credential |
    Set-ADUser -Title $JobTitle -MobilePhone $MobilePhone -OfficePhone $BusinessPhone -Department $($user.Department).Trim() -Manager "CN=$Manager,OU=AzureADUsers,DC=airbelgium,DC=com" -Company $($user.Company).Trim() -DisplayName $displayName

    (Get-Content $destinationFile) | ForEach-Object {
        $_.replace('[[FIRSTNAME]]', $firstNameWeb).replace('[[LASTNAME]]', $lastnameWeb).replace('[[OFFICEPHONE]]', $BusinessPhone).replace('[[MOBILEPHONE]]', $MobilePhone).replace('[[JOBTITLE]]', $JobTitle)
     } | Set-Content $destinationFile

#     (Get-Content $TEMPLATE_REPLY) | ForEach-Object {
#        $_.replace('[[FIRSTNAME]]', $firstNameWeb).replace('[[LASTNAME]]', $lastnameWeb).replace('[[OFFICEPHONE]]', $BusinessPhone).replace('[[MOBILEPHONE]]', $MobilePhone).replace('[[JOBTITLE]]', $JobTitle)
#     } | Set-Content $destinationFileReply

#    $identityName = $firstName + "." + $lastName + "@airbelgium.com"
    $sign = Get-Content -Path $destinationFile -ReadCount 0
#    $signReply = Get-Content -Path $destinationFileReply -ReadCount 0
    Set-MailboxMessageConfiguration -identity $identityName -SignatureHtml $sign -AutoAddSignature $true -AutoAddSignatureOnReply $true -AutoAddSignatureOnMobile $true -DefaultFontName Verdana
    #Set-MailboxMessageConfiguration -identity $identityName -SignatureHtml $signReply -AutoAddSignatureOnReply $true

}

try{
    Remove-PSSession $exchangeSession -ErrorAction SilentlyContinue

    #Remove-PSSession $sfboSession 
    Remove-PSSession $sfbSession -ErrorAction SilentlyContinue 
    Remove-PSSession $exchangeSession -ErrorAction SilentlyContinue 
    Remove-PSSession $ccSession -ErrorAction SilentlyContinue

    Disconnect-SPOService -ErrorAction SilentlyContinue

    Get-PSSession | Remove-PSSession
}catch{

}
