$logpath = "D:\Clients\PGA\PowerShell\TeamsTelephony\Log.csv"
$attachmentpath = "D:\Clients\PGA\PowerShell\TeamsTelephony\Welcome.pdf"
$WelcomeMailSender = "GOCGOClient.GOCShared@heidelbergcement.com"

$emailtext = "Dear colleagues, 
 
welcome to Teams Telephony! We are happy to have you here!

Please see the attached guide.

Wish you a pleasant day!
"


$me = whoami
$initiator = $me.Split("\") | select -last 1
$UPN = $initiator | Get-ADUser | select -ExpandProperty userprincipalname
if (-not(Test-Path -Path $logpath)) {New-Item -ItemType file -Path $logpath 1>$null}

function Add-Log {
    Param(
        [string]$result,
        [string]$message
    )
    $logmessage = [PSCustomObject]@{
        Time = get-date
        User = $user.userprincipalname
        Result = $result
        Message = $message
    }
    $logmessage | Export-Csv $logpath -Append
}

Add-log -message "Teams Activation script started."

function Get-MyCredential {
    if ($credential -eq $null -or $credential.UserName -ne $UPN) {
        $pwFilePath = "C:\\Users\$initiator\pw.txt"

        function Update-TTPassword {
            $cred = Get-Credential -UserName $initiator -Message "Enter your GIT password."
            $cred.Password | ConvertFrom-SecureString | Out-File -FilePath $pwFilePath
        }

        $pwFile = get-item -path $pwFilePath
        $pwdate = get-aduser $initiator -Properties PasswordLastSet | select -ExpandProperty PasswordLastSet
        if (-not $pwfile -or $pwFile.LastWriteTime -lt $pwdate) {
            $hostname = HOSTNAME
            $passwordupdate = "Update password of $upn on $hostname for Teams Telephony activation script!"
            $user = Get-ADUser $initiator -Properties emailaddress
            Send-MailMessage -From "GOCGOClient.GOCShared@heidelbergcement.com" -to $User.emailaddress -Subject "Password update needed!" -Priority High -Body $passwordupdate -SmtpServer DEUNonAuthSMTP.grouphc.net
            Add-Log -result "CRITICAL" -message "Password for connecting to Teams Online missing/expired."
            break
        }
        $PwdSecureString = Get-Content $pwFilePath | ConvertTo-SecureString
        $Credential = New-Object System.Management.Automation.PSCredential -ArgumentList $upn, $PwdSecureString
    }
    $credential
}

function Connect-ARS {
    $session = New-PSSession -ComputerName deusafran0161.grouphc.net
    Import-PSSession -Module ActiveRolesManagementShell -Session $session -AllowClobber 3>&1>$null
    Connect-QADService -Service deusafran0161.grouphc.net –Proxy 1>$null
}

function Connect-Teams {
    try {
        Get-PSSession -Name "SfBPowerShellSessionViaTeamsModule*" | where State -EQ "broken" | Remove-PSSession
        $TeamsSessionState = Get-PSSession -Name "SfBPowerShellSessionViaTeamsModule*" | select -ExpandProperty State
    }
    catch {}
    if ($TeamsSessionState -ne "Opened") {
        $credential = Get-MyCredential
        $TeamsSession = New-CsOnlineSession -Credential $credential -OverrideAdminDomain "hcgroupnet.onmicrosoft.com"
        Import-PSSession -Session $TeamsSession -AllowClobber 1>$null
    }
}

function Get-TeamsTelephonyUnifyNumber {
    Connect-Teams
    $allnumbers = get-csonlineuser -Filter "OnPremLineUri -like '*49622148141*'" | sort -property OnPremLineUri
    $firstnumber = $allnumbers | select -ExpandProperty OnPremLineUri -First 1
    $numbers = $allnumbers | select -ExpandProperty OnPremLineUri -Skip 1
    $position = $firstnumber.Length - 4
    $previousextension = $firstnumber.Substring($position,4).toint32($null)
    foreach ($number in $numbers) {
        $newnumber = [PSCustomObject]@{
            TTNumber = $null
            ADNumber = $null
        }
        $position = $number.Length - 4
        $extension = $number.Substring($position,4).toint32($null)
        if ($extension -ne $previousextension + 1) {
            $newnumber.TTNumber = "+4962214814"+($previousextension+1)
            $newnumber.ADNumber = "+49 6221 481 4"+($previousextension+1)
            break
        }
        else {
            $previousextension = $extension
        }
    }
    $newnumber
}

<# Option 1 for getting new users based on location
$today = get-date
$yesterday = $today.AddDays(-1)
get-aduser -Filter "City -eq 'Heidelberg' -or Citry -eq 'Leimen'" -Properties created | where created -gt $yesterday
#>

# Option 2 for getting users based on group
$Users = get-adgroupmember -Identity "DEU TeamsTelephony activation" | select -ExpandProperty samaccountname | Get-ADUser -Properties telephoneNumber, created

Connect-ARS #function above
Connect-Teams #function above

#select policies (expand with switch based on location and move to loop in future when adding other countries/locations)
$RoutingPolicy = "DEUHeidelberg-All"
$CallingPolicy = "AllowCallingBOB"
$DialPlan = "DP-DEUHeidelberg"


foreach ($user in $users) {
    # Check if it's a new user
    $Today = Get-Date
    $cutoff = $Today.AddMonths(-1)
    if ($user.Created -lt $cutoff) {
        Remove-QADGroupMember -Member $user.distinguishedname -Identity "CN=DEU TeamsTelephony activation,OU=Groups,OU=DEU,OU=EU,DC=grouphc,DC=net"
        Add-Log -result "FAILED" -message "User was created more than a month ago."
    }
    else {
        # Check if CSOnlineUser exists
        $UserReturned = $null
        $UserReturned = Get-CSOnlineUser -Identity $User.userprincipalname -ErrorAction SilentlyContinue
        $UserUPN = $user | select -ExpandProperty UserPrincipalName
        if ($UserReturned.EnterpriseVoiceEnabled -eq $true) {
            Remove-QADGroupMember -Member $user.distinguishedname -Identity "CN=DEU TeamsTelephony activation,OU=Groups,OU=DEU,OU=EU,DC=grouphc,DC=net"
            Add-Log -result "SUCCESS" -message "EnterpriceVoice is enabled. User removed from activation group."
        }
        elseif ($UserReturned) {
            $groups = Get-ADPrincipalGroupMembership $user.samaccountname | select -ExpandProperty name
            if (-not($groups -contains "GRP Microsoft Teams Telephony Users")) {
                Add-Log -message "User not member of Teams Telephony group. Adding now."
                Add-QADGroupMember -Member $user.distinguishedname -Identity "CN=GRP Microsoft Teams Telephony Users,OU=Groups,OU=GIT,DC=grouphc,DC=net"
            }
            else {
                $number = Get-TeamsTelephonyUnifyNumber #function above
                try {
                    Add-Log -Message "Attempting to enable enterprise voice with number $($number.TTNumber)"
                    $tel = $number.TTNumber
                    Set-CsUser -Identity "$UserUPN" -EnterpriseVoiceEnabled $true -HostedVoiceMail $true -OnPremLineURI tel:$tel -ErrorAction stop
                }
                catch {
                    if ($Error.exception -like "*MCOProfessional.") {
                        Add-Log -result "FAILED" -message "Could not enable Enterprise Voice. Waiting for License activation."
                    }
                    else {
                        Add-Log -result "CRITICAL" -message "Could not enable Enterprise Voice. Investigate Reason."
                    }
                }
                $EnterpriseVoiceEnabled = $false
                $EnterpriseVoiceEnabled = Get-CSOnlineUser -Identity $User.userprincipalname -ErrorAction SilentlyContinue | select -ExpandProperty EnterpriseVoiceEnabled        
                if ($EnterpriseVoiceEnabled -eq $true) { #set policies and send welcome mail
                    Grant-CsOnlineVoiceRoutingPolicy -Identity $UserUPN -PolicyName $RoutingPolicy
                    Add-Log -message "Online Voice routing [$RoutingPolicy] policy set. User can call all numbers, no restriction."
                    Grant-CsTeamsCallingPolicy -PolicyName "Tag:$CallingPolicy" -Identity  $UserUPN
                    Add-Log -message "Teams Calling policy [$CallingPolicy] set."
                    Grant-CsTenantDialPlan -PolicyName $DialPlan -Identity $UserUPN
                    Add-Log -message "Dial Plan [$DialPlan] was assigned."

                    Remove-QADGroupMember -Member $user.distinguishedname -Identity "CN=DEU TeamsTelephony activation,OU=Groups,OU=DEU,OU=EU,DC=grouphc,DC=net"
                    Add-Log -result "SUCCESS" -message "EnterpriceVoice is enabled. User removed from activation group."

                    Send-MailMessage -From $WelcomeMailSender -to $UserUPN -Subject "Welcome to Teams Telephony" -Body $emailtext -Attachments $attachmentpath -SmtpServer DEUNonAuthSMTP.grouphc.net
                    
                    Set-QADUser -Identity $user.DistinguishedName -ObjectAttributes @{telephoneNumber=$number.ADNumber}
                    Add-Log -message "Updated telephone number in AD. Changed from $($User.telephoneNumber) to $($number.ADNumber)."
                
                }
                if ($UserReturned.TeamsUpgradeEffectiveMode -ne 'TeamsOnly' ) {
                    Add-Log -Message "User is not in Teams Only Mode - Outgoing phone calls will not work."
                }
            }
        }
        else {
            Add-Log -result "CRITICAL" -message "CSOnlineUser not found."
        }
    }
}

Get-PSSession | Remove-PSSession


