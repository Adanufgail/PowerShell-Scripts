<#  
.NOTES
===========================================================================
Created on:     3/27/2018 7:37 PM
Created by:     Bradley Wyatt
Updated by:     Michael Lubert
Updated on:     12/27/2023 4:00 PM
Version:        2.0.1
Notes:          Updated for M365
===========================================================================
.DESCRIPTION
This script will send an e-mail notification to users where their password is set to expire soon. It includes step by step directions for them to 
change it on their own.
It will look for the users e-mail address in the emailaddress attribute and if it's empty it will use the proxyaddress attribute as a fail back. 
The script will log each run at $DirPath\log.txt
#>

##################################################################
##################################################################
########## BEGIN VARIABLES. CONFIGURATION OPTIONS BELOW ##########
##################################################################
##################################################################

$FromEmail = "administrator@contoso.com"                                        # Account in tenant to send email from
$TechEmail = "tech@contoso.com"                                                 # Gets a daily report
$DebugMode = $False                                                             # Send all emails to the $TechEmail instead of end users (for testing)
$TechPriority = "Normal"                                                        # Importance (High, Normal, Low) of email report to tech

$expireindays = 14                                                              # Number of days before expirtation to begin warning
$UserPriority = "High"                                                          # Importance (High, Normal, Low) of email alert to user

$AlertEmail = "support@contoso.com"                                             # Email to alert to when users let a password get to $AlertDays of fewer before expiration
$AlertDays = 3                                                                  # Number of days remaining to trigger an alert to the $AlertEmail
$AlertPriority = "High"                                                         # Importance (High, Normal, Low) of email alert to $AlertEmail

$DirPath = "C:\Scripts\PasswordResetNotification\"                              # Location that script is located

$CertThumb = "0000000000000000000000000000000000000000"                         # Thumbprint for Self-Signed Cert used for Entra Authentication
$EARAppID = "00000000-0000-0000-0000-000000000000"                              # Entra App Registration Application ID
$EARTenantID = "00000000-0000-0000-0000-000000000000"                           # Entra Tenant ID

#########################################
### BEGIN USER EMAIL MESSAGE TEMPLATE ###
#########################################

# USERDAYSPLACEHOLDER will be replaced by days until User's password expires.
# USERNAMEPLACEHOLDER will be replaced by User's name
$UserSubject = "Your password will expire USERDAYSPLACEHOLDER days"
$UserMessage =
"Dear USERNAMEPLACEHOLDER,</br>
Your Contoso password will expire in USERDAYSPLACEHOLDER days. <u>Please change it as soon as possible.</u></br>
To change your password, follow the method below:</br>
<ol>
<li>On your Windows computer
    <ol type='a'>
    <li>If you are not in the office, logon or connect to Citrix and follow that section below. 
    <li>Log onto your computer as usual and make sure you are connected to the internet.
    <li>Press Ctrl-Alt-Del and click on ""Change Password"".
    <li>Fill in your old password and set a new password.  See the password requirements below.
    <li>Press OK to return to your desktop.
    <li>Please note: this will take a variable amount of time to sync to Microsoft 365 (between a few minutes and an hour). Please contact Support if you need access to webmail and are locked out. Other services should take effect immediately.
    </ol>
<li>In Microsoft 365
    <ol type='a'>
    <li>Go to <a href='https://outlook.office.com'>https://outlook.office.com</a>
    <li>Log in with your Contoso username and password.
    <li>In the top right, click on your username to open a drop-down and select ""View Account.""
    <li>Select ""Change a password.""
    <li>Fill in your old password and set a new password.  See the password requirements below.
    <li>Please note: this will immediately push to other parts of the on-prem environment (Citrix, Sharefile), but will only sync with your computer if you are both in the office <b>and</b> if you manually lock or reboot your computer.
    </ol>
</ol>
<ul>
<li>The new password must meet the minimum requirements set forth in our corporate policies including:
    <ol type='a'>
    <li>It must be at least 12 characters long.
    <li>It must contain at least one character from 3 of the 4 following groups of characters:
        <ol type='i'>
        <li>Uppercase letters (A-Z)
        <li>Lowercase letters (a-z)
        <li>Numbers (0-9)
        <li>Symbols (!@#$%^&*...)
        </ol>
    <li>It cannot match any of your past 3 passwords.
    <li>It cannot contain characters which match 2 or more consecutive characters of your username.
    <li>You cannot change your password more often than once in a 24 hour period.
    </ol>
</ul>
<b>Please note: You will need to update your new password on your mobile device for Contoso email. It will typically prompt you to do so.</br></b>
</br>
If you have any questions please contact support at <a href='mailto:support@contoso.com'>support@contoso.com</a> or call us at 555-555-5555
"

#######################################
### END USER EMAIL MESSAGE TEMPLATE ###
#######################################

###########################################
### BEGIN REPORT EMAIL MESSAGE TEMPLATE ###
###########################################

# USERCOUNTPLACEHOLDER text will be replaced by the number of users expiring in the expiration window
$ReportSubject = "Contoso Password Reset Report for $(get-date -format "yyyy-MM-dd"): USERCOUNTPLACEHOLDER Users" 
$ReportMessage = "<h1>Contoso Password Report for $(get-date -format "yyyy-MM-dd")</h1></br>The following users have passwords expiring in the next $expireindays days:</br>"
$ReportNoUsers = "No user passwords are expiring in the next $expireindays days"

#########################################
### END REPORT EMAIL MESSAGE TEMPLATE ###
#########################################

#########################################
### BEGIN TEXT REPLACEMENT HASH TABLE ###
#########################################

# Define any placeholders
# FORMAT: TEXT-TO-BE-REMOVED = VARIABLE_NAME

$Placeholders = @{
    USERNAMEPLACEHOLDER = Name
    USEREMAILPLACEHOLDER = emailaddress
    USERPWDLASTSET = passwordSetDate
    USERDAYSPLACEHOLDER = daystoexpire
    USERCOUNTPLACEHOLDER = ReportUsers
}

#######################################
### END TEXT REPLACEMENT HASH TABLE ###
#######################################

#################################################################################################
#################################################################################################
#################################################################################################
########## END VARIABLES. THERE IS NOTHING MORE FOR YOU TO CONFIGURE BEYOND THIS POINT ##########
#################################################################################################
#################################################################################################
#################################################################################################

###############################
########## FUNCTIONS ##########
###############################

function SendMessage
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position=0)]
        [string]$From,
        [Parameter(Mandatory, Position=1)]
        [object]$To,
        [Parameter(Mandatory, Position=2)]
        [string]$Subject,
        [Parameter(Mandatory, Position=3)]
        [string]$Priority,
        [Parameter(Mandatory, Position=4)]
        [string]$Message
    )

    if($DebugMode){
        Write-Host "FROM: $From" -ForegroundColor Green
        Write-Host "TO: $From" -ForegroundColor Green
        Write-Host "SUBJECT: $Subject" -ForegroundColor Green
        Write-Host "PRIORITY: $Priority" -ForegroundColor Green
        Write-Host "MESSAGE: $Message" -ForegroundColor Green
    }
    $Email = @{
        Message = @{
            Subject = $Subject
            ToRecipients = $To # Array of @{EmailAddress = @{Address = $ToAddress}}
            Body = @{
                contentType = "HTML";
                content = $Message
            }
            Importance = $Priority
        }
        #SaveToSentItems = "true"
    }
    Send-MGUserMail -UserId $From -BodyParameter $Email
}

###################################
########## END FUNCTIONS ##########
###################################

########################################
########## CONFIRM LOG ACCESS ##########
########################################

$Date = Get-Date
#Check if program dir is present
$DirPathCheck = Test-Path -Path $DirPath
If (!($DirPathCheck))
{
    Try
    {
        #If not present then create the dir
        New-Item -ItemType Directory $DirPath -Force
    }
    Catch
    {
       $_ | Out-File ($DirPath + "\" + "Log.txt") -Append
    }
}

###############################
########## BEGIN RUN ##########
###############################

echo "----------------------START OF RUN----------------------"| Out-File ($DirPath + "\" + "Log.txt") -Append
if($DebugMode){
    Write-Host "DEBUG MODE ACTIVE, SENDING ALL MAIL TO $TechEmail)" -ForegroundColor Green;
    Write-Host "FROM EMAIL: $FromEmail" -ForegroundColor Green
    Write-Host "TECH EMAIL: $TechEmail" -ForegroundColor Green
    Write-Host "ALERT EMAIL: $AlertEmail" -ForegroundColor Green
    Write-Host "EXPIRATION THRESHOLD: $expireindays" -ForegroundColor Green
    Write-Host "ALERT THRESHOLD: $AlertDays" -ForegroundColor Green
    Write-Host "FROM EMAIL: $FromEmail" -ForegroundColor Green
}
else{Write-Host "DEBUG MODE DISABLED, SENDING ALL MAIL TO USERS)" -ForegroundColor Green}

#####################################
########## CONNECT TO M365 ##########
#####################################

Try
{
    Connect-MgGraph -CertificateThumbprint $CertThumb -AppId $EARAppId -TenantId $EARTenantID -nowelcome
}
Catch
{
   $_ | Out-File ($DirPath + "\" + "Log.txt") -Append
}

########################################
########## GET LOCAL AD USERS ##########
########################################

# Get Users From AD who are Enabled, Passwords Expire and are Not Currently Expired
"$Date - INFO: Importing AD Module" | Out-File ($DirPath + "\" + "Log.txt") -Append
Import-Module ActiveDirectory
"$Date - INFO: Getting users" | Out-File ($DirPath + "\" + "Log.txt") -Append
$users = Get-Aduser -properties Name, PasswordNeverExpires, PasswordExpired, PasswordLastSet, EmailAddress -filter { (Enabled -eq 'True') -and (PasswordNeverExpires -eq 'False') } | Where-Object { $_.PasswordExpired -eq $False }
$maxPasswordAge = (Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge
# Process Each User for Password Expiry
$ExpiringUsers=@()
$AlertSupport=0;
foreach ($user in $users)
{
    # Get User Details
    $Name = $user.Name
    Write-Host "Working on $Name..." -ForegroundColor White
    Write-Host "Getting e-mail address for $Name..." -ForegroundColor Yellow
    $emailaddress = $user.emailaddress
    If (!($emailaddress))
    {
        Write-Host "$Name has no E-Mail address listed, looking at their proxyaddresses attribute..." -ForegroundColor Red
        echo "$Name has no E-Mail address listed, looking at their proxyaddresses attribute, ignore errors as that means they have no emails whatsoever" | Out-File ($DirPath + "\" + "Log.txt") -Append
        Try
        {
           $emailaddress = (Get-ADUser $user -Properties proxyaddresses | Select-Object -ExpandProperty proxyaddresses | Where-Object { $_ -cmatch '^SMTP' }).Trim("SMTP:")
        }
        Catch
        {
            $_ | Out-File ($DirPath + "\" + "Log.txt") -Append
        }
        If (!($emailaddress))
        {
            Write-Host "$Name has no email addresses to send an e-mail to!" -ForegroundColor Red
            #Don't continue on as we can't email $Null, but if there is an e-mail found it will email that address
            "$Date - WARNING: No email found for $Name" | Out-File ($DirPath + "\" + "Log.txt") -Append
        }
    }
    
    # Get Password last set date
    $passwordSetDate = (Get-ADUser $user -properties * | ForEach-Object { $_.PasswordLastSet })
    #Check for Fine Grained Passwords
    $PasswordPol = (Get-ADUserResultantPasswordPolicy $user)
    if (($PasswordPol) -ne $null)
    {
        $maxPasswordAge = ($PasswordPol).MaxPasswordAge
    }
    $expireson = $passwordsetdate + $maxPasswordAge
    $today = (get-date)

    # Gets the count on how many days until the password expires and stores it in the $daystoexpire var
    $daystoexpire = (New-TimeSpan -Start $today -End $Expireson).Days
    $user.daystoexpire=$daystoexpire
    If (($daystoexpire -ge "0") -and ($daystoexpire -lt $expireindays))
    {
        $To=@()
        if($daystoexpire -le $AlertDays){$AlertSupport=1;} # Alert Support if daystoexpire is equal to or less than $AlertDays
        "$Date - INFO: Sending expiry notice email to $Name" | Out-File ($DirPath + "\" + "Log.txt") -Append
        Write-Host "Sending Password expiry email to $name" -ForegroundColor Yellow
        if($DebugMode)
        {
            $To+=@{EmailAddress = @{Address = $TechEmail}}      # SEND EACH TO TECHEMAIL FOR DEBUG PURPOSES
        }
        else
        {
            $To+=@{EmailAddress = @{Address = $emailaddress}}      # SEND EACH TO RESPECTIVE USERS
        }

        $Priority = $UserPriority

        $Message = $UserMessage
        $Subject = $UserSubject

        foreach($Replacement in $Placeholders.keys)
        {
            if($DebugMode){$FullName;$Replacement;$Placeholders[$Replacement];(Get-Variable -name $Placeholders[$Replacement]).value}
            try
            {
                $Message = $Message.replace($Replacement,(Get-Variable -name $Placeholders[$Replacement]).value)
                $Subject = $Subject.replace($Replacement,(Get-Variable -name $Placeholders[$Replacement]).value)
            }
            catch
            {
                echo $Placeholders[$Replacement] not present in current context | Out-File ($DirPath + "\" + "Log.txt") -Append
            }
        }
        

        if($DebugMode){Write-Host "Sending E-mail to $emailaddress...(DEBUG, ACTUALLY SENDING TO $TechEmail)" -ForegroundColor Green}
        else{Write-Host "Sending E-mail to $emailaddress..." -ForegroundColor Green}
        Try
        {
            SendMessage -From $FromEmail -To $To -Subject $Subject -Priority $Priority -Message $Message
        }
        Catch
        {
            $_ | Out-File ($DirPath + "\" + "Log.txt") -Append
        }
        $ExpiringUsers += $user
    }
    Else
    {
        "$Date - INFO: Password for $Name not expiring for $daystoexpire days, $daystoexpire remain" | Out-File ($DirPath + "\" + "Log.txt") -Append
        Write-Host "Password for $Name does not expire for $daystoexpire days, $daystoexpire remain" -ForegroundColor White
    }
}

############################################
########## PREPARE SUMMARY REPORT ##########
############################################

$To=@()
$Priority=$TechPriority

if($AlertSupport -ge 1){$To+=@{EmailAddress = @{Address = $AlertEmail}};$Priority=$AlertPriority} #If any users DaysToExpire is less than AlertDays, also alert the AlertEmail

$To+=@{EmailAddress = @{Address = $TechEmail}} #SEND REPORT TO $TechEmail

$ReportUsers = $ExpiringUsers.Count

$Message = $ReportMessage
$Subject = $ReportSubject

foreach($Replacement in $Placeholders.keys)
{
    try
    {
        $Message = $Message.replace($Replacement,(Get-Variable -name $Placeholders[$Replacement]).value)
        $Subject = $Subject.replace($Replacement,(Get-Variable -name $Placeholders[$Replacement]).value)
    }
    catch
    {
        echo $Placeholders[$Replacement] not present in current context | Out-File ($DirPath + "\" + "Log.txt") -Append
    }
}

foreach ($eu in $ExpiringUsers)
{
    $Message += $eu.Name + ": "+$eu.daystoexpire+" days<br>"
}
if($ExpiringUsers.Count -eq 0)
{
    $Message += $ReportNoUsers
}
Write-Host "Sending Report Email" -ForegroundColor Green
Try
{
    SendMessage -From $FromEmail -To $To -Subject $Subject -Priority $Priority -Message $Message
}
Catch
{
    $_ | Out-File ($DirPath + "\" + "Log.txt") -Append
}
echo "----------------------END OF RUN----------------------"| Out-File ($DirPath + "\" + "Log.txt") -Append