
###    Transport Rule Script v1.1
###    Found online somewhere
###    Modified by Michael Lubert
###    Last Updated 2024-12-26
###    This script creates a new M365 Exchange Rule for incoming messages.
###    If the sender name ("Michael Lubert," not "mlubert@domain.com")
###    matches any sender name of a user in the tenant, it will add a banner at
###    the top of the email to warn them it may be a possible phishing email.
###    Before it creates this, it asks the user to supply a group of accounts to skip.
###    This group can be named anything, but should include common external
###    account names (help, info, sales, etc). If users are added after this rule,
###    the script will need to be run again to include them.

###############################
##### BEGIN CONFIGURATION #####
###############################

##### Name of Transport Rule #####
$ruleName = "External Senders with matching Display Names"
##### Name of Transport Rule #####
### Get Org Name ###
$orgName = Read-host -prompt "Please enter name of Organization to fill into HTML"
##### HTML Message Code #####
### Use `" to escape quotation marks in code below ###
$ruleHtml = @"
<table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0 align=left width=`"100%`" style='width:100.0%;mso-cellspacing:0cm;mso-yfti-tbllook:1184; mso-table-lspace:2.25pt;mso-table-rspace:2.25pt;mso-table-anchor-vertical:paragraph;mso-table-anchor-horizontal:column;mso-table-left:left;mso-padding-alt:0cm 0cm 0cm 0cm'>
    <tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes;mso-yfti-lastrow:yes'>
        <td style='background:#910A19;padding:5.25pt 1.5pt 5.25pt 1.5pt'></td>
        <td width=`"100%`" style='width:100.0%;background:#FDF2F4;padding:5.25pt 3.75pt 5.25pt 11.25pt; word-wrap:break-word' cellpadding=`"7px 5px 7px 15px`" color=`"#212121`">
            <div>
            <p class=MsoNormal style='mso-element:frame;mso-element-frame-hspace:2.25pt; mso-element-wrap:around;mso-element-anchor-vertical:paragraph;mso-element-anchor-horizontal:column;mso-height-rule:exactly'>
            <span style='font-size:9.0pt;font-family: `"Segoe UI`",sans-serif;mso-fareast-font-family:`"Times New Roman`";color:#212121'>
                This message was sent from outside $orgName by someone with a display name matching a user at $orgName. Please do not click links or open attachments unless you recognize the sender address of this email and know the content is safe.
            </span>
            </p>
            </div>
        </td>
    </tr>
</table>
<hr>
"@
#############################
##### END CONFIGURATION #####
#############################
### Get M365 Credentials ###
#$credentials = Get-Credential -Message "Enter Office 365 Admin User (for tenant, not a partner account)"
### Get Name of Group to Skip ###
$skipGroup=Read-host -prompt "Please enter name of distribution group with members to skip (default NOALERT)"
if ($skipGroup -eq '') {$skipGroup="NOALERT"}
### Connect to M365 ###
## Test for Exchange Online Module ##
if(Get-Module -ListAvailable -Name ExchangeOnlineManagement)
{
    Write-Host "Importing ExchangeOnlineManagement module"
    Import-Module ExchangeOnlineManagement
}
else{
    try
    {
        Write-Host "Installing ExchangeOnlineManagement module"
        Install-Module ExchangeOnlineManagement -Scope CurrentUser
    }
    catch
    {
        Write-Host "ExchangeOnlineManagement module is not installed and the script encountered an error trying to install it."
        exit
    }
}
## Connect ##
Write-Host "Connecting to Exchange Online" -ForegroundColor Yellow
Connect-ExchangeOnline
## Test for Graph Module ##
if(Get-Module -ListAvailable -Name Microsoft.Graph)
{
    Write-Host "Importing Microsoft.Graph.Users module"
    Import-Module Microsoft.Graph.Users
    Write-Host "Importing Microsoft.Graph.Groups module"
    Import-Module Microsoft.Graph.Groups
}
else{
    try
    {
        Write-Host "Installing Microsoft.Graph module"
        Install-Module Microsoft.Graph -Scope CurrentUser
    }
    catch
    {
        Write-Host "Microsoft.Graph module is not installed and the script encountered an error trying to install it."
        exit
    }
}
## Connect ##
Write-Host "Connecting to Entra Graph" -ForegroundColor Yellow
Connect-MgGraph -Scopes 'User.Read.All', 'GroupMember.Read.All' -NoWelcome
### Get Transport Rules ###
Write-Host "Getting Transport Rules" -ForegroundColor Yellow
$rule = Get-TransportRule | Where-Object {$_.Name -contains $ruleName}
### Get All Users' Details ###
$userDetailsExtensions = Get-MgUser -Property "id,displayName,onPremisesExtensionAttributes,surname,givenname"
### Get Exclude Group and Remove Those Users ###
$mgSkipGroup=(Get-MgGroup | where-object {$_.DisplayName -like "*$skipGroup*"})
if ($mgSkipGroup -eq $null)
{
    Write-Host "WARNING: Skip Group $skipGroup was not found. No users will be excluded." -ForegroundColor Yellow
}
else
{
    $mgSkipGroupMembers=(Get-MgGroupMember -GroupID $mgSkipGroup.Id)
    foreach($mgSkipGroupMember in $mgSkipGroupMembers)
    {
        $userDetailsExtensions = $userDetailsExtensions | where-object {$_.Id -ne $mgSkipGroupMember.Id}
    }
}
### Create List of Names ###
Write-Host "Creating list of names" -ForegroundColor Yellow
$alertList = @()
foreach($userDetailsExtension in $userDetailsExtensions)
{
    if($userDetailsExtension.givenname -ne $null -AND $userDetailsExtension.surname -ne $null)
    {
        $alertList += "$($userDetailsExtension.givenname) $($userDetailsExtension.surname)".toUpper()
        $add="$($userDetailsExtension.givenname) $($userDetailsExtension.surname)" -replace "[, ]*Ph.D.",""
        $alertList += $add.toUpper()
    }
    if($userDetailsExtension.displayName -ne $null){
        $alertList += $userDetailsExtension.displayName.toUpper()
        $add=$userDetailsExtension.displayName -replace "[, ]*Ph.D.",""
        $alertList += $add.toUpper()
    }
    if($userDetailsExtension.OnPremisesExtensionAttributes.ExtensionAttribute1 -ne $null){$alertList += "$($userDetailsExtension.givenname) $($userDetailsExtension.OnPremisesExtensionAttributes.ExtensionAttribute1) $($userDetailsExtension.surname)".toUpper()}
    if($userDetailsExtension.OnPremisesExtensionAttributes.ExtensionAttribute2 -ne $null){$alertList += "$($userDetailsExtension.givenname) $($userDetailsExtension.OnPremisesExtensionAttributes.ExtensionAttribute2) $($userDetailsExtension.surname)".toUpper()}
    if($userDetailsExtension.OnPremisesExtensionAttributes.ExtensionAttribute3 -ne $null){$alertList += "$($userDetailsExtension.givenname) $($userDetailsExtension.OnPremisesExtensionAttributes.ExtensionAttribute3) $($userDetailsExtension.surname)".toUpper()}
    if($userDetailsExtension.OnPremisesExtensionAttributes.ExtensionAttribute4 -ne $null){$alertList += "$($userDetailsExtension.givenname) $($userDetailsExtension.OnPremisesExtensionAttributes.ExtensionAttribute4) $($userDetailsExtension.surname)".toUpper()}
    if($userDetailsExtension.OnPremisesExtensionAttributes.ExtensionAttribute5 -ne $null){$alertList += "$($userDetailsExtension.givenname) $($userDetailsExtension.OnPremisesExtensionAttributes.ExtensionAttribute5) $($userDetailsExtension.surname)".toUpper()}
    if($userDetailsExtension.OnPremisesExtensionAttributes.ExtensionAttribute6 -ne $null){$alertList += "$($userDetailsExtension.givenname) $($userDetailsExtension.OnPremisesExtensionAttributes.ExtensionAttribute6) $($userDetailsExtension.surname)".toUpper()}
    if($userDetailsExtension.OnPremisesExtensionAttributes.ExtensionAttribute7 -ne $null){$alertList += "$($userDetailsExtension.givenname) $($userDetailsExtension.OnPremisesExtensionAttributes.ExtensionAttribute7) $($userDetailsExtension.surname)".toUpper()}
    if($userDetailsExtension.OnPremisesExtensionAttributes.ExtensionAttribute8 -ne $null){$alertList += "$($userDetailsExtension.givenname) $($userDetailsExtension.OnPremisesExtensionAttributes.ExtensionAttribute8) $($userDetailsExtension.surname)".toUpper()}
    if($userDetailsExtension.OnPremisesExtensionAttributes.ExtensionAttribute9 -ne $null){$alertList += "$($userDetailsExtension.givenname) $($userDetailsExtension.OnPremisesExtensionAttributes.ExtensionAttribute9) $($userDetailsExtension.surname)".toUpper()}
    if($userDetailsExtension.OnPremisesExtensionAttributes.ExtensionAttribute10 -ne $null){$alertList += "$($userDetailsExtension.givenname) $($userDetailsExtension.OnPremisesExtensionAttributes.ExtensionAttribute10) $($userDetailsExtension.surname)".toUpper()}
    if($userDetailsExtension.OnPremisesExtensionAttributes.ExtensionAttribute11 -ne $null){$alertList += "$($userDetailsExtension.givenname) $($userDetailsExtension.OnPremisesExtensionAttributes.ExtensionAttribute11) $($userDetailsExtension.surname)".toUpper()}
    if($userDetailsExtension.OnPremisesExtensionAttributes.ExtensionAttribute12 -ne $null){$alertList += "$($userDetailsExtension.givenname) $($userDetailsExtension.OnPremisesExtensionAttributes.ExtensionAttribute12) $($userDetailsExtension.surname)".toUpper()}
    if($userDetailsExtension.OnPremisesExtensionAttributes.ExtensionAttribute13 -ne $null){$alertList += "$($userDetailsExtension.givenname) $($userDetailsExtension.OnPremisesExtensionAttributes.ExtensionAttribute13) $($userDetailsExtension.surname)".toUpper()}
    if($userDetailsExtension.OnPremisesExtensionAttributes.ExtensionAttribute14 -ne $null){$alertList += "$($userDetailsExtension.givenname) $($userDetailsExtension.OnPremisesExtensionAttributes.ExtensionAttribute14) $($userDetailsExtension.surname)".toUpper()}
    if($userDetailsExtension.OnPremisesExtensionAttributes.ExtensionAttribute15 -ne $null){$alertList += "$($userDetailsExtension.givenname) $($userDetailsExtension.OnPremisesExtensionAttributes.ExtensionAttribute15) $($userDetailsExtension.surname)".toUpper()}
}
$alertList = $alertList | sort-object -unique
Write-Host "List created. It contains $($alertList.count) items." -ForegroundColor Yellow
### Check if rule exists and update it, or create it ###
if (!$rule) {
    Write-Host "Rule not found, creating rule" -ForegroundColor Yellow
    New-TransportRule -Name $ruleName -Priority 0 -FromScope "NotInOrganization" -SentToScope "InOrganization" -ApplyHtmlDisclaimerLocation "Prepend" -HeaderMatchesMessageHeader From -HeaderMatchesPatterns $alertList  -ApplyHtmlDisclaimerText $ruleHtml -Enabled $False
    Write-Host "Rule Created. Please double check and enable it manually."
}
else {
    Write-Host "Rule found, updating rule" -ForegroundColor Yellow
    Set-TransportRule -Identity $ruleName -Priority 0 -FromScope "NotInOrganization" -SentToScope "InOrganization" -ApplyHtmlDisclaimerLocation "Prepend" -HeaderMatchesMessageHeader From -HeaderMatchesPatterns $alertList -ApplyHtmlDisclaimerText $ruleHtml
}

### Disconnect ###
Write-Host "Disconnecting Graph" -ForegroundColor Yellow
disconnect-mggraph
Write-Host "Disconnecting Exchange Online" -ForegroundColor Yellow
disconnect-exchangeonline -Confirm:$false
