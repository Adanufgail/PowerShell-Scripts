<#
.DESCRIPTION

###############################################################################################################
# Author: Francisco Manigrasso                                                                                #
# Version: 1.0 03/23/2023                                                                                     #
# https://techcommunity.microsoft.com/t5/exchange/get-permissions-for-all-folders-in-a-mailbox-ps/m-p/3776808 #
# This script will provide you all the folders permissions and permission level for a specific mailbox,       #
# avoiding the default ones. The script output is a csv file in C:\temp\FoldersPermissionsOutput.csv.         #
# If the script doesn't find any user permissions, it will not export any file and you'll see the output      #
# in PS: "There are no user permissions for this mailbox folders"                                             #
###############################################################################################################

------------------------------------------- DISCLAIMER ----------------------------------------------------

                  The sample script is provided AS IS without warranty of any kind.
      You are solely responsible for reviewing it and executing it if you deem it appropriate.

-----------------------------------------------------------------------------------------------------------


.PARAMETER mailbox
Specify the email address of the mailbox for which you would like to get all folders permissions.

.EXAMPLE

.\FolderPermissions.ps1 -mailbox email address removed for privacy reasons

#>

param (
  [Parameter(Position=0,Mandatory=$True,HelpMessage='Specifies the mailbox to be accessed')]
  [ValidateNotNullOrEmpty()]
  [string]$mailbox
  );

Write-Host "Getting folders information..." -ForegroundColor Green
$permissions = @()
$folders = Get-Mailboxfolderstatistics $mailbox | % {$_.folderpath} | % {$_.replace("/","\")}
$folderKey = $mailbox + ":" + "\"

Write-Host "Getting Permissions information..." -ForegroundColor Green
$permissions += Get-MailboxFolderPermission -identity $folderKey -ErrorAction SilentlyContinue | Where-Object {$_.User -notlike "Default" -and $_.User -notlike "Anonymous" -and $_.AccessRights -notlike "None" -and $_.User -notlike $mailbox}

$list = ForEach ($folder in $folders)
   {
    $folderKey = $mailbox + ":" + $folder
    $permissions += Get-MailboxFolderPermission -identity $folderKey -ErrorAction SilentlyContinue | Where-Object {$_.User -notlike "Default" -and $_.User -notlike "Anonymous" -and $_.AccessRights -notlike "None" -and $_.User -notlike $mailbox}
   }

if (!$Permissions) {Write-Host "There are no user permissions for this mailbox folders" -ForegroundColor Magenta}
if ($Permissions ) {
$permissions | Export-csv -path C:\temp\FoldersPermissionsOutput.csv -NoTypeInformation
Write-Host "Permissions file exported successfully" -ForegroundColor Green
}
