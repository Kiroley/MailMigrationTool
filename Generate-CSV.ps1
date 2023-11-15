<#
.Synopsis
   This script will generate a complete CSV file for a migration batch
.DESCRIPTION
   Starting with a CSV file containing a list of users this script will find mailbox information for each user and populate it in the CSV file.
   This information is used in the Export-PST script to generate PST files, which will then compare the exported item with the mailbox count to ensure the total amount matches.
.EXAMPLE
   Generate-CSV -Path \\Server1\Users.csv 
.INPUTS
   -Path (A path to a CSV file in a specified format, at a minimum it needs an ALIAS,NEWALIAS, and EXTEMAIL column)
.OUTPUTS
   Output from this cmdlet (if any)
.NOTES
   Requires the Exchange Management Shell to run
#>
# This parameter is looking for the filepath to the CSV file that contains the users to populate against. Simply needs an ALIAS column

Param
(
    [Parameter(Mandatory = $true, Position = 0, ParameterSetName = "Path",
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true,
        HelpMessage = "Path to the CSV file. Includes file name")]
    [ValidateNotNullOrEmpty()]
    [string]
    $Path
)

#region --------------- Collect Mailbox Info -------------------------------
<# **********************************************************************************************************************
Collect-MailboxInfo
This function will collect the mailbox information for all users, then append that information to a CSV file
************************************************************************************************************************* #>

Function Collect-MailboxInfo {
    try {
        Write-Output "Checking user $_"
        $user = Get-User -Identity $_.Alias -ErrorAction SilentlyContinue
        $mailstats = Get-MailboxStatistics -Identity $_.Alias -ErrorAction SilentlyContinue
        $mailfolder = Get-MailboxFolderStatistics -Identity $_.Alias -FolderScope DeletedItems | Where-Object {$_.Name -eq "Deleted Items"} -ErrorAction SilentlyContinue
        $obj = [PSCustomObject]@{

            Alias = $_.Alias
            NewAlias = $_.NewAlias
            FirstName = $user.FirstName
            LastName = $user.LastName
            MailboxSize = $mailstats.TotalItemSize
            LastAccess = $mailstats.LastLogonTime
            DisplayName = $user.DisplayName
            ExtEmail = $_.ExtEmail
            DeletedItemCount = $mailfolder.ItemsInFolderAndSubfolders
            PSTCheck = $null
        }  | Export-Csv -Path '.\Output.csv' -Append -ErrorAction SilentlyContinue -NoTypeInformation
    } catch {
        Write-Output "An error occurred while checking a mailbox"
    }
    Write-Output "User $_ complete"
}

#region --------------- MAIN -------------------------------
<# **********************************************************************************************************************
MAIN
Call all functions and run the script
************************************************************************************************************************* #>

#Import the CSV into memory
$CSV = Import-Csv -Path $Path
$CSV | foreach-object {Collect-MailboxInfo}
