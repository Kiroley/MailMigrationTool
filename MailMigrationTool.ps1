<# 
.Synopsis
   This script will take the users identified in a CSV file and export the PSTs for the user, placing them unto the UNC file path specified. 
   It can then check the PST folder structure (once exported) and identify how many mail items are in the 'Deleted Items' folder. 
   Finally, it can generate a mail contact with a specified email address, then set a forwarding rule for a specified mailbox.
.DESCRIPTION
   Given the inputs of a CSV file and the path to save PST files (UNC) this script will get the list of users, export the mailbox for each user in PST 
   format, then save it to the UNC path specified. It will name the file according to the users alias, and destination system alias ($newalias). 
   This is done with the -Export switch.

   Utilising an add-in for Microsoft Outlook the script is capable of inspecting the exported PST files. This is done as a yardstick of the 
   data integrity upon export to a file share. It then generates a csv file (PSTCheckOutput.csv) in the $path directory so you can compare
   the results with what is reported by the mailbox. This is performed with the -CheckPST switch

   The final function is to generate a mail contact and set it as the desitnation of a mail forwarder for a specified mailbox. It uses the
   NewAlias and ExtEmail fields in the CSV file to complete this function. This is performed with the -CreateContact switch.
.EXAMPLE
   MailMigrationTool.ps1 -Path \\Server1\SharedExportFolder -CSV C:\Temp\users.csv -Export -CheckPST -Createcontact

   The above command will export all users PST files specified in the users.csv file to the SharedExportFolder UNC path, it will then check 
   the PSTs, and create contacts for the users.
.EXAMPLE
   MailMigrationTool.ps1 -path C:\Temp\Exports -CSV C:\Temp\Output.csv -CheckPST

   The above command will look for all the PST files as it relates to the Output.csv file and check their integrity (by looking at
   deleted item count)
.INPUTS
   -Path (The UNC Path that the PST files will be saved to. The results of the export are also saved to this same path as a CSV)
   -CSV (The path to a CSV file in a specified format, at a minimum it needs ALIAS, NEWALIAS, and EXTEMAIL columns.)
   -Export (Will export the mailboxes listed in the CSV file to PST)
   -CheckPST (Will evaluate the PST file after being exported by looking at the DeletedItemsFolder)
   -CreateContact (Creates a mail contact in Exchange and sets mail forwarding rules, the name as the NEWALIAS and email as EXTEMAIL)
.OUTPUTS
   .\FailedItems.CSV (Any items that fail at any step in the process, along with details of what failed)
   .\PSTCheckOutput.csv (Results from the -CheckPST function)
   .\GenerateContactOutput.csv (Results from the -CreateContact function)
.NOTES
   Requires the Exchange Management Shell to run and MS Outlook to be installed on the machine running the script. You will need to create
   a 'blank' Outlook profile that is not connected to Exchange. To do this go to Control Panel > User Accounts > Mail > Show Profiles >
   Add... > then follow the prompts to create a profile that is NOT connected to Exchange. Back in the Show Profiles windows set this 
   to 'Always use this profile'.
 
   Specify the Organizational Unit you wish mail contacts to be created in by using the $OU parameter.
#>
 
#region --------------- Param Loading -------------------------------
<# **********************************************************************************************************************
Section for loading any required parameters or variables
************************************************************************************************************************* #>
 
# $Path is the path to place the exported PST files (as well as the generated CSV) NOTE::: Needs to be in UNC format
# $CSV is the full path to the CSV file that contains the users you wish to export (created by Generate-CSV)
param 

   (
       [Parameter(Mandatory = $true, Position = 0, ParameterSetName = "Default",
           ValueFromPipeline = $true,
           ValueFromPipelineByPropertyName = $true,
           HelpMessage = "Path to save the PST files and CSV files in")]
       [ValidateNotNullOrEmpty()]
       [string]
       $Path,

       [Parameter(Mandatory = $true, Position = 1, ParameterSetName = "Default",
           ValueFromPipeline = $true,
           ValueFromPipelineByPropertyName = $true,
           HelpMessage = "Path to the input CSV (list of users), created by Generate-CSV")]
       [ValidateNotNullOrEmpty()]
       [string]
       $CSV,
      
       [Parameter(Mandatory = $false, Position = 2, ParameterSetName = "Default",
           ValueFromPipeline = $true,
           ValueFromPipelineByPropertyName = $true,
           HelpMessage = "If Export is set to TRUE the script will generate the PST files, if this is set to FALSE it will conduct the PST checks")]
       [ValidateNotNullOrEmpty()]
       [Switch]
       $Export,

       [Parameter(Mandatory = $false, Position = 4, ParameterSetName = "Default",
           ValueFromPipeline = $true,
           ValueFromPipelineByPropertyName = $true,
           HelpMessage = "If CheckPST is true the script will check the PST files and export the results to CSV")]
       [ValidateNotNullOrEmpty()]
       [Switch]
       $CheckPST,

       [Parameter(Mandatory = $false, Position = 5, ParameterSetName = "Default",
           ValueFromPipeline = $true,
           ValueFromPipelineByPropertyName = $true,
           HelpMessage = "If CreateContact is set to TRUE the script will generate the contact card and forwarding rules")]
       [ValidateNotNullOrEmpty()]
       [Switch]
       $CreateContact
   )   
 
# This variable is used in the Main script to update the details of a CSV file with how many items were counted
# in the Test-PST function.
$global:PSTCount = $null
 
# This variable is used in the Generate-Contact function. This specifies what Organizational Unit the mail contacts are to be created in
# will need to be set before running this script
$ou = $null

#endregion

#region --------------- Function - Export-PST -------------------------------
<# **********************************************************************************************************************
This function will iterate through the list of users specified in the CSV file and create an export request for 
each mailbox. Once the function is called it will export all PST files listed. If an error is detected write the 
details to the console and a CSV file
************************************************************************************************************************* #>
function Export-PST {
 
   # Load the CSV param from the main script
 
   #NOTE: Maybe change this to be the params for the MailboxExportRequest and pipe the details for main so the CSV isn't imported twice
   param ($CSV)
 
   # Check the path, if it's confirmed create a mailbox export request for each mailbox
 
   If (Test-Path -Path $CSV) 
    {
        Write-Host "Importing CSV"
        $EL = Import-Csv $CSV
 
        $EL | ForEach-Object {
            Write-Host "Starting export process for" ($_.Alias)
            try {
               If (Get-MailboxExportRequest -Mailbox $_.Alias) {
                  #First, make sure there is no existing mailbox export requests for that user. If so, remove it
                  Write-Host "An existing export request has been found, removing it for" $_.Alias
                  Remove-MailboxExportRequest -identity $_.Alias -Confirm:$False
                  Start-Sleep 3
               }   
                  
               If ((Test-Path -Path ($path + '\' + $_.Alias + '-' + $_.NewAlias +'.pst')) -eq $false) {
                  #If an existing PST file is NOT found in the directory create the export request for that user
                  Write-Host "No duplicate PST file found, generating new reqest for" ($_.Alias)
                  New-MailboxExportRequest -Name $_.Alias -Mailbox $_.Alias -FilePath ($path + '\' + $_.Alias + '-' + $_.NewAlias + '.pst')
               }
               Else {
                  #This part of the script is executed if an existing PST file is already in the $Path dir
                  #Write a message to the log, then collect the details of the user and write it to a CSV file
                  Write-host "Please ensure that there aren't any existing PST files in the same directory"
                  $obj = [PSCustomObject]@{
                     Issue = "A duplicate PST file was found"
                     name = ($_.Alias)
                     newalias = ($_.NewAlias)
                     FirstName = ($_.FirstName)
                     LastName = ($_.LastName)
                  } | Export-Csv -Path ($path + '\FailedItems.csv') -Append -NoTypeInformation
               }
            }
   
            catch {
               # If we cannot export the user's mailbox to CSV write a message to the console then to a CSV file
               Write-Host "An error occured when exporting the PST for ($_.Alias) , results can be found in the FailedItems.csv file"
               $obj = [PSCustomObject]@{
                  Issue = "An error occured when creating the export request, there may be an existing export request"
                  name = ($_.Alias)
                  newalias = ($_.NewAlias)
                  FirstName = ($_.FirstName)
                  LastName = ($_.LastName)
               } | Export-Csv -Path ($path + '\FailedItems.csv') -Append -NoTypeInformation
            }
         }
    }
   else {
      Write-Host "A CSV file must be specified and accessible. No CSV file found"
    }
}
#endregion
 
#region --------------- Function - Get-ExportProgress  -------------------------------
<# **********************************************************************************************************************
This function checks for any queued or in-progress PST exports. If it finds any it begins looping in 60s increments
************************************************************************************************************************* #>
function Get-ExportProgress {
   # Check the mailbox export progress. If items are queued or in progress wait, then continue
   try {
      #If the amount of queued and in-progress export requests is greater than 1, write to the console then sleep
      while ((Get-MailboxExportRequest | where {($_.Status -eq 'Queued') -or ($_.Status -eq 'InProgress')}).Count -ge 1) {
         Write-Host "PST Exports are still queued or in progress, the number of requests pending is;"
         (Get-MailboxExportRequest | where {($_.Status -eq 'Queued') -or ($_.Status -eq 'InProgress')}).Count 
         Write-Host "Waiting a minute then checking again"
         Start-Sleep -Seconds 60
      }  
   }
   catch {
      Write-Host "There was an issue checking the mailbox export progress"
   }
}
#endregion
 
#region --------------- Function - Test-PST -------------------------------
<# **********************************************************************************************************************
This function will count the items in the PST file. Specifically it should return the Deleted Item folder count. Derived
from PowerCountPST written by Alexander Bilz.
 
.LINK https://github.com/lxndrblz/PowerCountPST/
.NOTES
    Author: Alexander Bilz
    Date:   April 18, 2020 
************************************************************************************************************************* #>
function Test-PST {
   param (
      $PSTPath
   )
   #Sets the error codes in case of an issue opening the PST
   $ERR_FILENOTFOUND = 1000
   $ERR_ROOTFOLDER = 1001
   $ERR_LOCKEDFILE = 1002
   $ERR_CANTACCESSPST = 1003
 
   #This function writes to the command line the results of each folder. It reports back the name of the folder, the path, and how many items are contained
   function AnalyzeFolder ($folder) {
 
      $count = $folder.items.count
      Write-Host ('Folder: {0} contains {1} items' -f $folder.FolderPath, $count)
      Foreach ($subfolder in $folder.Folders) {
          $subfolderitems = AnalyzeFolder $subfolder
          $count = $count + $subfolderitems
      }
 
      return $count
  }
 
  function CountElements ($strPSTPath) {
 
   $global:PSTCount = 0
 
   #Check if Outlook is installed 
   Get-ItemProperty HKLM:\SOFTWARE\Classes\Outlook.Application | Select-Object PSPath -OutVariable outlook 
 
   #Create Outlook COM Object 
   $objOutlook = New-Object -com Outlook.Application 
   $objNameSpace = $objOutlook.GetNamespace("MAPI")
 
   #Try to load the PST into Outlook 
   try { 
       $objNameSpace.AddStore($strPSTPath) 
   } 
   catch { 
       Write-Error "Could not load pst - usually this is because the file is locked by another process."
       $objNameSpace.RemoveStore($strPSTPath) 
       try {
         if (Get-Process -Name OUTLOOK) {
            Write-Host "Outlook is currently open, attempting to terminate it"
            Stop-Process -Name OUTLOOK
         }
         Write-Host "Attempting to add PST again."
         $objNameSpace.AddStore($strPSTPath)
       }
       catch {
         Write-Host "Still unable to mount PST file. Attempting to duplicate to remove stuck handles"
         $objNameSpace.RemoveStore($strPSTPath) 
         Stop-Process -Name OUTLOOK
         Start-Sleep 1
         try {
            Copy-Item -Path $PSTPath -Destination ($PSTPath + '.TMP')
            Start-Sleep 1
            Remove-Item -Path $PSTPath
            Rename-Item -Path ($pstpath + '.tmp') -NewName $pstpath
            Write-Host "Duplication complete, attempting PST mount a final time"
            Start-Sleep 1
            $objNameSpace.AddStore($strPSTPath)
         }
         catch {
            Write-Host "Unable to mount PST at all, moving on"
            $global:PSTCount = "LockedPST"
            Exit $ERR_LOCKEDFILE
         }
       }
   } 
   Write-Host "Attempting mount of folders"
 
   #Try to load the Outlook Folders 
   try { 
       $PSTpath = $objnamespace.stores | ? { $_.FilePath -eq $strPSTPath } 
   } 
   catch { 
      #Write an error to the CSV file and exit if a fault is caught
       Write-Error "You have another PST added to outlook that cannot be accessed or found, please remove then re-run this script."
       $global:PSTCount = "CantAccessPST"
       Exit $ERR_CANTACCESSPST
   }
 
   #Try accessing the PST root
   try { 
       #Browse to PST Root 
       Write-Host "Browsing to root path"
       $root = $PSTpath.GetRootFolder() 
 
       #Count Items in subfolders, but only select the Deleted Items folder
       Write-Host "Attempting count of deleted items"
       $DeletedItems = $root.Folders | ? { $_.Name -eq "Deleted Items" }
       $global:PSTCount = AnalyzeFolder $DeletedItems
 
       # Output total number of found elements
       Write-Host $strPSTPath
       Write-Host ('Total Items: {0}' -f $global:PSTCount)
 
       # Unmount PST
       $objNameSpace.RemoveStore($root) 
       $objOutlook.Quit() | out-null 
       [System.Runtime.Interopservices.Marshal]::ReleaseComObject($objOutlook)
       Remove-Variable outlook | out-null
   }
   catch { 
      #Write an error to the CSV file and exit if a fault is caught
       Write-Error "Could not access root folder"
       $global:PSTCount = "CannotAccessRoot"
       Exit $ERR_ROOTFOLDER
   }
   } 
   If (Test-Path $PSTPath) {
      #If the path to the PST is confirmed count the items in it
      CountElements -strPSTPath $PSTPath
   }
   Else {
      #Write an error to the CSV file and exit if a fault is caught
      Write-Error "Please provide a valid path to a pst file."
      $global:PSTCount = "InvalidPath"
      Exit $ERR_FILENOTFOUND
   }
}
 
#endregion
 
#region --------------- Function - Generate-Contact  -------------------------------
<# **********************************************************************************************************************
This function will generate the contact card for the user, set the forwarding address for their mailbox, then ???
 
************************************************************************************************************************* #>
function Generate-Contact {
   param (
      $alias,$newalias,$extemail,$firstname,$lastname,$displayname,$ou
   )
   try {
      #Look for an existing mail contact for the user
      $Contact = $Null
        Try {
            $Contact = Get-MailContact $alias -ErrorAction SilentlyContinue
        }
        Catch {}

      #If no mail contact is found, create one
      If ( $Contact -eq $Null ) {
            Write-Host "Creating Contact ..."
            New-MailContact -Name $NewAlias `
               -ExternalEmailAddress $ExtEmail `
               -FirstName $firstName `
               -LastName $lastName `
               -DisplayName $displayname `
               -OrganizationalUnit $ou `
               -Confirm:$False
               #Start looking for the mail contact that was just created (AD Replication can take a while)
               Start-Sleep 5
               While ( $Contact -eq $Null ) {
                     $Contact = Get-MailContact $NewAlias -ErrorAction SilentlyContinue
                  If ( $Contact -eq $Null ) {
                     Write-Host "Contact not found, trying again ..."
                     Start-Sleep 5
                  }
               } 
         }
      # Set a mail forwarding address for the user's mailbox, using the name of the contact we just created
      Write-Host "Setting forwarding ..." 
      Set-Mailbox $Alias -ForwardingAddress $Contact
      # User Notification
      Write-Host "Forwarding Address = $ExtEmail"
      $obj = [PSCustomObject]@{
         name = $Alias
         newalias = $NewAlias
         FirstName = $FirstName
         LastName = $LastName
         ExtEmail = $extemail
         OU = $ou
         } | Export-Csv -Path ($path + '\GenerateContactOutput.csv') -Append -NoTypeInformation
   }  
   catch {
      #If there is an issue creating the contact card or setting the forwarder on the mailbox write the user's details to a CSV file
      Write-host "There was an issue creating a mail contact for $alias"
      $obj = [PSCustomObject]@{
         Issue = "A contact card could not be created, or a forwarding email could not be set"
         name = ($Alias)
         newalias = ($NewAlias)
         FirstName = ($_.FirstName)
         LastName = ($_.LastName)
         } | Export-Csv -Path ($path + '\FailedItems.csv') -Append -NoTypeInformation
   }
}
#endregion

#region --------------- Main  -------------------------------
<# **********************************************************************************************************************
In Main the CSV is imported and then:
 
1. Call Export-PST (Generates the PST file export requests)
2. Call Get-ExportProgress (Creates a sleep loop while the PST files are being created)
3. Iterate through the users listed in the CSV file and run Test-PST against each user
 
At the end you will end up with two CSV files in the $Path directory. One (Output.CSV) contains the list of PST
exports. The other (FailedItems.CSV) contains all the PST exports that failed.
 
As a note, a PST export might succeed, but there may be a significant discrepancy between the reported number of mail items
in the mailbox (deleted items folder) and the number found in the PST. If the threshold is reached it will:
 
TBC:Try and export that PST file again?
TBC:Let you know?
TBC:Add that user detail to the FailedItems.CSV file (but if it did you'd have to remove it from the Output.CSV file?)
 
************************************************************************************************************************* #>

If (!$Path.Contains("\\")) {
   Write-Error "The path specified needs to be a UNC path in order to export PST files"
}

Write-Host "Executing script"

#The two commands below are only to be executed if the -Export switch is set to $True
IF ($Export){
   Export-PST -CSV $CSV
   Get-ExportProgress
   Write-Host "All PST file export requests have been completed"
}

#The following command will check the PST files and report on the amount of items contained 
IF ($CheckPST) {
   $PL = Import-CSV $CSV
   $PL | foreach-object {       
      $global:PSTCount = $null
      $PathToPst = ($Path + '\' + ($_.Alias) + '-' + ($_.NewAlias) + '.pst')
      #Run Test-PST for the specified user
      Test-PST -PSTPath $PathToPst
               $obj = [PSCustomObject]@{
                  name = $_.Alias
                  newalias = $_.NewAlias
                  FirstName = $_.FirstName
                  LastName = $_.LastName
                  MailboxSize = $_.MailboxSize
                  LastAccess = $_.LastAccess
                  ExtEmail = $_.ExtEmail
                  DeletedItemCount = $_.DeletedItemCount
                  PSTCheck = $global:PSTCount
                  } | Export-Csv -Path ($path + '\PSTCheckOutput.csv') -Append -NoTypeInformation
   } 
}

#If the generate contact flag is set to true create a contact card for each user in the CSV file, then set forwarding rule for that mailbox
IF ($CreateContact) {
   $CL = Import-CSV $CSV
   $CL | foreach-object {

      Try {
         Generate-Contact -alias $_.Alias -newalias $_.NewAlias -firstname $_.FirstName -lastname $_.LastName -extemail $_.ExtEmail -displayname $_.DisplayName -ou $ou
      }
      #If it fails for any reason, write the user's details to the FailedItems CSV file
      catch {
         Write-Host "Unable to generate contact for" ($_.Alias)
         $obj = [PSCustomObject]@{
            Issue = "There was an issue with the user's details when creating a contact"
            name = $_.Alias
            newalias = $_.NewAlias
            FirstName = $_.FirstName
            LastName = $_.LastName
            } | Export-Csv -Path ($path + '\FailedItems.csv') -Append -NoTypeInformation
      }
   }
}

#endregion 


