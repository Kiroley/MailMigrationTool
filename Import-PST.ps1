<#
.Synopsis
   Imports PST files that are labelled with the OLDID-NEWID.PST format to mailboxes that match the NEWID
.DESCRIPTION
   This script looks at a specified file path ($path), then tries to find all PSTs produced by the
   MailMigrationTool.ps1 script. As such the PSTS are named in the following format;

   OLDID-NEWID.PST
   E.G. JohnDoe-JDoe.PST (where JDoe is the new user account ID, and name of the new mailbox)
.EXAMPLE
   .\Import-PST -Path \\Server1\Exports -ImportPST
.OUTPUTS
   .\FailedItems.CSV (Any failed imports)
   .\Imported.CSV (All succesful imports)
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false, Position = 1, ParameterSetName = "Default",
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true,
        HelpMessage = "If ImportPST is enabled it will attempt to import the PSTs to the mailboxes belonging to the new aliases")]
       [ValidateNotNullOrEmpty()]
       [Switch]
       $ImportPst,

       [Parameter(Mandatory = $true, Position = 2, ParameterSetName = "Default",
           ValueFromPipeline = $true,
           ValueFromPipelineByPropertyName = $true,
           HelpMessage = "Path to import PST files from")]
       [ValidateNotNullOrEmpty()]
       [string]
       $Path
)

#region --------------- Function - OutputInfo  -------------------------------
<# **********************************************************************************************************************
This Exports failed item details to a CSV file
************************************************************************************************************************* #>
function OutputInfo {
   param (
      $name,$newalias,$Info,$CSVname
   )

   $obj = [PSCustomObject]@{
      Info = $info
      name = $name
      newalias = $newalias
   } | Export-Csv -Path ($path + '\' + $CSVname + '.csv') -Append -NoTypeInformation
   
}
#endregion

If ($ImportPst) {
    $ImportList = Get-ChildItem -Path $path -Include *.pst -Recurse
    
    ForEach ($item in $ImportList) {
      
      #Split out the name of the alias from the PST file in the folder
      Write-Host "Splitting" $item
      #The $IMalias is found by splitting out all the backslashes (\) in the path, selecting the last one 
      #(being the file name), then seperates the name of the two aliases, selects the second name, then
      #trims the .PST from the end.
      $IMalias = $item.ToString().Split('\')[-1].Split('-')[1].Split('.')[0]
      Write-Host "looking for mailbox for" $IMalias
      try {
         $MailboxPresent = Get-Mailbox -Identity $IMalias
      }
      catch {
         Write-Host "No mailbox found for" $IMalias
         OutputInfo -info "No mailbox found" -name $item.Name -newalias $IMalias -CSVname 'FailedItems'
      }
      
      If ({Test-Path $item} -and {$MailboxPresent}){
         #If the PST file is in the right spot, and a mailbox exists, create the import request
         try {
            New-MailboxImportRequest -Name $IMAlias -FilePath $item -Mailbox $IMalias
            Write-Host "Request generated for" $IMalias
            OutputInfo -Info "Import request created" -name $item.name -newalias $IMalias -CSVname 'Imported' 
         }
         catch {
            Write-host "There was an issue creating the mailbox import request for " $IMalias
            OutputInfo -Info "Could not create import request" -name $item.Name -newalias $IMalias -CSVname 'FailedItems'
         }
      }
        
   }
}
