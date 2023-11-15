
## Mail Migration Tool ##

This set of scripts has been created to solve the problem of moving mailboxes from an old exchange environment to a new one, complete with new identities for users.

A few challenges arose with this;

1. How can we make sure the exported PST files have retained their integrity?

    > It turns out that using the MailItemCount attribute to measure the mailbox items against the amount of items in the PST file are wildly different. However, when looking at the DeletedItems item  count it appears to be accurate within 2 items between what the mailbox reports and the PST file. Using PowerCountPST written by Alexander Bilz we can interrogate the PST file and return the item count of a given folder.  
https://github.com/lxndrblz/PowerCountPST/    
Author: Alexander Bilz  
Date:   April 18, 2020 

2. How can we map the PST file to a new mailbox?

    > By keeping track of the new alias/identity of the user from the very beginning of the process within our reports (CSV files), then naming the PST file in a certain way to track the old alias and the new one (e.G. JohnDoe-JD1268.pst, where JohnDoe was the old identity and JD117268 is the new target identity), which may contain a 4-digit employee number

3. How can we keep track of who has been migrated?

    > At each step of the process a CSV file is produced, allowing you to keep track of users. In terms of interrogating a live system if a mail contact has been set for a user, and/or a forwarding address is set for the mailbox then they have been migrated (or at least their mail contact created)

4. How do we route mail from the old mailbox to the new one?

    > By using a mail contact within the old system that contains the new email address (EXTEMAIL), then by setting a forwarding rule on the original mailbox

### How to Use It ###

1. Copy Generate-CSV.ps1 and MailMigrationTool.ps1 into the source domain and place in a directory of your choosing (ideally on an Exchange server)

2. Create a CSV for the users you wish to migrate (for this example we'll call it users.csv) file in the following format;

    `alias,newalias,extemail`  
    `johndoe,JD1174,jd1174@company.com`  
    `helpdesk,servicedesk,servicedesk@company.com`  

    NOTE: Alias is the OLD user ID, NEWALIAS is the target (or new) user ID, EXTEMAIL is the target (new) mailbox email address

3. From the Exchange Management Shell navigate to the directory where your scripts are kept and run the following command

    > `.\Generate-CSV.ps1 -path .\users.csv`

4. A new CSV (titled Output.csv) will be created in the specified -path. Review the information contained, such as the mailbox size and last logon time. This information can be used to help size your migration batches by user count, mailbox size, or your most frequent users. What is important is that the CSV file you use contains only the users you wish to export/migrate. 

5. Ensure you have Outlook installed on the server/client you are running this script from. No particular version of Outlook is required, but it is recommended you use 2016 or later. Open Outlook but do not connect to Exchange. You can do this by clicking 'Cancel' on the 'Add Account' popup, then click 'OK' to set up outlook without connecting. Once Outlook is open it should be showing 'Outlook Data File' on the left. You may need to go into your outlook profile settings and ensure that the newly added profile is the default (and is not set to 'prompt' for an account).

6. Designate a UNC path that is accessible to the machine you are executing the scripts from. You may need to set share security and file-level permissions. For this example we will use \\server1\exports. From the Exchange Management Shell navigate to the directory where your scripts are kept and run the following command:

    > `.\MailMigrationTool.ps1 -Path \\server1\exports -CSV .\output.csv -Export -CheckPST`

7. The script will now create mailbox export requests, once all the exports have completed it will then move on to check the PST files. At the end of this script you will end up with a new CSV file titled 'PSTCheckOutput.csv' that will contain all the same information as 'output.csv' except the CheckPST column will be populated. You can use this information to determine if the exported PST files have maintained the same integrity as the mailbox itself.  

8. Once you are happy, and you wish to redirect all future mail from the old mailbox to the new mailbox (once created in the target domain of course), run the script again but omit the -export and -CheckPST switches with -CreateContact as per the example below;

    > `.\MailMigrationTool.ps1 -Path \\server1\exports -CSV .\output.csv -CreateContact`

9. You will now have a new CSV file available called 'GenerateContactOutput.csv' that will show all the details of the contacts created. From this point any mail destined for the old mailboxes will be forwarded to the target email (as per the EXTEMAIL column for that user).

10. If any items fail during any of the scripts you will find their results in the 'FailedItems.csv' file

11. Move the PST files to your target domain/exchange system and ensure they are accessible from the Exchange server/client. You will note that the name of the pst files contains the old user ID and the new one. If there is no change to your user ID's they should both be the same (E.G. user1-user1.pst).

12. Ensure all the target mailboxes have been created and are ready for import. Once checked run the following command;

    > `.\Import-PST.ps1 -ImportPST -Path \\newserver\imports`

13. All failed items will appear in a CSV file called 'FailedItems', all successful import requests will be added to 'Imported'. From here sit back and watch the mail items roll into the mailboxes. Congratulations!
