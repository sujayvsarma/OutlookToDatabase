# OutlookToDatabase
Pull e-mail messages from your Outlook email into a SQL database. If you have a large mailbox and you need to analyze the messages in it for some reason, then this is a tool that can help you. 

Think about it this way -- you have a mailbox set up to receive automated system alerts and you need to analyze those alerts for something or take some other action on it.


## Pre-requisites

1. Machine with Outlook 2010 or above installed. You should have configured this Outlook already with your email profile and settings and downloaded all appropriate email. The tool will **not** directly download e-mail for you!

2. Machine with SQL Server. Also, have an account you can use that will have permissions to create a new database.

## Instructions

1. Find and install a tool called the "Office 2010 PIA Redistributable" (the file will be named "PIARedist_Office2010.exe"). 

   Link: https://www.microsoft.com/en-in/download/details.aspx?id=3508

2. Copy the files from this repository onto any folder on the system. Ensure both files are in the same folder, remember to right-click on both files one by one, and check ON "unblock" and click "Apply".

3. Now, open a PowerShell Command Prompt. CD to the folder you copied the two files to.

4. Run "OutlookDataMining.ps1" (remember to run it as "./OutlookDataMining.ps1"). Note the help, create the parameters as instructed, run it again, providing the parameters.

E-mail from the mailbox will be copied to the specified SQL database.
