#############################################################################################
#  PowerShell script to mine Outlook mailbox folders
#  Requires PowerShell v2.0
#############################################################################################

Add-Type -assembly "Microsoft.Office.Interop.Outlook" | Out-Null

function Recurse-Folder($CurrentFolder,$SearchFolder) {
    #Write-Host 'Search='$SearchFolder
    
    $interestedFolder = ''
    foreach($Folder in $CurrentFolder.Folders)
    {
        if ($Folder.FolderPath)
        {
            #Write-Host $Folder.FolderPath
            if ($Folder.FolderPath.StartsWith($SearchFolder))
            {
                $interestedFolder = $Folder
                break
            }

            if ($Folder.Folders -ne $NULL)
            {
                $interestedFolder = Recurse-Folder $Folder $interestedFolderName
                if ($interestedFolder) 
                {
                    if ($interestedFolder.FolderPath)
                    {
                        if ($interestedFolder.FolderPath.StartsWith($interestedFolderName))
                        {
                            break
                        }
                    }
                }
            }
        }
    }

    return $interestedFolder
}


$outlookFolders = "Microsoft.Office.Interop.Outlook.olDefaultFolders" -as [Type]
$outlookApplication = New-Object -ComObject Outlook.Application
$namespaceMAPI = $outlookApplication.GetNameSpace("MAPI")

$paramMailboxName = ''
$paramFolderName = ''
$paramDBConnection = ''
$paramMoveMessage = ''

$interestedMailbox = ''
$interestedFolderName = ''

$flagMoveMail = 0

$wvOutlookDataMiningScriptsFolder = Get-Location

if ($args)
{
    foreach($parameter in $args)
    {
        if ($parameter.ToLower().StartsWith('-mailbox'))
        {
            $paramMailboxName = $parameter.Substring(9)
        }

        if ($parameter.ToLower().StartsWith('-folder'))
        {
            $paramFolderName = $parameter.Substring(8)
        }

        if ($parameter.ToLower().StartsWith('-database'))
        {
            $paramDBConnection = $parameter.Substring(10)
        }

        if ($parameter.ToLower().StartsWith('-move'))
        {
            $paramMoveMessage = $parameter.Substring(6).ToLower()

            if (($paramMoveMessage -eq 'true') -or ($paramMoveMessage -eq 'yes') -or ($paramMoveMessage -eq '1'))
            {
                $flagMoveMail = 1
            }
            else
            {
                $flagMoveMail = 0
            }
        }
    }
}
else
{
    Write-Host '---------------------------------------'
    Write-Host 'Outlook Data Mining Tool v1.0'
    Write-Host '---------------------------------------'
    Write-Host 'This PowerShell script will import email messages from any Outlook MAPI folder on your system into a SQL Server database. Scripts to generate the SQL database are included in the original distribution of this script and will be automatically run when you run this script.'
    Write-Host ' '
    Write-Host ' Usage Instructions: '
    Write-Host '     ./OutlookDataMining.ps1 ''-mailbox=<MAILBOXNAME>'' ''-folder=<FOLDERNAME>'' ''-database=<CONNECTIONSTRING>'' ''-move=<<TRUE|FALSE>>'' '
    Write-Host ' '
    Write-Host ' Note that because of a bug in PowerShell, you will need to place single ('') or double (") quotation-marks around each parameter as shown above. Otherwise your value may not be fully read by the system. '
    Write-Host ' '
    Write-Host '     -mailbox   :    Specify the MAPI folder configured in the Mail or MAPI profile on this computer from which the emails have to be imported.'
    Write-Host '                     Example: -mailbox=johndoe@hotmail.com '
    Write-Host ' '
    Write-Host '     -folder    :    Name of the folder or subfolder to import the emails from.'
    Write-Host '                     Example: -folder=Inbox\AirTickets '
    Write-Host ' '
    Write-Host '     -database  :    Database connection string in standard format.'
    Write-Host '                     Example: -database=Data Source=(local);Initial Catalog=OutlookDataMining;Integrated Security=true; '
    Write-Host ' '
    Write-Host '     -move      :    Flag, if set will delete matching mails from the Outlook folder after moving to database.'
    Write-Host '                     Example: -move=TRUE '
    Write-Host ' '
    Write-Host ' Usage Example: '
    Write-Host '       ./OutlookDataMining.ps1 ''-mailbox=johndoe@hotmail.com'' ''-folder=Inbox\AirTickets'' ''-database=Data Source=(local);Initial Catalog=OutlookDataMining;Integrated Security=true;'' ''-move=TRUE'' '
    Write-Host ' '
    Write-Host '---------------------------------------'

    exit
}

if ($paramFolderName)
{
    $interestedFolderName = $paramFolderName
}
else
{
    $interestedFolderName = $namespaceMAPI.GetDefaultFolder($outlookFolders::olFolderInbox).Name
}

if ($paramMailboxName)
{
    $interestedMailbox = $paramMailboxName
}
else
{
    $interestedMailbox = $namespaceMAPI.DefaultStore.DisplayName
}


foreach($Store in $namespaceMAPI.Stores)
{
    if ($Store.DisplayName -eq $interestedMailbox)
    {
        $interestedMailbox = $Store.DisplayName
        $interestedFolderName = '\\' + $interestedMailbox + '\' + $interestedFolderName

        foreach($Folder in $namespaceMAPI.Folders.Item($interestedMailbox).Folders)
        {
            if ($Folder.FolderPath)
            {
                #Write-Host $Folder.FolderPath
                if ($Folder.FolderPath.StartsWith($interestedFolderName))
                {
                    $interestedFolder = $Folder
                    break
                }

                if ($Folder.Folders -ne $NULL)
                {
                    $interestedFolder = Recurse-Folder $Folder $interestedFolderName
                    if ($interestedFolder) 
                    {
                        if ($interestedFolder.FolderPath)
                        {
                            if ($interestedFolder.FolderPath.StartsWith($interestedFolderName))
                            {
                                break
                            }
                        }
                    }
                }
            }
        }
    }
}


if ((-not $paramDBConnection) -or ($paramDBConnection -eq $NULL))
{
    $errorMsg = 'The database parameter is mandatory.'
    Write-Error -message $errorMsg -category InvalidArgument -errorId  'ERRDatabaseConnection' -recommendedAction 'Please check the parameters and try again.'
    
    exit
}

if ($interestedFolder)
{
    Write-Host 'Found: ' + $interestedFolder.FolderPath
}
else
{
    $errorMsg = 'Requested folder ' + $interestedFolderName + ' was not found in currently configured MAPI profiles on this system.'
    Write-Error -message $errorMsg -category ObjectNotFound -errorId  'ERROutlookFolderNotFound' -recommendedAction 'Please check the parameters OR reconfigure the MAPI profiles on this system and try again.'
    
    exit
}

$SQLDatabase = New-Object -TypeName System.Data.SqlClient.SqlConnection
$SQLDatabase.ConnectionString = $paramDBConnection

Invoke-Sqlcmd -InputFile 'CreateSQLDatabase.sql' -ServerInstance $SQLDatabase.DataSource
Set-Location $wvOutlookDataMiningScriptsFolder
[Environment]::CurrentDirectory = $wvOutlookDataMiningScriptsFolder

Write-Host 'Enumerating messages in '$interestedFolder.FolderPath' and storing in '$SQLDatabase.DataSource

$PullTimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
$SQLCommand = $SQLDatabase.CreateCommand()
$FetchedMessageCount = 0

foreach($msg in $interestedFolder.Items)
{
    $FetchedMessageCount = ($FetchedMessageCount + 1)

    #$msgSender = $msg.SenderName
    #$msgTime = $msg.ReceivedTime.ToString("yyyy-MM-dd HH:mm:ss")

    Write-Host 'Processing message '$FetchedMessageCount'... WITH_MOVE='$flagMoveMail   #$msgSender' at '$msgTime
    
    $SQLInsertQuery = "INSERT INTO [MailMessages] VALUES(" + 
                        "'" + $interestedFolder.FolderPath + "'," + 
                        "'" + $msg.ReceivedTime.ToString("yyyy-MM-dd HH:mm:ss") + "'," + 
                        "'" + $msg.SenderName + "'," + 
                        "'" + $msg.To.Replace("'", "''") + "'," + 
                        "'" + $msg.CC.Replace("'", "''") + "'," + 
                        "'" + $msg.Subject.Replace("'", "''") + "'," + 
                        "'" + $msg.Body.Replace("'", "''") + "'," + 
                        "'" + $msg.UnRead.ToString() + "'," + 
                        "'" + $msg.Importance.ToString() + "'," + 
                        "'" + $PullTimeStamp  + "'" + 
                        ")"

    $SQLCommand.CommandText = $SQLInsertQuery

    if ($SQLDatabase.State -ne "Open")
    {
        $SQLDatabase.Open()
    }

    try {
        
        $SQLCommand.ExecuteNonQuery() | Out-Null

        #if ($flagMoveMail -eq 1)
        #{
        #    $msg.Delete()
        #}

    } 
    catch {
    }
}

if ($SQLDatabase.State -ne "Closed")
{
    $SQLDatabase.Close()
}

Write-Host 'Imported ['$FetchedMessageCount'] messages.'