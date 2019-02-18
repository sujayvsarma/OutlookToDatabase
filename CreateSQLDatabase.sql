---------------------------------------------------------------------------
---   SQLCMD script to create SQL database for OutlookDataMining script
---------------------------------------------------------------------------

USE master
GO

IF (NOT EXISTS (SELECT name from sys.databases WHERE ([name]='OutlookDataMining')))
BEGIN
    CREATE DATABASE [OutlookDataMining]
END

USE [OutlookDataMining]
GO

IF (OBJECT_ID('MailMessages') IS NULL)
BEGIN
    CREATE TABLE [MailMessages]
    (
        [MailMessageID]     [bigint]            IDENTITY(1,1)   NOT NULL,
        [MailFolder]        [nvarchar](500)     NULL,
        [MailDate]          [datetime]          NULL,
        [SenderInfoName]    [nvarchar](500)     NULL,
        [ToInfoList]        [nvarchar](500)     NULL,
        [CcInfoList]        [nvarchar](500)     NULL,
        [Subject]           [nvarchar](500)     NULL,
	[BodyText]	    [nvarchar](MAX)     NULL,
        [FlagIsUnread]      [bit],
        [FlagImportance]    [tinyint],
        [MessagePullDate]   [datetime]
    )
END
