SET NOEXEC OFF

USE MASTER
GO

DECLARE @db_id int;  
SET @db_id = DB_ID(N'uits_gst_db');  

IF @db_id IS NULL   
  BEGIN 
    PRINT N'MSG A:: uits_gst_db not found, continuing to create new ...';  
  END 
ELSE  
  BEGIN  
	DROP DATABASE uits_gst_db
	PRINT N'MSG B:: dropped existing uits_gst_db, continuing to create new ...';  
  END 
GO 

---###################################################################################
--- uits_gst_db Windows Version           ############################################
---###################################################################################
CREATE DATABASE [uits_gst_db]  ON (NAME = N'uits_gst_db', FILENAME = N'C:\uits_gst_db.mdf' , SIZE = 512, FILEGROWTH = 10%) LOG ON (NAME = N'uits_gst_db_Log', FILENAME = N'C:\uits_gst_db_log.LDF' , FILEGROWTH = 10%)  COLLATE SQL_Latin1_General_CP1_CI_AS

---###################################################################################
--- uits_gst_db Linux Version             ############################################
---###################################################################################
--- CREATE DATABASE [uits_gst_db]  ON (NAME = N'uits_gst_db', FILENAME = N'/var/opt/mssql/data/uits_gst_db.mdf' , SIZE = 512, FILEGROWTH = 10%) LOG ON (NAME = N'uits_gst_db_Log', FILENAME = N'/var/opt/mssql/data/uits_gst_db_log.LDF' , FILEGROWTH = 10%)  COLLATE SQL_Latin1_General_CP1_CI_AS
GO


SET QUOTED_IDENTIFIER OFF 
SET ANSI_NULLS ON 
use [uits_gst_db]

DECLARE @db_id int;  
SET @db_id = DB_ID(N'uits_gst_db');  
IF @db_id IS NULL   
BEGIN 
	PRINT N'MSG C:: Unable to select correct database, setting noexec ...';  
	set noexec on
END 
GO


---###################################################################################
--- uits_gst_db TABLES                    ############################################
---###################################################################################
CREATE TABLE [dbo].[DBKey] (
	[KeyOne] [nvarchar] (50)  NOT NULL  DEFAULT ('to.change'),
	[KeyTwo] [nvarchar] (50)  NOT NULL  DEFAULT ('to.change')
) ON [PRIMARY]

CREATE TABLE [dbo].[CMan] (
	[UserName] [nvarchar] (50)  NOT NULL DEFAULT ('scotty'),
	[LoginID] [nvarchar] (50)  NOT NULL DEFAULT ('scotty'),
	[Passwd] [nvarchar] (10)  NOT NULL DEFAULT ('scotty'),
	[Address] [nvarchar] (100)  NOT NULL DEFAULT ('user.address'),
	[Phones] [nvarchar] (50)  NOT NULL DEFAULT ('user.phone'),
	[Rights] [numeric](18, 0) NOT NULL DEFAULT (2)
) ON [PRIMARY]

CREATE TABLE [dbo].[Currency] (
	[Serial] [decimal](18, 0) NOT NULL DEFAULT (0),
	[Currency] [nvarchar] (20)  NOT NULL  DEFAULT ('INR'),
	[CurrPrice] [float] NOT NULL  DEFAULT (1),
	[Codes] [char] (10)  NOT NULL  DEFAULT ('Rs.')
) ON [PRIMARY]

CREATE TABLE [dbo].[GRP] (
	[GRP] [varchar] (20)  NOT NULL DEFAULT ('XXX'),
	[Description] [varchar] (50)  NOT NULL DEFAULT ('XXX'),
	[Initial] [varchar] (1)  NOT NULL DEFAULT ('X')
) ON [PRIMARY]

CREATE TABLE [dbo].[Letters] (
	[LetterRef] [numeric](18, 0) IDENTITY (1, 1) NOT NULL ,
	[LetterDate] [datetime] NOT NULL DEFAULT (getdate()),
	[Receipient] [nvarchar] (50)  NOT NULL DEFAULT ('Letter_Receipient'),
	[Subject] [nvarchar] (150)  NOT NULL DEFAULT ('Subject'),
	[LetterBody] [nvarchar] (900)  NOT NULL DEFAULT ('LetterBody'),
	[Sender] [nvarchar] (50)  NOT NULL DEFAULT ('Letter_Sender')
) ON [PRIMARY]

CREATE TABLE [dbo].[PERSONAL] (
	[ID] [nvarchar] (10)  NULL ,
	[Name] [nvarchar] (100)  NULL DEFAULT ('Name'),
	[Address] [nvarchar] (200)  NULL DEFAULT ('Address'),
	[City] [nvarchar] (100)  NULL DEFAULT ('City'),
	[State] [nvarchar] (50)  NULL DEFAULT ('BR'),
	[Phones] [nvarchar] (50)  NULL DEFAULT ('Phones'),
	[Email] [nvarchar] (100)  NULL DEFAULT ('r@example.com'),
	[PAN] [varchar] (20)  NULL DEFAULT ('PANIN2345X'),
	[GSTIN] [varchar] (20)  NULL DEFAULT ('GSTIN22333'),
	[ShipAddress] [varchar] (200)  NULL DEFAULT ('Ship.Address'),
	[TID] [varchar] (10)  NULL DEFAULT ('T0001'),
	[KID] [varchar] (10)  NULL DEFAULT ('K0001'),
	[PID] [varchar] (10)  NULL DEFAULT ('O0001'),
	[ACCOUNT] [varchar] (20)  NULL DEFAULT ('YES'),
	[OBDATE] [datetime] NULL DEFAULT (getdate()),
	[OB] [money] NULL DEFAULT (0),
	[CreditLimit] [money] NULL DEFAULT (0),
	[GRP] [varchar] (20)  NULL DEFAULT ('GENERAL'),
	[Type] [varchar] (20)  NULL DEFAULT ('Type'),
	[Code] [varchar] (10)  NULL DEFAULT ('New') ,
	[DiscGrp] [char] (10)  NULL DEFAULT ('DiscGrp01'),
	[DiscTplt] [nvarchar] (10)  NULL DEFAULT ('DiscTplt01'),
	[TEMP1] [varchar] (50)  NULL DEFAULT ('TEMP1'),
	[TEMP2] [varchar] (50)  NULL DEFAULT ('TEMP2'),
	[TEMP3] [varchar] (50)  NULL DEFAULT ('TEMP3')
) ON [PRIMARY]

---###################################################################################
CREATE TABLE [dbo].[PMT] (
	[Serial] [numeric](18, 0) IDENTITY (1, 1) NOT NULL ,
	[Date] [datetime] NOT NULL DEFAULT (getdate()),
	[ID] [varchar] (10)  NOT NULL DEFAULT ('R0002'),
	[Name] [varchar] (100)  NOT NULL DEFAULT ('CASH'),
	[City] [varchar] (50)  NULL DEFAULT ('PTN'),
	[Amount] [money] NOT NULL DEFAULT (0.00),
	[Mode] [varchar] (10)  NOT NULL DEFAULT ('R0002'),
	[ModeName] [varchar] (50)  NULL DEFAULT ('CASH'),
	[Narration] [varchar] (200)  NOT NULL DEFAULT ('PMT#')
) ON [PRIMARY]

CREATE TABLE [dbo].[RCT] (
	[Serial] [numeric](18, 0) IDENTITY (1, 1) NOT NULL ,
	[Date] [datetime] NOT NULL DEFAULT (getdate()),
	[ID] [varchar] (10)  NOT NULL DEFAULT ('R0002'),
	[Name] [varchar] (100)  NOT NULL DEFAULT ('CASH'),
	[City] [varchar] (50)  NULL DEFAULT ('PTN'),
	[Amount] [money] NOT NULL DEFAULT (0.00),
	[Mode] [varchar] (10)  NOT NULL DEFAULT ('R0002'),
	[ModeName] [varchar] (50)  NULL DEFAULT ('CASH'),
	[Narration] [varchar] (200)  NOT NULL DEFAULT ('RCT#')
) ON [PRIMARY]

CREATE TABLE [dbo].[TRF] (
	[Serial] [numeric](18, 0) IDENTITY (1, 1) NOT NULL ,
	[Date] [datetime] NOT NULL DEFAULT (getdate()),
	[DrID] [varchar] (10)  NOT NULL DEFAULT ('R0002'),
	[DrName] [varchar] (100)  NOT NULL DEFAULT ('TRF_DrName'),
	[DrCity] [varchar] (50)  NULL DEFAULT ('PTN'),
	[CrID] [varchar] (10)  NOT NULL DEFAULT ('R0002'),
	[CrName] [varchar] (100)  NOT NULL DEFAULT ('TRF_CrName'),
	[CrCity] [varchar] (50)  NULL DEFAULT ('PTN'),
	[Amount] [money] NOT NULL DEFAULT (0.00),
	[Narration] [varchar] (200)  NOT NULL DEFAULT ('TRF_Narration')
) ON [PRIMARY]

CREATE TABLE [dbo].[VOUCHERS] (
	[Serial] [numeric](18, 0) IDENTITY (1, 1) NOT NULL ,
	[Date] [datetime] NOT NULL DEFAULT (getdate()),
	[DrID] [varchar] (10)  NOT NULL DEFAULT ('R0002'),
	[DrName] [varchar] (100)  NOT NULL  DEFAULT ('VCH_DrName'),
	[DrCity] [varchar] (50)  NULL DEFAULT ('PTN'),
	[CrID] [varchar] (10)  NOT NULL DEFAULT ('R0002'),
	[CrName] [varchar] (100)  NOT NULL DEFAULT ('VCH_CrName'),
	[CrCity] [varchar] (50)  NULL DEFAULT ('PTN'),
	[Amount] [money] NOT NULL DEFAULT (0.00),
	[Narration] [varchar] (200)  NOT NULL DEFAULT ('MISC'),
	[VType] [varchar] (10)  NOT NULL DEFAULT ('MEMO') 
) ON [PRIMARY]

CREATE TABLE [dbo].[JOURNAL] (
	[Serial] [numeric](18, 0) IDENTITY (1, 1) NOT NULL ,
	[Date] [datetime] NOT NULL DEFAULT (getdate()),
	[DrAC] [varchar] (10)  NOT NULL DEFAULT ('R0002'),
	[DrName] [varchar] (100)  NULL DEFAULT ('DrName'),
	[CrAC] [varchar] (10)  NOT NULL DEFAULT ('R0002'),
	[CrName] [varchar] (100)  NULL DEFAULT ('CrName'),
	[Amount] [money] NOT NULL DEFAULT (0.00),
	[Narration] [varchar] (150)  NOT NULL DEFAULT ('JrNarration'),
	[UserID] [varchar] (10)  NOT NULL DEFAULT ('U000'),
	[MemoRef] [varchar] (15)  NOT NULL DEFAULT ('Memo#'),
	[AuthDate] [datetime] NOT NULL DEFAULT (getdate()),
	[AuthBy] [varchar] (50)  NOT NULL DEFAULT ('None')
) ON [PRIMARY]

---###################################################################################
CREATE TABLE [dbo].[Items] (
	[ItemID] [numeric](18, 0) IDENTITY (1, 1) NOT NULL ,
	[ItemCode] [nvarchar] (20)  NOT NULL DEFAULT ('00-000-0000'),
	[HSNCode]  [nvarchar] (20)  NOT NULL DEFAULT ('HSN_SAC_Code'),
	[ItemName] [nvarchar] (200)  NOT NULL DEFAULT ('ItemName'),
	[MakerAuthor] [nvarchar] (200)  NOT NULL DEFAULT ('Maker/Author'),
	[ProducerID] [nvarchar] (20)  NOT NULL DEFAULT ('P0000'),
	[ProducerName] [nvarchar] (200)  NOT NULL DEFAULT ('ProducerName'),
	[Packing] [nvarchar] (10)  NOT NULL DEFAULT ('Box'),
	[Unit] [nvarchar] (10)  NOT NULL DEFAULT ('Count'),
	[Currency] [nvarchar] (10)  NOT NULL DEFAULT ('INR'),
	[CurrMRP] [float] NOT NULL DEFAULT (100.00),	
	[CurrSRP] [float] NOT NULL DEFAULT (90.00),
	[PDisc] [nvarchar] (50)  NOT NULL DEFAULT ('0'),
	[SDisc] [nvarchar] (50)  NOT NULL DEFAULT ('0'),
	[GST] [float] NOT NULL DEFAULT (18),
	[Cess] [float] NOT NULL DEFAULT (0),
	[Version] [nvarchar] (50)  NULL DEFAULT ('V1'),
	[MfgDate] [datetime] NOT NULL DEFAULT (getdate()),
	[ExpDate] [datetime] NOT NULL DEFAULT (getdate()),
	[InitStock] [numeric](18, 0) NOT NULL DEFAULT (0),
	[WareLocation] [nvarchar] (20)  NOT NULL DEFAULT ('Cage0.Rac0'),
	[ItemID1] [numeric](18, 0) NOT NULL DEFAULT (0),
	[Misc1] [nvarchar] (20)  NOT NULL DEFAULT ('Misc1'),
	[Misc2] [nvarchar] (20)  NOT NULL DEFAULT ('Misc2'),
	[Misc3] [nvarchar] (20)  NOT NULL DEFAULT ('Misc3')
) ON [PRIMARY]

---###################################################################################
CREATE TABLE [dbo].[PURCHASE] (
	[Serial] [decimal](18, 0) NOT NULL DEFAULT (0),
	[DBRefX] [decimal](18, 0) NOT NULL DEFAULT (0),
	[ItemID] [numeric](18, 0) NOT NULL DEFAULT (0),
	[ItemCode] [nvarchar] (20)  NOT NULL DEFAULT ('ItemCode'),
	[HSNCode]  [nvarchar] (20)  NOT NULL DEFAULT ('HSN_SAC_Code'),
	[ItemName] [nvarchar] (200)  NOT NULL DEFAULT ('ItemName'),
	[MakerAuthor] [nvarchar] (200)  NOT NULL DEFAULT ('MakerORAuthor'),
	[ProducerID] [nvarchar] (50)  NOT NULL DEFAULT ('P0001'),
	[ProducerName] [nvarchar] (200) NOT NULL DEFAULT ('ProducerName'),
	[Version] [nvarchar] (50) NOT NULL DEFAULT ('V1'),
	[MfgDate] [datetime] NOT NULL DEFAULT (getdate()),
	[ExpDate] [datetime] NOT NULL DEFAULT (getdate()),
	[Packing] [nvarchar] (10)  NOT NULL DEFAULT ('BoxOrDrumOrPouch'),
	[Unit] [nvarchar] (10)  NOT NULL DEFAULT ('Count'),
	[Qty]  [numeric](18, 0) NOT NULL  DEFAULT (0),
	[Free] [numeric](18, 0) NOT NULL DEFAULT (0),
	[MRP] [float] NOT NULL DEFAULT (0.00),	
	[SRP] [float] NOT NULL DEFAULT (0.00),
	[Gross] [money] NOT NULL DEFAULT (0.00),
	[aDisc] [nvarchar] (50) NOT NULL DEFAULT ('0'),
	[aDiscAmt] [money] NOT NULL DEFAULT (0.00),
	[bDisc] [nvarchar] (50)  NOT NULL DEFAULT ('0'),
	[bDiscAmt] [money] NOT NULL DEFAULT (0.00),
	[cDisc] [nvarchar] (50)  NOT NULL DEFAULT ('0'),
	[cDiscAmt] [money] NOT NULL DEFAULT (0.00),
	[GST] [float] NOT NULL DEFAULT (0.00),
	[GSTAmt] [money] NOT NULL DEFAULT (0.00),
	[Cess] [float] NOT NULL DEFAULT (0.00),
	[CessAmt] [money] NOT NULL DEFAULT (0.00),
	[Amount] [money] NOT NULL DEFAULT (0.00),
	[Stock] [numeric](18, 0) NOT NULL DEFAULT (0)
) ON [PRIMARY]

CREATE TABLE [dbo].[SALE] (
	[Serial] [decimal](18, 0) NOT NULL DEFAULT (0),
	[DBRefX] [decimal](18, 0) NOT NULL DEFAULT (0),
	[ItemID] [numeric](18, 0) NOT NULL DEFAULT (0),
	[ItemCode] [nvarchar] (20)  NOT NULL DEFAULT ('ItemCode'),
	[HSNCode]  [nvarchar] (20)  NOT NULL DEFAULT ('HSN_SAC_Code'),
	[ItemName] [nvarchar] (200)  NOT NULL DEFAULT ('ItemName'),
	[MakerAuthor] [nvarchar] (200)  NOT NULL DEFAULT ('MakerORAuthor'),
	[ProducerID] [nvarchar] (50)  NOT NULL DEFAULT ('P0001'),
	[ProducerName] [nvarchar] (200) NOT NULL DEFAULT ('ProducerName'),
	[Version] [nvarchar] (50) NOT NULL DEFAULT ('V1'),
	[MfgDate] [datetime] NOT NULL DEFAULT (getdate()),
	[ExpDate] [datetime] NOT NULL DEFAULT (getdate()),
	[Packing] [nvarchar] (10)  NOT NULL DEFAULT ('BoxOrDrumOrPouch'),
	[Unit] [nvarchar] (10)  NOT NULL DEFAULT ('Count'),
	[Qty]  [numeric](18, 0) NOT NULL  DEFAULT (0),
	[Free] [numeric](18, 0) NOT NULL DEFAULT (0),
	[MRP] [float] NOT NULL DEFAULT (0.00),	
	[SRP] [float] NOT NULL DEFAULT (0.00),
	[Gross] [money] NOT NULL DEFAULT (0.00),
	[aDisc] [nvarchar] (50) NOT NULL DEFAULT ('0'),
	[aDiscAmt] [money] NOT NULL DEFAULT (0.00),
	[bDisc] [nvarchar] (50)  NOT NULL DEFAULT ('0'),
	[bDiscAmt] [money] NOT NULL DEFAULT (0.00),
	[cDisc] [nvarchar] (50)  NOT NULL DEFAULT ('0'),
	[cDiscAmt] [money] NOT NULL DEFAULT (0.00),
	[GST] [float] NOT NULL DEFAULT (0.00),
	[GSTAmt] [money] NOT NULL DEFAULT (0.00),
	[Cess] [float] NOT NULL DEFAULT (0.00),
	[CessAmt] [money] NOT NULL DEFAULT (0.00),
	[Amount] [money] NOT NULL DEFAULT (0.00),
	[Stock] [numeric](18, 0) NOT NULL DEFAULT (0)
) ON [PRIMARY]

CREATE TABLE [dbo].[PurchaseReturn] (
	[Serial] [decimal](18, 0) NOT NULL DEFAULT (0),
	[DBRefX] [decimal](18, 0) NOT NULL DEFAULT (0),
	[ItemID] [numeric](18, 0) NOT NULL DEFAULT (0),
	[ItemCode] [nvarchar] (20)  NOT NULL DEFAULT ('ItemCode'),
	[HSNCode]  [nvarchar] (20)  NOT NULL DEFAULT ('HSN_SAC_Code'),
	[ItemName] [nvarchar] (200)  NOT NULL DEFAULT ('ItemName'),
	[MakerAuthor] [nvarchar] (200)  NOT NULL DEFAULT ('MakerORAuthor'),
	[ProducerID] [nvarchar] (50)  NOT NULL DEFAULT ('P0001'),
	[ProducerName] [nvarchar] (200) NOT NULL DEFAULT ('ProducerName'),
	[Version] [nvarchar] (50) NOT NULL DEFAULT ('V1'),
	[MfgDate] [datetime] NOT NULL DEFAULT (getdate()),
	[ExpDate] [datetime] NOT NULL DEFAULT (getdate()),
	[Packing] [nvarchar] (10)  NOT NULL DEFAULT ('BoxOrDrumOrPouch'),
	[Unit] [nvarchar] (10)  NOT NULL DEFAULT ('Count'),
	[Qty]  [numeric](18, 0) NOT NULL  DEFAULT (0),
	[Free] [numeric](18, 0) NOT NULL DEFAULT (0),
	[MRP] [float] NOT NULL DEFAULT (0.00),	
	[SRP] [float] NOT NULL DEFAULT (0.00),
	[Gross] [money] NOT NULL DEFAULT (0.00),
	[aDisc] [nvarchar] (50) NOT NULL DEFAULT ('0'),
	[aDiscAmt] [money] NOT NULL DEFAULT (0.00),
	[bDisc] [nvarchar] (50)  NOT NULL DEFAULT ('0'),
	[bDiscAmt] [money] NOT NULL DEFAULT (0.00),
	[cDisc] [nvarchar] (50)  NOT NULL DEFAULT ('0'),
	[cDiscAmt] [money] NOT NULL DEFAULT (0.00),
	[GST] [float] NOT NULL DEFAULT (0.00),
	[GSTAmt] [money] NOT NULL DEFAULT (0.00),
	[Cess] [float] NOT NULL DEFAULT (0.00),
	[CessAmt] [money] NOT NULL DEFAULT (0.00),
	[Amount] [money] NOT NULL DEFAULT (0.00),
	[Stock] [numeric](18, 0) NOT NULL DEFAULT (0)
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[SaleReturn] (
	[Serial] [decimal](18, 0) NOT NULL DEFAULT (0),
	[DBRefX] [decimal](18, 0) NOT NULL DEFAULT (0),
	[ItemID] [numeric](18, 0) NOT NULL DEFAULT (0),
	[ItemCode] [nvarchar] (20)  NOT NULL DEFAULT ('ItemCode'),
	[HSNCode]  [nvarchar] (20)  NOT NULL DEFAULT ('HSN_SAC_Code'),
	[ItemName] [nvarchar] (200)  NOT NULL DEFAULT ('ItemName'),
	[MakerAuthor] [nvarchar] (200)  NOT NULL DEFAULT ('MakerORAuthor'),
	[ProducerID] [nvarchar] (50)  NOT NULL DEFAULT ('P0001'),
	[ProducerName] [nvarchar] (200) NOT NULL DEFAULT ('ProducerName'),
	[Version] [nvarchar] (50) NOT NULL DEFAULT ('V1'),
	[MfgDate] [datetime] NOT NULL DEFAULT (getdate()),
	[ExpDate] [datetime] NOT NULL DEFAULT (getdate()),
	[Packing] [nvarchar] (10)  NOT NULL DEFAULT ('BoxOrDrumOrPouch'),
	[Unit] [nvarchar] (10)  NOT NULL DEFAULT ('Count'),
	[Qty]  [numeric](18, 0) NOT NULL  DEFAULT (0),
	[Free] [numeric](18, 0) NOT NULL DEFAULT (0),
	[MRP] [float] NOT NULL DEFAULT (0.00),	
	[SRP] [float] NOT NULL DEFAULT (0.00),
	[Gross] [money] NOT NULL DEFAULT (0.00),
	[aDisc] [nvarchar] (50) NOT NULL DEFAULT ('0'),
	[aDiscAmt] [money] NOT NULL DEFAULT (0.00),
	[bDisc] [nvarchar] (50)  NOT NULL DEFAULT ('0'),
	[bDiscAmt] [money] NOT NULL DEFAULT (0.00),
	[cDisc] [nvarchar] (50)  NOT NULL DEFAULT ('0'),
	[cDiscAmt] [money] NOT NULL DEFAULT (0.00),
	[GST] [float] NOT NULL DEFAULT (0.00),
	[GSTAmt] [money] NOT NULL DEFAULT (0.00),
	[Cess] [float] NOT NULL DEFAULT (0.00),
	[CessAmt] [money] NOT NULL DEFAULT (0.00),
	[Amount] [money] NOT NULL DEFAULT (0.00),
	[Stock] [numeric](18, 0) NOT NULL DEFAULT (0)
) ON [PRIMARY]


CREATE TABLE [dbo].[TIN] (
	[Serial] [decimal](18, 0) NOT NULL DEFAULT (0),
	[DBRefX] [decimal](18, 0) NOT NULL DEFAULT (0),
	[ItemID] [numeric](18, 0) NOT NULL DEFAULT (0),
	[ItemCode] [nvarchar] (20)  NOT NULL DEFAULT ('ItemCode'),
	[HSNCode]  [nvarchar] (20)  NOT NULL DEFAULT ('HSN_SAC_Code'),
	[ItemName] [nvarchar] (200)  NOT NULL DEFAULT ('ItemName'),
	[MakerAuthor] [nvarchar] (200)  NOT NULL DEFAULT ('MakerORAuthor'),
	[ProducerID] [nvarchar] (50)  NOT NULL DEFAULT ('P0001'),
	[ProducerName] [nvarchar] (200) NOT NULL DEFAULT ('ProducerName'),
	[Version] [nvarchar] (50) NOT NULL DEFAULT ('V1'),
	[MfgDate] [datetime] NOT NULL DEFAULT (getdate()),
	[ExpDate] [datetime] NOT NULL DEFAULT (getdate()),
	[Packing] [nvarchar] (10)  NOT NULL DEFAULT ('BoxOrDrumOrPouch'),
	[Unit] [nvarchar] (10)  NOT NULL DEFAULT ('Count'),
	[Qty]  [numeric](18, 0) NOT NULL  DEFAULT (0),
	[Free] [numeric](18, 0) NOT NULL DEFAULT (0),
	[MRP] [float] NOT NULL DEFAULT (0.00),	
	[SRP] [float] NOT NULL DEFAULT (0.00),
	[Gross] [money] NOT NULL DEFAULT (0.00),
	[aDisc] [nvarchar] (50) NOT NULL DEFAULT ('0'),
	[aDiscAmt] [money] NOT NULL DEFAULT (0.00),
	[bDisc] [nvarchar] (50)  NOT NULL DEFAULT ('0'),
	[bDiscAmt] [money] NOT NULL DEFAULT (0.00),
	[cDisc] [nvarchar] (50)  NOT NULL DEFAULT ('0'),
	[cDiscAmt] [money] NOT NULL DEFAULT (0.00),
	[GST] [float] NOT NULL DEFAULT (0.00),
	[GSTAmt] [money] NOT NULL DEFAULT (0.00),
	[Cess] [float] NOT NULL DEFAULT (0.00),
	[CessAmt] [money] NOT NULL DEFAULT (0.00),
	[Amount] [money] NOT NULL DEFAULT (0.00),
	[Stock] [numeric](18, 0) NOT NULL DEFAULT (0)
) ON [PRIMARY]


CREATE TABLE [dbo].[TOUT] (
	[Serial] [decimal](18, 0) NOT NULL DEFAULT (0),
	[DBRefX] [decimal](18, 0) NOT NULL DEFAULT (0),
	[ItemID] [numeric](18, 0) NOT NULL DEFAULT (0),
	[ItemCode] [nvarchar] (20)  NOT NULL DEFAULT ('ItemCode'),
	[HSNCode]  [nvarchar] (20)  NOT NULL DEFAULT ('HSN_SAC_Code'),
	[ItemName] [nvarchar] (200)  NOT NULL DEFAULT ('ItemName'),
	[MakerAuthor] [nvarchar] (200)  NOT NULL DEFAULT ('MakerORAuthor'),
	[ProducerID] [nvarchar] (50)  NOT NULL DEFAULT ('P0001'),
	[ProducerName] [nvarchar] (200) NOT NULL DEFAULT ('ProducerName'),
	[Version] [nvarchar] (50) NOT NULL DEFAULT ('V1'),
	[MfgDate] [datetime] NOT NULL DEFAULT (getdate()),
	[ExpDate] [datetime] NOT NULL DEFAULT (getdate()),
	[Packing] [nvarchar] (10)  NOT NULL DEFAULT ('BoxOrDrumOrPouch'),
	[Unit] [nvarchar] (10)  NOT NULL DEFAULT ('Count'),
	[Qty]  [numeric](18, 0) NOT NULL  DEFAULT (0),
	[Free] [numeric](18, 0) NOT NULL DEFAULT (0),
	[MRP] [float] NOT NULL DEFAULT (0.00),	
	[SRP] [float] NOT NULL DEFAULT (0.00),
	[Gross] [money] NOT NULL DEFAULT (0.00),
	[aDisc] [nvarchar] (50) NOT NULL DEFAULT ('0'),
	[aDiscAmt] [money] NOT NULL DEFAULT (0.00),
	[bDisc] [nvarchar] (50)  NOT NULL DEFAULT ('0'),
	[bDiscAmt] [money] NOT NULL DEFAULT (0.00),
	[cDisc] [nvarchar] (50)  NOT NULL DEFAULT ('0'),
	[cDiscAmt] [money] NOT NULL DEFAULT (0.00),
	[GST] [float] NOT NULL DEFAULT (0.00),
	[GSTAmt] [money] NOT NULL DEFAULT (0.00),
	[Cess] [float] NOT NULL DEFAULT (0.00),
	[CessAmt] [money] NOT NULL DEFAULT (0.00),
	[Amount] [money] NOT NULL DEFAULT (0.00),
	[Stock] [numeric](18, 0) NOT NULL DEFAULT (0)
) ON [PRIMARY]


---###################################################################################
CREATE TABLE [dbo].[PMain] (
	[DBRef] [decimal](18, 0) IDENTITY (1, 1) NOT NULL ,
	[DBDate] [datetime] NOT NULL DEFAULT (getdate()),
	[Status] [nvarchar] (10) NOT NULL DEFAULT ('NEW'),
	[OrderRef] [nvarchar] (50) NOT NULL DEFAULT ('0'),
	[OrderDate] [datetime] NOT NULL DEFAULT (getdate()),
	[InvRef] [nvarchar] (50) NOT NULL DEFAULT ('InvRef'),
	[InvDate] [datetime] NOT NULL DEFAULT (getdate()),
	[ID] [nvarchar] (10) NOT NULL DEFAULT ('D0001'),
	[Name] [nvarchar] (100) NOT NULL DEFAULT ('Sample Seller'),
	[Address] [nvarchar] (200) NOT NULL DEFAULT ('Address'),
	[City] [nvarchar] (100) NOT NULL DEFAULT ('City'),
	[State] [nvarchar] (50) NOT NULL DEFAULT ('BR'),
	[Phones] [nvarchar] (50) NOT NULL DEFAULT ('Ph#'),
	[Email] [nvarchar] (100) NOT NULL DEFAULT ('r@example.com'),
	[ShipAddress] [nvarchar] (200) NOT NULL DEFAULT ('Ship.Address'),
	[ShipState] [nvarchar] (50) NOT NULL DEFAULT ('BR'),
	[TID] [nvarchar] (10) NOT NULL DEFAULT ('T0001'),
	[TName] [nvarchar] (100) NOT NULL DEFAULT ('Transporter.Name'),
	[TAddress] [nvarchar] (200) NOT NULL DEFAULT ('T.Address'),
	[TCity] [nvarchar] (100) NOT NULL DEFAULT ('T.City'),
	[TPhones] [nvarchar] (50) NOT NULL DEFAULT ('Ph#'),
	[TEmail] [nvarchar] (100) NOT NULL DEFAULT ('r@example'),
	[Terminal] [nvarchar] (50) NOT NULL DEFAULT ('Terminal1'),
	[KID] [nvarchar] (10) NOT NULL DEFAULT ('K0001'),
	[KName] [nvarchar] (100) NOT NULL DEFAULT ('K.Name'),
	[KRate] [money] NOT NULL DEFAULT (0.00),
	[KAmount] [money] NOT NULL DEFAULT (0.00),
	[PID] [nvarchar] (10) NOT NULL  DEFAULT ('O0001'),
	[PName] [nvarchar] (100) NOT NULL  DEFAULT ('Sample Postman or Courier'),
	[Postage] [money] NOT NULL DEFAULT (0.00),
	[SDID] [nvarchar] (10) NOT NULL DEFAULT('SD001'),
	[SDName] [nvarchar] (100) NOT NULL DEFAULT('SD.Name'),
	[SDCommission] [money] NOT NULL DEFAULT (0.00),
	[GRNo] [nvarchar] (50) NOT NULL DEFAULT ('to.change'),
	[GRDate] [datetime] NULL DEFAULT (getdate()),
	[GRMode] [nvarchar] (10) NOT NULL  DEFAULT ('to.change'),
	[GRAmount] [money] NOT NULL DEFAULT (0.00),
	[ToPayMode] [decimal](18, 0) NOT NULL DEFAULT (2),
	[BundleCount] [decimal](18, 0) NOT NULL DEFAULT (0),
	[BundleWeight] [float] NOT NULL DEFAULT (0.00),
	[ItemCount] [float] NOT NULL DEFAULT (0.00),
	[DiscAmt] [money] NOT NULL DEFAULT (0.00),
	[ItemAmount] [money] NOT NULL DEFAULT (0.00),
	[SplDisc] [float] NOT NULL DEFAULT (0.00),
	[BulkDisc] [float] NOT NULL DEFAULT (0.00),
	[AddMisc] [money] NOT NULL DEFAULT (0.00),
	[LessMisc] [money] NOT NULL DEFAULT (0.00),
	[AddFreight] [money] NOT NULL DEFAULT (0.00),
	[LessFreight] [money] NOT NULL DEFAULT (0.00),
	[NetGSTAmt]  [money] NOT NULL DEFAULT (0.00),
	[NetCGSTAmt] [money] NOT NULL DEFAULT (0.00),
	[NetSGSTAmt] [money] NOT NULL DEFAULT (0.00),
	[NetIGSTAmt] [money] NOT NULL DEFAULT (0.00),
	[NetCessAmt] [money] NOT NULL DEFAULT (0.00),
	[RoundOff] [money] NOT NULL DEFAULT (0.00),
	[NetAmount] [money] NOT NULL DEFAULT (0.00),
	[UserNo] [nvarchar] (200) NOT NULL DEFAULT ('U0001'),
	[Comments] [nvarchar] (200) NOT NULL DEFAULT ('Comments')
) ON [PRIMARY]


CREATE TABLE [dbo].[SMain] (
	[DBRef] [decimal](18, 0) IDENTITY (1, 1) NOT NULL ,
	[DBDate] [datetime] NOT NULL DEFAULT (getdate()),
	[Status] [nvarchar] (10) NOT NULL DEFAULT ('NEW'),
	[OrderRef] [nvarchar] (50) NOT NULL DEFAULT ('0'),
	[OrderDate] [datetime] NOT NULL DEFAULT (getdate()),
	[InvRef] [nvarchar] (50) NOT NULL DEFAULT ('InvRef'),
	[InvDate] [datetime] NOT NULL DEFAULT (getdate()),
	[ID] [nvarchar] (10) NOT NULL DEFAULT ('C0001'),
	[Name] [nvarchar] (100) NOT NULL DEFAULT ('Sample Buyer'),
	[Address] [nvarchar] (200) NOT NULL DEFAULT ('Address'),
	[City] [nvarchar] (100) NOT NULL DEFAULT ('City'),
	[State] [nvarchar] (50) NOT NULL DEFAULT ('BR'),
	[Phones] [nvarchar] (50) NOT NULL DEFAULT ('Ph#'),
	[Email] [nvarchar] (100) NOT NULL DEFAULT ('r@example.com'),
	[ShipAddress] [nvarchar] (200) NOT NULL DEFAULT ('Ship.Address'),
	[ShipState] [nvarchar] (50) NOT NULL DEFAULT ('BR'),
	[TID] [nvarchar] (10) NOT NULL DEFAULT ('T0001'),
	[TName] [nvarchar] (100) NOT NULL DEFAULT ('Transporter.Name'),
	[TAddress] [nvarchar] (200) NOT NULL DEFAULT ('T.Address'),
	[TCity] [nvarchar] (100) NOT NULL DEFAULT ('T.City'),
	[TPhones] [nvarchar] (50) NOT NULL DEFAULT ('Ph#'),
	[TEmail] [nvarchar] (100) NOT NULL DEFAULT ('r@example'),
	[Terminal] [nvarchar] (50) NOT NULL DEFAULT ('Terminal1'),
	[KID] [nvarchar] (10) NOT NULL DEFAULT ('K0001'),
	[KName] [nvarchar] (100) NOT NULL DEFAULT ('K.Name'),
	[KRate] [money] NOT NULL DEFAULT (0.00),
	[KAmount] [money] NOT NULL DEFAULT (0.00),
	[PID] [nvarchar] (10) NOT NULL  DEFAULT ('O0001'),
	[PName] [nvarchar] (100) NOT NULL  DEFAULT ('Sample Postman or Courier'),
	[Postage] [money] NOT NULL DEFAULT (0.00),
	[SDID] [nvarchar] (10) NOT NULL DEFAULT('SD001'),
	[SDName] [nvarchar] (100) NOT NULL DEFAULT('SD.Name'),
	[SDCommission] [money] NOT NULL DEFAULT (0.00),
	[GRNo] [nvarchar] (50) NOT NULL DEFAULT ('to.change'),
	[GRDate] [datetime] NULL DEFAULT (getdate()),
	[GRMode] [nvarchar] (10) NOT NULL  DEFAULT ('to.change'),
	[GRAmount] [money] NOT NULL DEFAULT (0.00),
	[ToPayMode] [decimal](18, 0) NOT NULL DEFAULT (2),
	[BundleCount] [decimal](18, 0) NOT NULL DEFAULT (0),
	[BundleWeight] [float] NOT NULL DEFAULT (0.00),
	[ItemCount] [float] NOT NULL DEFAULT (0.00),
	[DiscAmt] [money] NOT NULL DEFAULT (0.00),
	[ItemAmount] [money] NOT NULL DEFAULT (0.00),
	[SplDisc] [float] NOT NULL DEFAULT (0.00),
	[BulkDisc] [float] NOT NULL DEFAULT (0.00),
	[AddMisc] [money] NOT NULL DEFAULT (0.00),
	[LessMisc] [money] NOT NULL DEFAULT (0.00),
	[AddFreight] [money] NOT NULL DEFAULT (0.00),
	[LessFreight] [money] NOT NULL DEFAULT (0.00),
	[NetGSTAmt]  [money] NOT NULL DEFAULT (0.00),
	[NetCGSTAmt] [money] NOT NULL DEFAULT (0.00),
	[NetSGSTAmt] [money] NOT NULL DEFAULT (0.00),
	[NetIGSTAmt] [money] NOT NULL DEFAULT (0.00),
	[NetCessAmt] [money] NOT NULL DEFAULT (0.00),
	[RoundOff] [money] NOT NULL DEFAULT (0.00),
	[NetAmount] [money] NOT NULL DEFAULT (0.00),
	[UserNo] [nvarchar] (200) NOT NULL DEFAULT ('U0001'),
	[Comments] [nvarchar] (200) NOT NULL DEFAULT ('Comments')
) ON [PRIMARY]


CREATE TABLE [dbo].[PReturnMain] (
	[DBRef] [decimal](18, 0) IDENTITY (1, 1) NOT NULL ,
	[DBDate] [datetime] NOT NULL DEFAULT (getdate()),
	[Status] [nvarchar] (10) NOT NULL DEFAULT ('NEW'),
	[OrderRef] [nvarchar] (50) NOT NULL DEFAULT ('0'),
	[OrderDate] [datetime] NOT NULL DEFAULT (getdate()),
	[InvRef] [nvarchar] (50) NOT NULL DEFAULT ('InvRef'),
	[InvDate] [datetime] NOT NULL DEFAULT (getdate()),
	[ID] [nvarchar] (10) NOT NULL DEFAULT ('D0001'),
	[Name] [nvarchar] (100) NOT NULL DEFAULT ('Sample Seller'),
	[Address] [nvarchar] (200) NOT NULL DEFAULT ('Address'),
	[City] [nvarchar] (100) NOT NULL DEFAULT ('City'),
	[State] [nvarchar] (50) NOT NULL DEFAULT ('BR'),
	[Phones] [nvarchar] (50) NOT NULL DEFAULT ('Ph#'),
	[Email] [nvarchar] (100) NOT NULL DEFAULT ('r@example.com'),
	[ShipAddress] [nvarchar] (200) NOT NULL DEFAULT ('Ship.Address'),
	[ShipState] [nvarchar] (50) NOT NULL DEFAULT ('BR'),
	[TID] [nvarchar] (10) NOT NULL DEFAULT ('T0001'),
	[TName] [nvarchar] (100) NOT NULL DEFAULT ('Transporter.Name'),
	[TAddress] [nvarchar] (200) NOT NULL DEFAULT ('T.Address'),
	[TCity] [nvarchar] (100) NOT NULL DEFAULT ('T.City'),
	[TPhones] [nvarchar] (50) NOT NULL DEFAULT ('Ph#'),
	[TEmail] [nvarchar] (100) NOT NULL DEFAULT ('r@example'),
	[Terminal] [nvarchar] (50) NOT NULL DEFAULT ('Terminal1'),
	[KID] [nvarchar] (10) NOT NULL DEFAULT ('K0001'),
	[KName] [nvarchar] (100) NOT NULL DEFAULT ('K.Name'),
	[KRate] [money] NOT NULL DEFAULT (0.00),
	[KAmount] [money] NOT NULL DEFAULT (0.00),
	[PID] [nvarchar] (10) NOT NULL  DEFAULT ('O0001'),
	[PName] [nvarchar] (100) NOT NULL  DEFAULT ('Sample Postman or Courier'),
	[Postage] [money] NOT NULL DEFAULT (0.00),
	[SDID] [nvarchar] (10) NOT NULL DEFAULT('SD001'),
	[SDName] [nvarchar] (100) NOT NULL DEFAULT('SD.Name'),
	[SDCommission] [money] NOT NULL DEFAULT (0.00),
	[GRNo] [nvarchar] (50) NOT NULL  DEFAULT 'to.change' ,
	[GRDate] [datetime] NULL DEFAULT (getdate()),
	[GRMode] [nvarchar] (10) NOT NULL  DEFAULT ('to.change') ,
	[GRAmount] [money] NOT NULL DEFAULT (0.00),
	[ToPayMode] [decimal](18, 0) NOT NULL DEFAULT (2),
	[BundleCount] [decimal](18, 0) NOT NULL DEFAULT (0),
	[BundleWeight] [float] NOT NULL DEFAULT (0.00),
	[ItemCount] [float] NOT NULL DEFAULT (0.00),
	[DiscAmt] [money] NOT NULL DEFAULT (0.00),
	[ItemAmount] [money] NOT NULL DEFAULT (0.00),
	[SplDisc] [float] NOT NULL DEFAULT (0.00),
	[BulkDisc] [float] NOT NULL DEFAULT (0.00),
	[AddMisc] [money] NOT NULL DEFAULT (0.00),
	[LessMisc] [money] NOT NULL DEFAULT (0.00),
	[AddFreight] [money] NOT NULL DEFAULT (0.00),
	[LessFreight] [money] NOT NULL DEFAULT (0.00),
	[NetGSTAmt]  [money] NOT NULL DEFAULT (0.00),
	[NetCGSTAmt] [money] NOT NULL DEFAULT (0.00),
	[NetSGSTAmt] [money] NOT NULL DEFAULT (0.00),
	[NetIGSTAmt] [money] NOT NULL DEFAULT (0.00),
	[NetCessAmt] [money] NOT NULL DEFAULT (0.00),
	[RoundOff] [money] NOT NULL DEFAULT (0.00),
	[NetAmount] [money] NOT NULL DEFAULT (0.00),
	[UserNo] [nvarchar] (200) NOT NULL DEFAULT ('U0001'),
	[Comments] [nvarchar] (200) NOT NULL DEFAULT ('Comments')
) ON [PRIMARY]


CREATE TABLE [dbo].[SReturnMain] (
	[DBRef] [decimal](18, 0) IDENTITY (1, 1) NOT NULL ,
	[DBDate] [datetime] NOT NULL DEFAULT (getdate()),
	[Status] [nvarchar] (10) NOT NULL DEFAULT ('NEW'),
	[OrderRef] [nvarchar] (50) NOT NULL DEFAULT ('0'),
	[OrderDate] [datetime] NOT NULL DEFAULT (getdate()),
	[InvRef] [nvarchar] (50) NOT NULL DEFAULT ('InvRef'),
	[InvDate] [datetime] NOT NULL DEFAULT (getdate()),
	[ID] [nvarchar] (10) NOT NULL DEFAULT ('D0001'),
	[Name] [nvarchar] (100) NOT NULL DEFAULT ('Sample Seller'),
	[Address] [nvarchar] (200) NOT NULL DEFAULT ('Address'),
	[City] [nvarchar] (100) NOT NULL DEFAULT ('City'),
	[State] [nvarchar] (50) NOT NULL DEFAULT ('BR'),
	[Phones] [nvarchar] (50) NOT NULL DEFAULT ('Ph#'),
	[Email] [nvarchar] (100) NOT NULL DEFAULT ('r@example.com'),
	[ShipAddress] [nvarchar] (200) NOT NULL DEFAULT ('Ship.Address'),
	[ShipState] [nvarchar] (50) NOT NULL DEFAULT ('BR'),
	[TID] [nvarchar] (10) NOT NULL DEFAULT ('T0001'),
	[TName] [nvarchar] (100) NOT NULL DEFAULT ('Transporter.Name'),
	[TAddress] [nvarchar] (200) NOT NULL DEFAULT ('T.Address'),
	[TCity] [nvarchar] (100) NOT NULL DEFAULT ('T.City'),
	[TPhones] [nvarchar] (50) NOT NULL DEFAULT ('Ph#'),
	[TEmail] [nvarchar] (100) NOT NULL DEFAULT ('r@example'),
	[Terminal] [nvarchar] (50) NOT NULL DEFAULT ('Terminal1'),
	[KID] [nvarchar] (10) NOT NULL DEFAULT ('K0001'),
	[KName] [nvarchar] (100) NOT NULL DEFAULT ('K.Name'),
	[KRate] [money] NOT NULL DEFAULT (0.00),
	[KAmount] [money] NOT NULL DEFAULT (0.00),
	[PID] [nvarchar] (10) NOT NULL  DEFAULT ('O0001'),
	[PName] [nvarchar] (100) NOT NULL  DEFAULT ('Sample Postman or Courier'),
	[Postage] [money] NOT NULL DEFAULT (0.00),
	[SDID] [nvarchar] (10) NOT NULL DEFAULT('SD001'),
	[SDName] [nvarchar] (100) NOT NULL DEFAULT('SD.Name'),
	[SDCommission] [money] NOT NULL DEFAULT (0.00),
	[GRNo] [nvarchar] (50) NOT NULL  DEFAULT ('to.change') ,
	[GRDate] [datetime] NULL DEFAULT (getdate()),
	[GRMode] [nvarchar] (10) NOT NULL  DEFAULT ('to.change') ,
	[GRAmount] [money] NOT NULL DEFAULT (0.00),
	[ToPayMode] [decimal](18, 0) NOT NULL DEFAULT (2),
	[BundleCount] [decimal](18, 0) NOT NULL DEFAULT (0),
	[BundleWeight] [float] NOT NULL DEFAULT (0.00),
	[ItemCount] [float] NOT NULL DEFAULT (0.00),
	[DiscAmt] [money] NOT NULL DEFAULT (0.00),
	[ItemAmount] [money] NOT NULL DEFAULT (0.00),
	[SplDisc] [float] NOT NULL DEFAULT (0.00),
	[BulkDisc] [float] NOT NULL DEFAULT (0.00),
	[AddMisc] [money] NOT NULL DEFAULT (0.00),
	[LessMisc] [money] NOT NULL DEFAULT (0.00),
	[AddFreight] [money] NOT NULL DEFAULT (0.00),
	[LessFreight] [money] NOT NULL DEFAULT (0.00),
	[NetGSTAmt]  [money] NOT NULL DEFAULT (0.00),
	[NetCGSTAmt] [money] NOT NULL DEFAULT (0.00),
	[NetSGSTAmt] [money] NOT NULL DEFAULT (0.00),
	[NetIGSTAmt] [money] NOT NULL DEFAULT (0.00),
	[NetCessAmt] [money] NOT NULL DEFAULT (0.00),
	[RoundOff] [money] NOT NULL DEFAULT (0.00),
	[NetAmount] [money] NOT NULL DEFAULT (0.00),
	[UserNo] [nvarchar] (200) NOT NULL DEFAULT ('U0001'),
	[Comments] [nvarchar] (200) NOT NULL DEFAULT ('Comments') 
) ON [PRIMARY]


CREATE TABLE [dbo].[TINMAIN] (
	[DBRef] [decimal](18, 0) IDENTITY (1, 1) NOT NULL ,
	[DBDate] [datetime] NOT NULL DEFAULT (getdate()),
	[Status] [nvarchar] (10) NOT NULL DEFAULT ('NEW'),
	[OrderRef] [nvarchar] (50) NOT NULL DEFAULT ('0'),
	[OrderDate] [datetime] NOT NULL DEFAULT (getdate()),
	[InvRef] [nvarchar] (50) NOT NULL DEFAULT ('InvRef'),
	[InvDate] [datetime] NOT NULL DEFAULT (getdate()),
	[ID] [nvarchar] (10) NOT NULL DEFAULT ('D0001'),
	[Name] [nvarchar] (100) NOT NULL DEFAULT ('Sample Seller'),
	[Address] [nvarchar] (200) NOT NULL DEFAULT ('Address'),
	[City] [nvarchar] (100) NOT NULL DEFAULT ('City'),
	[State] [nvarchar] (50) NOT NULL DEFAULT ('BR'),
	[Phones] [nvarchar] (50) NOT NULL DEFAULT ('Ph#'),
	[Email] [nvarchar] (100) NOT NULL DEFAULT ('r@example.com'),
	[ShipAddress] [nvarchar] (200) NOT NULL DEFAULT ('Ship.Address'),
	[ShipState] [nvarchar] (50) NOT NULL DEFAULT ('BR'),
	[TID] [nvarchar] (10) NOT NULL DEFAULT ('T0001'),
	[TName] [nvarchar] (100) NOT NULL DEFAULT ('Transporter.Name'),
	[TAddress] [nvarchar] (200) NOT NULL DEFAULT ('T.Address'),
	[TCity] [nvarchar] (100) NOT NULL DEFAULT ('T.City'),
	[TPhones] [nvarchar] (50) NOT NULL DEFAULT ('Ph#'),
	[TEmail] [nvarchar] (100) NOT NULL DEFAULT ('r@example'),
	[Terminal] [nvarchar] (50) NOT NULL DEFAULT ('Terminal1'),
	[KID] [nvarchar] (10) NOT NULL DEFAULT ('K0001'),
	[KName] [nvarchar] (100) NOT NULL DEFAULT ('K.Name'),
	[KRate] [money] NOT NULL DEFAULT (0.00),
	[KAmount] [money] NOT NULL DEFAULT (0.00),
	[PID] [nvarchar] (10) NOT NULL  DEFAULT ('O0001'),
	[PName] [nvarchar] (100) NOT NULL  DEFAULT ('Sample Postman or Courier'),
	[Postage] [money] NOT NULL DEFAULT (0.00),
	[SDID] [nvarchar] (10) NOT NULL DEFAULT('SD001'),
	[SDName] [nvarchar] (100) NOT NULL DEFAULT('SD.Name'),
	[SDCommission] [money] NOT NULL DEFAULT (0.00),
	[GRNo] [nvarchar] (50) NOT NULL  DEFAULT ('to.change') ,
	[GRDate] [datetime] NULL DEFAULT (getdate()),
	[GRMode] [nvarchar] (10) NOT NULL  DEFAULT ('to.change') ,
	[GRAmount] [money] NOT NULL DEFAULT (0.00),
	[ToPayMode] [decimal](18, 0) NOT NULL DEFAULT (2),
	[BundleCount] [decimal](18, 0) NOT NULL DEFAULT (0),
	[BundleWeight] [float] NOT NULL DEFAULT (0.00),
	[ItemCount] [float] NOT NULL DEFAULT (0.00),
	[DiscAmt] [money] NOT NULL DEFAULT (0.00),
	[ItemAmount] [money] NOT NULL DEFAULT (0.00),
	[SplDisc] [float] NOT NULL DEFAULT (0.00),
	[BulkDisc] [float] NOT NULL DEFAULT (0.00),
	[AddMisc] [money] NOT NULL DEFAULT (0.00),
	[LessMisc] [money] NOT NULL DEFAULT (0.00),
	[AddFreight] [money] NOT NULL DEFAULT (0.00),
	[LessFreight] [money] NOT NULL DEFAULT (0.00),
	[NetGSTAmt]  [money] NOT NULL DEFAULT (0.00),
	[NetCGSTAmt] [money] NOT NULL DEFAULT (0.00),
	[NetSGSTAmt] [money] NOT NULL DEFAULT (0.00),
	[NetIGSTAmt] [money] NOT NULL DEFAULT (0.00),
	[NetCessAmt] [money] NOT NULL DEFAULT (0.00),
	[RoundOff] [money] NOT NULL DEFAULT (0.00),
	[NetAmount] [money] NOT NULL DEFAULT (0.00),
	[UserNo] [nvarchar] (200) NOT NULL DEFAULT ('U0001'),
	[Comments] [nvarchar] (200) NOT NULL DEFAULT ('Comments') 
) ON [PRIMARY]


CREATE TABLE [dbo].[TOUTMAIN] (
	[DBRef] [decimal](18, 0) IDENTITY (1, 1) NOT NULL ,
	[DBDate] [datetime] NOT NULL DEFAULT (getdate()),
	[Status] [nvarchar] (10) NOT NULL DEFAULT ('NEW'),
	[OrderRef] [nvarchar] (50) NOT NULL DEFAULT ('0'),
	[OrderDate] [datetime] NOT NULL DEFAULT (getdate()),
	[InvRef] [nvarchar] (50) NOT NULL DEFAULT ('InvRef'),
	[InvDate] [datetime] NOT NULL DEFAULT (getdate()),
	[ID] [nvarchar] (10) NOT NULL DEFAULT ('D0001'),
	[Name] [nvarchar] (100) NOT NULL DEFAULT ('Sample Seller'),
	[Address] [nvarchar] (200) NOT NULL DEFAULT ('Address'),
	[City] [nvarchar] (100) NOT NULL DEFAULT ('City'),
	[State] [nvarchar] (50) NOT NULL DEFAULT ('BR'),
	[Phones] [nvarchar] (50) NOT NULL DEFAULT ('Ph#'),
	[Email] [nvarchar] (100) NOT NULL DEFAULT ('r@example.com'),
	[ShipAddress] [nvarchar] (200) NOT NULL DEFAULT ('Ship.Address'),
	[ShipState] [nvarchar] (50) NOT NULL DEFAULT ('BR'),
	[TID] [nvarchar] (10) NOT NULL DEFAULT ('T0001'),
	[TName] [nvarchar] (100) NOT NULL DEFAULT ('Transporter.Name'),
	[TAddress] [nvarchar] (200) NOT NULL DEFAULT ('T.Address'),
	[TCity] [nvarchar] (100) NOT NULL DEFAULT ('T.City'),
	[TPhones] [nvarchar] (50) NOT NULL DEFAULT ('Ph#'),
	[TEmail] [nvarchar] (100) NOT NULL DEFAULT ('r@example'),
	[Terminal] [nvarchar] (50) NOT NULL DEFAULT ('Terminal1'),
	[KID] [nvarchar] (10) NOT NULL DEFAULT ('K0001'),
	[KName] [nvarchar] (100) NOT NULL DEFAULT ('K.Name'),
	[KRate] [money] NOT NULL DEFAULT (0.00),
	[KAmount] [money] NOT NULL DEFAULT (0.00),
	[PID] [nvarchar] (10) NOT NULL  DEFAULT ('O0001'),
	[PName] [nvarchar] (100) NOT NULL  DEFAULT ('Sample Postman or Courier'),
	[Postage] [money] NOT NULL DEFAULT (0.00),
	[SDID] [nvarchar] (10) NOT NULL DEFAULT('SD001'),
	[SDName] [nvarchar] (100) NOT NULL DEFAULT('SD.Name'),
	[SDCommission] [money] NOT NULL DEFAULT (0.00),
	[GRNo] [nvarchar] (50) NOT NULL  DEFAULT ('to.change') ,
	[GRDate] [datetime] NULL DEFAULT (getdate()),
	[GRMode] [nvarchar] (10) NOT NULL  DEFAULT ('to.change') ,
	[GRAmount] [money] NOT NULL DEFAULT (0.00),
	[ToPayMode] [decimal](18, 0) NOT NULL DEFAULT (2),
	[BundleCount] [decimal](18, 0) NOT NULL DEFAULT (0),
	[BundleWeight] [float] NOT NULL DEFAULT (0.00),
	[ItemCount] [float] NOT NULL DEFAULT (0.00),
	[DiscAmt] [money] NOT NULL DEFAULT (0.00),
	[ItemAmount] [money] NOT NULL DEFAULT (0.00),
	[SplDisc] [float] NOT NULL DEFAULT (0.00),
	[BulkDisc] [float] NOT NULL DEFAULT (0.00),
	[AddMisc] [money] NOT NULL DEFAULT (0.00),
	[LessMisc] [money] NOT NULL DEFAULT (0.00),
	[AddFreight] [money] NOT NULL DEFAULT (0.00),
	[LessFreight] [money] NOT NULL DEFAULT (0.00),
	[NetGSTAmt]  [money] NOT NULL DEFAULT (0.00),
	[NetCGSTAmt] [money] NOT NULL DEFAULT (0.00),
	[NetSGSTAmt] [money] NOT NULL DEFAULT (0.00),
	[NetIGSTAmt] [money] NOT NULL DEFAULT (0.00),
	[NetCessAmt] [money] NOT NULL DEFAULT (0.00),
	[RoundOff] [money] NOT NULL DEFAULT (0.00),
	[NetAmount] [money] NOT NULL DEFAULT (0.00),
	[UserNo] [nvarchar] (200) NOT NULL DEFAULT ('U0001'),
	[Comments] [nvarchar] (200) NOT NULL DEFAULT ('Comments')
) ON [PRIMARY]


CREATE TABLE [dbo].[XBD] (
	[ItemID] [numeric](18, 0) NOT NULL DEFAULT (0),
	[ProducerID] [nvarchar] (50)  NULL DEFAULT ('P0001'),
	[ProducerName] [nvarchar] (200)  NULL DEFAULT ('ProducerName'),
	[PDisc] [nvarchar] (50) NOT NULL DEFAULT ('0'),
	[SDisc] [nvarchar] (50) NOT NULL DEFAULT ('0'),
	[Disc] [nvarchar] (50) NOT NULL DEFAULT ('0'),
) ON [PRIMARY]
GO

---###################################################################################
--- uits_gst_db VIEWS                     ############################################
---###################################################################################
CREATE VIEW SABIG 
AS
SELECT M.*,S.* FROM SMAIN M, SALE S WHERE M.DBREF=S.DBREFX
GO

CREATE VIEW SRBIG 
AS
SELECT M.*,S.* FROM SRETURNMAIN M, SALERETURN S WHERE M.DBREF=S.DBREFX
GO

CREATE VIEW PUBIG 
AS
SELECT M.*,S.* FROM PMAIN M, PURCHASE S WHERE M.DBREF=S.DBREFX
GO

CREATE VIEW PRBIG 
AS
SELECT M.*,S.* FROM PRETURNMAIN M, PURCHASERETURN S WHERE M.DBREF=S.DBREFX
GO

CREATE VIEW TIBIG 
AS
SELECT M.*,S.* FROM TINMAIN M, TIN S WHERE M.DBREF=S.DBREFX
GO

CREATE VIEW TOBIG 
AS
SELECT M.*,S.* FROM TOUTMAIN M, TOUT S WHERE M.DBREF=S.DBREFX
GO

---###################################################################################
CREATE VIEW STOCK_FULL AS
(
SELECT A.Version, A.ItemID, A.INITSTOCK AS QTY FROM Items A
UNION ALL
SELECT A.Version, A.ItemID, A.QTY+A.FREE AS QTY FROM PURCHASE A, PMAIN B WHERE (B.STATUS='CASH' OR B.STATUS='CREDIT' OR B.STATUS='CHALLAN') AND B.DBREF=A.DBREFX AND ( A.Version<>'DAMAGE'  AND A.Version<>'OLD')
UNION ALL
SELECT A.Version, A.ItemID, -(A.QTY+A.FREE) AS QTY FROM PURCHASERETURN A, PRETURNMAIN B WHERE (B.STATUS='CASH' OR B.STATUS='CREDIT' OR B.STATUS='CHALLAN')  AND B.DBREF=A.DBREFX AND ( A.Version<>'DAMAGE'  AND A.Version<>'OLD')
UNION ALL
SELECT A.Version, A.ItemID, -(A.QTY+A.FREE) AS QTY FROM SALE A, SMAIN B WHERE (B.STATUS='CASH' OR B.STATUS='CREDIT' OR B.STATUS='CHALLAN')  AND B.DBREF=A.DBREFX AND ( A.Version<>'DAMAGE'  AND A.Version<>'OLD')
UNION ALL
SELECT A.Version, A.ItemID, A.QTY+A.FREE AS QTY FROM SALERETURN A,  SRETURNMAIN B  WHERE (B.STATUS='CASH' OR B.STATUS='CREDIT' OR B.STATUS='CHALLAN')  AND B.DBREF=A.DBREFX AND ( A.Version<>'DAMAGE'  AND A.Version<>'OLD')
UNION ALL
SELECT A.Version, A.ItemID, -(A.QTY+A.FREE) AS QTY FROM TOUT A, TOUTMAIN B WHERE (B.STATUS='CASH' OR B.STATUS='CREDIT' OR B.STATUS='CHALLAN')  AND B.DBREF=A.DBREFX AND ( A.Version<>'DAMAGE'  AND A.Version<>'OLD')
UNION ALL
SELECT A.Version,  A.ItemID, A.QTY+A.FREE AS QTY FROM TIN A,  TINMAIN B  WHERE (B.STATUS='CASH' OR B.STATUS='CREDIT' OR B.STATUS='CHALLAN') AND B.DBREF=A.DBREFX AND ( A.Version<>'DAMAGE'  AND A.Version<>'OLD')
)
GO

CREATE VIEW STOCK_FULL_DATEWISE AS
(
SELECT '01-JAN-2000' as StockDate, A.Version, A.ItemID, A.ItemName, A.CurrMRP as MRP, A.ProducerName, A.INITSTOCK AS QTY FROM Items A
UNION ALL
SELECT B.DBDate as StockDate, A.Version, A.ItemID, A.ItemName, A.MRP, A.ProducerName, A.QTY+A.FREE AS QTY FROM PURCHASE A, PMAIN B WHERE (B.STATUS='CASH' OR B.STATUS='CREDIT' OR B.STATUS='CHALLAN') AND B.DBREF=A.DBREFX AND ( A.Version<>'DAMAGE'  AND A.Version<>'OLD')
UNION ALL
SELECT B.DBDate as StockDate, A.Version, A.ItemID, A.ItemName, A.MRP, A.ProducerName, -(A.QTY+A.FREE) AS QTY FROM PURCHASERETURN A, PRETURNMAIN B WHERE (B.STATUS='CASH' OR B.STATUS='CREDIT' OR B.STATUS='CHALLAN')  AND B.DBREF=A.DBREFX AND ( A.Version<>'DAMAGE'  AND A.Version<>'OLD')
UNION ALL
SELECT B.DBDate as StockDate, A.Version, A.ItemID, A.ItemName, A.MRP, A.ProducerName, -(A.QTY+A.FREE) AS QTY FROM SALE A, SMAIN B WHERE (B.STATUS='CASH' OR B.STATUS='CREDIT' OR B.STATUS='CHALLAN')  AND B.DBREF=A.DBREFX AND ( A.Version<>'DAMAGE'  AND A.Version<>'OLD')
UNION ALL
SELECT B.DBDate as StockDate, A.Version, A.ItemID, A.ItemName, A.MRP, A.ProducerName, A.QTY+A.FREE AS QTY FROM SALERETURN A,  SRETURNMAIN B  WHERE (B.STATUS='CASH' OR B.STATUS='CREDIT' OR B.STATUS='CHALLAN')  AND B.DBREF=A.DBREFX AND ( A.Version<>'DAMAGE'  AND A.Version<>'OLD')
UNION ALL
SELECT B.DBDate as StockDate, A.Version, A.ItemID, A.ItemName, A.MRP, A.ProducerName, -(A.QTY+A.FREE) AS QTY FROM TOUT A, TOUTMAIN B WHERE (B.STATUS='CASH' OR B.STATUS='CREDIT' OR B.STATUS='CHALLAN')  AND B.DBREF=A.DBREFX AND ( A.Version<>'DAMAGE'  AND A.Version<>'OLD')
UNION ALL
SELECT B.DBDate as StockDate, A.Version,  A.ItemID, A.ItemName, A.MRP, A.ProducerName, A.QTY+A.FREE AS QTY FROM TIN A,  TINMAIN B  WHERE (B.STATUS='CASH' OR B.STATUS='CREDIT' OR B.STATUS='CHALLAN') AND B.DBREF=A.DBREFX AND ( A.Version<>'DAMAGE'  AND A.Version<>'OLD')
)
GO

CREATE VIEW appview_Stock AS 
(SELECT ItemID, SUM(QTY) AS AVLBL FROM STOCK_FULL  GROUP BY ItemID)
GO

---###################################################################################
CREATE VIEW EXTVIEW_JOURNAL AS 
(
SELECT SERIAL, DATE, DRAC AS AC, DRNAME AS NAME, AMOUNT, 'DR' AS DRCR FROM JOURNAL
UNION ALL
SELECT SERIAL, DATE, CRAC AS AC, CRNAME AS NAME, AMOUNT, 'CR' AS DRCR FROM JOURNAL
)
GO

CREATE VIEW EXTVIEW_LEDGER AS 
SELECT MONTH(DATE) AS MONTH, AC, NAME, SUM(AMOUNT) AS AMOUNT, DRCR FROM EXTVIEW_JOURNAL GROUP BY MONTH(DATE), AC, NAME, DRCR
GO

CREATE VIEW EXTVIEW_ProducerWISESALEVOLUME AS 
SELECT DATEDIFF(D, DBDATE, GETDATE()) AS DaysBeforeToday, ProducerNAME, COUNT(ProducerNAME) as ProducerCount FROM SABIG 
GROUP BY ProducerNAME, DATEDIFF(D, DBDATE, GETDATE())
GO

CREATE VIEW JOURNALNEW AS 
(
SELECT DATE, DAY(DATE) AS DD, MONTH(DATE) AS MM, YEAR(DATE) AS YYYY, DRAC, DRNAME, AMOUNT, 'DR' AS TYPE, MEMOREF FROM JOURNAL 
UNION ALL
SELECT DATE, DAY(DATE) AS DD, MONTH(DATE) AS MM, YEAR(DATE) AS YYYY, CRAC, CRNAME, AMOUNT, 'CR' AS TYPE, MEMOREF FROM JOURNAL 
)
GO

CREATE VIEW appview_AllAccounts AS
SELECT ID, Name, Address, City, State, Phones, Email, ShipAddress, TID, KID, PID, ACCOUNT, OBDATE, OB, CreditLimit, GRP, TYPE FROM PERSONAL WHERE ID NOT LIKE 'P%'
GO

CREATE VIEW appview_Damages AS
(
SELECT 'SA' as Type, * FROM SABIG WHERE Version='DAMAGE' 
UNION ALL
SELECT 'SR' as Type, * FROM SRBIG WHERE Version='DAMAGE' 
UNION ALL
SELECT 'PU' as Type, * FROM PUBIG WHERE Version='DAMAGE' 
UNION ALL
SELECT 'PR' as Type, * FROM PRBIG WHERE Version='DAMAGE' 
UNION ALL
SELECT 'TI' as Type, * FROM TIBIG WHERE Version='DAMAGE' 
UNION ALL
SELECT 'TO' as Type, * FROM TOBIG WHERE Version='DAMAGE' 
)
GO

CREATE VIEW appview_Journal AS 
SELECT * FROM JOURNAL 
GO

CREATE VIEW appview_Old AS
(
SELECT 'SA' as Type, * FROM SABIG WHERE Version='OLD' 
UNION ALL
SELECT 'SR' as Type, * FROM SRBIG WHERE Version='OLD' 
UNION ALL
SELECT 'PU' as Type, * FROM PUBIG WHERE Version='OLD' 
UNION ALL
SELECT 'PR' as Type, * FROM PRBIG WHERE Version='OLD' 
UNION ALL
SELECT 'TI' as Type, * FROM TIBIG WHERE Version='OLD' 
UNION ALL
SELECT 'TO' as Type, * FROM TOBIG WHERE Version='OLD' 
)
GO

---###################################################################################
CREATE VIEW appview_PMain_Select_View AS
SELECT DBRef, DBDate, Status, ID, Name, City, ItemCount, BundleCount, TName, NetAmount, GRNo, InvRef, Comments, UserNo FROM PMAIN 
GO

CREATE VIEW appview_PReturnMain_Select_View AS
SELECT DBRef, DBDate, Status, ID, Name, City, ItemCount, BundleCount, TName, NetAmount, GRNo, InvRef, Comments, UserNo FROM PRETURNMAIN 
GO

CREATE VIEW appview_SMain_Select_View AS
SELECT DBRef, DBDate, Status, ID, Name, City, ItemCount, BundleCount, TName,ItemAmount, NetAmount, GRNo, InvRef, Comments, UserNo FROM SMAIN 
GO

CREATE VIEW appview_SReturnMain_Select_View AS
SELECT DBRef, DBDate, Status, ID, Name, City, ItemCount, BundleCount, TName, NetAmount, GRNo, InvRef, Comments, UserNo FROM SRETURNMAIN 
GO

--- ### USE FOLLOWING TWO VIEWS FOR PURCHASE/ PURCHASERETURN/ SALE/ SALERETURN/ TIN/ TOUT
CREATE VIEW appview_SelectItemPurchase AS
SELECT 0 AS Serial, 0 AS DBRefX, B.ItemID, B.ItemCode, B.HSNCode, B.ItemName, B.MakerAuthor, B.ProducerID, B.ProducerName, B.Version, B.MfgDate, B.ExpDate, B.Packing, B.Unit, 0 AS Qty, 0 AS Free, B.CurrMRP*C.CurrPrice as MRP, B.CurrSRP*C.CurrPrice as SRP, 0 as Gross, B.PDisc as aDisc,  0.00 as aDiscAmt, 0 as bDisc,  0.00 as bDiscAmt, 0 as cDisc,  0.00 as cDiscAmt, B.GST AS GST, 0.00 AS GSTAmt,  B.Cess AS Cess, 0.00 AS CessAmt, 0.00 AS Amount, A.AVLBL as Stock
FROM appview_Stock A, Items B, Currency C
WHERE A.ItemID = B.ItemID
AND B.CURRENCY=C.CURRENCY
GO

CREATE VIEW appview_SelectItemSale AS
SELECT 0 AS Serial, 0 AS DBRefX, B.ItemID, B.ItemCode, B.HSNCode, B.ItemName, B.MakerAuthor, B.ProducerID, B.ProducerName, B.Version, B.MfgDate, B.ExpDate, B.Packing, B.Unit, 0 AS Qty, 0 AS Free, B.CurrMRP*C.CurrPrice as MRP, B.CurrSRP*C.CurrPrice as SRP, 0 as Gross, B.SDisc as aDisc,  0.00 as aDiscAmt, 0 as bDisc,  0.00 as bDiscAmt, 0 as cDisc,  0.00 as cDiscAmt, B.GST AS GST, 0.00 AS GSTAmt,  B.Cess AS Cess, 0.00 AS CessAmt, 0.00 AS Amount, A.AVLBL as Stock
FROM appview_Stock A, Items B, Currency C
WHERE A.ItemID = B.ItemID
AND B.CURRENCY=C.CURRENCY
GO

---###################################################################################
CREATE VIEW appview_StockDamage AS 
(SELECT ItemID, SUM(QTY) AS AVLBL FROM STOCK_FULL  WHERE Version='DAMAGE' OR Version='OLD' GROUP BY ItemID)
GO

CREATE VIEW appview_StockExtended AS
SELECT 	B.ItemID, B.ItemCode, B.HSNCode, B.ItemNAME, B.ProducerID, B.ProducerNAME, B.CurrMRP*C.CurrPrice AS MRP, 
		B.PDisc, B.SDisc, S.AVLBL, '0' AS PSTOCKVAL, '0'  AS SSTOCKVAL  
FROM 		Items B, APPVIEW_STOCK S, CURRENCY C
WHERE 	B.ItemID=S.ItemID AND B.CURRENCY=C.CURRENCY
GO

CREATE VIEW appview_TINMain_Select_View AS
SELECT DBREF, DBDATE, STATUS, ID, NAME, CITY, GRNO, GRDATE, BUNDLECOUNT, BUNDLEWEIGHT, INVREF, INVDATE, NETAMOUNT, TNAME, GRAMOUNT, TOPAYMODE, COMMENTS, USERNO FROM TINMAIN 
GO

CREATE VIEW appview_TOUTMain_Select_View AS
SELECT DBREF, DBDATE, STATUS, ID, NAME, CITY, GRNO, GRDATE, BUNDLECOUNT, BUNDLEWEIGHT, INVREF, INVDATE, NETAMOUNT, TNAME, GRAMOUNT, TOPAYMODE, COMMENTS, USERNO  FROM TOUTMAIN 
GO

---###################################################################################
--- uits_gst_db PROCEDURES                 ##############################################
---###################################################################################
CREATE PROC ADDNEW_PMAIN  @USERNO nvarchar(10)='UNKN' AS
BEGIN
	INSERT INTO PMAIN (DBDate,UserNo) VALUES (GETDATE(), @USERNO)
	SELECT MAX(DBRef) AS DBRef FROM PMAIN
END
GO

CREATE PROC ADDNEW_PMT AS
BEGIN
	INSERT INTO PMT (Date) VALUES (getdate())
	SELECT MAX(Serial) AS MaxSerial FROM PMT
END
GO

CREATE PROC ADDNEW_PRETURNMAIN  @USERNO nvarchar(10)='UNKN' AS
BEGIN
	INSERT INTO PRETURNMAIN (DBDate,UserNo) VALUES (GETDATE(), @USERNO)
	SELECT MAX(DBRef) AS DBRef FROM PRETURNMAIN 
END
GO

CREATE PROC ADDNEW_RCT AS
BEGIN
	INSERT INTO RCT (Date) VALUES (getdate())
	SELECT MAX(Serial) AS MaxSerial FROM RCT
END
GO

CREATE PROC ADDNEW_SMAIN @USERNO nvarchar(10)='UNKN' AS
BEGIN
	INSERT INTO SMAIN (DBDate,UserNo) VALUES (GETDATE(), @USERNO)
	SELECT MAX(DBRef) AS DBRef FROM SMAIN
END
GO

CREATE PROC ADDNEW_SRETURNMAIN  @USERNO nvarchar(10)='UNKN' AS
BEGIN
	INSERT INTO SRETURNMAIN (DBDate,UserNo) VALUES (GETDATE(), @USERNO)
	SELECT MAX(DBRef) AS DBRef FROM SRETURNMAIN
END
GO

CREATE PROC ADDNEW_TINMAIN @USERNO nvarchar(10)='UNKN' AS
BEGIN
	INSERT INTO TINMAIN (DBDate,UserNo) VALUES (GETDATE(), @USERNO)
	SELECT MAX(DBRef) AS DBRef FROM TINMAIN
END
GO

CREATE PROC ADDNEW_TOUTMAIN @USERNO nvarchar(10)='UNKN' AS
BEGIN
	INSERT INTO TOUTMAIN (DBDate,UserNo) VALUES (GETDATE(), @USERNO)
	SELECT MAX(DBRef) AS DBRef FROM TOUTMAIN
END
GO

CREATE PROC ADDNEW_TRF AS
BEGIN
	INSERT INTO TRF (Date) VALUES (getdate())
	SELECT MAX(Serial) AS MaxSerial FROM TRF
END
GO

CREATE PROC ADDNEW_USER  @LoginID varchar(20) AS
BEGIN
	INSERT INTO Users (LoginID) VALUES (@LoginID)
	SELECT * FROM Users WHERE LOGINID=@LOGINID
END
GO

CREATE PROC ADDNEW_VOUCHERS @VType as nvarchar(10) AS
BEGIN
	INSERT INTO VOUCHERS (Date, VType) VALUES (getdate(), @VType)
	SELECT MAX(Serial) AS MaxSerial FROM VOUCHERS
END
GO

CREATE PROC appproc_CreateItemID_UPTO @ItemID NUMERIC, @ItemCode VARCHAR(10)='PATNA' AS
BEGIN
WHILE ((SELECT MAX(ItemID) FROM Items) < @ItemID) INSERT INTO Items (ItemCode) VALUES (@ItemCode)
END
GO


CREATE PROC appproc_ReportAllTransactions @DT DATETIME, @Opt Varchar(1)='D' AS
BEGIN
IF @Opt='D' 
    SELECT 'SALE' AS TYPE, STATUS, DBDATE, DBREF AS DBR, ID, NAME, CITY, NETAMOUNT FROM SMAIN WHERE DATEDIFF(D, DBDATE, @DT)=0
    Union All
    SELECT 'SRET'AS TYPE, STATUS, DBDATE, DBREF AS DBR, ID, NAME, CITY, NETAMOUNT FROM SRETURNMAIN WHERE DATEDIFF(D,DBDATE, @DT)=0
    Union All
    SELECT 'PURC'AS TYPE, STATUS, DBDATE, DBREF AS DBR, ID, NAME, CITY, NETAMOUNT FROM PMAIN WHERE DATEDIFF(D,DBDATE, @DT)=0
    Union All
    SELECT 'PRET'AS TYPE, STATUS, DBDATE, DBREF AS DBR, ID, NAME, CITY, NETAMOUNT FROM PRETURNMAIN WHERE DATEDIFF(D,DBDATE, @DT)=0
    Union All
    SELECT 'TIN.' AS TYPE, STATUS, DBDATE, DBREF AS DBR, ID, NAME, CITY, NETAMOUNT FROM TINMAIN WHERE DATEDIFF(D,DBDATE, @DT)=0
    Union All
    SELECT 'TOUT' AS TYPE, STATUS, DBDATE, DBREF AS DBR, ID, NAME, CITY, NETAMOUNT FROM TOUTMAIN WHERE DATEDIFF(D,DBDATE, @DT)=0
    Union All
    SELECT 'PYMT' AS TYPE, MODE AS STATUS, DATE, SERIAL, ID, NAME, '', AMOUNT FROM PMT WHERE DATEDIFF(D,DATE, @DT)=0
    Union All
    SELECT 'RCPT' AS TYPE, MODE AS STATUS, DATE, SERIAL, ID, NAME, '', AMOUNT FROM RCT WHERE DATEDIFF(D,DATE, @DT)=0
    ORDER BY 1, 2,3
Else
    SELECT 'SALE' AS TYPE, STATUS, DBDATE, DBREF AS DBR, ID, NAME, CITY, NETAMOUNT FROM SMAIN WHERE DATEDIFF(M, DBDATE, @DT)=0
    Union All
    SELECT 'SRET' AS TYPE, STATUS, DBDATE, DBREF AS DBR, ID, NAME, CITY, NETAMOUNT FROM SRETURNMAIN WHERE DATEDIFF(M,DBDATE, @DT)=0
    Union All
    SELECT 'PURC' AS TYPE, STATUS, DBDATE, DBREF AS DBR, ID, NAME, CITY, NETAMOUNT FROM PMAIN WHERE DATEDIFF(M,DBDATE, @DT)=0
    Union All
    SELECT 'PRET' AS TYPE, STATUS, DBDATE, DBREF AS DBR, ID, NAME, CITY, NETAMOUNT FROM PRETURNMAIN WHERE DATEDIFF(M,DBDATE, @DT)=0
    Union All
    SELECT 'TIN.' AS TYPE, STATUS, DBDATE, DBREF AS DBR, ID, NAME, CITY, NETAMOUNT FROM TINMAIN WHERE DATEDIFF(M,DBDATE, @DT)=0
    Union All
    SELECT 'TOUT' AS TYPE, STATUS, DBDATE, DBREF AS DBR, ID, NAME, CITY, NETAMOUNT FROM TOUTMAIN WHERE DATEDIFF(M,DBDATE, @DT)=0
    Union All
    SELECT 'PYMT' AS TYPE, MODE AS STATUS, DATE, SERIAL, ID, NAME, '', AMOUNT FROM PMT WHERE DATEDIFF(M,DATE, @DT)=0
    Union All
    SELECT 'RCPT' AS TYPE, MODE AS STATUS, DATE, SERIAL, ID, NAME, '', AMOUNT FROM RCT WHERE DATEDIFF(M,DATE, @DT)=0
    ORDER BY 1, 2,3
END
GO

CREATE PROC appproc_StockReport AS
BEGIN
SELECT B.ItemID, B.ItemNAME, B.MakerAuthor, B.ProducerNAME, B.CurrMRP as MRP, S.AVLBL FROM Items B, appview_Stock S WHERE B.ItemID=S.ItemID
ORDER BY B.ProducerNAME, B.ItemNAME, B.ItemID
END
GO

CREATE PROC appproc_ShowMergedStock @ItemID numeric AS
BEGIN
SELECT SUM(AVLBL) AS AVLBL, COUNT(ItemID) AS CountOfItemID FROM appview_Stock 
WHERE 
ItemID IN (SELECT ItemID FROM Items WHERE ItemID1 =((SELECT ItemID1 FROM Items WHERE ItemID=@ItemID)))
SELECT ItemID, ItemCode, ItemNAME, ProducerNAME, CurrMRP as MRP FROM Items WHERE ItemID1 =((SELECT ItemID1 FROM Items WHERE ItemID=@ItemID))
END
GO

CREATE PROC UPDATE_PMT @Serial As Numeric, @Date AS datetime, @ID AS nvarchar(5), @Name as nvarchar(100), @City as nvarchar(50), @Amount as decimal, @Mode as nvarchar(10), @ModeName as nvarchar(50), @Narration as nvarchar(100) AS
BEGIN
	UPDATE PMT SET Date=@Date, ID=@ID, Name=@Name, City=@City, Amount=@Amount, Mode=@Mode, ModeName=@ModeName, Narration=@Narration
	WHERE Serial=@Serial
END
GO

CREATE PROC UPDATE_RCT @Serial As Numeric, @Date AS datetime, @ID AS nvarchar(5), @Name as nvarchar(100), @City as nvarchar(50), @Amount as decimal, @Mode as nvarchar(10), @ModeName as nvarchar(50), @Narration as nvarchar(100) AS
BEGIN
	UPDATE RCT SET Date=@Date, ID=@ID, Name=@Name, City=@City, Amount=@Amount, Mode=@Mode, ModeName=@ModeName, Narration=@Narration
	WHERE Serial=@Serial
END
GO

CREATE PROC UPDATE_TRF @Serial As Numeric, @Date AS datetime, @DRID AS nvarchar(5), @DRName as nvarchar(50), @CRID AS nvarchar(5), @CRName as nvarchar(50), @Amount as decimal, @Narration as nvarchar(100) AS
BEGIN
	UPDATE TRF SET Date=@Date, DRID=@DRID, DRName=@DRName, CRID=@CRID, CRName=@CRName, Amount=@Amount, Narration=@Narration
	WHERE Serial=@Serial
END
GO

CREATE PROC UPDATE_VOUCHERS @Serial As Numeric, @Date AS datetime, 
@DrID AS nvarchar(5), @DrName as nvarchar(100), @DrCity as nvarchar(50), 
@CrID AS nvarchar(5), @CrName as nvarchar(100), @CrCity as nvarchar(50), 
@Amount as decimal, @Narration as nvarchar(100) , @VType as nvarchar(10)
AS
BEGIN
	UPDATE VOUCHERS SET Date=@Date, 
	DrID=@DrID, DrName=@DrName, DrCity=@DrCity, 
	CrID=@CrID, CrName=@CrName, CrCity=@CrCity, 
	Amount=@Amount, Narration=@Narration, VType=@VType
	WHERE Serial=@Serial
END
GO

CREATE PROC appproc_ItemReturnEnhancer @ID nvarchar(10),  @PubID  nvarchar(10), @TBL nvarchar(20),  @DT1 AS datetime, @DT2 AS datetime AS
DECLARE @BALANCE FLOAT
DECLARE @DT DATETIME

IF @DT1 > @DT2
BEGIN
	SET @DT=@DT1
	SET @DT1=@DT2
	SET @DT2=@DT
END

IF @TBL = 'SABIG'
BEGIN
SELECT(isnull((SELECT SUM(Amount) AS Amt FROM SABIG where ID=@ID AND ProducerID=@PubID AND (INVDATE BETWEEN @DT1 AND @DT2)  AND STATUS<>'CANCELLED'    ),0) )
END

IF @TBL = 'SRBIG'
BEGIN
SELECT(isnull((SELECT SUM(Amount) AS Amt FROM SRBIG where ID=@ID AND ProducerID=@PubID AND (INVDATE BETWEEN @DT1 AND @DT2)  AND STATUS<>'CANCELLED'    ),0) )
END

IF @TBL = 'PUBIG'
BEGIN
SELECT(isnull((SELECT SUM(Amount) AS Amt FROM PUBIG where ID=@ID AND ProducerID=@PubID AND (INVDATE BETWEEN @DT1 AND @DT2)  AND STATUS<>'CANCELLED'    ),0) )
END

IF @TBL = 'PRBIG'
BEGIN
SELECT(isnull((SELECT SUM(Amount) AS Amt FROM PRBIG where ID=@ID AND ProducerID=@PubID AND (INVDATE BETWEEN @DT1 AND @DT2)  AND STATUS<>'CANCELLED'    ),0) )
END
GO

CREATE PROC appproc_LEDGER @ID nvarchar(10), @DT1 AS datetime, @DT2 AS datetime AS
DECLARE @BALANCE FLOAT
DECLARE @DT DATETIME

IF @DT1 > @DT2
BEGIN
	SET @DT=@DT1
	SET @DT1=@DT2
	SET @DT2=@DT
END

SET @BALANCE = (SELECT(
isnull((SELECT SUM(OB) FROM appview_AllAccounts WHERE ID=@id),0)  
+
isnull((SELECT SUM(AMOUNT) FROM JOURNAL WHERE DRAC=@id AND AUTHDATE < @DT1),0)  
- 
isnull((SELECT SUM(AMOUNT) FROM JOURNAL WHERE CRAC=@id AND AUTHDATE < @DT1),0)
))
IF @BALANCE>=0 
	SELECT @DT1 AS DATE, "By BALANCE " AS PARTICULARS , @BALANCE AS DR, 0 AS CR, @BALANCE AS BALANCE, 'X' AS MEMOREF, '-' as SATISFACTION, @BALANCE AS DRSUM, 0 AS CRSUM, 'OPENING BAL' as Narration
	UNION ALL
	SELECT AUTHDATE AS DT, "To " + CRNAME + " (" + MemoRef + ")" AS NAME, AMOUNT AS DR, 0 AS CR, AMOUNT-0 AS BALANCE, MEMOREF, '-' as SATISFACTION, AMOUNT AS DRSUM, 0 AS CRSUM , Narration FROM JOURNAL WHERE DRAC=@ID  AND (AUTHDATE BETWEEN @DT1 AND @DT2+1)
	UNION ALL
	SELECT AUTHDATE AS DT, "By " + DRNAME + " (" + MemoRef + ")" AS NAME, 0 AS DR, AMOUNT AS CR, 0-AMOUNT AS BALANCE, MEMOREF, '-' as SATISFACTION, 0 AS DRSUM, AMOUNT AS CRSUM, Narration  FROM JOURNAL WHERE CRAC=@ID  AND (AUTHDATE BETWEEN @DT1 AND @DT2+1)
	ORDER BY 1
ELSE
	SELECT @DT1 AS DATE, "By BALANCE " AS PARTICULARS  , 0 AS DR, -@BALANCE  AS CR, @BALANCE  AS BALANCE, 'X' AS MEMOREF, '-' as SATISFACTION, 0 AS DRSUM, -@BALANCE  AS CRSUM, 'OPENING BAL' as Narration
	UNION ALL
	SELECT AUTHDATE AS DT, "To " + CRNAME + " (" + MemoRef + ")" AS NAME, AMOUNT AS DR, 0 AS CR, AMOUNT-0 AS BALANCE, MEMOREF, '-' as SATISFACTION, AMOUNT AS DRSUM, 0 AS CRSUM, Narration  FROM JOURNAL WHERE DRAC=@ID  AND (AUTHDATE BETWEEN @DT1 AND @DT2+1)
	UNION ALL
	SELECT AUTHDATE AS DT, "By " + DRNAME + " (" + MemoRef + ")" AS NAME, 0 AS DR, AMOUNT AS CR, 0-AMOUNT AS BALANCE, MEMOREF, '-' as SATISFACTION, 0 AS DRSUM, AMOUNT AS CRSUM, Narration  FROM JOURNAL WHERE CRAC=@ID  AND (AUTHDATE BETWEEN @DT1 AND @DT2+1)
	ORDER BY 1
GO


CREATE PROC appproc_CREDITCHECK @ID VARCHAR(10), @AMT AS MONEY  AS 
IF (ISNULL((SELECT SUM(CL) FROM APPVIEW_ACHEADS WHERE ID=@ID AND ACGROUP="PERSONAL"),0)) > 0 
	SELECT ( 
		ISNULL((SELECT SUM(CreditLimit) - SUM(OB) FROM appview_AllAccounts WHERE ID=@ID),0) 
		-
		ISNULL((SELECT SUM(AMOUNT) FROM JOURNAL WHERE DRAC=@ID),0) 
		+ 
		ISNULL((SELECT SUM(AMOUNT) FROM JOURNAL WHERE CRAC=@ID),0)
		-
		@AMT
		) AS BAL

ELSE
		SELECT 0 AS BAL
GO

CREATE PROC appproc_CashReport @DT AS datetime AS
EXEC appproc_Ledger 'R0002', @DT, @DT
GO


CREATE PROC appproc_CashReportPeriod @DT1 AS datetime, @DT2 AS datetime AS
EXEC appproc_Ledger 'R0002', @DT1, @DT2
GO


CREATE PROC appproc_ChallanDeletedReport @SDate AS datetime, @EDate AS datetime AS
SELECT DBREF,  DBDATE, STATUS, ID, NAME, NETAMOUNT FROM Ctmain WHERE DBDate >= @SDate and DBDate < @EDate GROUP BY DBREF, DBDATE, STATUS, ID, NAME, NETAMOUNT ORDER BY DBREF Desc
GO


CREATE PROC appproc_CustItemDiscReturn @id varchar(10), @Itemid numeric, @DiscType varchar(50) AS
BEGIN
	SELECT ISNULL((SELECT TOP 1 Disc from  CUST_Item_Disc Where id=@ID AND ItemID=@ItemID AND DiscType=@DiscType), '0') AS Disc
END
GO

CREATE PROC appproc_CustItemDiscSave @id varchar(10), @Itemid numeric, @DiscType varchar(100), @Disc varchar(50) AS
BEGIN
	DELETE FROM CUST_Item_Disc WHERE id=@id AND Itemid=@Itemid AND DiscType=@DiscType
	IF @Disc <> '0'
		INSERT INTO CUST_Item_Disc (ID, ItemID, DiscType, Disc) VALUES (@ID, @ItemID, @DiscType, @Disc)
END
GO


CREATE PROC appproc_DuplicateItem @ItemID numeric, @MRP float = 0 AS 
BEGIN
	BEGIN
	INSERT INTO Items ( ItemCode, HSNCode, ItemName, MakerAuthor, ProducerID, ProducerName,  Version, Packing, Currency, CurrMRP, CurrSRP, PDisc, SDisc, GST, Cess ) 
	SELECT ItemCode, HSNCode, ItemName, MakerAuthor, ProducerID, ProducerName,  Version, Packing, Currency, @MRP as CurrMRP, @MRP as CurrSRP, PDisc, SDisc, GST, Cess FROM Items WHERE ItemID=@ItemID
	SELECT MAX(ItemID) as NewItemID FROM Items
	END
END
GO


CREATE PROC appproc_JournalDate @Dt datetime AS
SELECT SERIAL, AUTHDATE, 'By ' + DRNAME , 'To ' + CRNAME, AMOUNT, MemoRef FROM JOURNAL WHERE DATEDIFF(D, AUTHDATE, @Dt)=0 ORDER BY SERIAL
GO

CREATE PROC appproc_LEDGERBAL @id varchar(10), @dt datetime = NULL AS 
IF @DT IS NULL
BEGIN
	SET @DT = GETDATE()
END
SELECT (
	isnull((SELECT SUM(OB) FROM appview_AllAccounts WHERE ID=@id AND DATEDIFF(D, OBDATE, @dt)>=0),0) 
	+
	isnull((SELECT SUM(AMOUNT) FROM JOURNAL WHERE DRAC=@id  AND DATEDIFF(D, AUTHDATE, @dt)>=0),0  ) 
	- 
	isnull((SELECT SUM(AMOUNT) FROM JOURNAL WHERE CRAC=@id AND DATEDIFF(D, AUTHDATE, @dt)>=0),0)) 
as BAL
GO

CREATE PROC appproc_LedgerChart @ID varchar(10) AS
SELECT AUTHDATE AS DT, SUM(AMOUNT) AS DR, 0 AS CR FROM JOURNAL WHERE DRAC=@ID GROUP BY AUTHDATE
UNION ALL
SELECT AUTHDATE AS DT, 0 AS DR, SUM(AMOUNT) AS CR FROM JOURNAL WHERE CRAC=@ID GROUP BY AUTHDATE
ORDER BY 1
GO

CREATE PROC appproc_PRTrace @ItemID NUMERIC, @ID Varchar(10)  AS 
SELECT TOP 3 DBREF, INVREF, INVDATE, QTY FROM PUBIG WHERE ID=@ID AND ItemID=@ItemID ORDER BY DBREF DESC, DBDATE DESC
GO

CREATE PROC appproc_ProfitReport_ItemWise @DT1 as Datetime, @DT2 as DateTime, @PubName as Varchar(100) AS
DECLARE @DT DATETIME
IF @DT1 > @DT2
BEGIN
	SET @DT=@DT1
	SET @DT1=@DT2
	SET @DT2=@DT
END
SELECT B.ItemID, B.ItemName, B.SRP, SUM(QTY) As Qty, SUM(Amount) As Amount, AVG(A.PDISC) AS PDISC, AVG(A.SDISC) AS SDISC, SUM(A.PROFIT) As Profit
FROM APPVIEWProfit_SALE_PROFIT A, Items B
WHERE A.ItemID=B.ItemID AND DateDiff(D, A.DBDate,  @DT1)<=0 AND DateDiff(D, A.DBDate,  @DT2)>=0 AND B.ProducerName=@PubName
GROUP BY B.ItemID, B.ItemName, B.SRP
ORDER BY 1,2
GO

CREATE PROC appproc_ProfitReport_ProducerWise @DT1 as Datetime, @DT2 as DateTime AS
DECLARE @DT DATETIME
IF @DT1 > @DT2
BEGIN
	SET @DT=@DT1
	SET @DT1=@DT2
	SET @DT2=@DT
END
SELECT B.ProducerNAME, SUM(QTY) As Qty, SUM(Amount) As Amount, SUM(A.PROFIT) As Profit
FROM APPVIEWProfit_SALE_PROFIT A, Items B
WHERE A.ItemID=B.ItemID AND DateDiff(D, A.DBDate,  @DT1)<=0 AND DateDiff(D, A.DBDate,  @DT2)>=0
GROUP BY B.ProducerNAME
ORDER BY 1,2
GO

CREATE PROC appproc_PurchaseReportDate @DT AS datetime AS
EXEC appproc_Ledger 'N0001', @DT, @DT
GO

CREATE PROC appproc_PurchaseReportNetAmount @SDate AS datetime, @EDate AS datetime AS
SELECT DBDATE, DBREF, STATUS, NAME, CITY, NETAMOUNT FROM PUBIG WHERE DBDate >= @SDate and DBDate < @EDate GROUP BY DBREF, DBDATE, STATUS, NAME, CITY, NETAMOUNT ORDER BY DBREF Desc
GO

CREATE PROC appproc_PurchaseReturnReportDate @DT AS datetime AS
EXEC appproc_Ledger 'N0002', @DT, @DT
GO

CREATE PROC appproc_PurchaseReturnReportNetAmount @SDate AS datetime, @EDate AS datetime AS
SELECT DBDATE, DBREF, STATUS, NAME, CITY, NETAMOUNT FROM PRBIG WHERE DBDate >= @SDate and DBDate < @EDate GROUP BY DBREF, DBDATE, STATUS, NAME, CITY, NETAMOUNT ORDER BY DBREF Desc
GO

CREATE PROC appproc_ReportBillAmountPeriodCustwise @SDate AS datetime, @EDate AS datetime AS
SELECT DBRef as BillNo, Status, ID as CustID, Name, NetAmount as BillAmount from smain where DBDate >= @SDate and DBDate < @EDate and status !='order' and status!='cancelled' order by DBRef desc
GO

CREATE PROC appproc_ReportPeriod_ItemsIN @SDate datetime, @EDate datetime AS
SELECT ItemID, ItemName, ProducerName, MRP FROM TIBIG
WHERE DBDATE >= @SDate AND DBDATE < @EDate AND ItemID NOT IN 
(SELECT ItemID FROM SABIG WHERE DBDATE >= @SDate AND DBDATE < @EDate)
ORDER BY ProducerName
GO

CREATE PROC appproc_ReturnDisc @ID numeric, @ItemID numeric, @Opt varchar(10), @Type varchar(20)='SALE' AS
DECLARE @ProducerID varchar(10)
DECLARE @PDisc varchar(20)
DECLARE @SDisc varchar(20)
Set @ProducerID=(SELECT ProducerID FROM Items WHERE ItemID=@ItemID)
Set @PDisc=(SELECT PDisc FROM Items WHERE ItemID=@ItemID)
Set @SDisc=(SELECT SDisc FROM Items WHERE ItemID=@ItemID)

IF @Type='SALE' OR @Type='SALERETURN' 
    BEGIN
        if @Opt='SPL'
            BEGIN
                IF (SELECT ISNULL((SELECT TOP 1 X.Disc FROM XBD X WHERE X.ItemID=@ID AND X.ProducerID=@ProducerID AND X.PDisc=@PDisc AND X.SDisc=@SDisc), '0.00'))<>'0.00'
                    SELECT ISNULL((SELECT TOP 1 X.Disc FROM XBD X WHERE X.ItemID=@ID AND X.ProducerID=@ProducerID AND X.PDisc=@PDisc AND X.SDisc=@SDisc), '0.00') AS Disc
                Else
                    SELECT ISNULL((SELECT TOP 1 X.SDisc FROM Items X WHERE X.ProducerID=@ProducerID AND X.PDisc=@PDisc AND X.SDisc=@SDisc), '0.00') AS Disc
            End
        Else
                    SELECT ISNULL((SELECT TOP 1 X.SDisc FROM Items X WHERE X.ProducerID=@ProducerID AND X.PDisc=@PDisc AND X.SDisc=@SDisc), '0.00') AS Disc
    End
Else
    BEGIN
        if @Opt='SPL'
            BEGIN
                IF (SELECT ISNULL((SELECT TOP 1 X.Disc FROM XBD X WHERE X.ItemID=@ID AND X.ProducerID=@ProducerID AND X.PDisc=@PDisc AND X.SDisc=@SDisc), '0.00'))<>'0.00'
                    SELECT ISNULL((SELECT TOP 1 X.Disc FROM XBD X WHERE X.ItemID=@ID AND X.ProducerID=@ProducerID AND X.PDisc=@PDisc AND X.SDisc=@SDisc), '0.00') AS Disc
                Else
                    SELECT ISNULL((SELECT TOP 1 X.PDisc FROM Items X WHERE X.ProducerID=@ProducerID AND X.PDisc=@PDisc AND X.SDisc=@SDisc), '0.00') AS Disc
            End
        Else
                    SELECT ISNULL((SELECT TOP 1 X.PDisc FROM Items X WHERE X.ProducerID=@ProducerID AND X.PDisc=@PDisc AND X.SDisc=@SDisc), '0.00') AS Disc
    End
GO

CREATE PROC appproc_ReturnDiscNew @ID numeric, @ItemID numeric, @Opt varchar(10), @Type varchar(20)='SALE' AS
declare @DiscTplt nvarchar(10)
declare @ProducerID nvarchar(10)
declare @DiscGroup nvarchar(10)

SELECT @DiscTplt=DiscTplt  FROM PERSONAL WHERE ID=@ID
SELECT @ProducerID=ProducerID FROM Items WHERE ItemID=@ItemID
if @Type='SALE' OR @Type='SALERETURN' 
    BEGIN
        if @Opt='SPL'
            BEGIN
                IF (SELECT ISNULL((SELECT TOP 1 Disc FROM DiscTemplate X WHERE X.DiscTplt =@DiscTplt AND X.ProducerID=@ProducerID), '0'))<>'0'
                    SELECT ISNULL((SELECT TOP 1Disc FROM DiscTemplate X WHERE X.DiscTplt =@DiscTplt AND X.ProducerID=@ProducerID), '0') AS Disc
                Else
                    SELECT ISNULL((SELECT TOP 1 X.SDisc FROM Items X WHERE X.ItemID=@ItemID), '0') AS Disc
            End
        Else
                    SELECT ISNULL((SELECT TOP 1 X.SDisc FROM Items X WHERE X.ItemID=@ItemID), '0') AS Disc
    End
Else
    BEGIN
        if @Opt='SPL'
            BEGIN
                IF (SELECT ISNULL((SELECT TOP 1 Disc FROM DiscTemplate X WHERE X.DiscTplt =@DiscTplt AND X.ProducerID=@ProducerID), '0'))<>'0'
                    SELECT ISNULL((SELECT TOP 1Disc FROM DiscTemplate X WHERE X.DiscTplt =@DiscTplt AND X.ProducerID=@ProducerID), '0') AS Disc
                Else
                    SELECT ISNULL((SELECT TOP 1 X.PDisc FROM Items X WHERE X.ItemID=@ItemID), '0') AS Disc
            End
        Else
                    SELECT ISNULL((SELECT TOP 1 X.PDisc FROM Items X WHERE X.ItemID=@ItemID), '0') AS Disc
    End
GO

CREATE PROC appproc_ReturnStock @ItemID Numeric AS
SELECT ISNULL((SELECT AVLBL FROM appview_Stock WHERE ItemID=@ItemID),0)  as AVLBL
GO

CREATE PROC appproc_SaleReportDate @DT AS datetime AS
EXEC appproc_Ledger 'N0003', @DT, @DT
GO

CREATE PROC appproc_SaleReportNetAmount @SDate AS datetime, @EDate AS datetime AS
SELECT DBDATE, DBREF, STATUS, ID, NAME, CITY, NETAMOUNT FROM SABIG WHERE DBDate >= @SDate and DBDate <= @EDate GROUP BY DBREF, DBDATE, STATUS, ID, NAME, CITY, NETAMOUNT ORDER BY DBREF Desc
GO

CREATE PROC appproc_SaleReturnReportDate @DT AS datetime AS
EXEC appproc_Ledger 'N0004', @DT, @DT
GO

CREATE PROC appproc_SaleReturnReportNetAmount @SDate AS datetime, @EDate AS datetime AS
SELECT DBDATE, DBREF, STATUS, NAME, CITY, NETAMOUNT FROM SRBIG WHERE DBDate >= @SDate and DBDate < @EDate GROUP BY DBREF, DBDATE, STATUS, NAME, CITY, NETAMOUNT ORDER BY DBREF Desc
GO

CREATE PROC appproc_SaveDisc @id varchar(10), @Producerid varchar(10), @Producername varchar(100), @PDisc varchar(50), @SDisc varchar(50), @Disc varchar(50) AS
BEGIN
	DELETE FROM XBD WHERE ItemID=@id AND Producerid=@Producerid AND PDisc=@PDisc  AND SDisc=@SDisc
	IF @Disc <> '0'
		INSERT INTO XBD (ItemID, ProducerID, ProducerNAME, PDisc, SDisc, Disc) VALUES (@ID, @ProducerID, @ProducerNAME, @PDisc, @SDisc, @Disc)
END
GO

CREATE PROC appproc_SaveDiscNew @DiscTplt varchar(10), @Producerid varchar(10), @Producername varchar(100), @DiscGroup varchar(50), @Disc varchar(50) AS
BEGIN
	DELETE FROM DiscTemplate WHERE DiscTplt=@DiscTplt AND ProducerID=@ProducerID AND DiscGroup=@DiscGroup
	IF @Disc <> '0'
		INSERT INTO DiscTemplate (DiscTplt, ProducerID, ProducerNAME, DiscGroup, DISC) VALUES (@DiscTplt, @ProducerID, @ProducerNAME, @DiscGroup, @DISC)
END
GO

CREATE PROC appproc_SendItemsToBin AS
BEGIN
	INSERT INTO ItemsBIN 
	SELECT * FROM BASEItems WHERE ItemID IN
	(SELECT Min(ItemID) FROM Items GROUP BY ItemCode HAVING COUNT(ItemCode)>2)

	DELETE BASEItems WHERE ItemID IN
	(SELECT Min(ItemID) FROM Items GROUP BY ItemCode HAVING COUNT(ItemCode)>2)
END
GO

CREATE PROC appproc_SetInvNo @DBRef decimal, @Table varchar(20) AS 
IF @TABLE='SMAIN'
BEGIN
	IF (SELECT InvRef from SMAIN WHERE DBREF=@DBREF)='0'
		UPDATE SMAIN SET INVREF='NEW' WHERE DBREF=@DBREF
END
---		UPDATE SMAIN SET INVREF=(SELECT (CAST(MAX(INVREF) AS DECIMAL)+ 1) FROM SMAIN WHERE STATUS=(SELECT STATUS FROM SMAIN WHERE DBREF=@DBREF)) WHERE DBREF=@DBREF
GO

CREATE PROC appproc_StockINfromBranch @SDate AS datetime, @EDate AS datetime AS
SELECT DBREF, DBDATE, INVREF, NAME, BUNDLECOUNT AS BUNDLE, NETAMOUNT FROM TIBIG WHERE DBDate >= @SDate and DBDate < @EDate and ID LIKE 'W%' GROUP BY DBREF, DBDATE, INVREF, NAME, BUNDLECOUNT, NETAMOUNT ORDER BY DBREF
GO

CREATE PROC appproc_StockReportPeriod @DT1 AS datetime, @DT2 AS datetime AS
SELECT ItemID, ItemName, MRP, ProducerName, Sum(Qty) as Stock FROM Stock_Full_Datewise WHERE StockDate between @DT1 and @DT2 Group by ItemID, ItemName, MRP, ProducerName order by ItemId
GO

CREATE PROC appproc_StockTIBranch @SDate AS datetime, @EDate AS datetime AS
SELECT DBDATE, DBREF, INVREF, NAME, BUNDLECOUNT AS BUNDLE, STATUS, NETAMOUNT, ORDERREF AS TIN_TOUT FROM TIBIG WHERE DBDate >= @SDate and DBDate < @EDate and ID LIKE 'W%' GROUP BY DBDATE, DBREF, INVREF, NAME, BUNDLECOUNT, STATUS, NETAMOUNT, ORDERREF ORDER BY DBREF
GO

CREATE PROC appproc_StockTOBranch @SDate AS datetime, @EDate AS datetime AS
SELECT DBDATE, DBREF, INVREF, NAME, BUNDLECOUNT AS BUNDLE, STATUS, NETAMOUNT, ORDERREF AS TIN_TOUT FROM TOBIG WHERE DBDate >= @SDate and DBDate < @EDate and ID LIKE 'W%' GROUP BY DBDATE, DBREF, INVREF, NAME, BUNDLECOUNT, STATUS, NETAMOUNT, ORDERREF ORDER BY DBREF
GO

CREATE PROC appproc_UpdateItems_ItemName @ItemID numeric, @ItemName nvarchar(200) AS
BEGIN
	Update Items Set ItemName = @ItemName WHERE ItemID=@ItemID
END
GO

CREATE PROC appproc_UpdateItems_InitStock @ItemID numeric, @CurrentStock numeric AS
BEGIN
	DECLARE @AVLBL as Numeric
	SELECT @AVLBL=AVLBL FROM appview_Stock WHERE ItemID=@ItemID
	Update Items Set InitStock = InitStock - (@AVLBL - @CurrentStock) WHERE ItemID=@ItemID
END
GO

CREATE PROC appproc_UpdateItems_PDisc @ItemID numeric, @Disc nvarchar(50) AS
BEGIN
	Update Items Set PDisc = @Disc WHERE ItemID=@ItemID
END
GO

CREATE PROC appproc_UpdateItems_PubID @ItemID numeric, @PubID nvarchar(10) AS
BEGIN
	Update Items Set ProducerID = @PubID WHERE ItemID=@ItemID
END
GO

CREATE PROC appproc_UpdateItems_SDisc @ItemID numeric, @Disc nvarchar(50) AS
BEGIN
	Update Items Set SDisc = @Disc WHERE ItemID=@ItemID
END

GO

---###################################################################################
---###################################################################################
---###################################################################################



---###################################################################################
--- INSERTING SEED DATA                      #########################################
---###################################################################################
use [uits_gst_db]
GO

INSERT INTO dbkey (keyone,keytwo) values ('sbck1#','sbck1#')

INSERT INTO CMan (username,loginid,passwd,address,phones,rights) values ('UITS','scott','©žœš§','pat','99xxx99xxx','0')

INSERT INTO Currency (Serial, Currency, CurrPrice, Codes) values ('1','INR','1','INR')
INSERT INTO Currency (Serial, Currency, CurrPrice, Codes) values ('1','Rs.','1','Rs.')
INSERT INTO Currency (Serial, Currency, CurrPrice, Codes) values ('2','$','65','$')
INSERT INTO Currency (Serial, Currency, CurrPrice, Codes) values ('3','£','94','£')
INSERT INTO Currency (Serial, Currency, CurrPrice, Codes) values ('4','€','82','€')
INSERT INTO Currency (Serial, Currency, CurrPrice, Codes) values ('4','¥','10','¥')
GO

INSERT INTO GRP (grp,description,initial) values ('REAL','Group for Real Accounts','R')
INSERT INTO GRP (grp,description,initial) values ('NOMINAL','Group for Nominal Accounts','N')
INSERT INTO GRP (grp,description,initial) values ('CUSTOMER','Group for Customers','C')
INSERT INTO GRP (grp,description,initial) values ('PRODUCER','Group for Producers','P')
INSERT INTO GRP (grp,description,initial) values ('DISTRIBUTOR','Group for Distributors','D')
INSERT INTO GRP (grp,description,initial) values ('SUBDIST','Group for Subdistributors','U')
INSERT INTO GRP (grp,description,initial) values ('TRANSPORTER','Group for Transporters','T')
INSERT INTO GRP (grp,description,initial) values ('KARTER','Group for Karters','K')
INSERT INTO GRP (grp,description,initial) values ('GENERAL','Group for General','G')
INSERT INTO GRP (grp,description,initial) values ('POSTAGE','Group for Postage','O')
INSERT INTO GRP (grp,description,initial) values ('WAREHOUSES','Group for Warehouses','W')
INSERT INTO GRP (grp,description,initial) values ('DAMAGE','Group for Damage Accounts','M')
INSERT INTO GRP (grp,description,initial) values ('TAXMAN','Group for Tax Authorities','X')
GO

INSERT INTO PERSONAL (ID,NAME,GRP) VALUES ('N0001','CGST','NOMINAL')
INSERT INTO PERSONAL (ID,NAME,GRP) VALUES ('N0002','SGST','NOMINAL')
INSERT INTO PERSONAL (ID,NAME,GRP) VALUES ('N0003','IGST','NOMINAL')
INSERT INTO PERSONAL (ID,NAME,GRP) VALUES ('N0004','Cess','NOMINAL')

INSERT INTO PERSONAL (ID,NAME,GRP) VALUES ('N0011','PURCHASES','NOMINAL')
INSERT INTO PERSONAL (ID,NAME,GRP) VALUES ('N0012','PURCHASE RETURNS','NOMINAL')
INSERT INTO PERSONAL (ID,NAME,GRP) VALUES ('N0013','SALES','NOMINAL')
INSERT INTO PERSONAL (ID,NAME,GRP) VALUES ('N0014','SALE RETURNS','NOMINAL')
INSERT INTO PERSONAL (ID,NAME,GRP) VALUES ('N0015','TRANSFER-OUT','NOMINAL')
INSERT INTO PERSONAL (ID,NAME,GRP) VALUES ('N0016','TRANSFER-IN','NOMINAL')
INSERT INTO PERSONAL (ID,NAME,GRP) VALUES ('N0017','TRANSPORTATION','NOMINAL')
INSERT INTO PERSONAL (ID,NAME,GRP) VALUES ('N0018','KARTAGE','NOMINAL')
INSERT INTO PERSONAL (ID,NAME,GRP) VALUES ('N0019','POSTAGE','NOMINAL')
INSERT INTO PERSONAL (ID,NAME,GRP) VALUES ('N0020','ASSORTED','NOMINAL')

INSERT INTO PERSONAL (ID,NAME,GRP) VALUES ('R0001','BANK','REAL')
INSERT INTO PERSONAL (ID,NAME,GRP) VALUES ('R0002','CASH','REAL')

INSERT INTO PERSONAL (ID,NAME,GRP) VALUES ('X0001','CGST COLLECTOR','TAXMAN')
INSERT INTO PERSONAL (ID,NAME,GRP) VALUES ('X0002','SGST COLLECTOR','TAXMAN')
INSERT INTO PERSONAL (ID,NAME,GRP) VALUES ('X0003','IGST COLLECTOR','TAXMAN')
INSERT INTO PERSONAL (ID,NAME,GRP) VALUES ('X0004','CESS COLLECTOR','TAXMAN')

INSERT INTO PERSONAL (ID,NAME,GRP) VALUES ('C0001','Sample Buyer','CUSTOMER')
INSERT INTO PERSONAL (ID,NAME,GRP) VALUES ('C0002','Cash Buyer','CUSTOMER')
INSERT INTO PERSONAL (ID,NAME,GRP) VALUES ('P0001','Sample Producer','PRODUCER')
INSERT INTO PERSONAL (ID,NAME,GRP) VALUES ('D0001','Sample Distributor','DISTRIBUTOR')
INSERT INTO PERSONAL (ID,NAME,GRP) VALUES ('T0001','Sample Transporter','TRANSPORTER')
INSERT INTO PERSONAL (ID,NAME,GRP) VALUES ('K0001','Sample Karter','KARTER')
INSERT INTO PERSONAL (ID,NAME,GRP) VALUES ('O0001','Sample Postman or Courier','KARTER')
GO

INSERT INTO Items (ItemName,ProducerId, ProducerName, Currency, CurrMRP, CurrSRP, PDisc, SDisc, ItemId1) VALUES ('Sample Item', 'P001', 'Sample Producer', 'Rs.', '100', '100', '40', '20','1')
GO

