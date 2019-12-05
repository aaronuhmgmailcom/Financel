USE [RSDataV2]
GO
/****** Object:  Table [dbo].[rsGroups]    Script Date: 01/22/2016 17:59:47 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[rsGroups](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[GroupName] [nvarchar](150) NULL,
	[GroupDisable] [bit] NULL,
	[Remark] [nvarchar](4000) NULL,
	[AddInTabName] [nvarchar](max) NULL,
 CONSTRAINT [PK_rsGroups] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[rsSystemProcesses]    Script Date: 01/22/2016 17:59:47 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[rsSystemProcesses](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[ProcessName] [nvarchar](100) NULL,
 CONSTRAINT [PK_rsSystemProcesses] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[rsTemplateActions]    Script Date: 01/22/2016 17:59:47 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[rsTemplateActions](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[ButtonName] [nvarchar](50) NULL,
	[ButtonText] [nvarchar](50) NULL,
	[ButtonIcon] [nvarchar](150) NULL,
	[ButtonGroup] [nvarchar](50) NULL,
	[ButtonSize] [nvarchar](50) NULL,
	[ButtonOrder] [int] NULL,
	[GroupOrder] [int] NULL,
	[StopOnError] [bit] NULL,
	[TemplateID] [nvarchar](max) NULL,
	[ProcessID] [nvarchar](max) NULL,
	[MacroName] [nvarchar](50) NULL,
	[ProcessMacroOrder] [int] NULL,
	[Type] [int] NULL,
 CONSTRAINT [PK_FinTools_Buttons] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[rsTemplateAllocationMakerUpdate]    Script Date: 01/22/2016 17:59:47 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[rsTemplateAllocationMakerUpdate](
	[GUID] [int] IDENTITY(1,1) NOT NULL,
	[JournalNumber] [nvarchar](50) NULL,
	[Ledger] [nvarchar](50) NULL,
	[ft_Account] [nvarchar](50) NULL,
	[Period] [nvarchar](50) NULL,
	[TransactionDate] [nvarchar](50) NULL,
	[JrnlType] [nvarchar](50) NULL,
	[TransRef] [nvarchar](50) NULL,
	[AlloctnMarker] [nvarchar](50) NULL,
	[LA1] [nvarchar](50) NULL,
	[LA2] [nvarchar](50) NULL,
	[LA3] [nvarchar](50) NULL,
	[LA4] [nvarchar](50) NULL,
	[LA5] [nvarchar](50) NULL,
	[LA6] [nvarchar](50) NULL,
	[LA7] [nvarchar](50) NULL,
	[LA8] [nvarchar](50) NULL,
	[LA9] [nvarchar](50) NULL,
	[LA10] [nvarchar](50) NULL,
	[TemplateID] [nvarchar](max) NOT NULL,
	[LineIndicator] [nvarchar](50) NOT NULL,
	[StartinginCell] [nvarchar](50) NOT NULL,
	[inputFields] [nvarchar](max) NOT NULL,
	[updateFields] [nvarchar](max) NOT NULL,
	[JournalLineNumber] [nvarchar](50) NULL,
 CONSTRAINT [PK_FinTools_Settings_AllocationMarkerUpdate] PRIMARY KEY CLUSTERED 
(
	[GUID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[rsTemplateConsolidation]    Script Date: 01/22/2016 17:59:47 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[rsTemplateConsolidation](
	[ft_id] [int] IDENTITY(1,1) NOT NULL,
	[Ledger] [nvarchar](50) NULL,
	[ft_Account] [nvarchar](50) NULL,
	[Period] [nvarchar](50) NULL,
	[TransDate] [nvarchar](50) NULL,
	[DueDate] [nvarchar](50) NULL,
	[JrnlType] [nvarchar](50) NULL,
	[JrnlSource] [nvarchar](50) NULL,
	[TransRef] [nvarchar](50) NULL,
	[Description] [nvarchar](50) NULL,
	[AlloctnMarker] [nvarchar](50) NULL,
	[LA1] [nvarchar](50) NULL,
	[LA2] [nvarchar](50) NULL,
	[LA3] [nvarchar](50) NULL,
	[LA4] [nvarchar](50) NULL,
	[LA5] [nvarchar](50) NULL,
	[LA6] [nvarchar](50) NULL,
	[LA7] [nvarchar](50) NULL,
	[LA8] [nvarchar](50) NULL,
	[LA9] [nvarchar](50) NULL,
	[LA10] [nvarchar](50) NULL,
	[GenDesc1] [nvarchar](50) NULL,
	[GenDesc2] [nvarchar](50) NULL,
	[GenDesc3] [nvarchar](50) NULL,
	[GenDesc4] [nvarchar](50) NULL,
	[GenDesc5] [nvarchar](50) NULL,
	[GenDesc6] [nvarchar](50) NULL,
	[GenDesc7] [nvarchar](50) NULL,
	[GenDesc8] [nvarchar](50) NULL,
	[GenDesc9] [nvarchar](50) NULL,
	[GenDesc10] [nvarchar](50) NULL,
	[GenDesc11] [nvarchar](50) NULL,
	[GenDesc12] [nvarchar](50) NULL,
	[GenDesc13] [nvarchar](50) NULL,
	[GenDesc14] [nvarchar](50) NULL,
	[GenDesc15] [nvarchar](50) NULL,
	[GenDesc16] [nvarchar](50) NULL,
	[GenDesc17] [nvarchar](50) NULL,
	[GenDesc18] [nvarchar](50) NULL,
	[GenDesc19] [nvarchar](50) NULL,
	[GenDesc20] [nvarchar](50) NULL,
	[GenDesc21] [nvarchar](50) NULL,
	[GenDesc22] [nvarchar](50) NULL,
	[GenDesc23] [nvarchar](50) NULL,
	[GenDesc24] [nvarchar](50) NULL,
	[GenDesc25] [nvarchar](50) NULL,
	[TransAmount] [nvarchar](50) NULL,
	[Currency] [nvarchar](50) NULL,
	[BaseAmount] [nvarchar](50) NULL,
	[2ndBase] [nvarchar](50) NULL,
	[4thAmount] [nvarchar](50) NULL,
	[TemplateID] [nvarchar](max) NOT NULL,
	[LineIndicator] [nvarchar](50) NOT NULL,
	[StartinginCell] [nvarchar](50) NOT NULL,
	[ConsolidateBy1] [nvarchar](50) NOT NULL,
	[ConsolidateBy2] [nvarchar](50) NOT NULL,
	[ConsolidateBy3] [nvarchar](50) NOT NULL,
	[ConsolidateBy4] [nvarchar](50) NOT NULL,
	[BalanceBy] [nvarchar](50) NULL,
	[PopWithJNNumber] [nvarchar](50) NULL,
 CONSTRAINT [PK_FinTools_OutPutConsolidation] PRIMARY KEY CLUSTERED 
(
	[ft_id] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[rsTemplateContainer]    Script Date: 01/22/2016 17:59:47 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[rsTemplateContainer](
	[ft_id] [uniqueidentifier] ROWGUIDCOL  NOT NULL,
	[TemplateID] [nvarchar](max) NOT NULL,
	[ft_relatefilepath] [nvarchar](max) NOT NULL,
	[column] [nvarchar](5) NULL,
	[FromDB] [bit] NULL,
 CONSTRAINT [PK_FinTools_Settings_ContainerSetting] PRIMARY KEY CLUSTERED 
(
	[ft_id] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[rsTemplateDrillDown]    Script Date: 01/22/2016 17:59:47 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[rsTemplateDrillDown](
	[GUID] [uniqueidentifier] ROWGUIDCOL  NOT NULL,
	[SunField] [nvarchar](50) NULL,
	[InputStatus] [nvarchar](50) NULL,
	[CellName] [nvarchar](50) NULL,
	[OutputStatus] [nvarchar](50) NULL,
	[TemplateID] [nvarchar](500) NULL,
	[order] [int] NULL,
 CONSTRAINT [PK_FinTools_Settings_DrillDown] PRIMARY KEY CLUSTERED 
(
	[GUID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[rsTemplateCreateXMLTextFile]    Script Date: 01/22/2016 17:59:47 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[rsTemplateCreateXMLTextFile](
	[ft_id] [int] IDENTITY(1,1) NOT NULL,
	[Column1] [nvarchar](150) NULL,
	[Column2] [nvarchar](150) NULL,
	[Column3] [nvarchar](150) NULL,
	[Column4] [nvarchar](150) NULL,
	[Column5] [nvarchar](150) NULL,
	[Column6] [nvarchar](150) NULL,
	[Column7] [nvarchar](150) NULL,
	[Column8] [nvarchar](150) NULL,
	[Column9] [nvarchar](150) NULL,
	[Column10] [nvarchar](150) NULL,
	[Column11] [nvarchar](150) NULL,
	[Column12] [nvarchar](150) NULL,
	[Column13] [nvarchar](150) NULL,
	[Column14] [nvarchar](150) NULL,
	[Column15] [nvarchar](150) NULL,
	[Column16] [nvarchar](150) NULL,
	[Column17] [nvarchar](150) NULL,
	[Column18] [nvarchar](150) NULL,
	[Column19] [nvarchar](150) NULL,
	[Column20] [nvarchar](150) NULL,
	[Column21] [nvarchar](150) NULL,
	[Column22] [nvarchar](150) NULL,
	[Column23] [nvarchar](150) NULL,
	[Column24] [nvarchar](150) NULL,
	[Column25] [nvarchar](150) NULL,
	[Column26] [nvarchar](150) NULL,
	[Column27] [nvarchar](150) NULL,
	[Column28] [nvarchar](150) NULL,
	[Column29] [nvarchar](150) NULL,
	[Column30] [nvarchar](150) NULL,
	[Column31] [nvarchar](150) NULL,
	[Column32] [nvarchar](150) NULL,
	[Column33] [nvarchar](150) NULL,
	[Column34] [nvarchar](150) NULL,
	[Column35] [nvarchar](150) NULL,
	[Column36] [nvarchar](150) NULL,
	[Column37] [nvarchar](150) NULL,
	[Column38] [nvarchar](150) NULL,
	[Column39] [nvarchar](150) NULL,
	[Column40] [nvarchar](150) NULL,
	[Column41] [nvarchar](150) NULL,
	[Column42] [nvarchar](150) NULL,
	[Column43] [nvarchar](150) NULL,
	[Column44] [nvarchar](150) NULL,
	[Column45] [nvarchar](150) NULL,
	[Column46] [nvarchar](150) NULL,
	[Column47] [nvarchar](150) NULL,
	[Column48] [nvarchar](150) NULL,
	[Column49] [nvarchar](150) NULL,
	[Column50] [nvarchar](150) NULL,
	[HeaderTextes] [nvarchar](max) NULL,
	[StartinginCell] [nvarchar](50) NULL,
	[LineIndicator] [nvarchar](50) NULL,
	[TemplateID] [nvarchar](max) NOT NULL,
	[IncludeHeaderRow] [bit] NULL,
	[SavePath] [nvarchar](max) NULL,
	[SaveName] [nvarchar](150) NULL,
	[ReferenceNumber] [nchar](10) NULL,
	[SunComponent] [nvarchar](max) NULL,
	[SunMethod] [nvarchar](max) NULL,
 CONSTRAINT [PK_FinTools_OutPutCreateTextFile] PRIMARY KEY CLUSTERED 
(
	[ft_id] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  StoredProcedure [dbo].[rsTemplateCreateTextFile_Ins]    Script Date: 01/22/2016 17:59:47 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[rsTemplateCreateTextFile_Ins]
@Column1 nvarchar(150) ,
@Column2 nvarchar(150) ,
@Column3 nvarchar(150) ,
@Column4 nvarchar(150) ,
@Column5 nvarchar(150) ,
@Column6 nvarchar(150) ,
@Column7 nvarchar(150) ,
@Column8 nvarchar(150) ,
@Column9 nvarchar(150) ,
@Column10 nvarchar(150) ,
@Column11 nvarchar(150) ,
@Column12 nvarchar(150) ,
@Column13 nvarchar(150) ,
@Column14 nvarchar(150) ,
@Column15 nvarchar(150) ,
@Column16 nvarchar(150) ,
@Column17 nvarchar(150) ,
@Column18 nvarchar(150) ,
@Column19 nvarchar(150) ,
@Column20 nvarchar(150) ,
@Column21 nvarchar(150) ,
@Column22 nvarchar(150) ,
@Column23 nvarchar(150) ,
@Column24 nvarchar(150) ,
@Column25 nvarchar(150) ,
@Column26 nvarchar(150) ,
@Column27 nvarchar(150) ,
@Column28 nvarchar(150) ,
@Column29 nvarchar(150) ,
@Column30 nvarchar(150) ,
@Column31 nvarchar(150) ,
@Column32 nvarchar(150) ,
@Column33 nvarchar(150) ,
@Column34 nvarchar(150) ,
@Column35 nvarchar(150) ,
@Column36 nvarchar(150) ,
@Column37 nvarchar(150) ,
@Column38 nvarchar(150) ,
@Column39 nvarchar(150) ,
@Column40 nvarchar(150) ,
@Column41 nvarchar(150) ,
@Column42 nvarchar(150) ,
@Column43 nvarchar(150) ,
@Column44 nvarchar(150) ,
@Column45 nvarchar(150) ,
@Column46 nvarchar(150) ,
@Column47 nvarchar(150) ,
@Column48 nvarchar(150) ,
@Column49 nvarchar(150) ,
@Column50 nvarchar(150) ,
@HeaderTextes nvarchar(max),
@StartinginCell nvarchar(50),
@LineIndicator nvarchar(50) ,
@TemplateID nvarchar(max),
@IncludeHeaderRow bit,
@SavePath nvarchar(max),
@SaveName nvarchar(150)

as
begin
        insert into dbo.rsTemplateCreateTextFile
        (Column1,Column2,Column3,Column4,Column5    ,
Column6    ,Column7    ,Column8  ,Column9  ,Column10  ,
Column11,Column12,Column13,Column14,Column15    ,
Column16    ,Column17    ,Column18  ,Column19  ,Column20  ,
Column21,Column22,Column23,Column24,Column25    ,
Column26    ,Column27    ,Column28  ,Column29  ,Column30  ,
Column31,Column32,Column33,Column34,Column35    ,
Column36    ,Column37    ,Column38  ,Column39  ,Column40  ,
Column41,Column42,Column43,Column44,Column45    ,
Column46    ,Column47    ,Column48  ,Column49  ,Column50  ,
HeaderTextes ,StartinginCell ,LineIndicator ,TemplateID ,IncludeHeaderRow ,SavePath,SaveName) 
        values(@Column1,@Column2,@Column3,@Column4,@Column5    ,
@Column6    ,@Column7    ,@Column8  ,@Column9  ,@Column10  ,
@Column11,@Column12,@Column13,@Column14,@Column15    ,
@Column16    ,@Column17    ,@Column18  ,@Column19  ,@Column20  ,
@Column21,@Column22,@Column23,@Column24,@Column25    ,
@Column26    ,@Column27    ,@Column28  ,@Column29  ,@Column30  ,
@Column31,@Column32,@Column33,@Column34,@Column35    ,
@Column36    ,@Column37    ,@Column38  ,@Column39  ,@Column40  ,
@Column41,@Column42,@Column43,@Column44,@Column45    ,
@Column46    ,@Column47    ,@Column48  ,@Column49  ,@Column50  ,
@HeaderTextes ,@StartinginCell ,@LineIndicator ,@TemplateID ,@IncludeHeaderRow,@SavePath,@SaveName ) 
end
GO
/****** Object:  StoredProcedure [dbo].[rsTemplateCreateTextFile_Del]    Script Date: 01/22/2016 17:59:47 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[rsTemplateCreateTextFile_Del]
@TemplateID nvarchar(max)

as
begin
        delete dbo.rsTemplateCreateTextFile where TemplateID=@TemplateID 
end
GO
/****** Object:  Table [dbo].[rsTemplateGenDescFields]    Script Date: 01/22/2016 17:59:47 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[rsTemplateGenDescFields](
	[GUID] [uniqueidentifier] ROWGUIDCOL  NOT NULL,
	[FieldGroup] [nvarchar](50) NULL,
	[SunField] [nvarchar](50) NULL,
	[UserFriendlyName] [nvarchar](50) NULL,
	[Output] [nvarchar](50) NULL,
	[Input] [nvarchar](50) NULL,
	[XML_Query] [nvarchar](max) NULL,
	[TemplateID] [nvarchar](max) NULL,
	[version] [int] NULL,
 CONSTRAINT [PK_FinTools_Settings_GenDescFields] PRIMARY KEY CLUSTERED 
(
	[GUID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[rsTemplateHelp]    Script Date: 01/22/2016 17:59:47 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[rsTemplateHelp](
	[ft_id] [int] IDENTITY(1,1) NOT NULL,
	[TemplateID] [nvarchar](max) NULL,
	[helpFileData] [image] NULL,
	[helpFileType] [nvarchar](50) NULL,
 CONSTRAINT [PK_FinTools_Settings_Help] PRIMARY KEY CLUSTERED 
(
	[ft_id] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[rsTemplateJournal]    Script Date: 01/22/2016 17:59:47 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[rsTemplateJournal](
	[ft_id] [int] IDENTITY(1,1) NOT NULL,
	[Ledger] [nvarchar](50) NULL,
	[ft_Account] [nvarchar](50) NULL,
	[Period] [nvarchar](50) NULL,
	[TransDate] [nvarchar](50) NULL,
	[DueDate] [nvarchar](50) NULL,
	[JrnlType] [nvarchar](50) NULL,
	[JrnlSource] [nvarchar](50) NULL,
	[TransRef] [nvarchar](50) NULL,
	[Description] [nvarchar](50) NULL,
	[AlloctnMarker] [nvarchar](50) NULL,
	[LA1] [nvarchar](50) NULL,
	[LA2] [nvarchar](50) NULL,
	[LA3] [nvarchar](50) NULL,
	[LA4] [nvarchar](50) NULL,
	[LA5] [nvarchar](50) NULL,
	[LA6] [nvarchar](50) NULL,
	[LA7] [nvarchar](50) NULL,
	[LA8] [nvarchar](50) NULL,
	[LA9] [nvarchar](50) NULL,
	[LA10] [nvarchar](50) NULL,
	[GenDesc1] [nvarchar](50) NULL,
	[GenDesc2] [nvarchar](50) NULL,
	[GenDesc3] [nvarchar](50) NULL,
	[GenDesc4] [nvarchar](50) NULL,
	[GenDesc5] [nvarchar](50) NULL,
	[GenDesc6] [nvarchar](50) NULL,
	[GenDesc7] [nvarchar](50) NULL,
	[GenDesc8] [nvarchar](50) NULL,
	[GenDesc9] [nvarchar](50) NULL,
	[GenDesc10] [nvarchar](50) NULL,
	[GenDesc11] [nvarchar](50) NULL,
	[GenDesc12] [nvarchar](50) NULL,
	[GenDesc13] [nvarchar](50) NULL,
	[GenDesc14] [nvarchar](50) NULL,
	[GenDesc15] [nvarchar](50) NULL,
	[GenDesc16] [nvarchar](50) NULL,
	[GenDesc17] [nvarchar](50) NULL,
	[GenDesc18] [nvarchar](50) NULL,
	[GenDesc19] [nvarchar](50) NULL,
	[GenDesc20] [nvarchar](50) NULL,
	[GenDesc21] [nvarchar](50) NULL,
	[GenDesc22] [nvarchar](50) NULL,
	[GenDesc23] [nvarchar](50) NULL,
	[GenDesc24] [nvarchar](50) NULL,
	[GenDesc25] [nvarchar](50) NULL,
	[TransAmount] [nvarchar](50) NULL,
	[Currency] [nvarchar](50) NULL,
	[BaseAmount] [nvarchar](50) NULL,
	[2ndBase] [nvarchar](50) NULL,
	[4thAmount] [nvarchar](50) NULL,
	[TemplateID] [nvarchar](max) NOT NULL,
	[LineIndicator] [nvarchar](50) NOT NULL,
	[StartinginCell] [nvarchar](50) NOT NULL,
	[BalanceBy] [nvarchar](50) NULL,
	[PopWithJNNumber] [nvarchar](50) NULL,
 CONSTRAINT [PK_FinTools_OutPutProfile] PRIMARY KEY CLUSTERED 
(
	[ft_id] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[rsTemplates]    Script Date: 01/22/2016 17:59:47 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[rsTemplates](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[TemplateData] [image] NULL,
	[TemplateName] [nvarchar](max) NULL,
	[OriginTemplatePath] [nvarchar](max) NULL,
	[FileType] [nvarchar](50) NULL,
	[Description] [nvarchar](max) NULL,
 CONSTRAINT [PK_rsTemplates] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[rsTemplateSequenceNumbering]    Script Date: 01/22/2016 17:59:47 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[rsTemplateSequenceNumbering](
	[ft_id] [uniqueidentifier] ROWGUIDCOL  NOT NULL,
	[UseSequenceNum] [int] NULL,
	[SequencePrefix] [nvarchar](50) NULL,
	[PostToField] [nvarchar](50) NULL,
	[PopulateCell] [nvarchar](50) NULL,
	[TemplateID] [nvarchar](500) NULL,
 CONSTRAINT [PK_FinTools_Settings_SequenceNumbering] PRIMARY KEY CLUSTERED 
(
	[ft_id] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[rsTemplateSetting]    Script Date: 01/22/2016 17:59:47 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[rsTemplateSetting](
	[ft_id] [uniqueidentifier] ROWGUIDCOL  NOT NULL,
	[CriteriaName] [nvarchar](50) NULL,
	[CellReference] [nvarchar](50) NULL,
	[TemplateID] [nvarchar](500) NULL,
	[orderNum] [int] NULL,
	[OpenTransUponSave] [bit] NULL,
 CONSTRAINT [PK_FinTools_Settings_ReportSetting] PRIMARY KEY CLUSTERED 
(
	[ft_id] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[rsTemplateTransactions]    Script Date: 01/22/2016 17:59:47 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[rsTemplateTransactions](
	[GUID] [uniqueidentifier] ROWGUIDCOL  NOT NULL,
	[TemplateName] [nvarchar](50) NULL,
	[Criteria1] [nvarchar](50) NULL,
	[Criteria2] [nvarchar](50) NULL,
	[Criteria3] [nvarchar](50) NULL,
	[Criteria4] [nvarchar](50) NULL,
	[Criteria5] [nvarchar](50) NULL,
	[Value1] [nvarchar](50) NULL,
	[Value2] [nvarchar](50) NULL,
	[Value3] [nvarchar](50) NULL,
	[Value4] [nvarchar](50) NULL,
	[Value5] [nvarchar](50) NULL,
	[Data] [image] NULL,
	[DataType] [nvarchar](50) NULL,
	[PDFData] [image] NULL,
	[XMLData] [varchar](8000) NULL,
	[TemplateID] [nvarchar](150) NULL,
	[maxNum] [int] NULL,
	[TransactionName] [nvarchar](50) NULL,
	[Prefix] [nvarchar](50) NULL,
	[SunJournalNumber] [nvarchar](100) NULL,
 CONSTRAINT [PK_FinTools_Invoicing] PRIMARY KEY CLUSTERED 
(
	[GUID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[rsTemplateTransactionUpdate]    Script Date: 01/22/2016 17:59:47 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[rsTemplateTransactionUpdate](
	[GUID] [int] IDENTITY(1,1) NOT NULL,
	[JournalNumber] [nvarchar](50) NULL,
	[JournalLineNumber] [nvarchar](50) NULL,
	[Ledger] [nvarchar](50) NULL,
	[ft_Account] [nvarchar](50) NULL,
	[Period] [nvarchar](50) NULL,
	[TransDate] [nvarchar](50) NULL,
	[DueDate] [nvarchar](50) NULL,
	[JrnlType] [nvarchar](50) NULL,
	[JrnlSource] [nvarchar](50) NULL,
	[TransRef] [nvarchar](50) NULL,
	[Description] [nvarchar](50) NULL,
	[AlloctnMarker] [nvarchar](50) NULL,
	[LA1] [nvarchar](50) NULL,
	[LA2] [nvarchar](50) NULL,
	[LA3] [nvarchar](50) NULL,
	[LA4] [nvarchar](50) NULL,
	[LA5] [nvarchar](50) NULL,
	[LA6] [nvarchar](50) NULL,
	[LA7] [nvarchar](50) NULL,
	[LA8] [nvarchar](50) NULL,
	[LA9] [nvarchar](50) NULL,
	[LA10] [nvarchar](50) NULL,
	[GenDesc1] [nvarchar](50) NULL,
	[GenDesc2] [nvarchar](50) NULL,
	[GenDesc3] [nvarchar](50) NULL,
	[GenDesc4] [nvarchar](50) NULL,
	[GenDesc5] [nvarchar](50) NULL,
	[GenDesc6] [nvarchar](50) NULL,
	[GenDesc7] [nvarchar](50) NULL,
	[GenDesc8] [nvarchar](50) NULL,
	[GenDesc9] [nvarchar](50) NULL,
	[GenDesc10] [nvarchar](50) NULL,
	[GenDesc11] [nvarchar](50) NULL,
	[GenDesc12] [nvarchar](50) NULL,
	[GenDesc13] [nvarchar](50) NULL,
	[GenDesc14] [nvarchar](50) NULL,
	[GenDesc15] [nvarchar](50) NULL,
	[GenDesc16] [nvarchar](50) NULL,
	[GenDesc17] [nvarchar](50) NULL,
	[GenDesc18] [nvarchar](50) NULL,
	[GenDesc19] [nvarchar](50) NULL,
	[GenDesc20] [nvarchar](50) NULL,
	[GenDesc21] [nvarchar](50) NULL,
	[GenDesc22] [nvarchar](50) NULL,
	[GenDesc23] [nvarchar](50) NULL,
	[GenDesc24] [nvarchar](50) NULL,
	[GenDesc25] [nvarchar](50) NULL,
	[TransAmount] [nvarchar](50) NULL,
	[Currency] [nvarchar](50) NULL,
	[BaseAmount] [nvarchar](50) NULL,
	[2ndBase] [nvarchar](50) NULL,
	[4thAmount] [nvarchar](50) NULL,
	[TemplateID] [nvarchar](max) NOT NULL,
	[LineIndicator] [nvarchar](50) NOT NULL,
	[StartinginCell] [nvarchar](50) NOT NULL,
	[inputFields] [nvarchar](max) NOT NULL,
	[updateFields] [nvarchar](max) NOT NULL,
 CONSTRAINT [PK_FinTools_Settings_TransactionUpdate] PRIMARY KEY CLUSTERED 
(
	[GUID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[rsTemplateVisible]    Script Date: 01/22/2016 17:59:47 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[rsTemplateVisible](
	[ft_id] [uniqueidentifier] ROWGUIDCOL  NOT NULL,
	[TemplateID] [nvarchar](500) NULL,
	[OutputPaneVisiable] [nvarchar](5) NULL,
	[UserID] [nvarchar](50) NULL,
 CONSTRAINT [PK_FinTools_Settings_VisiableRemembers] PRIMARY KEY CLUSTERED 
(
	[ft_id] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[rsUpgrade]    Script Date: 01/22/2016 17:59:47 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[rsUpgrade](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[RootPath] [nvarchar](max) NULL,
	[V1IP] [nvarchar](max) NULL,
	[V1UserID] [nvarchar](max) NULL,
	[V1Password] [nvarchar](max) NULL,
	[V2IP] [nvarchar](max) NULL,
	[V2UserID] [nvarchar](max) NULL,
	[V2Password] [nvarchar](max) NULL,
 CONSTRAINT [PK_rsUpgrade] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[rsUsers]    Script Date: 01/22/2016 17:59:47 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[rsUsers](
	[ft_id] [uniqueidentifier] ROWGUIDCOL  NOT NULL,
	[FormUserID] [nvarchar](50) NULL,
	[FormUserPassword] [nvarchar](50) NULL,
	[WindowsUserID] [nvarchar](50) NULL,
	[MachineName] [nvarchar](50) NULL,
	[LoginType] [int] NULL,
	[SUNUserIP] [nvarchar](50) NULL,
	[SUNUserID] [nvarchar](50) NULL,
	[SUNUserPass] [nvarchar](50) NULL,
	[AddInTabName] [nvarchar](max) NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[rsGlobalDocumentViews]    Script Date: 01/22/2016 17:59:47 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[rsGlobalDocumentViews](
	[vd_id] [uniqueidentifier] NOT NULL,
	[vd_type] [nvarchar](50) NOT NULL,
	[vd_prefix] [nchar](10) NOT NULL,
	[vd_folder] [nvarchar](max) NULL,
	[vd_use_ref_as_name] [bit] NOT NULL,
	[vd_file] [nvarchar](50) NULL,
	[vd_filetype] [varchar](4) NULL,
	[vd_macro01] [nvarchar](50) NULL,
 CONSTRAINT [PK_VIEW_DOC] PRIMARY KEY CLUSTERED 
(
	[vd_id] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[rsUsersTemplatesVisible]    Script Date: 01/22/2016 17:59:47 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[rsUsersTemplatesVisible](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[TemplateID] [nvarchar](max) NULL,
	[UserID] [nvarchar](max) NULL,
	[Visible] [nchar](10) NULL,
 CONSTRAINT [PK_rsUsersTemplatesVisible] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[rsUserGroup]    Script Date: 01/22/2016 17:59:47 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[rsUserGroup](
	[ft_id] [int] IDENTITY(1,1) NOT NULL,
	[UserID] [nvarchar](150) NULL,
	[GroupID] [nvarchar](50) NULL,
 CONSTRAINT [PK_rsUserGroup] PRIMARY KEY CLUSTERED 
(
	[ft_id] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[rsTemplateCreateXMLTextProfile]    Script Date: 01/22/2016 17:59:47 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[rsTemplateCreateXMLTextProfile](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[Field] [nvarchar](50) NULL,
	[FriendlyName] [nvarchar](50) NULL,
	[Visible] [bit] NULL,
	[DefaultValue] [nvarchar](max) NULL,
	[SunComponentName] [nvarchar](max) NULL,
	[SunMethod] [nvarchar](max) NULL,
	[Mandatory] [bit] NULL,
	[Separator] [nvarchar](max) NULL,
	[TextLength] [nvarchar](50) NULL,
	[Prefix] [nvarchar](50) NULL,
	[Suffix] [nvarchar](50) NULL,
	[RemoveCharacters] [nvarchar](max) NULL,
	[TextFileName] [nvarchar](max) NULL,
	[Parent] [nvarchar](max) NULL,
 CONSTRAINT [PK_rsTemplateCreateXMLTextProfile] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[rsTemplateXMLTEXTFiles]    Script Date: 01/22/2016 17:59:47 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[rsTemplateXMLTEXTFiles](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[FileContent] [nvarchar](max) NULL,
	[RelatedName] [nvarchar](max) NULL,
	[FileType] [int] NULL,
 CONSTRAINT [PK_rsTemplateXMLTEXTFiles] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[rsGlobalFields]    Script Date: 01/22/2016 17:59:47 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[rsGlobalFields](
	[GUID] [uniqueidentifier] ROWGUIDCOL  NOT NULL,
	[FieldGroup] [nvarchar](50) NULL,
	[SunField] [nvarchar](50) NULL,
	[UserFriendlyName] [nvarchar](50) NULL,
	[Output] [nvarchar](50) NULL,
	[Input] [nvarchar](50) NULL,
	[XML_Query] [nvarchar](max) NULL,
	[version] [int] NULL,
 CONSTRAINT [PK_FinTools_Settings_Fields] PRIMARY KEY CLUSTERED 
(
	[GUID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[rsGroupPermissions]    Script Date: 01/22/2016 17:59:47 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[rsGroupPermissions](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[GroupID] [nvarchar](150) NULL,
	[PermissionID] [nvarchar](150) NULL,
	[PermissionGroupName] [nvarchar](max) NULL,
 CONSTRAINT [PK_rsGroupPermissions] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[rsPermissions]    Script Date: 01/22/2016 17:59:47 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[rsPermissions](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[PermissionName] [nvarchar](150) NULL,
	[TemplateID] [nvarchar](150) NULL,
	[ActionID] [nvarchar](150) NULL,
	[Per_Type] [int] NULL,
	[remark] [nvarchar](3000) NULL,
 CONSTRAINT [PK_rsPermissions] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  StoredProcedure [dbo].[rsGroups_Upd]    Script Date: 01/22/2016 17:59:47 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
create procedure [dbo].[rsGroups_Upd]
@id INT,
@GroupName nvarchar(150),
@GroupDisable bit,
@Remark nvarchar(4000),
@AddInTabName nvarchar(max)
as
begin
        Update dbo.rsGroups set GroupName=@GroupName,GroupDisable=@GroupDisable,Remark=@Remark, AddInTabName=@AddInTabName where id=@id
end
GO
/****** Object:  StoredProcedure [dbo].[rsGroups_Ins]    Script Date: 01/22/2016 17:59:47 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[rsGroups_Ins]
@GroupName nvarchar(150),
@GroupDisable bit,
@Remark nvarchar(4000),
@AddInTabName nvarchar(max)
as
begin
DECLARE @ReturnValue INT ;
        insert into dbo.rsGroups(GroupName,GroupDisable,Remark,AddInTabName) values(@GroupName,@GroupDisable,@Remark,@AddInTabName)

 set @ReturnValue= @@IDENTITY;
         return @ReturnValue
end
GO
/****** Object:  StoredProcedure [dbo].[rsGroups_Del]    Script Date: 01/22/2016 17:59:47 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
create procedure [dbo].[rsGroups_Del]
	@id int
	
as
begin
       Delete rsGroups where ID=@id
     
end
GO
/****** Object:  StoredProcedure [dbo].[rsSystemProcesses_Ins]    Script Date: 01/22/2016 17:59:47 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[rsSystemProcesses_Ins]
@ProcessName nvarchar(100)

as
begin
        insert into dbo.rsSystemProcesses(ProcessName) values(@ProcessName)
end
GO
/****** Object:  StoredProcedure [dbo].[rsTemplateActions_UpdGroupName]    Script Date: 01/22/2016 17:59:47 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[rsTemplateActions_UpdGroupName]

@ButtonGroup nvarchar(50),
@GroupOrder int,
@id int

as
begin
        Update dbo.rsTemplateActions set ButtonGroup=@ButtonGroup,GroupOrder=@GroupOrder where ID=@id
end
GO
/****** Object:  StoredProcedure [dbo].[rsTemplateActions_UpdActionOrder]    Script Date: 01/22/2016 17:59:47 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[rsTemplateActions_UpdActionOrder]

@ButtonOrder int,
@id int

as
begin
        Update dbo.rsTemplateActions set ButtonOrder=@ButtonOrder where ID=@id
end
GO
/****** Object:  StoredProcedure [dbo].[rsTemplateActions_Ins]    Script Date: 01/22/2016 17:59:47 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[rsTemplateActions_Ins]

@Name nvarchar(50),
@Text nvarchar(50),
@ButtonIcon nvarchar(150),
@ButtonGroup nvarchar(50),
@ButtonSize nvarchar(50),
@ButtonOrder int,
@GroupOrder int,
@StopOnError bit,
@templateID nvarchar(MAX),
@ProcessID nvarchar(MAX),
@MacroName nvarchar(50),
@ProcessMacroOrder INT,
@Type INT

as
begin
        insert into dbo.rsTemplateActions(ButtonName,ButtonText,ButtonIcon,ButtonGroup,ButtonSize,ButtonOrder,GroupOrder,StopOnError,templateID,ProcessID,MacroName,ProcessMacroOrder,[Type]) values(@Name,@Text,@ButtonIcon,@ButtonGroup,@ButtonSize,@ButtonOrder,@GroupOrder,@StopOnError,@templateID,@ProcessID,@MacroName,@ProcessMacroOrder,@Type)
end
GO
/****** Object:  StoredProcedure [dbo].[rsTemplateActions_Del]    Script Date: 01/22/2016 17:59:47 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[rsTemplateActions_Del]
	@id int
	
as
begin
       Delete rsTemplateActions where ID=@id
     
end
GO
/****** Object:  StoredProcedure [dbo].[rsTemplateAllocationMakerUpdate_Ins]    Script Date: 01/22/2016 17:59:47 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[rsTemplateAllocationMakerUpdate_Ins]
@JournalNumber nvarchar(50) ,
@Ledger nvarchar(50),
@ft_Account nvarchar(50),
@Period nvarchar(50),
@TransactionDate nvarchar(50),
@JrnlType nvarchar(50),
@TransRef nvarchar(50),
@AlloctnMarker nvarchar(50),
@LA1 nvarchar(50),
@LA2 nvarchar(50),
@LA3 nvarchar(50),
@LA4 nvarchar(50),
@LA5 nvarchar(50),
@LA6 nvarchar(50),
@LA7 nvarchar(50),
@LA8 nvarchar(50),
@LA9 nvarchar(50),
@LA10 nvarchar(50),
@TemplateID nvarchar(max),
@LineIndicator nvarchar(50) ,
@StartinginCell nvarchar(50),
@inputFields nvarchar(max),
@updateFields nvarchar(max),
@JournalLineNumber nvarchar(50)
as
begin
        insert into dbo.rsTemplateAllocationMakerUpdate
        (JournalNumber  ,
	Ledger,ft_Account,Period,TransactionDate,
        JrnlType,TransRef,
        AlloctnMarker,LA1,
        LA2,LA3,LA4,LA5,LA6,LA7,LA8,LA9,LA10,
        
        TemplateID,LineIndicator,StartinginCell,inputFields,updateFields,JournalLineNumber) 
        values(@JournalNumber  ,
	@Ledger,@ft_Account,@Period,@TransactionDate,@JrnlType,
        @TransRef,
        @AlloctnMarker,@LA1,@LA2,@LA3,@LA4,
        @LA5,@LA6,@LA7,@LA8,@LA9,@LA10,
        
        @TemplateID,@LineIndicator,@StartinginCell,@inputFields,
@updateFields ,@JournalLineNumber)
end
GO
/****** Object:  StoredProcedure [dbo].[rsTemplateAllocationMakerUpdate_Del]    Script Date: 01/22/2016 17:59:47 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[rsTemplateAllocationMakerUpdate_Del]
@TemplateID nvarchar(MAX)

as
begin
        delete dbo.rsTemplateAllocationMakerUpdate where TemplateID=@TemplateID
end
GO
/****** Object:  StoredProcedure [dbo].[rsTemplateConsolidation_Ins]    Script Date: 01/22/2016 17:59:47 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[rsTemplateConsolidation_Ins]
@Ledger nvarchar(50),
@ft_Account nvarchar(50),
@Period nvarchar(50),
@TransDate nvarchar(50),
@DueDate nvarchar(50),
@JrnlType nvarchar(50),
@JrnlSource nvarchar(50),
@TransRef nvarchar(50),
@Description nvarchar(50),
@AlloctnMarker nvarchar(50),
@LA1 nvarchar(50),
@LA2 nvarchar(50),
@LA3 nvarchar(50),
@LA4 nvarchar(50),
@LA5 nvarchar(50),
@LA6 nvarchar(50),
@LA7 nvarchar(50),
@LA8 nvarchar(50),
@LA9 nvarchar(50),
@LA10 nvarchar(50),
@GenDesc1 nvarchar(50),
@GenDesc2 nvarchar(50),
@GenDesc3 nvarchar(50),
@GenDesc4 nvarchar(50),
@GenDesc5 nvarchar(50),
@GenDesc6 nvarchar(50),
@GenDesc7 nvarchar(50),
@GenDesc8 nvarchar(50) ,
@GenDesc9 nvarchar(50) ,
@GenDesc10 nvarchar(50) ,
@GenDesc11 nvarchar(50) ,
@GenDesc12 nvarchar(50) ,
@GenDesc13 nvarchar(50) ,
@GenDesc14 nvarchar(50) ,
@GenDesc15 nvarchar(50) ,
@GenDesc16 nvarchar(50) ,
@GenDesc17 nvarchar(50) ,
@GenDesc18 nvarchar(50) ,
@GenDesc19 nvarchar(50) ,
@GenDesc20 nvarchar(50) ,
@GenDesc21 nvarchar(50) ,
@GenDesc22 nvarchar(50) ,
@GenDesc23 nvarchar(50) ,
@GenDesc24 nvarchar(50) ,
@GenDesc25 nvarchar(50) ,
@TransAmount nvarchar(50),
@Currency nvarchar(50),
@BaseAmount nvarchar(50),
@2ndBase nvarchar(50),
@4thAmount nvarchar(50),
@TemplateID nvarchar(max),
@LineIndicator nvarchar(50) ,
@StartinginCell nvarchar(50),
@ConsolidateBy1 nvarchar(50) ,
@ConsolidateBy2 nvarchar(50),
@ConsolidateBy3 nvarchar(50) ,
@ConsolidateBy4 nvarchar(50),
@BalanceBy nvarchar(50),
@PopWithJNNumber nvarchar(50)

as
begin
        insert into dbo.rsTemplateConsolidation
        (Ledger,ft_Account,Period,TransDate,DueDate,
        JrnlType,JrnlSource,TransRef,
        Description,AlloctnMarker,LA1,
        LA2,LA3,LA4,LA5,LA6,LA7,LA8,LA9,LA10,
        GenDesc1,GenDesc2,GenDesc3,GenDesc4,GenDesc5,GenDesc6,GenDesc7,
GenDesc8,
GenDesc9,
GenDesc10,
GenDesc11,
GenDesc12,
GenDesc13,
GenDesc14,
GenDesc15,
GenDesc16,
GenDesc17,
GenDesc18,
GenDesc19,
GenDesc20,
GenDesc21,
GenDesc22,
GenDesc23,
GenDesc24,
GenDesc25,
        TransAmount,Currency,BaseAmount,[2ndBase],[4thAmount],TemplateID,LineIndicator,StartinginCell,ConsolidateBy1,ConsolidateBy2,ConsolidateBy3,ConsolidateBy4,BalanceBy,PopWithJNNumber) 
        values(@Ledger,@ft_Account,@Period,@TransDate,@DueDate,@JrnlType,
        @JrnlSource,@TransRef,@Description,
        @AlloctnMarker,@LA1,@LA2,@LA3,@LA4,
        @LA5,@LA6,@LA7,@LA8,@LA9,@LA10,
        @GenDesc1,@GenDesc2,@GenDesc3,
        @GenDesc4,@GenDesc5,@GenDesc6,@GenDesc7,
@GenDesc8,
@GenDesc9,
@GenDesc10,
@GenDesc11,
@GenDesc12,
@GenDesc13,
@GenDesc14,
@GenDesc15,
@GenDesc16,
@GenDesc17,
@GenDesc18,
@GenDesc19,
@GenDesc20,
@GenDesc21,
@GenDesc22,
@GenDesc23,
@GenDesc24,
@GenDesc25,
        @TransAmount,@Currency,@BaseAmount,
        @2ndBase,@4thAmount,@TemplateID,@LineIndicator,@StartinginCell,@ConsolidateBy1,@ConsolidateBy2,@ConsolidateBy3,@ConsolidateBy4,@BalanceBy,@PopWithJNNumber)
end
GO
/****** Object:  StoredProcedure [dbo].[rsTemplateConsolidation_Del]    Script Date: 01/22/2016 17:59:47 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[rsTemplateConsolidation_Del]
@TemplateID nvarchar(max)

as
begin
        delete dbo.rsTemplateConsolidation where TemplateID=@TemplateID 
end
GO
/****** Object:  StoredProcedure [dbo].[rsTemplateContainer_Ins]    Script Date: 01/22/2016 17:59:47 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[rsTemplateContainer_Ins]
@TemplateID  nvarchar(max),
@ft_relatefilepath nvarchar(max),
@column nvarchar(5),
@FromDB bit

as
begin
        insert into dbo.rsTemplateContainer(ft_id,TemplateID,ft_relatefilepath,[column],FromDB) values(newid(),@TemplateID,@ft_relatefilepath,@column,@FromDB)
end
GO
/****** Object:  StoredProcedure [dbo].[rsTemplateContainer_Del]    Script Date: 01/22/2016 17:59:47 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[rsTemplateContainer_Del]
@TemplateID nvarchar(max)

as
begin
        delete dbo.rsTemplateContainer where TemplateID=@TemplateID 
end
GO
/****** Object:  StoredProcedure [dbo].[rsTemplateDrillDown_Ins]    Script Date: 01/22/2016 17:59:48 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[rsTemplateDrillDown_Ins]

@SunField nvarchar(50) ,
	@InputStatus nvarchar(50) ,
	@CellName nvarchar(50) ,
	@OutputStatus nvarchar(50) ,
	@TemplateID nvarchar(500) ,
	@order int
	
as
begin
        insert into dbo.rsTemplateDrillDown
        (SunField,InputStatus,CellName,OutputStatus,TemplateID,[order]) 
        values(@SunField,@InputStatus,@CellName,@OutputStatus,@TemplateID,@order)
end
GO
/****** Object:  StoredProcedure [dbo].[rsTemplateDrillDown_Del]    Script Date: 01/22/2016 17:59:47 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[rsTemplateDrillDown_Del]
@TemplateID nvarchar(500)

as
begin
        delete dbo.rsTemplateDrillDown where TemplateID=@TemplateID 
end
GO
/****** Object:  StoredProcedure [dbo].[rsTemplateGenDescFields_Ins]    Script Date: 01/22/2016 17:59:48 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[rsTemplateGenDescFields_Ins]

@FieldGroup nvarchar(50),
@SunField nvarchar(50),
@UserFriendlyName	nvarchar(50),
@Output nvarchar(50),
@Input nvarchar(50),
@XML_Query nvarchar(max),
@TemplateID nvarchar(max),
@version int

as
begin
        insert into dbo.rsTemplateGenDescFields([GUID],FieldGroup,SunField,UserFriendlyName,[Output],Input,XML_Query,TemplateID,[version]) 
        
        values(newid(),@FieldGroup,@SunField,@UserFriendlyName,@Output,@Input,@XML_Query,@TemplateID,@version)
end
GO
/****** Object:  StoredProcedure [dbo].[rsTemplateGenDescFields_Del]    Script Date: 01/22/2016 17:59:48 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[rsTemplateGenDescFields_Del]

@TemplateID nvarchar(max)

as
begin
        delete dbo.rsTemplateGenDescFields where TemplateID=@TemplateID
end
GO
/****** Object:  StoredProcedure [dbo].[rsTemplateHelp_Ins]    Script Date: 01/22/2016 17:59:48 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[rsTemplateHelp_Ins]


	@TemplateID nvarchar(max) ,
	@helpFileData image ,
	@helpFileType nvarchar(50) 
	
as
begin
        insert into dbo.rsTemplateHelp(TemplateID,helpFileData,helpFileType) values(@TemplateID,@helpFileData,@helpFileType)
end
GO
/****** Object:  StoredProcedure [dbo].[rsTemplateHelp_Del]    Script Date: 01/22/2016 17:59:48 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[rsTemplateHelp_Del]
	@TemplateID nvarchar(max) 
	
as
begin
       Delete rsTemplateHelp where TemplateID=@TemplateID
     
end
GO
/****** Object:  StoredProcedure [dbo].[rsTemplateJournal_Ins]    Script Date: 01/22/2016 17:59:48 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[rsTemplateJournal_Ins]
@Ledger nvarchar(50),
@ft_Account nvarchar(50),
@Period nvarchar(50),
@TransDate nvarchar(50),
@DueDate nvarchar(50),
@JrnlType nvarchar(50),
@JrnlSource nvarchar(50),
@TransRef nvarchar(50),
@Description nvarchar(50),
@AlloctnMarker nvarchar(50),
@LA1 nvarchar(50),
@LA2 nvarchar(50),
@LA3 nvarchar(50),
@LA4 nvarchar(50),
@LA5 nvarchar(50),
@LA6 nvarchar(50),
@LA7 nvarchar(50),
@LA8 nvarchar(50),
@LA9 nvarchar(50),
@LA10 nvarchar(50),
@GenDesc1 nvarchar(50),
@GenDesc2 nvarchar(50),
@GenDesc3 nvarchar(50),
@GenDesc4 nvarchar(50),
@GenDesc5 nvarchar(50),
@GenDesc6 nvarchar(50),
@GenDesc7 nvarchar(50),
@GenDesc8 nvarchar(50) ,
@GenDesc9 nvarchar(50) ,
@GenDesc10 nvarchar(50) ,
@GenDesc11 nvarchar(50) ,
@GenDesc12 nvarchar(50) ,
@GenDesc13 nvarchar(50) ,
@GenDesc14 nvarchar(50) ,
@GenDesc15 nvarchar(50) ,
@GenDesc16 nvarchar(50) ,
@GenDesc17 nvarchar(50) ,
@GenDesc18 nvarchar(50) ,
@GenDesc19 nvarchar(50) ,
@GenDesc20 nvarchar(50) ,
@GenDesc21 nvarchar(50) ,
@GenDesc22 nvarchar(50) ,
@GenDesc23 nvarchar(50) ,
@GenDesc24 nvarchar(50) ,
@GenDesc25 nvarchar(50) ,
@TransAmount nvarchar(50),
@Currency nvarchar(50),
@BaseAmount nvarchar(50),
@2ndBase nvarchar(50),
@4thAmount nvarchar(50),
@TemplateID nvarchar(max),
@LineIndicator nvarchar(50) ,
@StartinginCell nvarchar(50),
@BalanceBy nvarchar(50),
@PopWithJNNumber nvarchar(50)

as
begin
        insert into dbo.rsTemplateJournal
        (Ledger,ft_Account,Period,TransDate,DueDate,
        JrnlType,JrnlSource,TransRef,
        Description,AlloctnMarker,LA1,
        LA2,LA3,LA4,LA5,LA6,LA7,LA8,LA9,LA10,
        GenDesc1,GenDesc2,GenDesc3,GenDesc4,GenDesc5,GenDesc6,GenDesc7,
GenDesc8,
GenDesc9,
GenDesc10,
GenDesc11,
GenDesc12,
GenDesc13,
GenDesc14,
GenDesc15,
GenDesc16,
GenDesc17,
GenDesc18,
GenDesc19,
GenDesc20,
GenDesc21,
GenDesc22,
GenDesc23,
GenDesc24,
GenDesc25,
        TransAmount,Currency,BaseAmount,[2ndBase],[4thAmount],TemplateID,LineIndicator,StartinginCell,BalanceBy,PopWithJNNumber) 
        values(@Ledger,@ft_Account,@Period,@TransDate,@DueDate,@JrnlType,
        @JrnlSource,@TransRef,@Description,
        @AlloctnMarker,@LA1,@LA2,@LA3,@LA4,
        @LA5,@LA6,@LA7,@LA8,@LA9,@LA10,
        @GenDesc1,@GenDesc2,@GenDesc3,
        @GenDesc4,@GenDesc5,@GenDesc6,@GenDesc7,
@GenDesc8,
@GenDesc9,
@GenDesc10,
@GenDesc11,
@GenDesc12,
@GenDesc13,
@GenDesc14,
@GenDesc15,
@GenDesc16,
@GenDesc17,
@GenDesc18,
@GenDesc19,
@GenDesc20,
@GenDesc21,
@GenDesc22,
@GenDesc23,
@GenDesc24,
@GenDesc25,
        @TransAmount,@Currency,@BaseAmount,
        @2ndBase,@4thAmount,@TemplateID,@LineIndicator,@StartinginCell,@BalanceBy,@PopWithJNNumber)
end
GO
/****** Object:  StoredProcedure [dbo].[rsTemplateJournal_Del]    Script Date: 01/22/2016 17:59:48 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[rsTemplateJournal_Del]
@TemplateID nvarchar(max)

as
begin
        delete dbo.rsTemplateJournal where TemplateID=@TemplateID 
end
GO
/****** Object:  View [dbo].[View_TemplatesPermissions]    Script Date: 01/22/2016 17:59:47 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE VIEW [dbo].[View_TemplatesPermissions]
AS
SELECT     dbo.rsPermissions.ID, dbo.rsPermissions.PermissionName, dbo.rsPermissions.TemplateID, dbo.rsPermissions.ActionID, dbo.rsPermissions.Per_Type, 
                      dbo.rsPermissions.remark, dbo.rsTemplates.ID AS Expr1, dbo.rsTemplates.TemplateData, dbo.rsTemplates.TemplateName, dbo.rsTemplates.OriginTemplatePath, 
                      dbo.rsTemplates.FileType, dbo.rsTemplates.Description, 
                      CASE dbo.rsPermissions.Per_Type WHEN '0' THEN 'Save/Amend' WHEN '1' THEN 'Write' WHEN '2' THEN 'Read' WHEN '3' THEN 'Global' ELSE 'New Action' END AS PerType
FROM         dbo.rsPermissions LEFT OUTER JOIN
                      dbo.rsTemplates ON dbo.rsPermissions.TemplateID = dbo.rsTemplates.ID
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[40] 4[20] 2[20] 3) )"
      End
      Begin PaneConfiguration = 1
         NumPanes = 3
         Configuration = "(H (1 [50] 4 [25] 3))"
      End
      Begin PaneConfiguration = 2
         NumPanes = 3
         Configuration = "(H (1 [50] 2 [25] 3))"
      End
      Begin PaneConfiguration = 3
         NumPanes = 3
         Configuration = "(H (4 [30] 2 [40] 3))"
      End
      Begin PaneConfiguration = 4
         NumPanes = 2
         Configuration = "(H (1 [56] 3))"
      End
      Begin PaneConfiguration = 5
         NumPanes = 2
         Configuration = "(H (2 [66] 3))"
      End
      Begin PaneConfiguration = 6
         NumPanes = 2
         Configuration = "(H (4 [50] 3))"
      End
      Begin PaneConfiguration = 7
         NumPanes = 1
         Configuration = "(V (3))"
      End
      Begin PaneConfiguration = 8
         NumPanes = 3
         Configuration = "(H (1[56] 4[18] 2) )"
      End
      Begin PaneConfiguration = 9
         NumPanes = 2
         Configuration = "(H (1 [75] 4))"
      End
      Begin PaneConfiguration = 10
         NumPanes = 2
         Configuration = "(H (1[66] 2) )"
      End
      Begin PaneConfiguration = 11
         NumPanes = 2
         Configuration = "(H (4 [60] 2))"
      End
      Begin PaneConfiguration = 12
         NumPanes = 1
         Configuration = "(H (1) )"
      End
      Begin PaneConfiguration = 13
         NumPanes = 1
         Configuration = "(V (4))"
      End
      Begin PaneConfiguration = 14
         NumPanes = 1
         Configuration = "(V (2))"
      End
      ActivePaneConfig = 0
   End
   Begin DiagramPane = 
      Begin Origin = 
         Top = 0
         Left = 0
      End
      Begin Tables = 
         Begin Table = "rsPermissions"
            Begin Extent = 
               Top = 6
               Left = 38
               Bottom = 192
               Right = 204
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "rsTemplates"
            Begin Extent = 
               Top = 6
               Left = 242
               Bottom = 187
               Right = 425
            End
            DisplayFlags = 280
            TopColumn = 0
         End
      End
   End
   Begin SQLPane = 
   End
   Begin DataPane = 
      Begin ParameterDefaults = ""
      End
   End
   Begin CriteriaPane = 
      Begin ColumnWidths = 11
         Column = 1440
         Alias = 900
         Table = 1170
         Output = 720
         Append = 1400
         NewValue = 1170
         SortType = 1350
         SortOrder = 1410
         GroupBy = 1350
         Filter = 1350
         Or = 1350
         Or = 1350
         Or = 1350
      End
   End
End
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'View_TemplatesPermissions'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=1 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'View_TemplatesPermissions'
GO
/****** Object:  StoredProcedure [dbo].[rsTemplates_Ins]    Script Date: 01/22/2016 17:59:48 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[rsTemplates_Ins]
@TemplateData image,
@TemplateName	nvarchar(max),
@OriginTemplatePath	nvarchar(max),
@FileType nvarchar(50),
@Description nvarchar(max)

as
begin
	DECLARE @ReturnValue INT ;
        insert into dbo.rsTemplates(TemplateData,TemplateName,OriginTemplatePath,FileType,[Description]) 
        
        values(@TemplateData,@TemplateName,@OriginTemplatePath,@FileType,@Description);
        
         set @ReturnValue= @@IDENTITY;
         return @ReturnValue
end
GO
/****** Object:  StoredProcedure [dbo].[rsTemplateSequenceNumbering_Ins]    Script Date: 01/22/2016 17:59:48 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[rsTemplateSequenceNumbering_Ins]

@UseSequenceNum int ,
	@SequencePrefix nvarchar(50) ,
	@PostToField nvarchar(50) ,
	@PopulateCell nvarchar(50) ,
	@TemplateID nvarchar(500) 


as
begin
        insert into dbo.rsTemplateSequenceNumbering(ft_id,UseSequenceNum,SequencePrefix,PostToField,PopulateCell,TemplateID) values(newid(),@UseSequenceNum ,@SequencePrefix ,@PostToField ,@PopulateCell ,@TemplateID)
end
GO
/****** Object:  StoredProcedure [dbo].[rsTemplateSequenceNumbering_Del]    Script Date: 01/22/2016 17:59:48 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[rsTemplateSequenceNumbering_Del]
@TemplateID varchar(500)

as
begin
        delete dbo.rsTemplateSequenceNumbering where TemplateID=@TemplateID
end
GO
/****** Object:  StoredProcedure [dbo].[rsTemplateSetting_Ins]    Script Date: 01/22/2016 17:59:48 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[rsTemplateSetting_Ins]

@CriteriaName nvarchar(50) ,
@CellReference nvarchar(50) ,
@TemplateID nvarchar(500) ,
@orderNum int,
@OpenTransUponSave bit
	

as
begin
        insert into dbo.rsTemplateSetting(ft_id,CriteriaName,CellReference,TemplateID, orderNum,OpenTransUponSave) values(newid(),@CriteriaName,@CellReference,@TemplateID,@orderNum,@OpenTransUponSave)
end
GO
/****** Object:  StoredProcedure [dbo].[rsTemplateSetting_Del]    Script Date: 01/22/2016 17:59:48 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[rsTemplateSetting_Del]
@TemplateID varchar(500)

as
begin
        delete dbo.rsTemplateSetting where TemplateID=@TemplateID
end
GO
/****** Object:  StoredProcedure [dbo].[rsTemplateTransactions_UpdateCriterias]    Script Date: 01/22/2016 17:59:48 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[rsTemplateTransactions_UpdateCriterias]
	@Criteria1 nvarchar(50) ,
	@Criteria2 nvarchar(50) ,
	@Criteria3 nvarchar(50) ,
	@Criteria4 nvarchar(50) ,
	@Criteria5 nvarchar(50) ,
	@Value1 nvarchar(50) ,
	@Value2 nvarchar(50) ,
	@Value3 nvarchar(50) ,
	@Value4 nvarchar(50) ,
	@Value5 nvarchar(50) ,
	@TemplateID nvarchar(150) 


as
begin

        update  rsTemplateTransactions
        set    	
	Criteria1=@Criteria1  ,
	Criteria2 =@Criteria2 ,
	Criteria3 =@Criteria3 ,
	Criteria4 =@Criteria4 ,
	Criteria5 =@Criteria5 ,
	Value1  =@Value1,
	Value2  =@Value2 ,
	Value3   =@Value3,
	Value4  =@Value4 ,
	Value5   =@Value5
	where TemplateID=@TemplateID
	  
	
end
GO
/****** Object:  StoredProcedure [dbo].[rsTemplateTransactions_Ins]    Script Date: 01/22/2016 17:59:48 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[rsTemplateTransactions_Ins]
	@TemplateName nvarchar(50) ,
	@Criteria1 nvarchar(50) ,
	@Criteria2 nvarchar(50) ,
	@Criteria3 nvarchar(50) ,
	@Criteria4 nvarchar(50) ,
	@Criteria5 nvarchar(50) ,
	@Value1 nvarchar(50) ,
	@Value2 nvarchar(50) ,
	@Value3 nvarchar(50) ,
	@Value4 nvarchar(50) ,
	@Value5 nvarchar(50) ,
	@Data image ,
	@DataType nvarchar(50) ,
	@PDFData image,
	@XMLData varchar(8000) ,
	@TemplateID nvarchar(150) ,
	@maxNum int ,
	@TransactionName  nvarchar(50),
	@Prefix nvarchar(50),
	@SunJournalNumber nvarchar(100)

as
begin
        insert into rsTemplateTransactions
        (    	TemplateName  ,
	Criteria1  ,
	Criteria2  ,
	Criteria3  ,
	Criteria4  ,
	Criteria5  ,
	Value1  ,
	Value2  ,
	Value3  ,
	Value4  ,
	Value5  ,
	Data  ,
	DataType  ,
	PDFData ,
	XMLData ,
	TemplateID  ,
	maxNum,TransactionName ,Prefix,SunJournalNumber) 
        values(
        
       	@TemplateName  ,
	@Criteria1  ,
	@Criteria2  ,
	@Criteria3,
	@Criteria4,
	@Criteria5,
	@Value1  ,
	@Value2  ,
	@Value3  ,
	@Value4  ,
	@Value5  ,
	@Data  ,
	@DataType  ,
	@PDFData ,
	@XMLData ,
	@TemplateID  ,
	@maxNum,@TransactionName ,@Prefix ,@SunJournalNumber)
end
GO
/****** Object:  StoredProcedure [dbo].[rsTemplateTransactions_Del]    Script Date: 01/22/2016 17:59:48 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[rsTemplateTransactions_Del]
	@TemplateID nvarchar(150) ,
	@TransactionName  nvarchar(50)
as
begin
       Delete rsTemplateTransactions where TemplateID=@TemplateID and  TransactionName =@TransactionName 
     
end
GO
/****** Object:  StoredProcedure [dbo].[rsTemplateTransactionUpdate_Ins]    Script Date: 01/22/2016 17:59:48 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[rsTemplateTransactionUpdate_Ins]
@JournalNumber nvarchar(50) ,
@JournalLineNumber nvarchar(50) ,
@Ledger nvarchar(50),
@ft_Account nvarchar(50),
@Period nvarchar(50),
@TransDate nvarchar(50),
@DueDate nvarchar(50),
@JrnlType nvarchar(50),
@JrnlSource nvarchar(50),
@TransRef nvarchar(50),
@Description nvarchar(50),
@AlloctnMarker nvarchar(50),
@LA1 nvarchar(50),
@LA2 nvarchar(50),
@LA3 nvarchar(50),
@LA4 nvarchar(50),
@LA5 nvarchar(50),
@LA6 nvarchar(50),
@LA7 nvarchar(50),
@LA8 nvarchar(50),
@LA9 nvarchar(50),
@LA10 nvarchar(50),
@GenDesc1 nvarchar(50),
@GenDesc2 nvarchar(50),
@GenDesc3 nvarchar(50),
@GenDesc4 nvarchar(50),
@GenDesc5 nvarchar(50),
@GenDesc6 nvarchar(50),
@GenDesc7 nvarchar(50),
@GenDesc8 nvarchar(50) ,
@GenDesc9 nvarchar(50) ,
@GenDesc10 nvarchar(50) ,
@GenDesc11 nvarchar(50) ,
@GenDesc12 nvarchar(50) ,
@GenDesc13 nvarchar(50) ,
@GenDesc14 nvarchar(50) ,
@GenDesc15 nvarchar(50) ,
@GenDesc16 nvarchar(50) ,
@GenDesc17 nvarchar(50) ,
@GenDesc18 nvarchar(50) ,
@GenDesc19 nvarchar(50) ,
@GenDesc20 nvarchar(50) ,
@GenDesc21 nvarchar(50) ,
@GenDesc22 nvarchar(50) ,
@GenDesc23 nvarchar(50) ,
@GenDesc24 nvarchar(50) ,
@GenDesc25 nvarchar(50) ,
@TransAmount nvarchar(50),
@Currency nvarchar(50),
@BaseAmount nvarchar(50),
@2ndBase nvarchar(50),
@4thAmount nvarchar(50),
@TemplateID nvarchar(max),
@LineIndicator nvarchar(50) ,
@StartinginCell nvarchar(50),
@inputFields nvarchar(max),
@updateFields nvarchar(max)


as
begin
        insert into dbo.rsTemplateTransactionUpdate
        (JournalNumber  ,
	JournalLineNumber ,Ledger,ft_Account,Period,TransDate,DueDate,
        JrnlType,JrnlSource,TransRef,
        Description,AlloctnMarker,LA1,
        LA2,LA3,LA4,LA5,LA6,LA7,LA8,LA9,LA10,
        GenDesc1,GenDesc2,GenDesc3,GenDesc4,GenDesc5,GenDesc6,GenDesc7,
GenDesc8,
GenDesc9,
GenDesc10,
GenDesc11,
GenDesc12,
GenDesc13,
GenDesc14,
GenDesc15,
GenDesc16,
GenDesc17,
GenDesc18,
GenDesc19,
GenDesc20,
GenDesc21,
GenDesc22,
GenDesc23,
GenDesc24,
GenDesc25,
        TransAmount,Currency,BaseAmount,[2ndBase],[4thAmount],TemplateID,LineIndicator,StartinginCell,inputFields,updateFields) 
        values(@JournalNumber  ,
	@JournalLineNumber ,@Ledger,@ft_Account,@Period,@TransDate,@DueDate,@JrnlType,
        @JrnlSource,@TransRef,@Description,
        @AlloctnMarker,@LA1,@LA2,@LA3,@LA4,
        @LA5,@LA6,@LA7,@LA8,@LA9,@LA10,
        @GenDesc1,@GenDesc2,@GenDesc3,
        @GenDesc4,@GenDesc5,@GenDesc6,@GenDesc7,
@GenDesc8,
@GenDesc9,
@GenDesc10,
@GenDesc11,
@GenDesc12,
@GenDesc13,
@GenDesc14,
@GenDesc15,
@GenDesc16,
@GenDesc17,
@GenDesc18,
@GenDesc19,
@GenDesc20,
@GenDesc21,
@GenDesc22,
@GenDesc23,
@GenDesc24,
@GenDesc25,
        @TransAmount,@Currency,@BaseAmount,
        @2ndBase,@4thAmount,@TemplateID,@LineIndicator,@StartinginCell,@inputFields,
@updateFields )
end
GO
/****** Object:  StoredProcedure [dbo].[rsTemplateTransactionUpdate_Del]    Script Date: 01/22/2016 17:59:48 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[rsTemplateTransactionUpdate_Del]
@TemplateID nvarchar(MAX)

as
begin
        delete dbo.rsTemplateTransactionUpdate where TemplateID=@TemplateID 
end
GO
/****** Object:  StoredProcedure [dbo].[rsTemplateVisible_Ins]    Script Date: 01/22/2016 17:59:48 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[rsTemplateVisible_Ins]
@TemplateID  nvarchar(500),
@OutputPaneVisiable nvarchar(5),
@UserID nvarchar(50)

as
begin
        insert into dbo.rsTemplateVisible(ft_id,TemplateID,OutputPaneVisiable,UserID) values(newid(),@TemplateID,@OutputPaneVisiable,@UserID)
end
GO
/****** Object:  StoredProcedure [dbo].[rsTemplateVisible_Del]    Script Date: 01/22/2016 17:59:48 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[rsTemplateVisible_Del]

@TemplateID nvarchar(500),
@UserID nvarchar(50)

as
begin
        delete dbo.rsTemplateVisible where TemplateID=@TemplateID and UserID=@UserID
end
GO
/****** Object:  StoredProcedure [dbo].[rsUpgrade_Ins]    Script Date: 01/22/2016 17:59:48 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[rsUpgrade_Ins]
@RootPath nvarchar(max),
@V1IP	nvarchar(max),
@V1UserID	nvarchar(max),
@V1Password	nvarchar(max),
@V2IP	nvarchar(max),
@V2UserID	nvarchar(max),
@V2Password	nvarchar(max)

as
begin
        insert into dbo.rsUpgrade(RootPath,V1IP,V1UserID,V1Password,V2IP,V2UserID,V2Password) 
        
        values(@RootPath,@V1IP,@V1UserID,@V1Password,@V2IP,@V2UserID,@V2Password)
end
GO
/****** Object:  StoredProcedure [dbo].[rsUsers_UserSunInfo_Upd]    Script Date: 01/22/2016 17:59:48 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[rsUsers_UserSunInfo_Upd]

	@SUNUserIP nvarchar(50) ,
	@SUNUserID nvarchar(50) ,
	@SUNUserPass nvarchar(50) ,
	@id uniqueidentifier

as
begin
        Update dbo.rsUsers set SUNUserIP=@SUNUserIP,
        SUNUserID=@SUNUserID,SUNUserPass=@SUNUserPass where ft_id=@id
        
end
GO
/****** Object:  StoredProcedure [dbo].[rsUsers_UserInfo_Upd]    Script Date: 01/22/2016 17:59:48 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[rsUsers_UserInfo_Upd]

@FormUserID nvarchar(50) ,
	@FormUserPassword nvarchar(50) ,
	@WindowsUserID nvarchar(50) ,
	@MachineName nvarchar(50) ,
	@LoginType int ,
	@SUNUserIP nvarchar(50) ,
	@SUNUserID nvarchar(50) ,
	@SUNUserPass nvarchar(50) ,
	@id uniqueidentifier,
	@AddInTabName nvarchar(MAX) 
as
begin
        Update dbo.rsUsers set FormUserPassword=@FormUserPassword,WindowsUserID=@WindowsUserID,MachineName=@MachineName,LoginType=@LoginType,SUNUserIP=@SUNUserIP,
        SUNUserID=@SUNUserID,SUNUserPass=@SUNUserPass , AddInTabName=@AddInTabName where ft_id=@id
        
end
GO
/****** Object:  StoredProcedure [dbo].[rsUsers_Ins]    Script Date: 01/22/2016 17:59:48 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[rsUsers_Ins]

@WindowsUserID nvarchar(50)

as
begin
        insert into dbo.rsUsers(ft_id,FormUserID,FormUserPassword,WindowsUserID,MachineName,LoginType,SUNUserIP,SUNUserID,SUNUserPass,AddInTabName) values(newid(),null,null,@WindowsUserID,null,null,null,null,null,NULL)
end
GO
/****** Object:  StoredProcedure [dbo].[rsUsers_DelByWindowsUserID]    Script Date: 01/22/2016 17:59:48 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[rsUsers_DelByWindowsUserID]

@WindowsUserID varchar(50)

as
begin
        delete dbo.rsUsers where WindowsUserID=@WindowsUserID 
end
GO
/****** Object:  StoredProcedure [dbo].[rsUsers_AddInTabName_Upd]    Script Date: 01/22/2016 17:59:48 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[rsUsers_AddInTabName_Upd]


	@id uniqueidentifier,
	@AddInTabName nvarchar(MAX) 
as
begin
        Update dbo.rsUsers set AddInTabName=@AddInTabName where ft_id=@id
        
end
GO
/****** Object:  StoredProcedure [dbo].[rsGlobalDocumentViews_Ins]    Script Date: 01/22/2016 17:59:47 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[rsGlobalDocumentViews_Ins]
@vd_file nvarchar(50),
@vd_filetype varchar(4),
@vd_folder nvarchar(max),
@vd_macro01 nvarchar(50),
@vd_prefix nchar(10),
@vd_type	nvarchar(50),
@vd_use_ref_as_name bit

as
begin
        insert into dbo.rsGlobalDocumentViews(vd_id,vd_type,vd_prefix,vd_folder,vd_use_ref_as_name,vd_file,vd_filetype,vd_macro01) values(newid(),@vd_type,@vd_prefix,@vd_folder,@vd_use_ref_as_name,@vd_file,@vd_filetype,@vd_macro01)
end
GO
/****** Object:  StoredProcedure [dbo].[rsUsersTemplatesVisible_Upd]    Script Date: 01/22/2016 17:59:48 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
create procedure [dbo].[rsUsersTemplatesVisible_Upd]
@TemplateID nvarchar(MAX),
@UserID nvarchar(MAX),
@Visible nchar(10)

as
begin
        Update dbo.rsUsersTemplatesVisible set Visible=@Visible where TemplateID =@TemplateID and UserID=@UserID
end
GO
/****** Object:  StoredProcedure [dbo].[rsUsersTemplatesVisible_Ins]    Script Date: 01/22/2016 17:59:48 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[rsUsersTemplatesVisible_Ins]
@TemplateID nvarchar(MAX),
@UserID nvarchar(MAX),
@Visible nchar(10)

as
begin
        insert into dbo.rsUsersTemplatesVisible(TemplateID,UserID,Visible) values(@TemplateID,@UserID,@Visible)
end
GO
/****** Object:  StoredProcedure [dbo].[rsUsersTemplatesVisible_Del]    Script Date: 01/22/2016 17:59:48 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
create procedure [dbo].[rsUsersTemplatesVisible_Del]
@UserID nvarchar(max)

as
begin
        delete dbo.rsUsersTemplatesVisible where UserID= @UserID
end
GO
/****** Object:  StoredProcedure [dbo].[rsUserGroup_Ins]    Script Date: 01/22/2016 17:59:48 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[rsUserGroup_Ins]
@UserID nvarchar(150),
@GroupID nvarchar(50)


as
begin
        insert into dbo.rsUserGroup(UserID,GroupID) values(@UserID,@GroupID)
end
GO
/****** Object:  StoredProcedure [dbo].[rsUserGroup_Del]    Script Date: 01/22/2016 17:59:48 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[rsUserGroup_Del]
@GroupID int

as
begin
        delete dbo.rsUserGroup WHERE  GroupID=@GroupID
end
GO
/****** Object:  StoredProcedure [dbo].[rsTemplateCreateXMLTextProfile_Ins]    Script Date: 01/22/2016 17:59:47 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[rsTemplateCreateXMLTextProfile_Ins]
@Field [nvarchar](50) ,
	@FriendlyName [nvarchar](50) ,
	@Visible [bit] ,
	@DefaultValue [nvarchar](max) ,
	@SunComponentName [nvarchar](max) ,
	@SunMethod [nvarchar](max) ,
	@Mandatory [bit] ,
	@Separator [nvarchar](max) ,
	@TextLength [nvarchar](50) ,
	@Prefix [nvarchar](50) ,
	@Suffix [nvarchar](50) ,
	@RemoveCharacters [nvarchar](max) ,
	@TextFileName [nvarchar](max) ,
	@Parent [nvarchar](max)
as
begin
        insert into dbo.rsTemplateCreateXMLTextProfile(Field,FriendlyName,Visible,DefaultValue,SunComponentName,SunMethod,Mandatory,Separator,TextLength,Prefix,Suffix,RemoveCharacters,TextFileName,Parent) 
        
        values(@Field,@FriendlyName,@Visible,@DefaultValue,@SunComponentName,@SunMethod,@Mandatory,@Separator,@TextLength,@Prefix,@Suffix,@RemoveCharacters,@TextFileName,@Parent)
end
GO
/****** Object:  StoredProcedure [dbo].[rsTemplateCreateXMLTextProfile_Del]    Script Date: 01/22/2016 17:59:47 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[rsTemplateCreateXMLTextProfile_Del]
@textfilename nvarchar(max)

as
begin
        delete dbo.rsTemplateCreateXMLTextProfile where TextFileName=@textfilename 
end
GO
/****** Object:  StoredProcedure [dbo].[rsTemplateCreateXMLTextProfile_DelXML]    Script Date: 01/22/2016 17:59:47 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[rsTemplateCreateXMLTextProfile_DelXML]
@SunComponentName NVARCHAR(MAX),
@SunMethod NVARCHAR(MAX)


as
begin
        delete dbo.rsTemplateCreateXMLTextProfile where SunComponentName=@SunComponentName AND  SunMethod=@SunMethod
end
GO
/****** Object:  StoredProcedure [dbo].[rsTemplateXMLTEXTFiles_Ins]    Script Date: 01/22/2016 17:59:48 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[rsTemplateXMLTEXTFiles_Ins]
@FileContent [nvarchar](MAX) ,
	@RelatedName [nvarchar](MAX) ,
	@FileType INT
	
as
begin
        insert into dbo.rsTemplateXMLTEXTFiles(FileContent,RelatedName,FileType) 
        
        values(@FileContent,@RelatedName,@FileType)
end
GO
/****** Object:  StoredProcedure [dbo].[rsTemplateXMLTEXTFiles_Del]    Script Date: 01/22/2016 17:59:48 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[rsTemplateXMLTEXTFiles_Del]
@RelatedName NVARCHAR(MAX),
@FileType INT
as
begin
        delete dbo.rsTemplateXMLTEXTFiles where RelatedName=@RelatedName AND FileType=@FileType
end
GO
/****** Object:  StoredProcedure [dbo].[rsGlobalFields_Ins]    Script Date: 01/22/2016 17:59:47 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[rsGlobalFields_Ins]
@FieldGroup nvarchar(50),
@SunField nvarchar(50),
@UserFriendlyName	nvarchar(50),
@Output nvarchar(50),
@Input nvarchar(50),
@XML_Query nvarchar(max),
@version int
as
begin
        insert into dbo.rsGlobalFields([GUID],FieldGroup,SunField,UserFriendlyName,[Output],Input,XML_Query,[version]) 
        
        values(newid(),@FieldGroup,@SunField,@UserFriendlyName,@Output,@Input,@XML_Query,@version)
end
GO
/****** Object:  StoredProcedure [dbo].[rsGlobalFields_Del]    Script Date: 01/22/2016 17:59:47 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[rsGlobalFields_Del]

as
begin
        delete dbo.rsGlobalFields 
end
GO
/****** Object:  StoredProcedure [dbo].[rsGroupPermissions_Ins]    Script Date: 01/22/2016 17:59:47 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[rsGroupPermissions_Ins]
@GroupID nvarchar(150),
@PermissionID nvarchar(150),
@PermissionGroupName nvarchar(max)

as
begin
        insert into dbo.rsGroupPermissions(GroupID,PermissionID,PermissionGroupName) values(@GroupID,@PermissionID,@PermissionGroupName)
end
GO
/****** Object:  StoredProcedure [dbo].[rsGroupPermissions_Del]    Script Date: 01/22/2016 17:59:47 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[rsGroupPermissions_Del]
	@GroupID int
	
as
begin
       Delete rsGroupPermissions where GroupID=@GroupID
     
end
GO
/****** Object:  StoredProcedure [dbo].[rsPermissions_Upd]    Script Date: 01/22/2016 17:59:47 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[rsPermissions_Upd]
	@id INT,
	@PermissionName nvarchar(150) ,
	@remark nvarchar(3000) 
as
begin
        Update dbo.rsPermissions set PermissionName=@PermissionName,remark=@remark where id=@id
        
end
GO
/****** Object:  StoredProcedure [dbo].[rsPermissions_Ins]    Script Date: 01/22/2016 17:59:47 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[rsPermissions_Ins]
@PermissionName nvarchar(150),
@TemplateID nvarchar(150),
@ActionID nvarchar(150),
@Per_Type int,
@remark nvarchar(3000)

as
begin
        insert into dbo.rsPermissions(PermissionName,TemplateID,ActionID,Per_Type,remark) values(@PermissionName,@TemplateID,@ActionID,@Per_Type,@remark)
end
GO
/****** Object:  StoredProcedure [dbo].[rsPermissions_Del]    Script Date: 01/22/2016 17:59:47 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
create procedure [dbo].[rsPermissions_Del]
	@id int
	
as
begin
       Delete rsPermissions where ID=@id
     
end
GO
/****** Object:  Default [DF_FinTools_Settings_Fields_GUID]    Script Date: 01/22/2016 17:59:47 ******/
ALTER TABLE [dbo].[rsGlobalFields] ADD  CONSTRAINT [DF_FinTools_Settings_Fields_GUID]  DEFAULT (newid()) FOR [GUID]
GO
/****** Object:  Default [DF_FinTools_Settings_ContainerSetting_ft_id]    Script Date: 01/22/2016 17:59:47 ******/
ALTER TABLE [dbo].[rsTemplateContainer] ADD  CONSTRAINT [DF_FinTools_Settings_ContainerSetting_ft_id]  DEFAULT (newid()) FOR [ft_id]
GO
/****** Object:  Default [DF_FinTools_Settings_DrillDown_GUID]    Script Date: 01/22/2016 17:59:47 ******/
ALTER TABLE [dbo].[rsTemplateDrillDown] ADD  CONSTRAINT [DF_FinTools_Settings_DrillDown_GUID]  DEFAULT (newid()) FOR [GUID]
GO
/****** Object:  Default [DF_FinTools_Settings_GenDescFields_GUID]    Script Date: 01/22/2016 17:59:47 ******/
ALTER TABLE [dbo].[rsTemplateGenDescFields] ADD  CONSTRAINT [DF_FinTools_Settings_GenDescFields_GUID]  DEFAULT (newid()) FOR [GUID]
GO
/****** Object:  Default [DF_FinTools_Invoicing_GUID]    Script Date: 01/22/2016 17:59:47 ******/
ALTER TABLE [dbo].[rsTemplateTransactions] ADD  CONSTRAINT [DF_FinTools_Invoicing_GUID]  DEFAULT (newid()) FOR [GUID]
GO
/****** Object:  Default [DF_FinTools_Settings_VisiableRemembers_ft_id]    Script Date: 01/22/2016 17:59:47 ******/
ALTER TABLE [dbo].[rsTemplateVisible] ADD  CONSTRAINT [DF_FinTools_Settings_VisiableRemembers_ft_id]  DEFAULT (newid()) FOR [ft_id]
GO
/****** Object:  Default [DF_FinTools_Users_ft_id]    Script Date: 01/22/2016 17:59:47 ******/
ALTER TABLE [dbo].[rsUsers] ADD  CONSTRAINT [DF_FinTools_Users_ft_id]  DEFAULT (newid()) FOR [ft_id]
GO
