USE [RSDataV2]
GO
/****** Object:  Table [dbo].[FinTools_Users]    Script Date: 09/16/2015 08:08:01 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[FinTools_Users](
	[ft_id] [uniqueidentifier] ROWGUIDCOL  NOT NULL,
	[FormUserID] [nvarchar](50) NULL,
	[FormUserPassword] [nvarchar](50) NULL,
	[WindowsUserID] [nvarchar](50) NULL,
	[MachineName] [nvarchar](50) NULL,
	[LoginType] [int] NULL,
	[SUNUserIP] [nvarchar](50) NULL,
	[SUNUserID] [nvarchar](50) NULL,
	[SUNUserPass] [nvarchar](50) NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[FinTools_UserGroup]    Script Date: 09/16/2015 08:08:01 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[FinTools_UserGroup](
	[ft_id] [int] IDENTITY(1,1) NOT NULL,
	[UserID] [nvarchar](150) NULL,
	[GroupID] [nvarchar](50) NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[FinTools_Templates]    Script Date: 09/16/2015 08:08:01 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[FinTools_Templates](
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
	[TemplatePath] [nvarchar](150) NULL,
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
/****** Object:  Table [dbo].[FinTools_TemplateButtons]    Script Date: 09/16/2015 08:08:01 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[FinTools_TemplateButtons](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[ButtonID] [nvarchar](50) NULL,
	[TemplatePath] [nvarchar](max) NULL,
 CONSTRAINT [PK_FinTools_TemplateButtons] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[FinTools_Settings_VisiableRemembers]    Script Date: 09/16/2015 08:08:01 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[FinTools_Settings_VisiableRemembers](
	[ft_id] [uniqueidentifier] ROWGUIDCOL  NOT NULL,
	[TemplatePath] [nvarchar](500) NULL,
	[OutputPaneVisiable] [nvarchar](5) NULL,
	[PDFPaneVisiable] [nvarchar](5) NULL,
	[UserID] [nvarchar](50) NULL,
 CONSTRAINT [PK_FinTools_Settings_VisiableRemembers] PRIMARY KEY CLUSTERED 
(
	[ft_id] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[FinTools_Settings_TransactionUpdate]    Script Date: 09/16/2015 08:08:01 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[FinTools_Settings_TransactionUpdate](
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
	[ft_filepath] [nvarchar](max) NOT NULL,
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
/****** Object:  Table [dbo].[FinTools_Settings_TemplateSetting]    Script Date: 09/16/2015 08:08:01 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[FinTools_Settings_TemplateSetting](
	[ft_id] [uniqueidentifier] ROWGUIDCOL  NOT NULL,
	[CriteriaName] [nvarchar](50) NULL,
	[CellReference] [nvarchar](50) NULL,
	[TemplatePath] [nvarchar](500) NULL,
	[orderNum] [int] NULL,
	[OpenTransUponSave] [bit] NULL,
 CONSTRAINT [PK_FinTools_Settings_ReportSetting] PRIMARY KEY CLUSTERED 
(
	[ft_id] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[FinTools_Settings_SequenceNumbering]    Script Date: 09/16/2015 08:08:01 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[FinTools_Settings_SequenceNumbering](
	[ft_id] [uniqueidentifier] ROWGUIDCOL  NOT NULL,
	[UseSequenceNum] [int] NULL,
	[SequencePrefix] [nvarchar](50) NULL,
	[PostToField] [nvarchar](50) NULL,
	[PopulateCell] [nvarchar](50) NULL,
	[TemplatePath] [nvarchar](500) NULL,
 CONSTRAINT [PK_FinTools_Settings_SequenceNumbering] PRIMARY KEY CLUSTERED 
(
	[ft_id] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[FinTools_Settings_Help]    Script Date: 09/16/2015 08:08:01 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[FinTools_Settings_Help](
	[ft_id] [int] IDENTITY(1,1) NOT NULL,
	[ft_filepath] [nvarchar](max) NULL,
	[helpFileData] [image] NULL,
	[helpFileType] [nvarchar](50) NULL,
 CONSTRAINT [PK_FinTools_Settings_Help] PRIMARY KEY CLUSTERED 
(
	[ft_id] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[FinTools_Settings_GenDescFields]    Script Date: 09/16/2015 08:08:01 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[FinTools_Settings_GenDescFields](
	[GUID] [uniqueidentifier] ROWGUIDCOL  NOT NULL,
	[FieldGroup] [nvarchar](50) NULL,
	[SunField] [nvarchar](50) NULL,
	[UserFriendlyName] [nvarchar](50) NULL,
	[Output] [nvarchar](50) NULL,
	[Input] [nvarchar](50) NULL,
	[XML_Query] [nvarchar](max) NULL,
	[TemplatePath] [nvarchar](max) NULL,
	[version] [int] NULL,
 CONSTRAINT [PK_FinTools_Settings_GenDescFields] PRIMARY KEY CLUSTERED 
(
	[GUID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[FinTools_Settings_Fields]    Script Date: 09/16/2015 08:08:01 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[FinTools_Settings_Fields](
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
/****** Object:  Table [dbo].[FinTools_Settings_DrillDown]    Script Date: 09/16/2015 08:08:01 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[FinTools_Settings_DrillDown](
	[GUID] [uniqueidentifier] ROWGUIDCOL  NOT NULL,
	[SunField] [nvarchar](50) NULL,
	[InputStatus] [nvarchar](50) NULL,
	[CellName] [nvarchar](50) NULL,
	[OutputStatus] [nvarchar](50) NULL,
	[TemplatePath] [nvarchar](500) NULL,
	[order] [int] NULL,
 CONSTRAINT [PK_FinTools_Settings_DrillDown] PRIMARY KEY CLUSTERED 
(
	[GUID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[FinTools_Settings_CriteriaPaneRemembers]    Script Date: 09/16/2015 08:08:01 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[FinTools_Settings_CriteriaPaneRemembers](
	[ft_id] [uniqueidentifier] ROWGUIDCOL  NOT NULL,
	[CriteriaPaneVisiable] [nvarchar](5) NULL,
	[UserID] [nvarchar](50) NULL,
 CONSTRAINT [PK_FinTools_Settings_CriteriaPaneRemembers] PRIMARY KEY CLUSTERED 
(
	[ft_id] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[FinTools_Settings_ContainerSetting]    Script Date: 09/16/2015 08:08:01 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[FinTools_Settings_ContainerSetting](
	[ft_id] [uniqueidentifier] ROWGUIDCOL  NOT NULL,
	[ft_filepath] [nvarchar](max) NOT NULL,
	[ft_relatefilepath] [nvarchar](max) NOT NULL,
	[column] [nvarchar](5) NULL,
	[FromDB] [bit] NULL,
 CONSTRAINT [PK_FinTools_Settings_ContainerSetting] PRIMARY KEY CLUSTERED 
(
	[ft_id] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[FinTools_Settings_AllocationMarkerUpdate]    Script Date: 09/16/2015 08:08:01 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[FinTools_Settings_AllocationMarkerUpdate](
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
	[ft_filepath] [nvarchar](max) NOT NULL,
	[LineIndicator] [nvarchar](50) NOT NULL,
	[StartinginCell] [nvarchar](50) NOT NULL,
	[inputFields] [nvarchar](max) NOT NULL,
	[updateFields] [nvarchar](max) NOT NULL,
 CONSTRAINT [PK_FinTools_Settings_AllocationMarkerUpdate] PRIMARY KEY CLUSTERED 
(
	[GUID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[FinTools_Settings]    Script Date: 09/16/2015 08:08:01 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[FinTools_Settings](
	[ft_id] [uniqueidentifier] NOT NULL,
	[ft_folder] [nvarchar](max) NOT NULL,
 CONSTRAINT [PK_FinTools_Settings] PRIMARY KEY CLUSTERED 
(
	[ft_id] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[FinTools_ProcessesMacros]    Script Date: 09/16/2015 08:08:01 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[FinTools_ProcessesMacros](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[ProcessMacroName] [nvarchar](50) NULL,
	[Type] [nvarchar](50) NULL,
 CONSTRAINT [PK_FinTools_FunctionMacros] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[FinTools_Processes_Approval]    Script Date: 09/16/2015 08:08:01 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[FinTools_Processes_Approval](
	[GUID] [uniqueidentifier] NOT NULL,
	[ProcessesID] [nvarchar](50) NULL,
	[Criteria1] [nvarchar](50) NULL,
	[Criteria2] [nvarchar](50) NULL,
	[Value1] [nvarchar](50) NULL,
	[Value2] [nvarchar](50) NULL,
	[ApproverID] [nvarchar](50) NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[FinTools_Permissions]    Script Date: 09/16/2015 08:08:01 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[FinTools_Permissions](
	[ft_id] [int] IDENTITY(1,1) NOT NULL,
	[PermissionName] [nvarchar](150) NULL,
	[Per_ClassName] [nvarchar](150) NULL,
	[Per_ControlID] [nvarchar](150) NULL,
	[Per_Type] [int] NULL,
	[remark] [nvarchar](3000) NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[FinTools_OutPutProfile]    Script Date: 09/16/2015 08:08:01 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[FinTools_OutPutProfile](
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
	[ft_filepath] [nvarchar](max) NOT NULL,
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
/****** Object:  Table [dbo].[FinTools_OutPutCreateTextFile]    Script Date: 09/16/2015 08:08:01 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[FinTools_OutPutCreateTextFile](
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
	[ft_filepath] [nvarchar](max) NOT NULL,
	[IncludeHeaderRow] [bit] NULL,
	[SavePath] [nvarchar](max) NULL,
	[SaveName] [nvarchar](150) NULL,
 CONSTRAINT [PK_FinTools_OutPutCreateTextFile] PRIMARY KEY CLUSTERED 
(
	[ft_id] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[FinTools_OutPutConsolidation]    Script Date: 09/16/2015 08:08:01 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[FinTools_OutPutConsolidation](
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
	[ft_filepath] [nvarchar](max) NOT NULL,
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
/****** Object:  Table [dbo].[FinTools_Groups]    Script Date: 09/16/2015 08:08:01 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[FinTools_Groups](
	[ft_id] [int] IDENTITY(1,1) NOT NULL,
	[GroupName] [nvarchar](150) NULL,
	[GroupDisable] [bit] NULL,
	[Remark] [nvarchar](4000) NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[FinTools_GroupPermissions]    Script Date: 09/16/2015 08:08:01 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[FinTools_GroupPermissions](
	[ft_id] [int] IDENTITY(1,1) NOT NULL,
	[GroupID] [nvarchar](150) NULL,
	[PermissionID] [nvarchar](150) NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[FinTools_CSL]    Script Date: 09/16/2015 08:08:01 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[FinTools_CSL](
	[ft_id] [uniqueidentifier] ROWGUIDCOL  NOT NULL,
	[CSLTableName] [nvarchar](50) NULL,
	[CSLColumnName] [nvarchar](50) NULL,
	[CSLFilter] [nvarchar](50) NULL,
	[CSLOperator] [nvarchar](50) NULL,
	[CSLColumnName2] [nvarchar](50) NULL,
	[CSLFilter2] [nvarchar](50) NULL,
	[CSLOperator2] [nvarchar](50) NULL,
	[CSLColumnName3] [nvarchar](50) NULL,
	[CSLFilter3] [nvarchar](50) NULL,
	[CSLOperator3] [nvarchar](50) NULL,
	[TemplatePath] [nvarchar](150) NULL,
	[SheetName] [nvarchar](150) NULL,
	[CellName] [nvarchar](50) NULL,
	[OutPut] [nvarchar](50) NULL,
	[TemplateName] [nvarchar](50) NULL,
	[test] [money] NULL,
	[test2] [nvarchar](50) NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[FinTools_Buttons]    Script Date: 09/16/2015 08:08:01 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[FinTools_Buttons](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[ButtonName] [nvarchar](50) NULL,
	[ButtonText] [nvarchar](50) NULL,
	[ButtonIcon] [nvarchar](150) NULL,
	[ButtonGroup] [nvarchar](50) NULL,
	[ButtonSize] [nvarchar](50) NULL,
	[ButtonOrder] [int] NULL,
	[GroupOrder] [int] NULL,
	[StopOnError] [bit] NULL,
 CONSTRAINT [PK_FinTools_Buttons] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[FinTools_ButtonProcessesMacros]    Script Date: 09/16/2015 08:08:01 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[FinTools_ButtonProcessesMacros](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[ButtonID] [nvarchar](50) NULL,
	[ProcessMacroID] [nvarchar](50) NULL,
	[ExecOrder] [int] NULL,
 CONSTRAINT [PK_FinTools_ButtonFunctionMacros] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[VIEW_DOC]    Script Date: 09/16/2015 08:08:01 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[VIEW_DOC](
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
/****** Object:  View [dbo].[View_TemplateButtons]    Script Date: 09/16/2015 08:08:01 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE VIEW [dbo].[View_TemplateButtons]
AS
SELECT     dbo.FinTools_Buttons.ID AS Expr1, dbo.FinTools_Buttons.ButtonName, dbo.FinTools_Buttons.ButtonText, dbo.FinTools_Buttons.ButtonIcon, 
                      dbo.FinTools_Buttons.ButtonGroup, dbo.FinTools_Buttons.ButtonSize, dbo.FinTools_Buttons.ButtonOrder, dbo.FinTools_Buttons.GroupOrder, 
                      dbo.FinTools_Buttons.StopOnError, dbo.FinTools_TemplateButtons.ID, dbo.FinTools_TemplateButtons.ButtonID, dbo.FinTools_TemplateButtons.TemplatePath
FROM         dbo.FinTools_Buttons INNER JOIN
                      dbo.FinTools_TemplateButtons ON dbo.FinTools_Buttons.ID = dbo.FinTools_TemplateButtons.ButtonID
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[27] 4[20] 2[34] 3) )"
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
         Begin Table = "FinTools_Buttons"
            Begin Extent = 
               Top = 6
               Left = 38
               Bottom = 236
               Right = 198
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "FinTools_TemplateButtons"
            Begin Extent = 
               Top = 6
               Left = 236
               Bottom = 190
               Right = 396
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
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'View_TemplateButtons'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=1 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'View_TemplateButtons'
GO
/****** Object:  StoredProcedure [dbo].[spVIEWDOC_Ins]    Script Date: 09/16/2015 08:08:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[spVIEWDOC_Ins]


@vd_file nvarchar(50),
@vd_filetype varchar(4),
@vd_folder nvarchar(max),
@vd_macro01 nvarchar(50),
@vd_prefix nchar(10),
@vd_type	nvarchar(50),
@vd_use_ref_as_name bit
--@sname varchar(10),
--@sex char(2)
as
begin
        insert into dbo.VIEW_DOC(vd_id,vd_type,vd_prefix,vd_folder,vd_use_ref_as_name,vd_file,vd_filetype,vd_macro01) values(newid(),@vd_type,@vd_prefix,@vd_folder,@vd_use_ref_as_name,@vd_file,@vd_filetype,@vd_macro01)
end
GO
/****** Object:  StoredProcedure [dbo].[FT_Users_UserSunInfo_Upd]    Script Date: 09/16/2015 08:08:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[FT_Users_UserSunInfo_Upd]

	@SUNUserIP nvarchar(50) ,
	@SUNUserID nvarchar(50) ,
	@SUNUserPass nvarchar(50) ,
	@id uniqueidentifier
--@sname varchar(10),
--@sex char(2)
as
begin
        Update dbo.FinTools_Users set SUNUserIP=@SUNUserIP,
        SUNUserID=@SUNUserID,SUNUserPass=@SUNUserPass where ft_id=@id
        
end
GO
/****** Object:  StoredProcedure [dbo].[FT_Users_UserInfo_Upd]    Script Date: 09/16/2015 08:08:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[FT_Users_UserInfo_Upd]

@FormUserID nvarchar(50) ,
	@FormUserPassword nvarchar(50) ,
	@WindowsUserID nvarchar(50) ,
	@MachineName nvarchar(50) ,
	@LoginType int ,
	@SUNUserIP nvarchar(50) ,
	@SUNUserID nvarchar(50) ,
	@SUNUserPass nvarchar(50) ,
	@id uniqueidentifier
--@sname varchar(10),
--@sex char(2)
as
begin
        Update dbo.FinTools_Users set FormUserPassword=@FormUserPassword,WindowsUserID=@WindowsUserID,MachineName=@MachineName,LoginType=@LoginType,SUNUserIP=@SUNUserIP,
        SUNUserID=@SUNUserID,SUNUserPass=@SUNUserPass where ft_id=@id
        
end
GO
/****** Object:  StoredProcedure [dbo].[FT_Users_Ins]    Script Date: 09/16/2015 08:08:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[FT_Users_Ins]

@WindowsUserID nvarchar(50)
--@sname varchar(10),
--@sex char(2)
as
begin
        insert into dbo.FinTools_Users(ft_id,FormUserID,FormUserPassword,WindowsUserID,MachineName,LoginType,SUNUserIP,SUNUserID,SUNUserPass) values(newid(),null,null,@WindowsUserID,null,null,null,null,null)
end
GO
/****** Object:  StoredProcedure [dbo].[FT_Users_DelByWindowsUserID]    Script Date: 09/16/2015 08:08:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[FT_Users_DelByWindowsUserID]

@WindowsUserID varchar(50)

as
begin
        delete dbo.FinTools_Users where WindowsUserID=@WindowsUserID 
end
GO
/****** Object:  StoredProcedure [dbo].[FT_TransactionUpdate_Ins]    Script Date: 09/16/2015 08:08:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[FT_TransactionUpdate_Ins]
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
@ft_filepath nvarchar(max),
@LineIndicator nvarchar(50) ,
@StartinginCell nvarchar(50),
@inputFields nvarchar(max),
@updateFields nvarchar(max)


as
begin
        insert into dbo.FinTools_Settings_TransactionUpdate
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
        TransAmount,Currency,BaseAmount,[2ndBase],[4thAmount],ft_filepath,LineIndicator,StartinginCell,inputFields,updateFields) 
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
        @2ndBase,@4thAmount,@ft_filepath,@LineIndicator,@StartinginCell,@inputFields,
@updateFields )
end
GO
/****** Object:  StoredProcedure [dbo].[FT_TransactionUpdate_Del]    Script Date: 09/16/2015 08:08:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[FT_TransactionUpdate_Del]
@ft_filepath nvarchar(500)
--@sname varchar(10),
--@sex char(2)
as
begin
        delete dbo.FinTools_Settings_TransactionUpdate where ft_filepath=@ft_filepath 
end
GO
/****** Object:  StoredProcedure [dbo].[FT_Templates_UpdateCriterias]    Script Date: 09/16/2015 08:08:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
create procedure [dbo].[FT_Templates_UpdateCriterias]
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
	@TemplatePath nvarchar(150) 


as
begin

        update  FinTools_Templates
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
	where TemplatePath=@TemplatePath
	  
	
end
GO
/****** Object:  StoredProcedure [dbo].[FT_Templates_Ins]    Script Date: 09/16/2015 08:08:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[FT_Templates_Ins]
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
	@TemplatePath nvarchar(150) ,
	@maxNum int ,
	@TransactionName  nvarchar(50),
	@Prefix nvarchar(50),
	@SunJournalNumber nvarchar(100)
--@sname varchar(10),
--@sex char(2)
as
begin
        insert into FinTools_Templates
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
	TemplatePath  ,
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
	@TemplatePath  ,
	@maxNum,@TransactionName ,@Prefix ,@SunJournalNumber)
end
GO
/****** Object:  StoredProcedure [dbo].[FT_Templates_Del]    Script Date: 09/16/2015 08:08:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[FT_Templates_Del]
	@TemplatePath nvarchar(500) ,
	@TransactionName  nvarchar(500)
as
begin
       Delete FinTools_Templates where TemplatePath=@TemplatePath and  TransactionName =@TransactionName 
     
end
GO
/****** Object:  StoredProcedure [dbo].[FT_TemplateButtons_Ins]    Script Date: 09/16/2015 08:08:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
create procedure [dbo].[FT_TemplateButtons_Ins]

@ButtonID nvarchar(50),
@TemplatePath nvarchar(max)
--@sname varchar(10),
--@sex char(2)
as
begin
        insert into dbo.FinTools_TemplateButtons(ButtonID,TemplatePath) values(@ButtonID,@TemplatePath)
end
GO
/****** Object:  StoredProcedure [dbo].[FT_TemplateButtons_Del]    Script Date: 09/16/2015 08:08:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
create procedure [dbo].[FT_TemplateButtons_Del]
	@Buttonid nvarchar(50)
	
as
begin
       Delete FinTools_TemplateButtons where ButtonID=@Buttonid
     
end
GO
/****** Object:  StoredProcedure [dbo].[FT_Settings_VisiableRemembers_Ins]    Script Date: 09/16/2015 08:08:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[FT_Settings_VisiableRemembers_Ins]
@TemplatePath  nvarchar(500),
@OutputPaneVisiable nvarchar(5),
@PDFPaneVisiable nvarchar(5),
@UserID nvarchar(50)
--@sname varchar(10),
--@sex char(2)
as
begin
        insert into dbo.FinTools_Settings_VisiableRemembers(ft_id,TemplatePath,OutputPaneVisiable,PDFPaneVisiable,UserID) values(newid(),@TemplatePath,@OutputPaneVisiable,@PDFPaneVisiable,@UserID)
end
GO
/****** Object:  StoredProcedure [dbo].[FT_Settings_VisiableRemembers_Del]    Script Date: 09/16/2015 08:08:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[FT_Settings_VisiableRemembers_Del]

@TemplatePath nvarchar(500),
@UserID nvarchar(50)
--@sname varchar(10),
--@sex char(2)
as
begin
        delete dbo.FinTools_Settings_VisiableRemembers where TemplatePath=@TemplatePath and UserID=@UserID
end
GO
/****** Object:  StoredProcedure [dbo].[FT_Settings_TemplateSetting_Ins]    Script Date: 09/16/2015 08:08:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[FT_Settings_TemplateSetting_Ins]

@CriteriaName nvarchar(50) ,
@CellReference nvarchar(50) ,

	@TemplatePath nvarchar(500) ,
@orderNum int,
@OpenTransUponSave bit
	
--@sname varchar(10),
--@sex char(2)
as
begin
        insert into dbo.FinTools_Settings_TemplateSetting(ft_id,CriteriaName,CellReference,TemplatePath, orderNum,OpenTransUponSave) values(newid(),@CriteriaName,@CellReference,@TemplatePath,@orderNum,@OpenTransUponSave)
end
GO
/****** Object:  StoredProcedure [dbo].[FT_Settings_TemplateSetting_Del]    Script Date: 09/16/2015 08:08:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[FT_Settings_TemplateSetting_Del]


@path varchar(500)
--@sname varchar(10),
--@sex char(2)
as
begin
        delete dbo.FinTools_Settings_TemplateSetting where TemplatePath=@path
end
GO
/****** Object:  StoredProcedure [dbo].[FT_Settings_SequenceNumbering_Ins]    Script Date: 09/16/2015 08:08:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
cREATE procedure [dbo].[FT_Settings_SequenceNumbering_Ins]

@UseSequenceNum int ,
	@SequencePrefix nvarchar(50) ,
	@PostToField nvarchar(50) ,
	@PopulateCell nvarchar(50) ,

	@TemplatePath nvarchar(500) 

	
--@sname varchar(10),
--@sex char(2)
as
begin
        insert into dbo.FinTools_Settings_SequenceNumbering(ft_id,UseSequenceNum,SequencePrefix,PostToField,PopulateCell,TemplatePath) values(newid(),@UseSequenceNum ,@SequencePrefix ,@PostToField ,@PopulateCell ,@TemplatePath)
end
GO
/****** Object:  StoredProcedure [dbo].[FT_Settings_SequenceNumbering_Del]    Script Date: 09/16/2015 08:08:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[FT_Settings_SequenceNumbering_Del]


@path varchar(500)
--@sname varchar(10),
--@sex char(2)
as
begin
        delete dbo.FinTools_Settings_SequenceNumbering where TemplatePath=@path
end
GO
/****** Object:  StoredProcedure [dbo].[FT_Settings_Help_Ins]    Script Date: 09/16/2015 08:08:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
create procedure [dbo].[FT_Settings_Help_Ins]


	@ft_filepath nvarchar(max) ,
	@helpFileData image ,
	@helpFileType nvarchar(50) 
	
as
begin
        insert into dbo.FinTools_Settings_Help(ft_filepath,helpFileData,helpFileType) values(@ft_filepath,@helpFileData,@helpFileType)
end
GO
/****** Object:  StoredProcedure [dbo].[FT_Settings_Help_Del]    Script Date: 09/16/2015 08:08:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[FT_Settings_Help_Del]
	@ft_filepath nvarchar(max) 
	
as
begin
       Delete FinTools_Settings_Help where ft_filepath=@ft_filepath
     
end
GO
/****** Object:  StoredProcedure [dbo].[FT_Settings_GenDescFields_Ins]    Script Date: 09/16/2015 08:08:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[FT_Settings_GenDescFields_Ins]

@FieldGroup nvarchar(50),
@SunField nvarchar(50),
@UserFriendlyName	nvarchar(50),
@Output nvarchar(50),
@Input nvarchar(50),
@XML_Query nvarchar(max),
@TemplatePath nvarchar(max),
@version int

--@sname varchar(10),
--@sex char(2)
as
begin
        insert into dbo.FinTools_Settings_GenDescFields([GUID],FieldGroup,SunField,UserFriendlyName,[Output],Input,XML_Query,TemplatePath,[version]) 
        
        values(newid(),@FieldGroup,@SunField,@UserFriendlyName,@Output,@Input,@XML_Query,@TemplatePath,@version)
end
GO
/****** Object:  StoredProcedure [dbo].[FT_Settings_GenDescFields_Del]    Script Date: 09/16/2015 08:08:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[FT_Settings_GenDescFields_Del]

@TemplatePath nvarchar(max)
--@sname varchar(10),
--@sex char(2)
as
begin
        delete dbo.FinTools_Settings_GenDescFields where TemplatePath=@TemplatePath
end
GO
/****** Object:  StoredProcedure [dbo].[FT_Settings_Fields_Ins]    Script Date: 09/16/2015 08:08:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[FT_Settings_Fields_Ins]

@FieldGroup nvarchar(50),
@SunField nvarchar(50),
@UserFriendlyName	nvarchar(50),
@Output nvarchar(50),
@Input nvarchar(50),
@XML_Query nvarchar(max),
@version int

--@sname varchar(10),
--@sex char(2)
as
begin
        insert into dbo.FinTools_Settings_Fields([GUID],FieldGroup,SunField,UserFriendlyName,[Output],Input,XML_Query,[version]) 
        
        values(newid(),@FieldGroup,@SunField,@UserFriendlyName,@Output,@Input,@XML_Query,@version)
end
GO
/****** Object:  StoredProcedure [dbo].[FT_Settings_Fields_Del]    Script Date: 09/16/2015 08:08:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[FT_Settings_Fields_Del]


--@sname varchar(10),
--@sex char(2)
as
begin
        delete dbo.FinTools_Settings_Fields 
end
GO
/****** Object:  StoredProcedure [dbo].[FT_Settings_CriteriaPaneRemembers_Ins]    Script Date: 09/16/2015 08:08:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[FT_Settings_CriteriaPaneRemembers_Ins]

@CriteriaPaneVisiable nvarchar(5),
@UserID nvarchar(50)
--@sname varchar(10),
--@sex char(2)
as
begin
        insert into dbo.FinTools_Settings_CriteriaPaneRemembers(ft_id,CriteriaPaneVisiable,UserID) values(newid(),@CriteriaPaneVisiable,@UserID)
end
GO
/****** Object:  StoredProcedure [dbo].[FT_Settings_CriteriaPaneRemembers_Del]    Script Date: 09/16/2015 08:08:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[FT_Settings_CriteriaPaneRemembers_Del]


@UserID nvarchar(50)
--@sname varchar(10),
--@sex char(2)
as
begin
        delete dbo.FinTools_Settings_CriteriaPaneRemembers where UserID=@UserID
end
GO
/****** Object:  StoredProcedure [dbo].[FT_Settings_Container_Ins]    Script Date: 09/16/2015 08:08:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[FT_Settings_Container_Ins]
@ft_filepath  nvarchar(max),
@ft_relatefilepath nvarchar(max),
@column nvarchar(5),
@FromDB bit
--@sname varchar(10),
--@sex char(2)
as
begin
        insert into dbo.FinTools_Settings_ContainerSetting(ft_id,ft_filepath,ft_relatefilepath,[column],FromDB) values(newid(),@ft_filepath,@ft_relatefilepath,@column,@FromDB)
end
GO
/****** Object:  StoredProcedure [dbo].[FT_Settings_Container_Del]    Script Date: 09/16/2015 08:08:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[FT_Settings_Container_Del]

@ft_filepath nvarchar(max)
--@sname varchar(10),
--@sex char(2)
as
begin
        delete dbo.FinTools_Settings_ContainerSetting where ft_filepath=@ft_filepath 
end
GO
/****** Object:  StoredProcedure [dbo].[FT_ProcessesMacros_Ins]    Script Date: 09/16/2015 08:08:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
create procedure [dbo].[FT_ProcessesMacros_Ins]

@ProcessMacroName nvarchar(50),
@type nvarchar(50)
--@sname varchar(10),
--@sex char(2)
as
begin
        insert into dbo.FinTools_ProcessesMacros(ProcessMacroName,[Type]) values(@ProcessMacroName,@type)
end
GO
/****** Object:  StoredProcedure [dbo].[FT_OutPutProfile_Ins]    Script Date: 09/16/2015 08:08:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[FT_OutPutProfile_Ins]
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
@ft_filepath nvarchar(max),
@LineIndicator nvarchar(50) ,
@StartinginCell nvarchar(50),
@BalanceBy nvarchar(50),
@PopWithJNNumber nvarchar(50)
--@sname varchar(10),
--@sex char(2)
as
begin
        insert into dbo.FinTools_OutPutProfile
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
        TransAmount,Currency,BaseAmount,[2ndBase],[4thAmount],ft_filepath,LineIndicator,StartinginCell,BalanceBy,PopWithJNNumber) 
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
        @2ndBase,@4thAmount,@ft_filepath,@LineIndicator,@StartinginCell,@BalanceBy,@PopWithJNNumber)
end
GO
/****** Object:  StoredProcedure [dbo].[FT_OutPutProfile_Del]    Script Date: 09/16/2015 08:08:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[FT_OutPutProfile_Del]
@ft_filepath nvarchar(max)
--@sname varchar(10),
--@sex char(2)
as
begin
        delete dbo.FinTools_OutPutProfile where ft_filepath=@ft_filepath 
end
GO
/****** Object:  StoredProcedure [dbo].[FT_OutPutCreateTextFile_Ins]    Script Date: 09/16/2015 08:08:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[FT_OutPutCreateTextFile_Ins]
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
@ft_filepath nvarchar(max),
@IncludeHeaderRow bit,
@SavePath nvarchar(max),
@SaveName nvarchar(150)

as
begin
        insert into dbo.FinTools_OutPutCreateTextFile
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
HeaderTextes ,StartinginCell ,LineIndicator ,ft_filepath ,IncludeHeaderRow ,SavePath,SaveName) 
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
@HeaderTextes ,@StartinginCell ,@LineIndicator ,@ft_filepath ,@IncludeHeaderRow,@SavePath,@SaveName ) 
end
GO
/****** Object:  StoredProcedure [dbo].[FT_OutPutCreateTextFile_Del]    Script Date: 09/16/2015 08:08:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
create procedure [dbo].[FT_OutPutCreateTextFile_Del]
@ft_filepath nvarchar(max)
--@sname varchar(10),
--@sex char(2)
as
begin
        delete dbo.FinTools_OutPutCreateTextFile where ft_filepath=@ft_filepath 
end
GO
/****** Object:  StoredProcedure [dbo].[FT_OutPutConsolidation_Ins]    Script Date: 09/16/2015 08:08:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[FT_OutPutConsolidation_Ins]
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
@ft_filepath nvarchar(max),
@LineIndicator nvarchar(50) ,
@StartinginCell nvarchar(50),
@ConsolidateBy1 nvarchar(50) ,
@ConsolidateBy2 nvarchar(50),
@ConsolidateBy3 nvarchar(50) ,
@ConsolidateBy4 nvarchar(50),
@BalanceBy nvarchar(50),
@PopWithJNNumber nvarchar(50)
--@sname varchar(10),
--@sex char(2)
as
begin
        insert into dbo.FinTools_OutPutConsolidation
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
        TransAmount,Currency,BaseAmount,[2ndBase],[4thAmount],ft_filepath,LineIndicator,StartinginCell,ConsolidateBy1,ConsolidateBy2,ConsolidateBy3,ConsolidateBy4,BalanceBy,PopWithJNNumber) 
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
        @2ndBase,@4thAmount,@ft_filepath,@LineIndicator,@StartinginCell,@ConsolidateBy1,@ConsolidateBy2,@ConsolidateBy3,@ConsolidateBy4,@BalanceBy,@PopWithJNNumber)
end
GO
/****** Object:  StoredProcedure [dbo].[FT_OutPutConsolidation_Del]    Script Date: 09/16/2015 08:08:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[FT_OutPutConsolidation_Del]
@ft_filepath nvarchar(max)
--@sname varchar(10),
--@sex char(2)
as
begin
        delete dbo.FinTools_OutPutConsolidation where ft_filepath=@ft_filepath 
end
GO
/****** Object:  StoredProcedure [dbo].[FT_DrillDown_Ins]    Script Date: 09/16/2015 08:08:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
create procedure [dbo].[FT_DrillDown_Ins]


@SunField nvarchar(50) ,
	@InputStatus nvarchar(50) ,
	@CellName nvarchar(50) ,
	@OutputStatus nvarchar(50) ,
	@TemplatePath nvarchar(500) ,
	@order int
	


as
begin
        insert into dbo.FinTools_Settings_DrillDown
        (SunField,InputStatus,CellName,OutputStatus,TemplatePath,[order]) 
        values(@SunField,@InputStatus,@CellName,@OutputStatus,@TemplatePath,@order)
end
GO
/****** Object:  StoredProcedure [dbo].[FT_DrillDown_Del]    Script Date: 09/16/2015 08:08:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[FT_DrillDown_Del]
@ft_filepath nvarchar(500)
--@sname varchar(10),
--@sex char(2)
as
begin
        delete dbo.FinTools_Settings_DrillDown where TemplatePath=@ft_filepath 
end
GO
/****** Object:  StoredProcedure [dbo].[FT_CSL_Ins]    Script Date: 09/16/2015 08:08:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
create procedure [dbo].[FT_CSL_Ins]


           @CSLTableName nvarchar(50),
           @CSLColumnName nvarchar(50),
           @CSLFilter nvarchar(50),
           @CSLOperator nvarchar(50),
           @CSLColumnName2 nvarchar(50),
           @CSLFilter2 nvarchar(50),
           @CSLOperator2 nvarchar(50),
           @CSLColumnName3 nvarchar(50),
           @CSLFilter3 nvarchar(50),
           @CSLOperator3 nvarchar(50),
           @TemplatePath nvarchar(150),
           @SheetName nvarchar(150),
           @CellName nvarchar(50),
           @OutPut nvarchar(50),
           @TemplateName nvarchar(50)
as
begin
       INSERT INTO [RSDataV2].[dbo].[FinTools_CSL]
           (
           [CSLTableName]
           ,[CSLColumnName]
           ,[CSLFilter]
           ,[CSLOperator]
           ,[CSLColumnName2]
           ,[CSLFilter2]
           ,[CSLOperator2]
           ,[CSLColumnName3]
           ,[CSLFilter3]
           ,[CSLOperator3]
           ,[TemplatePath]
           ,[SheetName]
           ,[CellName]
           ,[OutPut]
           ,[TemplateName])
     VALUES
           (@CSLTableName
           ,@CSLColumnName
           ,@CSLFilter
           ,@CSLOperator
           ,@CSLColumnName2
           ,@CSLFilter2
           ,@CSLOperator2
           ,@CSLColumnName3
           ,@CSLFilter3
           ,@CSLOperator3
           ,@TemplatePath
           ,@SheetName
           ,@CellName
           ,@OutPut
           ,@TemplateName)


end
GO
/****** Object:  StoredProcedure [dbo].[FT_CSL_Delete]    Script Date: 09/16/2015 08:08:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[FT_CSL_Delete]
@ft_filepath nvarchar(150),
@ft_sheet nvarchar(50),
@ft_cell nvarchar(50)

as
begin
        delete dbo.FinTools_CSL where TemplatePath=@ft_filepath and SheetName=@ft_sheet and CellName=@ft_cell
end
GO
/****** Object:  StoredProcedure [dbo].[FT_Buttons_UpdGroupName]    Script Date: 09/16/2015 08:08:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[FT_Buttons_UpdGroupName]

@ButtonGroup nvarchar(50),
@GroupOrder int,
@id int

as
begin
        Update dbo.FinTools_Buttons set ButtonGroup=@ButtonGroup,GroupOrder=@GroupOrder where ID=@id
end
GO
/****** Object:  StoredProcedure [dbo].[FT_Buttons_UpdButtonOrder]    Script Date: 09/16/2015 08:08:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
create procedure [dbo].[FT_Buttons_UpdButtonOrder]

@ButtonOrder int,
@id int

as
begin
        Update dbo.FinTools_Buttons set ButtonOrder=@ButtonOrder where ID=@id
end
GO
/****** Object:  StoredProcedure [dbo].[FT_Buttons_Ins]    Script Date: 09/16/2015 08:08:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[FT_Buttons_Ins]

@Name nvarchar(50),
@Text nvarchar(50),
@ButtonIcon nvarchar(150),
@ButtonGroup nvarchar(50),
@ButtonSize nvarchar(50),
@ButtonOrder int,
@GroupOrder int,
@StopOnError bit
--@sname varchar(10),
--@sex char(2)
as
begin
        insert into dbo.FinTools_Buttons(ButtonName,ButtonText,ButtonIcon,ButtonGroup,ButtonSize,ButtonOrder,GroupOrder,StopOnError) values(@Name,@Text,@ButtonIcon,@ButtonGroup,@ButtonSize,@ButtonOrder,@GroupOrder,@StopOnError)
end
GO
/****** Object:  StoredProcedure [dbo].[FT_Buttons_Del]    Script Date: 09/16/2015 08:08:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
create procedure [dbo].[FT_Buttons_Del]
	@id int
	
as
begin
       Delete FinTools_Buttons where ID=@id
     
end
GO
/****** Object:  StoredProcedure [dbo].[FT_ButtonProcessesMacro_Ins]    Script Date: 09/16/2015 08:08:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
create procedure [dbo].[FT_ButtonProcessesMacro_Ins]

@ButtonID nvarchar(50),
@ProcessMacroID nvarchar(50),
@ExecOrder int
--@sname varchar(10),
--@sex char(2)
as
begin
        insert into dbo.FinTools_ButtonProcessesMacros(ButtonID,ProcessMacroID,ExecOrder) values(@ButtonID,@ProcessMacroID,@ExecOrder)
end
GO
/****** Object:  StoredProcedure [dbo].[FT_ButtonProcessesMacro_Del]    Script Date: 09/16/2015 08:08:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
create procedure [dbo].[FT_ButtonProcessesMacro_Del]
	@Buttonid nvarchar(50)
	
as
begin
       Delete FinTools_ButtonProcessesMacros where ButtonID=@Buttonid
     
end
GO
/****** Object:  StoredProcedure [dbo].[FT_AllocationMarkerUpdate_Ins]    Script Date: 09/16/2015 08:08:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[FT_AllocationMarkerUpdate_Ins]
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
@ft_filepath nvarchar(max),
@LineIndicator nvarchar(50) ,
@StartinginCell nvarchar(50),
@inputFields nvarchar(max),
@updateFields nvarchar(max)

as
begin
        insert into dbo.FinTools_Settings_AllocationMarkerUpdate
        (JournalNumber  ,
	Ledger,ft_Account,Period,TransactionDate,
        JrnlType,TransRef,
        AlloctnMarker,LA1,
        LA2,LA3,LA4,LA5,LA6,LA7,LA8,LA9,LA10,
        
        ft_filepath,LineIndicator,StartinginCell,inputFields,updateFields) 
        values(@JournalNumber  ,
	@Ledger,@ft_Account,@Period,@TransactionDate,@JrnlType,
        @TransRef,
        @AlloctnMarker,@LA1,@LA2,@LA3,@LA4,
        @LA5,@LA6,@LA7,@LA8,@LA9,@LA10,
        
        @ft_filepath,@LineIndicator,@StartinginCell,@inputFields,
@updateFields )
end
GO
/****** Object:  StoredProcedure [dbo].[FT_AllocationMarkerUpdate_Del]    Script Date: 09/16/2015 08:08:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
create procedure [dbo].[FT_AllocationMarkerUpdate_Del]
@ft_filepath nvarchar(500)
--@sname varchar(10),
--@sex char(2)
as
begin
        delete dbo.FinTools_Settings_AllocationMarkerUpdate where ft_filepath=@ft_filepath 
end
GO
/****** Object:  Default [DF_FinTools_CSL_ft_id]    Script Date: 09/16/2015 08:08:01 ******/
ALTER TABLE [dbo].[FinTools_CSL] ADD  CONSTRAINT [DF_FinTools_CSL_ft_id]  DEFAULT (newid()) FOR [ft_id]
GO
/****** Object:  Default [DF_FinTools_Settings_ContainerSetting_ft_id]    Script Date: 09/16/2015 08:08:01 ******/
ALTER TABLE [dbo].[FinTools_Settings_ContainerSetting] ADD  CONSTRAINT [DF_FinTools_Settings_ContainerSetting_ft_id]  DEFAULT (newid()) FOR [ft_id]
GO
/****** Object:  Default [DF_FinTools_Settings_CriteriaPaneRemembers_ft_id]    Script Date: 09/16/2015 08:08:01 ******/
ALTER TABLE [dbo].[FinTools_Settings_CriteriaPaneRemembers] ADD  CONSTRAINT [DF_FinTools_Settings_CriteriaPaneRemembers_ft_id]  DEFAULT (newid()) FOR [ft_id]
GO
/****** Object:  Default [DF_FinTools_Settings_DrillDown_GUID]    Script Date: 09/16/2015 08:08:01 ******/
ALTER TABLE [dbo].[FinTools_Settings_DrillDown] ADD  CONSTRAINT [DF_FinTools_Settings_DrillDown_GUID]  DEFAULT (newid()) FOR [GUID]
GO
/****** Object:  Default [DF_FinTools_Settings_Fields_GUID]    Script Date: 09/16/2015 08:08:01 ******/
ALTER TABLE [dbo].[FinTools_Settings_Fields] ADD  CONSTRAINT [DF_FinTools_Settings_Fields_GUID]  DEFAULT (newid()) FOR [GUID]
GO
/****** Object:  Default [DF_FinTools_Settings_GenDescFields_GUID]    Script Date: 09/16/2015 08:08:01 ******/
ALTER TABLE [dbo].[FinTools_Settings_GenDescFields] ADD  CONSTRAINT [DF_FinTools_Settings_GenDescFields_GUID]  DEFAULT (newid()) FOR [GUID]
GO
/****** Object:  Default [DF_FinTools_Settings_VisiableRemembers_ft_id]    Script Date: 09/16/2015 08:08:01 ******/
ALTER TABLE [dbo].[FinTools_Settings_VisiableRemembers] ADD  CONSTRAINT [DF_FinTools_Settings_VisiableRemembers_ft_id]  DEFAULT (newid()) FOR [ft_id]
GO
/****** Object:  Default [DF_FinTools_Invoicing_GUID]    Script Date: 09/16/2015 08:08:01 ******/
ALTER TABLE [dbo].[FinTools_Templates] ADD  CONSTRAINT [DF_FinTools_Invoicing_GUID]  DEFAULT (newid()) FOR [GUID]
GO
/****** Object:  Default [DF_FinTools_Users_ft_id]    Script Date: 09/16/2015 08:08:01 ******/
ALTER TABLE [dbo].[FinTools_Users] ADD  CONSTRAINT [DF_FinTools_Users_ft_id]  DEFAULT (newid()) FOR [ft_id]
GO
