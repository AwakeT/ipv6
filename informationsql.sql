USE [shuihu]
GO
/****** Object:  Table [dbo].[pluginInformation]    Script Date: 10/21/2016 01:13:26 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[pluginInformation](
	[PID] [int] IDENTITY(1,1) NOT NULL,
	[CID] [int] NULL,
	[tagName] [nvarchar](32) NOT NULL,
	[macAddress] [nvarchar](50) NOT NULL,
	[memo] [nvarchar](512) NULL
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[pluginInformation] DISABLE CHANGE_TRACKING
GO
EXEC sys.sp_addextendedproperty @name=N'MS_Description', @value=N'所属采集器ID' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'pluginInformation', @level2type=N'COLUMN',@level2name=N'CID'
GO
SET IDENTITY_INSERT [dbo].[pluginInformation] ON
INSERT [dbo].[pluginInformation] ([PID], [CID], [tagName], [macAddress], [memo]) VALUES (2, 2, N'插座1111', N'18FE34B01C22', NULL)
INSERT [dbo].[pluginInformation] ([PID], [CID], [tagName], [macAddress], [memo]) VALUES (3, 2, N'插座1112', N'18FE34B07C19', NULL)
INSERT [dbo].[pluginInformation] ([PID], [CID], [tagName], [macAddress], [memo]) VALUES (4, 2, N'插座1113', N'18FE34D465CD', NULL)
INSERT [dbo].[pluginInformation] ([PID], [CID], [tagName], [macAddress], [memo]) VALUES (5, 2, N'插座1114', N'18FE34D465A1', NULL)
INSERT [dbo].[pluginInformation] ([PID], [CID], [tagName], [macAddress], [memo]) VALUES (7, 3, N'插座1121', N'00012   ', NULL)
INSERT [dbo].[pluginInformation] ([PID], [CID], [tagName], [macAddress], [memo]) VALUES (8, 7, N'插座1211', N'000213  ', NULL)
SET IDENTITY_INSERT [dbo].[pluginInformation] OFF
/****** Object:  Table [dbo].[collectorInformation]    Script Date: 10/21/2016 01:13:26 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[collectorInformation](
	[CID] [int] IDENTITY(1,1) NOT NULL,
	[AID] [int] NULL,
	[tagName] [nvarchar](32) NOT NULL,
	[macAddress] [nvarchar](50) NOT NULL,
	[memo] [nvarchar](512) NULL
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[collectorInformation] DISABLE CHANGE_TRACKING
GO
EXEC sys.sp_addextendedproperty @name=N'MS_Description', @value=N'所在区域的ID' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'collectorInformation', @level2type=N'COLUMN',@level2name=N'AID'
GO
SET IDENTITY_INSERT [dbo].[collectorInformation] ON
INSERT [dbo].[collectorInformation] ([CID], [AID], [tagName], [macAddress], [memo]) VALUES (2, 1, N'101采集器', N'192.168.2.112   ', NULL)
INSERT [dbo].[collectorInformation] ([CID], [AID], [tagName], [macAddress], [memo]) VALUES (3, 1, N'102采集器', N'00002   ', NULL)
INSERT [dbo].[collectorInformation] ([CID], [AID], [tagName], [macAddress], [memo]) VALUES (7, 2, N'201采集器', N'02      ', NULL)
SET IDENTITY_INSERT [dbo].[collectorInformation] OFF
/****** Object:  Table [dbo].[buildingInformation]    Script Date: 10/21/2016 01:13:26 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[buildingInformation](
	[BID] [int] IDENTITY(1,1) NOT NULL,
	[tagName] [nvarchar](32) NOT NULL,
	[memo] [nvarchar](512) NULL
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[buildingInformation] DISABLE CHANGE_TRACKING
GO
SET IDENTITY_INSERT [dbo].[buildingInformation] ON
INSERT [dbo].[buildingInformation] ([BID], [tagName], [memo]) VALUES (1, N'教学A楼', NULL)
SET IDENTITY_INSERT [dbo].[buildingInformation] OFF
/****** Object:  Table [dbo].[areaInformation]    Script Date: 10/21/2016 01:13:26 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[areaInformation](
	[AID] [int] IDENTITY(1,1) NOT NULL,
	[BID] [int] NULL,
	[tagName] [nvarchar](50) NOT NULL,
	[memo] [nvarchar](512) NULL
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[areaInformation] DISABLE CHANGE_TRACKING
GO
EXEC sys.sp_addextendedproperty @name=N'MS_Description', @value=N'所在建筑物的ID' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'areaInformation', @level2type=N'COLUMN',@level2name=N'BID'
GO
SET IDENTITY_INSERT [dbo].[areaInformation] ON
INSERT [dbo].[areaInformation] ([AID], [BID], [tagName], [memo]) VALUES (1, 1, N'11一楼区域', NULL)
INSERT [dbo].[areaInformation] ([AID], [BID], [tagName], [memo]) VALUES (2, 1, N'12二楼区域', NULL)
SET IDENTITY_INSERT [dbo].[areaInformation] OFF
