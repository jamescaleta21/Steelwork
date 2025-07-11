CREATE TABLE [dbo].[GUIA_REMISION_CAB](
	[CODCIA] [char](2) NOT NULL,
	[SERIE] [char](4) NOT NULL,
	[NUMERO] [bigint] NOT NULL,
	[FECHAEMISION] [date] NOT NULL,
	[HORAEMISION] [time](7) NOT NULL,
	[FECHATRASLADO] [date] NOT NULL,
	[CODIGO_MOTIVO_TRASLADO] [char](2) NOT NULL,
	[CODIGO_MODALIDAD_TRASLADO] [char](2) NOT NULL,
	[IDTRANSPORTISTA] [int] NOT NULL,
	[DESTINATARIO_TIPO] [int] NULL,
	[DESTINATARIO_NUMERO] [varchar](11) NULL,
	[DESTINATARIO_RAZONSOCIAL] [varchar](100) NULL,
	[PESO_TOTAL] [decimal](10, 2) NOT NULL,
	[BULTOS] [int] NOT NULL,
	[respuesta_sunat_codigo] [varchar](20) NULL,
	[respuesta_sunat_descripcion] [varchar](500) NULL,
	[guia_nro_ticket] [varchar](100) NULL,
	[guia_fec_recepcion] [datetime] NULL,
	[IDCLIENTE] [bigint] NULL,
 CONSTRAINT [PK_GUIA_REMISION_CAB] PRIMARY KEY CLUSTERED 
(
	[CODCIA] ASC,
	[SERIE] ASC,
	[NUMERO] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[GUIA_REMISION_DET]    Script Date: 5/06/2024 08:57:18 p.m. ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[GUIA_REMISION_DET](
	[CODCIA] [char](2) NOT NULL,
	[SERIE] [char](4) NOT NULL,
	[NUMERO] [bigint] NOT NULL,
	[IDPRODUCTO] [int] NOT NULL,
	[CANTIDAD] [decimal](9, 2) NOT NULL,
 CONSTRAINT [PK_GUIA_REMISION_DET] PRIMARY KEY CLUSTERED 
(
	[CODCIA] ASC,
	[SERIE] ASC,
	[NUMERO] ASC,
	[IDPRODUCTO] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[MODALIDAD_TRANSPORTE]    Script Date: 5/06/2024 08:57:18 p.m. ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[MODALIDAD_TRANSPORTE](
	[Codigo] [char](2) NOT NULL,
	[Descripcion] [varchar](200) NOT NULL,
	[Activo] [bit] NOT NULL,
	[CurrentUser] [varchar](20) NOT NULL,
 CONSTRAINT [PK_MODALIDAD_TRANSPORTE] PRIMARY KEY CLUSTERED 
(
	[Codigo] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[MOTIVO_TRASLADO]    Script Date: 5/06/2024 08:57:18 p.m. ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[MOTIVO_TRASLADO](
	[Codigo] [char](2) NOT NULL,
	[Descripcion] [varchar](200) NOT NULL,
	[Activo] [bit] NOT NULL,
	[CurrentUser] [varchar](20) NOT NULL,
 CONSTRAINT [PK_MOTIVO_TRASLADO] PRIMARY KEY CLUSTERED 
(
	[Codigo] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[NUMERACION_DOCUMENTOS]    Script Date: 5/06/2024 08:57:18 p.m. ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[NUMERACION_DOCUMENTOS](
	[CODCIA] [char](2) NOT NULL,
	[TIPO_DOCUMENTO] [char](1) NOT NULL,
	[NUMERO_ACTUAL] [bigint] NOT NULL,
	[SERIE] [char](4) NOT NULL,
 CONSTRAINT [PK_NUMERACION_DOCUMENTOS] PRIMARY KEY CLUSTERED 
(
	[CODCIA] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO
INSERT [dbo].[MODALIDAD_TRANSPORTE] ([Codigo], [Descripcion], [Activo], [CurrentUser]) VALUES (N'02', N'Transporte privado', 1, N'SYSTEM')
GO
INSERT [dbo].[MOTIVO_TRASLADO] ([Codigo], [Descripcion], [Activo], [CurrentUser]) VALUES (N'01', N'Venta', 1, N'SYSTEM')
GO
INSERT [dbo].[MOTIVO_TRASLADO] ([Codigo], [Descripcion], [Activo], [CurrentUser]) VALUES (N'02', N'Compra', 1, N'SYSTEM')
GO
INSERT [dbo].[MOTIVO_TRASLADO] ([Codigo], [Descripcion], [Activo], [CurrentUser]) VALUES (N'03', N'Venta con entrega a terceros', 1, N'SYSTEM')
GO
INSERT [dbo].[MOTIVO_TRASLADO] ([Codigo], [Descripcion], [Activo], [CurrentUser]) VALUES (N'04', N'Traslado entre establecimientos de la misma empresa', 1, N'SYSTEM')
GO
INSERT [dbo].[MOTIVO_TRASLADO] ([Codigo], [Descripcion], [Activo], [CurrentUser]) VALUES (N'05', N'Consignación', 1, N'SYSTEM')
GO
INSERT [dbo].[MOTIVO_TRASLADO] ([Codigo], [Descripcion], [Activo], [CurrentUser]) VALUES (N'06', N'Devolución', 1, N'SYSTEM')
GO
INSERT [dbo].[MOTIVO_TRASLADO] ([Codigo], [Descripcion], [Activo], [CurrentUser]) VALUES (N'07', N'Recojo de bienes transformados', 1, N'SYSTEM')
GO
INSERT [dbo].[MOTIVO_TRASLADO] ([Codigo], [Descripcion], [Activo], [CurrentUser]) VALUES (N'08', N'Importación', 1, N'SYSTEM')
GO
INSERT [dbo].[MOTIVO_TRASLADO] ([Codigo], [Descripcion], [Activo], [CurrentUser]) VALUES (N'09', N'Exportación', 1, N'SYSTEM')
GO
INSERT [dbo].[MOTIVO_TRASLADO] ([Codigo], [Descripcion], [Activo], [CurrentUser]) VALUES (N'13', N'Otros', 1, N'SYSTEM')
GO
INSERT [dbo].[MOTIVO_TRASLADO] ([Codigo], [Descripcion], [Activo], [CurrentUser]) VALUES (N'14', N'Venta sujeta a confirmación del comprador', 1, N'SYSTEM')
GO
INSERT [dbo].[MOTIVO_TRASLADO] ([Codigo], [Descripcion], [Activo], [CurrentUser]) VALUES (N'17', N'Traslado de bienes para transformación', 1, N'SYSTEM')
GO
INSERT [dbo].[MOTIVO_TRASLADO] ([Codigo], [Descripcion], [Activo], [CurrentUser]) VALUES (N'18', N'Traslado emisor itinerante CP', 1, N'SYSTEM')
GO
INSERT [dbo].[NUMERACION_DOCUMENTOS] ([CODCIA], [TIPO_DOCUMENTO], [NUMERO_ACTUAL]) VALUES (N'01', N'G', 0)
GO
ALTER TABLE [dbo].[MODALIDAD_TRANSPORTE] ADD  CONSTRAINT [DF_MODALIDAD_TRANSPORTE_Activo]  DEFAULT ((1)) FOR [Activo]
GO
ALTER TABLE [dbo].[MOTIVO_TRASLADO] ADD  CONSTRAINT [DF_MOTIVO_TRASLADO_Activo]  DEFAULT ((1)) FOR [Activo]
GO
