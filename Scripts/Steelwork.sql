CREATE SCHEMA sw AUTHORIZATION dbo;
GO


CREATE TABLE [sw].[ACTIVO] (
  [codCia] char(2) NOT NULL,
  [activoId] integer NOT NULL,
  [codigoActivo] varchar(20) NOT NULL,
  [descripcion] varchar(100) NOT NULL,
  [numeroSerie] varchar(100),
  [categoriaId] integer NOT NULL,
  [proveedorId] integer NOT NULL,
  [costoInicial] money NOT NULL,
  [fechaIngreso] date NOT NULL DEFAULT (getdate()),
  [responsableId] integer NOT NULL,
  [ubicacionId] integer NOT NULL,
  [activo] bit NOT NULL DEFAULT (1),
  [feRegistro] datetime NOT NULL DEFAULT (getdate()),
  [cuRegistro] varchar(20) NOT NULL,
  [eliminado] bit NOT NULL DEFAULT (0),
  [feEliminado] datetime,
  [cuEliminado] varchar(20),
  PRIMARY KEY ([codCia], [activoId])
)
GO

CREATE TABLE [sw].[ACTIVOLOG] (
  [codCia] char(2) NOT NULL,
  [activoId] integer NOT NULL,
  [logId] integer NOT NULL,
  [codigoActivo] varchar(20) NOT NULL,
  [descripcion] varchar(100) NOT NULL,
  [numeroSerie] varchar(100),
  [categoriaId] integer NOT NULL,
  [proveedorId] integer NOT NULL,
  [costoInicial] money NOT NULL,
  [fechaIngreso] date NOT NULL,
   [responsableId] integer NOT NULL,
  [ubicacionId] integer NOT NULL,
  [activo] bit NOT NULL,
  [feRegistro] datetime NOT NULL,
  [cuRegistro] varchar(20) NOT NULL,
  [eliminado] bit NOT NULL,
  [feEliminado] datetime,
  [cuEliminado] varchar(20),
  PRIMARY KEY ([codCia], [activoId], [logId])
)
GO

CREATE TABLE [sw].[RESPONSABLE] (
  [codCia] char(2) NOT NULL,
  [responsableId] integer NOT NULL,
  [nombres] varchar(100) NOT NULL,
  [apellidos] varchar(200) NOT NULL,
  [activo] bit NOT NULL DEFAULT (1),
  [feRegistro] datetime NOT NULL DEFAULT (getdate()),
  [cuRegistro] varchar(20) NOT NULL,
  [eliminado] bit NOT NULL DEFAULT (0),
  [feEliminado] datetime,
  [cuEliminado] varchar(20),
  PRIMARY KEY ([codCia], [responsableId])
)
GO

CREATE TABLE [sw].[RESPONSABLELOG] (
  [codCia] char(2) NOT NULL,
  [responsableId] integer NOT NULL,
  [logId] integer NOT NULL,
  [nombres] varchar(100) NOT NULL,
  [apellidos] varchar(200) NOT NULL,
  [activo] bit NOT NULL DEFAULT (1),
  [feRegistro] datetime NOT NULL DEFAULT (getdate()),
  [cuRegistro] varchar(20) NOT NULL,
  [eliminado] bit NOT NULL DEFAULT (0),
  [feEliminado] datetime,
  [cuEliminado] varchar(20),
  PRIMARY KEY ([codCia], [responsableId], [logId])
)
GO

CREATE TABLE [sw].[MOVIMIENTO] (
  [codCia] char(2) NOT NULL,
  [movimientoId] integer NOT NULL,
  [activoId] integer NOT NULL,
  [fechaMovimiento] date NOT NULL,
  [responsableIdOrigen] integer NOT NULL,
  [responsableIdDestino] integer NOT NULL,
  [ubicacionId] integer NOT NULL,
  [observacion] varchar(300),
  [feRegistro] datetime NOT NULL DEFAULT (getdate()),
  [cuRegistro] varchar(20) NOT NULL,
  PRIMARY KEY ([codCia], [movimientoId])
)
GO

CREATE TABLE [sw].[BAJA] (
  [codCia] char(2) NOT NULL,
  [bajaId] integer NOT NULL,
  [activoId] integer NOT NULL,
  [fechaBaja] date NOT NULL,
  [motivoBaja] varchar(200) NOT NULL,
  [ubicacionId] integer NOT NULL,
  [fechaRegistro] datetime NOT NULL DEFAULT (getdate()),
  [cuRegistro] varchar(20) NOT NULL,
  PRIMARY KEY ([codCia], [bajaId])
)
GO

CREATE TABLE [sw].[CATEGORIA] (
  [codCia] char(2) NOT NULL,
  [categoriaId] integer NOT NULL,
  [descripcion] varchar(80) NOT NULL,
  [activo] bit NOT NULL DEFAULT (1),
  [feRegistro] datetime NOT NULL DEFAULT (getdate()),
  [cuRegistro] varchar(20) NOT NULL,
  [eliminado] bit NOT NULL DEFAULT (0),
  [feEliminado] datetime,
  [cuEliminado] varchar(20),
  PRIMARY KEY ([codCia], [categoriaId])
)
GO

CREATE TABLE [sw].[CATEGORIALOG] (
  [codCia] char(2) NOT NULL,
  [categoriaId] integer NOT NULL,
  [logId] integer NOT NULL,
  [descripcion] varchar(80) NOT NULL,
  [activo] bit NOT NULL,
  [feRegistro] datetime NOT NULL,
  [cuRegistro] varchar(20) NOT NULL,
  [eliminado] bit NOT NULL,
  [feEliminado] datetime,
  [cuEliminado] varchar(20),
  PRIMARY KEY ([codCia], [categoriaId], [logId])
)
GO

CREATE TABLE [sw].[PROVEEDOR] (
  [codCia] char(2) NOT NULL,
  [proveedorId] integer NOT NULL,
  [razonSocial] varchar(100) NOT NULL,
  [activo] bit NOT NULL DEFAULT (1),
  [feRegistro] datetime NOT NULL DEFAULT (getdate()),
  [cuRegistro] varchar(20) NOT NULL,
  [eliminado] bit NOT NULL DEFAULT (0),
  [cuEliminado] varchar(20),
  [feEliminado] datetime,
  PRIMARY KEY ([codCia], [proveedorId])
)
GO

CREATE TABLE [sw].[PROVEEDORLOG] (
  [codCia] char(2) NOT NULL,
  [proveedorId] integer NOT NULL,
  [logId] integer NOT NULL,
  [razonSocial] varchar(100) NOT NULL,
  [activo] bit NOT NULL DEFAULT (1),
  [feRegistro] datetime NOT NULL DEFAULT (getdate()),
  [cuRegistro] varchar(20) NOT NULL,
  [eliminado] bit NOT NULL DEFAULT (0),
  [cuEliminado] varchar(20),
  [feEliminado] datetime,
  PRIMARY KEY ([codCia], [proveedorId], [logId])
)
GO

CREATE TABLE [sw].[UBICACION] (
  [codCia] char(2) NOT NULL,
  [ubicacionId] integer NOT NULL,
  [denominacion] varchar(100) NOT NULL,
  [activo] bit NOT NULL DEFAULT (1),
  [feRegistro] datetime NOT NULL DEFAULT (getdate()),
  [cuRegistro] varchar(20) NOT NULL,
  [eliminado] bit NOT NULL DEFAULT (0),
  [feEliminado] datetime,
  [cuEliminado] varchar(20),
  PRIMARY KEY ([codCia], [ubicacionId])
)
GO

CREATE TABLE [sw].[UBICACIONLOG] (
  [codCia] char(2) NOT NULL,
  [ubicacionId] integer NOT NULL,
  [logId] integer NOT NULL,
  [denominacion] varchar(100) NOT NULL,
  [activo] bit NOT NULL,
  [feRegistro] datetime NOT NULL,
  [cuRegistro] varchar(20) NOT NULL,
  [eliminado] bit NOT NULL,
  [feEliminado] datetime,
  [cuEliminado] varchar(20),
  PRIMARY KEY ([codCia], [ubicacionId], [logId])
)
GO

CREATE INDEX [IX_Activo_activo] ON [sw].[Activo] ("activo")
GO

CREATE INDEX [Activo_codigoActivo_descripcion] ON [sw].[Activo] ("codigoActivo", "descripcion")
GO

CREATE INDEX [IX_Activo_responsable_ubicacion] ON [sw].[Activo] ("responsableId", "ubicacionId")
GO

CREATE INDEX [IX_Activo_codigoActivo] ON [sw].[Activo] ("codigoActivo")
GO

CREATE INDEX [IX_ActivoLog_activo] ON [sw].[ActivoLog] ("activo")
GO

CREATE INDEX [IX_ActivoLog_codigoActivo_descripcion] ON [sw].[ActivoLog] ("codigoActivo", "descripcion")
GO

CREATE INDEX [IX_Responsable_nombres_apellidos] ON [sw].[Responsable] ("nombres", "apellidos")
GO

CREATE INDEX [IX_Responsable_activo_fechaRegistro] ON [sw].[Responsable] ("activo", "feRegistro")
GO

CREATE INDEX [IX_Movimiento_activoId_fechaMovimiento] ON [sw].[Movimiento] ("activoId", "fechaMovimiento")
GO

CREATE INDEX [IX_Movimiento_fechaMovimiento] ON [sw].[Movimiento] ("fechaMovimiento")
GO

CREATE INDEX [IX_Movimiento_activoId_responsableIdDestino] ON [sw].[Movimiento] ("activoId", "responsableIdDestino")
GO

CREATE INDEX [IX_Baja_fechaBaja_cuRegistro] ON [sw].[Baja] ("fechaBaja", "cuRegistro")
GO

CREATE INDEX [IX_Baja_activoId_fechaBaja] ON [sw].[Baja] ("activoId", "fechaBaja")
GO

CREATE INDEX [IX_Categoria_activo_descripcion] ON [sw].[Categoria] ("activo", "descripcion")
GO

CREATE INDEX [IX_Proveedor_activo_razonsocial] ON [sw].[Proveedor] ("activo", "razonSocial")
GO

CREATE INDEX [IX_Proveedor_razonSocial] ON [sw].[Proveedor] ("razonSocial")
GO

CREATE INDEX [IX_Ubicacion_activo_feRegistro] ON [sw].[Ubicacion] ("activo", "feRegistro")
GO

CREATE INDEX [IX_UbicacionLog_activo_feRegistro] ON [sw].[UbicacionLog] ("activo", "feRegistro")
GO

ALTER TABLE [sw].[Baja] ADD FOREIGN KEY ([codCia], [activoId]) REFERENCES [sw].[Activo] ([codCia], [activoId])
GO

ALTER TABLE [sw].[ActivoLog] ADD FOREIGN KEY ([codCia], [activoId]) REFERENCES [sw].[Activo] ([codCia], [activoId])
GO

ALTER TABLE [sw].[Activo] ADD FOREIGN KEY ([codCia], [categoriaId]) REFERENCES [sw].[Categoria] ([codCia], [categoriaId])
GO

ALTER TABLE [sw].[CategoriaLog] ADD FOREIGN KEY ([codCia], [categoriaId]) REFERENCES [sw].[Categoria] ([codCia], [categoriaId])
GO

ALTER TABLE [sw].[Activo] ADD FOREIGN KEY ([codCia], [proveedorId]) REFERENCES [sw].[Proveedor] ([codCia], [proveedorId])
GO

ALTER TABLE [sw].[ProveedorLog] ADD FOREIGN KEY ([codCia], [proveedorId]) REFERENCES [sw].[Proveedor] ([codCia], [proveedorId])
GO

ALTER TABLE [sw].[Movimiento] ADD FOREIGN KEY ([codCia], [activoId]) REFERENCES [sw].[Activo] ([codCia], [activoId])
GO

ALTER TABLE [sw].[ResponsableLog] ADD FOREIGN KEY ([codCia], [responsableId]) REFERENCES [sw].[Responsable] ([codCia], [responsableId])
GO

ALTER TABLE [sw].[Baja] ADD FOREIGN KEY ([codCia], [ubicacionId]) REFERENCES [sw].[Ubicacion] ([codCia], [ubicacionId])
GO

ALTER TABLE [sw].[Movimiento] ADD FOREIGN KEY ([codCia], [ubicacionId]) REFERENCES [sw].[Ubicacion] ([codCia], [ubicacionId])
GO

ALTER TABLE [sw].[Activo] ADD FOREIGN KEY ([codCia], [responsableId]) REFERENCES [sw].[Responsable] ([codCia], [responsableId])
GO

ALTER TABLE [sw].[Activo] ADD FOREIGN KEY ([codCia], [ubicacionId]) REFERENCES [sw].[Ubicacion] ([codCia], [ubicacionId])
GO

ALTER TABLE [sw].[UbicacionLog] ADD FOREIGN KEY ([codCia], [ubicacionId]) REFERENCES [sw].[Ubicacion] ([codCia], [ubicacionId])
GO

ALTER TABLE [sw].[Movimiento] ADD FOREIGN KEY ([codCia], [responsableIdOrigen]) REFERENCES [sw].[Responsable] ([codCia], [responsableId])
GO

ALTER TABLE [sw].[Movimiento] ADD FOREIGN KEY ([codCia], [responsableIdDestino]) REFERENCES [sw].[Responsable] ([codCia], [responsableId])
GO

