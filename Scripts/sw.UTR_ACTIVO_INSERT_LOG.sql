IF EXISTS
(
    SELECT 1
    FROM sys.triggers t
    WHERE t.name = 'UTR_ACTIVO_INSERT_LOG'
          AND OBJECT_SCHEMA_NAME(t.object_id) = 'sw' -- validación de esquema en SQL 2008
)
BEGIN
    DROP TRIGGER [sw].[UTR_ACTIVO_INSERT_LOG];
END;
GO
CREATE TRIGGER [sw].[UTR_ACTIVO_INSERT_LOG]
ON [sw].[ACTIVO]
AFTER INSERT
AS
BEGIN
    SET NOCOUNT ON;

    INSERT INTO [sw].[ACTIVOLOG]
    (
        codCia,
        activoId,
        logId,
        codigoActivo,
        descripcion,
        numeroSerie,
        categoriaId,
        proveedorId,
        costoInicial,
        fechaIngreso,
		responsableId,
		ubicacionId,
        activo,
        feRegistro,
        cuRegistro,
        eliminado,
        feEliminado,
        cuEliminado
    )
    SELECT i.codCia,
           i.activoId,
           ISNULL(
           (
               SELECT MAX(logId)
               FROM [sw].[ACTIVOLOG]
               WHERE codCia = i.codCia
                     AND activoId = i.activoId
           ),
           0
                 ) + 1 AS logId,
           i.codigoActivo,
           i.descripcion,
           i.numeroSerie,
           i.categoriaId,
           i.proveedorId,
           i.costoInicial,
           i.fechaIngreso,
		   i.responsableId,
		   i.ubicacionId,
           i.activo,
           i.feRegistro,
           i.cuRegistro,
           i.eliminado,
           i.feEliminado,
           i.cuEliminado
    FROM INSERTED i;
END;
GO
