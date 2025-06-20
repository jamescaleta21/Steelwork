IF EXISTS
(
    SELECT 1
    FROM sys.triggers t
    WHERE t.name = 'UTR_PROVEEDOR_INSERT_LOG'
          AND OBJECT_SCHEMA_NAME(t.object_id) = 'sw' -- validación de esquema en SQL 2008
)
BEGIN
    DROP TRIGGER [sw].[UTR_PROVEEDOR_INSERT_LOG];
END;
GO
CREATE TRIGGER [sw].[UTR_PROVEEDOR_INSERT_LOG]
ON [sw].[PROVEEDOR]
AFTER INSERT
AS
BEGIN
    SET NOCOUNT ON;

    INSERT INTO sw.PROVEEDORLOG
    (
        codCia,
        proveedorId,
        logId,
        razonSocial,
        activo,
        feRegistro,
        cuRegistro,
        eliminado,
        cuEliminado,
        feEliminado
    )
    SELECT i.codCia,
           i.proveedorId,
           ISNULL(
           (
               SELECT MAX(logId)
               FROM [sw].[PROVEEDORLOG] pl
               WHERE pl.codCia = i.codCia
                     AND pl.proveedorId = i.proveedorId
           ),
           0
                 ) + 1 AS logId,
           i.razonSocial,
           i.activo,
           i.feRegistro,
           i.cuRegistro,
           i.eliminado,
           i.cuEliminado,
           i.feEliminado
    FROM INSERTED i;
END;
GO
