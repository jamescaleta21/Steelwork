IF EXISTS
(
    SELECT 1
    FROM sys.triggers t
    WHERE t.name = 'UTR_PROVEEDOR_UPDATE_LOG'
          AND OBJECT_SCHEMA_NAME(t.object_id) = 'sw' -- validación de esquema en SQL 2008
)
BEGIN
    DROP TRIGGER [sw].[UTR_PROVEEDOR_UPDATE_LOG];
END;
GO
CREATE TRIGGER [sw].[UTR_PROVEEDOR_UPDATE_LOG]
ON [sw].[PROVEEDOR]
AFTER UPDATE
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
               FROM sw.PROVEEDORLOG pl
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
    FROM INSERTED i
        INNER JOIN DELETED d
            ON i.codCia = d.codCia
               AND i.proveedorId = d.proveedorId
    WHERE  ISNULL(i.razonSocial, '') <> ISNULL(d.razonSocial, '')
          OR i.eliminado <> d.eliminado
          OR i.feEliminado <> d.feEliminado
          OR i.cuEliminado <> d.cuEliminado
		  OR i.activo <> d.activo;
END;
GO
