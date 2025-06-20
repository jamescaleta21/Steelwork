IF EXISTS
(
    SELECT 1
    FROM sys.triggers t
    WHERE t.name = 'UTR_CATEGORIA_INSERT_LOG'
          AND OBJECT_SCHEMA_NAME(t.object_id) = 'sw' -- validación de esquema en SQL 2008
)
BEGIN
    DROP TRIGGER [sw].[UTR_CATEGORIA_INSERT_LOG];
END;
GO
CREATE TRIGGER [sw].[UTR_CATEGORIA_INSERT_LOG]
ON [sw].[CATEGORIA]
AFTER INSERT
AS
BEGIN
    SET NOCOUNT ON;

	INSERT INTO sw.CATEGORIALOG
	(
	    codCia,
	    categoriaId,
	    logId,
	    descripcion,
	    activo,
	    feRegistro,
	    cuRegistro,
	    eliminado,
	    feEliminado,
	    cuEliminado
	)
    SELECT i.codCia,
           i.categoriaId,
           ISNULL(
           (
               SELECT MAX(logId)
               FROM [sw].[CATEGORIALOG]
               WHERE codCia = i.codCia
                     AND categoriaid = i.categoriaId
           ),
           0
                 ) + 1 AS logId,
           i.descripcion,
           i.activo,
           i.feRegistro,
           i.cuRegistro,
           i.eliminado,
           i.feEliminado,
           i.cuEliminado
    FROM INSERTED i;
END;
GO
