IF EXISTS
(
    SELECT 1
    FROM sys.triggers t
    WHERE t.name = 'UTR_CATEGORIA_UPDATE_LOG'
          AND OBJECT_SCHEMA_NAME(t.object_id) = 'sw' -- validación de esquema en SQL 2008
)
BEGIN
    DROP TRIGGER [sw].[UTR_CATEGORIA_UPDATE_LOG];
END;
GO
CREATE TRIGGER [sw].[UTR_CATEGORIA_UPDATE_LOG]
ON [sw].[CATEGORIA]
AFTER UPDATE
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
		I.categoriaId,
           ISNULL(
           (
               SELECT MAX(logId)
               FROM [sw].[CATEGORIALOG]
               WHERE codCia = i.codCia
                     AND categoriaId = i.categoriaId
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
    FROM INSERTED i
        INNER JOIN DELETED d
            ON i.codCia = d.codCia
               AND i.categoriaId = d.categoriaId
    WHERE  ISNULL(i.descripcion, '') <> ISNULL(d.descripcion, '')
          OR i.eliminado <> d.eliminado
          OR i.feEliminado <> d.feEliminado
          OR i.cuEliminado <> d.cuEliminado
		  OR i.activo <> d.activo;
END;
GO
