IF EXISTS
(
    SELECT 1
    FROM sys.triggers t
    WHERE t.name = 'UTR_UBICACION_UPDATE_LOG'
          AND OBJECT_SCHEMA_NAME(t.object_id) = 'sw' -- validación de esquema en SQL 2008
)
BEGIN
    DROP TRIGGER [sw].[UTR_UBICACION_UPDATE_LOG];
END;
GO
CREATE TRIGGER [sw].[UTR_UBICACION_UPDATE_LOG]
ON [sw].[UBICACION]
AFTER UPDATE
AS
BEGIN
    SET NOCOUNT ON;

    INSERT INTO sw.UBICACIONLOG
    (
        codCia,
        ubicacionId,
        logId,
        denominacion,
        activo,
        feRegistro,
        cuRegistro,
        eliminado,
        feEliminado,
        cuEliminado
    )
    SELECT i.codCia,
           i.ubicacionId,
           ISNULL(
           (
               SELECT MAX(logId)
               FROM [sw].[UBICACIONLOG] ul
               WHERE ul.codCia = i.codCia
                     AND ul.ubicacionId = i.ubicacionId
           ),
           0
                 ) + 1 AS logId,
           i.denominacion,
           i.activo,
           i.feRegistro,
           i.cuRegistro,
           i.eliminado,
           i.feEliminado,
           i.cuEliminado
    FROM INSERTED i
        INNER JOIN DELETED d
            ON i.codCia = d.codCia
               AND i.ubicacionId = d.ubicacionId
    WHERE ISNULL(i.denominacion, '') <> ISNULL(d.denominacion, '')
          OR i.eliminado <> d.eliminado
          OR i.feEliminado <> d.feEliminado
          OR i.cuEliminado <> d.cuEliminado
          OR i.activo <> d.activo;
END;
GO
