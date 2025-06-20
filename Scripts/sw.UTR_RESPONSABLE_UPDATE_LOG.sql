IF EXISTS
(
    SELECT 1
    FROM sys.triggers t
    WHERE t.name = 'UTR_RESPONSABLE_UPDATE_LOG'
          AND OBJECT_SCHEMA_NAME(t.object_id) = 'sw' -- validación de esquema en SQL 2008
)
BEGIN
    DROP TRIGGER [sw].[UTR_RESPONSABLE_UPDATE_LOG];
END;
GO
CREATE TRIGGER [sw].[UTR_RESPONSABLE_UPDATE_LOG]
ON [sw].[RESPONSABLE]
AFTER UPDATE
AS
BEGIN
    SET NOCOUNT ON;

    INSERT INTO sw.RESPONSABLELOG
    (
        codCia,
        responsableId,
        logId,
        nombres,
        apellidos,
        activo,
        feRegistro,
        cuRegistro,
        eliminado,
        feEliminado,
        cuEliminado
    )
    SELECT i.codCia,
           i.responsableId,
           ISNULL(
           (
               SELECT MAX(logId)
               FROM sw.RESPONSABLELOG pl
               WHERE pl.codCia = i.codCia
                     AND pl.responsableId = i.responsableId
           ),
           0
                 ) + 1 AS logId,
           i.nombres,
           i.apellidos,
           i.activo,
           i.feRegistro,
           i.cuRegistro,
           i.eliminado,
           i.feEliminado,
           i.cuEliminado
    FROM INSERTED i
        INNER JOIN DELETED d
            ON i.codCia = d.codCia
               AND i.responsableId = d.responsableId
    WHERE ISNULL(i.nombres, '') <> ISNULL(d.nombres, '')
          OR ISNULL(i.apellidos, '') <> ISNULL(d.apellidos, '')
          OR i.eliminado <> d.eliminado
          OR i.feEliminado <> d.feEliminado
          OR i.cuEliminado <> d.cuEliminado
          OR i.activo <> d.activo;
END;
GO
