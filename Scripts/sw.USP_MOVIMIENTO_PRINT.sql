IF EXISTS
(
    SELECT TOP 1
           s.SPECIFIC_NAME
    FROM INFORMATION_SCHEMA.ROUTINES s
    WHERE s.ROUTINE_TYPE = 'PROCEDURE' -- Validación del tipo
          AND ROUTINE_SCHEMA = 'sw' -- Validación del esquema
          AND s.ROUTINE_NAME = 'USP_MOVIMIENTO_PRINT'
) -- Validación del nombre
BEGIN
    DROP PROC [sw].[USP_MOVIMIENTO_PRINT];
END;
GO
/*
sw.USP_MOVIMIENTO_PRINT '01',3
*/
CREATE PROCEDURE [sw].[USP_MOVIMIENTO_PRINT]
    @codcia CHAR(2),
    @activoid INT
WITH ENCRYPTION
AS
BEGIN
    SET NOCOUNT ON;


    SELECT CONVERT(VARCHAR(10), al.fechaIngreso, 103) AS fechaMovimiento,
           '' AS ResponsableOrigen,
           '' AS ResponsableDestino,
           u.denominacion AS ubicacion,
           'UBICACIÓN INICIAL' AS tipoMovimiento,
           al.feRegistro,
           al.cuRegistro,
           '' AS obs
    FROM sw.ACTIVOLOG al
        LEFT JOIN sw.UBICACION u
            ON u.codCia = al.codCia
               AND u.ubicacionId = al.ubicacionId
    WHERE al.codCia = @codcia
          AND al.activoId = @activoid
          AND al.logId = 1
    UNION ALL

    -- Movimientos posteriores
    SELECT CONVERT(VARCHAR(10), m.fechaMovimiento, 103),
           r.apellidos + ' ' + r.nombres AS ResponsableOrigen,
           r2.apellidos + ' ' + r2.nombres AS ResponsableDestino,
           u.denominacion AS ubicacion,
           'TRASLADO' AS tipoMovimiento,
           m.feRegistro,
           m.cuRegistro,
           COALESCE(m.observacion, '') AS obs
    FROM sw.MOVIMIENTO m WITH (NOLOCK)
        INNER JOIN sw.ACTIVO a WITH (NOLOCK)
            ON m.codCia = a.codCia
               AND m.activoId = a.activoId
        INNER JOIN sw.RESPONSABLE r WITH (NOLOCK)
            ON m.codCia = r.codCia
               AND m.responsableIdOrigen = r.responsableId
        INNER JOIN sw.RESPONSABLE r2 WITH (NOLOCK)
            ON m.codCia = r2.codCia
               AND m.responsableIdDestino = r2.responsableId
        INNER JOIN sw.UBICACION u WITH (NOLOCK)
            ON m.codCia = u.codCia
               AND m.ubicacionId = u.ubicacionId
    WHERE m.codCia = @codcia
          AND m.activoId = @activoid
    ORDER BY al.feRegistro;
END;
GO