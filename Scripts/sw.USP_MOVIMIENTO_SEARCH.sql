IF EXISTS
(
    SELECT TOP 1
           s.SPECIFIC_NAME
    FROM INFORMATION_SCHEMA.ROUTINES s
    WHERE s.ROUTINE_TYPE = 'PROCEDURE' -- Validación del tipo
          AND ROUTINE_SCHEMA = 'sw' -- Validación del esquema
          AND s.ROUTINE_NAME = 'USP_MOVIMIENTO_SEARCH'
) -- Validación del nombre
BEGIN
    DROP PROC [sw].[USP_MOVIMIENTO_SEARCH];
END;
GO
/*
sw.USP_MOVIMIENTO_SEARCH '01',3
*/
CREATE PROCEDURE [sw].[USP_MOVIMIENTO_SEARCH]
    @codcia CHAR(2),
    @activoId INT
WITH ENCRYPTION
AS
BEGIN
    SET NOCOUNT ON;

    SELECT u.denominacion AS 'ubicacion'
    FROM sw.ACTIVO a WITH (NOLOCK)
        INNER JOIN sw.UBICACION u WITH (NOLOCK)
            ON u.codCia = a.codCia
               AND u.ubicacionId = a.ubicacionId
    WHERE a.codCia = @codcia
          AND a.activoId = @activoId;

    -- Ubicación inicial
    SELECT NULL AS movimientoId,
           al.fechaIngreso AS fechaMovimiento,
           NULL AS ResponsableOrigen,
           NULL AS ResponsableDestino,
           u.denominacion AS ubicacion,
           'UBICACIÓN INICIAL' AS tipoMovimiento,
           al.feRegistro,
           al.cuRegistro
    FROM sw.ACTIVOLOG al
        LEFT JOIN sw.UBICACION u
            ON u.codCia = al.codCia
               AND u.ubicacionId = al.ubicacionId
    WHERE al.codCia = @codcia
          AND al.activoId = @activoId
          AND al.logId = 1
    UNION ALL

    -- Movimientos posteriores
    SELECT m.movimientoId,
           m.fechaMovimiento,
           r.apellidos + ' ' + r.nombres AS ResponsableOrigen,
           r2.apellidos + ' ' + r2.nombres AS ResponsableDestino,
           u.denominacion AS ubicacion,
           'TRASLADO' AS tipoMovimiento,
           m.feRegistro,
           m.cuRegistro
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
          AND m.activoId = @activoId
    ORDER BY al.feRegistro,
             movimientoId;



END;
GO