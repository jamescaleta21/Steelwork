IF EXISTS
(
    SELECT TOP 1
           s.SPECIFIC_NAME
    FROM INFORMATION_SCHEMA.ROUTINES s
    WHERE s.ROUTINE_TYPE = 'PROCEDURE' -- Validación del tipo
          AND ROUTINE_SCHEMA = 'sw' -- Validación del esquema
          AND s.ROUTINE_NAME = 'USP_BAJA_UBICACIONACTUAL_ACTIVO'
) -- Validación del nombre
BEGIN
    DROP PROC [sw].[USP_BAJA_UBICACIONACTUAL_ACTIVO];
END;
GO
/*
sw.USP_BAJA_UBICACIONACTUAL_ACTIVO '01',3
*/
CREATE PROCEDURE [sw].[USP_BAJA_UBICACIONACTUAL_ACTIVO]
    @codcia CHAR(2),
    @activoid INT
WITH ENCRYPTION
AS
BEGIN
    SET NOCOUNT ON;

    SELECT u.denominacion AS ubicacion
    FROM sw.ACTIVO a WITH (NOLOCK)
        INNER JOIN sw.UBICACION u WITH (NOLOCK)
            ON u.codCia = a.codCia
               AND u.ubicacionId = a.ubicacionId
    WHERE a.activo = 1
          AND a.eliminado = 0
          AND a.activoId = @activoid
          AND a.codCia = @codcia;
END;
GO