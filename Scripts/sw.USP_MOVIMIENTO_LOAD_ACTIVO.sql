IF EXISTS
(
    SELECT TOP 1
           s.SPECIFIC_NAME
    FROM INFORMATION_SCHEMA.ROUTINES s
    WHERE s.ROUTINE_TYPE = 'PROCEDURE' -- Validación del tipo
          AND ROUTINE_SCHEMA = 'sw' -- Validación del esquema
          AND s.ROUTINE_NAME = 'USP_MOVIMIENTO_LOAD_ACTIVO'
) -- Validación del nombre
BEGIN
    DROP PROC [sw].[USP_MOVIMIENTO_LOAD_ACTIVO];
END;
GO
/*
sw.USP_MOVIMIENTO_LOAD_ACTIVO '01'
*/
CREATE PROCEDURE [sw].[USP_MOVIMIENTO_LOAD_ACTIVO] @codcia CHAR(2)
WITH ENCRYPTION
AS
BEGIN
    SET NOCOUNT ON;


    SELECT -1 AS activoId,
           '.: SELECCIONE :.' AS descripcion
    UNION
    SELECT a.activoId,
           a.descripcion
    FROM sw.ACTIVO a
    WHERE a.codCia = @codcia
          AND a.activo = 1
    ORDER BY 2;
END;
GO