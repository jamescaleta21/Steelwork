IF EXISTS
(
    SELECT TOP 1
           s.SPECIFIC_NAME
    FROM INFORMATION_SCHEMA.ROUTINES s
    WHERE s.ROUTINE_TYPE = 'PROCEDURE' -- Validación del tipo
          AND ROUTINE_SCHEMA = 'sw' -- Validación del esquema
          AND s.ROUTINE_NAME = 'USP_MOVIMIENTO_ACTIVO_SEARCH'
) -- Validación del nombre
BEGIN
    DROP PROC [sw].[USP_MOVIMIENTO_ACTIVO_SEARCH];
END;
GO
/*
sw.USP_MOVIMIENTO_ACTIVO_SEARCH '01','DE'
*/
CREATE PROCEDURE [sw].[USP_MOVIMIENTO_ACTIVO_SEARCH]
    @codcia CHAR(2),
    @SEARCH VARCHAR(100)
WITH ENCRYPTION
AS
BEGIN
    SET NOCOUNT ON;



    SELECT a.activoId,
           a.descripcion
    FROM sw.ACTIVO a
    WHERE a.codCia = @codcia
          AND a.descripcion LIKE '%' + @SEARCH + '%'
          AND a.activo = 1
    ORDER BY 2;
END;
GO