IF EXISTS
(
    SELECT TOP 1
           s.SPECIFIC_NAME
    FROM INFORMATION_SCHEMA.ROUTINES s
    WHERE s.ROUTINE_TYPE = 'PROCEDURE' -- Validación del tipo
          AND ROUTINE_SCHEMA = 'sw' -- Validación del esquema
          AND s.ROUTINE_NAME = 'USP_ACTIVO_SEARCH'
) -- Validación del nombre
BEGIN
    DROP PROC [sw].[USP_ACTIVO_SEARCH];
END;
GO
/*
sw.USP_ACTIVO_SEARCH '01','1'
*/
CREATE PROCEDURE [sw].[USP_ACTIVO_SEARCH]
    @codcia CHAR(2),
    @search VARCHAR(100) = NULL
WITH ENCRYPTION
AS
SET NOCOUNT ON;

SELECT c.activoId,
       c.codigoActivo,
	   C.descripcion,
	   C.costoInicial,
       CASE c.activo
           WHEN 1 THEN
               'SI'
           ELSE
               'NO'
       END AS activo
FROM sw.ACTIVO c
WHERE c.codCia = @codcia
      AND c.eliminado = 0
      AND
      (
          c.codigoActivo LIKE '%' + @search + '%'
		  OR C.descripcion LIKE '%' + @search + '%' 
          OR @search IS NULL
      );

GO