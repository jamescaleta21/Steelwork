IF EXISTS
(
    SELECT TOP 1
           s.SPECIFIC_NAME
    FROM INFORMATION_SCHEMA.ROUTINES s
    WHERE s.ROUTINE_TYPE = 'PROCEDURE' -- Validación del tipo
          AND ROUTINE_SCHEMA = 'sw' -- Validación del esquema
          AND s.ROUTINE_NAME = 'USP_CATEGORIA_SEARCH'
) -- Validación del nombre
BEGIN
    DROP PROC [sw].[USP_CATEGORIA_SEARCH];
END;
GO
/*
sw.USP_CATEGORIA_SEARCH '01','1'
*/
CREATE PROCEDURE [sw].[USP_CATEGORIA_SEARCH]
    @codcia CHAR(2),
    @search VARCHAR(100) = NULL
WITH ENCRYPTION
AS
SET NOCOUNT ON;

SELECT c.categoriaId,
       c.descripcion,
       CASE c.activo
           WHEN 1 THEN
               'SI'
           ELSE
               'NO'
       END AS activo
FROM sw.CATEGORIA c
WHERE c.codCia = @codcia AND C.eliminado = 0
      AND
      (
          c.descripcion LIKE '%' + @search + '%'
          OR @search IS NULL
      );

GO