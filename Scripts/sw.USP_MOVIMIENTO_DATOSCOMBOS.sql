IF EXISTS
(
    SELECT TOP 1
           s.SPECIFIC_NAME
    FROM INFORMATION_SCHEMA.ROUTINES s
    WHERE s.ROUTINE_TYPE = 'PROCEDURE' -- Validación del tipo
          AND ROUTINE_SCHEMA = 'sw' -- Validación del esquema
          AND s.ROUTINE_NAME = 'USP_MOVIMIENTO_DATOSCOMBOS'
) -- Validación del nombre
BEGIN
    DROP PROC [sw].[USP_MOVIMIENTO_DATOSCOMBOS];
END;
GO
/*
sw.USP_MOVIMIENTO_DATOSCOMBOS '01'
*/
CREATE PROCEDURE [sw].[USP_MOVIMIENTO_DATOSCOMBOS] @codcia CHAR(2)
WITH ENCRYPTION
AS
SET NOCOUNT ON;

SELECT -1 AS activoId,
       '.: SELECCIONE :.' AS descripcion
UNION
SELECT a.activoId,
       a.descripcion
FROM sw.ACTIVO a
WHERE a.activo = 1
      AND a.eliminado = 0
ORDER BY 2;


SELECT -1 AS responsableId,
       '.: SELECCIONE :.' AS nombres
UNION
SELECT a.responsableId,
       a.apellidos + ' ' + a.nombres AS nombres
FROM sw.RESPONSABLE a
WHERE a.activo = 1
      AND a.eliminado = 0
ORDER BY 2;


SELECT -1 AS ubicacionId,
       '.: SELECCIONE :.' AS denominacion
UNION
SELECT a.ubicacionId,
       a.denominacion
FROM sw.UBICACION a
WHERE a.activo = 1
      AND a.eliminado = 0
ORDER BY 2;
GO