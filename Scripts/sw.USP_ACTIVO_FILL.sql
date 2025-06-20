IF EXISTS
(
    SELECT TOP 1
           s.SPECIFIC_NAME
    FROM INFORMATION_SCHEMA.ROUTINES s
    WHERE s.ROUTINE_TYPE = 'PROCEDURE' -- Validaci�n del tipo
          AND ROUTINE_SCHEMA = 'sw' -- Validaci�n del esquema
          AND s.ROUTINE_NAME = 'USP_ACTIVO_FILL'
) -- Validaci�n del nombre
BEGIN
    DROP PROC [sw].[USP_ACTIVO_FILL];
END;
GO
/*
sw.USP_ACTIVO_FILL '01',1
*/
CREATE PROCEDURE [sw].[USP_ACTIVO_FILL]
    @codcia CHAR(2),
    @activoId INT
WITH ENCRYPTION
AS
SET NOCOUNT ON;

SELECT a.categoriaId,
       a.proveedorId,
       a.ubicacionId,
       a.responsableId,
       a.numeroSerie,
       a.fechaIngreso
FROM sw.ACTIVO a
WHERE a.codCia = @codcia
      AND a.activoId = @activoId;
GO