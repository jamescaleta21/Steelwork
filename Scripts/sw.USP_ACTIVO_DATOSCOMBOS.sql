IF EXISTS ( SELECT TOP 1
                    S.SPECIFIC_NAME
            FROM    information_schema.routines s
            WHERE   s.ROUTINE_TYPE = 'PROCEDURE'	-- Validación del tipo
            		AND ROUTINE_SCHEMA = 'sw'		-- Validación del esquema
                    AND S.ROUTINE_NAME = 'USP_ACTIVO_DATOSCOMBOS' )		-- Validación del nombre
    BEGIN
        DROP PROC [sw].[USP_ACTIVO_DATOSCOMBOS]
    END
GO
/*
sw.USP_ACTIVO_DATOSCOMBOS '01'
*/
CREATE PROCEDURE [sw].[USP_ACTIVO_DATOSCOMBOS]
@codcia CHAR(2)
WITH ENCRYPTION
AS
SET NOCOUNT ON

SELECT -1 AS categoriaId,'.: SELECCIONE :.' AS descripcion
UNION
SELECT c.categoriaId,c.descripcion
FROM sw.CATEGORIA c
WHERE c.activo = 1 AND c.eliminado = 0 AND c.codCia = @codcia
ORDER BY 2;

SELECT -1 AS proveedorId,'.: SELECCIONE :.' AS razonSocial
UNION
SELECT p.proveedorId, p.razonSocial FROM sw.PROVEEDOR p
WHERE p.activo = 1 AND p.eliminado = 0 AND p.codCia = @codcia
ORDER BY 2;

SELECT -1 AS proveedorId,'.: SELECCIONE :.' AS responsable
UNION
SELECT r.responsableId,r.apellidos + ' ' + r.nombres AS responsable
FROM sw.RESPONSABLE r
WHERE r.activo = 1 AND r.eliminado = 0 AND r.codCia = @codcia
ORDER BY 2;

SELECT -1 AS ubicacionId,'.: SELECCIONE :.' AS denominacion
UNION
SELECT u.ubicacionId,u.denominacion FROM sw.UBICACION u
WHERE u.activo = 1 AND u.eliminado = 0 AND u.codCia = @codcia
GO