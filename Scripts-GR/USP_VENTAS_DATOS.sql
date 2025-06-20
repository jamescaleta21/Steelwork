IF EXISTS ( SELECT TOP 1
                    S.SPECIFIC_NAME
            FROM    information_schema.routines s
            WHERE   s.ROUTINE_TYPE = 'PROCEDURE'
                    AND S.ROUTINE_NAME = 'USP_VENTAS_DATOS' )
    BEGIN
        DROP PROC [dbo].[USP_VENTAS_DATOS]
    END
GO
/*
USP_VENTAS_DATOS '01' ,'F-001-2727'
USP_VENTAS_DATOS '01' ,'001-2727'
*/
CREATE PROCEDURE [dbo].[USP_VENTAS_DATOS]
@CODCIA CHAR(2),
@VALORES VARCHAR(MAX)
WITH ENCRYPTION
AS
SET NOCOUNT ON

DECLARE @TBLDATOS TABLE(TIPO CHAR(1),SERIE CHAR(3), NUMERO BIGINT)

INSERT INTO @TBLDATOS
(
TIPO,
    SERIE,
    NUMERO
)
SELECT * FROM DBO.UFN_CONVERT_TEXT_TABLE(@VALORES,',') uctt



SELECT a.ART_KEY AS 'IDPRODUCTO', f.FAR_CANTIDAD AS 'CANTIDAD',
       a.ART_NOMBRE AS 'PRODUCTO',
       p.PRE_PESO AS 'PESO',
       f.FAR_CANTIDAD * p.PRE_PESO AS 'PESOTOTAL'
	   FROM dbo.FACART f WITH (NOLOCK)
INNER JOIN @TBLDATOS t ON f.FAR_FBG = T.TIPO 
  INNER JOIN dbo.ARTI a WITH (NOLOCK)
        ON f.FAR_CODCIA = a.ART_CODCIA
           AND f.FAR_CODART = a.ART_KEY
    INNER JOIN dbo.PRECIOS p WITH (NOLOCK)
        ON a.ART_CODCIA = p.PRE_CODCIA
           AND a.ART_KEY = p.PRE_CODART
           AND p.PRE_FLAG_UNIDAD = 'A'
AND RIGHT('000' + RTRIM(LTRIM(f.FAR_NUMSER)),3) = t.SERIE 
AND f.FAR_NUMFAC = t.NUMERO
WHERE f.FAR_CODCIA= @CODCIA
GO