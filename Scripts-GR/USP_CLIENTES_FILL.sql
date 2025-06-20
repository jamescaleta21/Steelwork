USE [BDATOS]
GO
/****** Object:  StoredProcedure [dbo].[USP_CLIENTES_FILL]    Script Date: 06/19/2024 23:46:15 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/*
USP_CLIENTES_FILL 'DI'
*/
ALTER PROC [dbo].[USP_CLIENTES_FILL] @SEARCH VARCHAR(80), @codcia char(2) = null
AS
    SET NOCOUNT ON

    SELECT  C.CLI_CODCLIE AS 'CODIGO' ,
            C.CLI_NOMBRE AS 'CLIENTE'
    FROM    dbo.CLIENTES c
    WHERE   C.CLI_CP = 'C'
            AND( C.CLI_NOMBRE LIKE @SEARCH + '%'
            OR @SEARCH IS NULL) and c.CLI_CODCIA = @codcia
            ORDER BY C.CLI_NOMBRE

