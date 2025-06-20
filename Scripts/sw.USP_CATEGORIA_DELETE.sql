IF EXISTS
(
    SELECT TOP 1
           s.SPECIFIC_NAME
    FROM INFORMATION_SCHEMA.ROUTINES s
    WHERE s.ROUTINE_TYPE = 'PROCEDURE' -- Validación del tipo
          AND ROUTINE_SCHEMA = 'sw' -- Validación del esquema
          AND s.ROUTINE_NAME = 'USP_CATEGORIA_DELETE' -- Validación del nombre
)
BEGIN
    DROP PROC [sw].[USP_CATEGORIA_DELETE];
END;
GO
/*
SW.USP_CATEGORIA_DELETE '01',1,'PEPITO'
*/
CREATE PROCEDURE [sw].[USP_CATEGORIA_DELETE]
(
    @codCia CHAR(2),
    @categoriaId INT,
    @cuRegistro VARCHAR(20)
)
WITH ENCRYPTION
AS
BEGIN
    SET NOCOUNT ON;


    DECLARE @code INT,
            @message VARCHAR(300);

    SET @message = 'Eliminado Satisfactoriamente.';
    SET @code = 0;



    --BEGIN TRAN;
    BEGIN TRY

        UPDATE sw.CATEGORIA
        SET eliminado = 1,
            cuEliminado = @cuRegistro,
            feEliminado = GETDATE()
        WHERE codCia = @codCia
              AND categoriaId = @categoriaId;

    END TRY
    BEGIN CATCH
        SET @code = ERROR_NUMBER();
        SET @message = ERROR_MESSAGE();
        --ROLLBACK TRAN;
        GOTO Terminar;
    END CATCH;

    --IF @@TRANCOUNT > 0
    --    COMMIT;

    Terminar:
    SELECT @code AS 'code',
           @message AS 'message';

END;
GO