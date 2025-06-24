IF EXISTS
(
    SELECT TOP 1
           s.SPECIFIC_NAME
    FROM INFORMATION_SCHEMA.ROUTINES s
    WHERE s.ROUTINE_TYPE = 'PROCEDURE' -- Validación del tipo
          AND ROUTINE_SCHEMA = 'sw' -- Validación del esquema
          AND s.ROUTINE_NAME = 'USP_RESPONSABLE_UPDATE' -- Validación del nombre
)
BEGIN
    DROP PROC [sw].[USP_RESPONSABLE_UPDATE];
END;
GO
/*

*/
CREATE PROCEDURE [sw].[USP_RESPONSABLE_UPDATE]
(
    @codCia CHAR(2),
    @responsableId INT,
    @nombres VARCHAR(100),
    @apellidos VARCHAR(200),
    @cuRegistro VARCHAR(20)
)
WITH ENCRYPTION
AS
BEGIN
    SET NOCOUNT ON;


    DECLARE @code INT,
            @message VARCHAR(300);

    SET @message = 'Actualizado Satisfactoriamente.';
    SET @code = 0;

    IF EXISTS
    (
        SELECT TOP 1
               'X'
        FROM sw.RESPONSABLE p
        WHERE p.codCia = @codCia
              AND p.nombres = @nombres
              AND p.apellidos = @apellidos
              AND COALESCE(p.eliminado, 0) = 0
              AND p.responsableId <> @responsableId
    )
    BEGIN
        SET @message = 'Nombres y Apellidos ingresados existentes.';
        SET @code = -1;
        GOTO Terminar;
    END;


    --BEGIN TRAN;
    BEGIN TRY

        UPDATE sw.RESPONSABLE
        SET nombres = @nombres,
            apellidos = @apellidos,
            cuRegistro = @cuRegistro,
            feRegistro = GETDATE()
        WHERE codCia = @codCia
              AND responsableId = @responsableId;

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