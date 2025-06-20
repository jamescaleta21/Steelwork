IF EXISTS
(
    SELECT TOP 1
           s.SPECIFIC_NAME
    FROM INFORMATION_SCHEMA.ROUTINES s
    WHERE s.ROUTINE_TYPE = 'PROCEDURE' -- Validación del tipo
          AND ROUTINE_SCHEMA = 'sw' -- Validación del esquema
          AND s.ROUTINE_NAME = 'USP_RESPONSABLE_STATUS' -- Validación del nombre
)
BEGIN
    DROP PROC [sw].[USP_RESPONSABLE_STATUS];
END;
GO
/*
SW.USP_RESPONSABLE_STATUS '01',1,0,'WOFLS'
*/
CREATE PROCEDURE [sw].[USP_RESPONSABLE_STATUS]
(
    @codCia CHAR(2),
    @responsableId INT,
    @activo BIT,
    @cuRegistro VARCHAR(20)
)
WITH ENCRYPTION
AS
BEGIN
    SET NOCOUNT ON;


    DECLARE @code INT,
            @message VARCHAR(300);

    IF @activo = 1
        SET @message = 'Habilitado Satisfactoriamente.';
    ELSE
        SET @message = 'Desabilitado Satisfactoriamente.';

    SET @code = 0;

    IF EXISTS
    (
        SELECT TOP 1
               'X'
        FROM sw.RESPONSABLE p
        WHERE p.codCia = @codCia
              AND p.responsableId = @responsableId
              AND p.activo = 1
    )
       AND (@activo = 1)
    BEGIN
        SET @message = 'Responsable ya se encuentra Habilitado.';
        SET @code = -1;
        GOTO Terminar;
    END;

    IF EXISTS
    (
        SELECT TOP 1
               'X'
        FROM sw.RESPONSABLE p
        WHERE p.codCia = @codCia
              AND p.responsableId = @responsableId
              AND p.activo = 0
    )
       AND (@activo = 0)
    BEGIN
        SET @message = 'Responsable ya se encuentra Desabilitado.';
        SET @code = -2;
        GOTO Terminar;
    END;


    --BEGIN TRAN;
    BEGIN TRY

        UPDATE sw.RESPONSABLE
        SET activo = @activo,
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