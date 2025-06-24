IF EXISTS
(
    SELECT TOP 1
           s.SPECIFIC_NAME
    FROM INFORMATION_SCHEMA.ROUTINES s
    WHERE s.ROUTINE_TYPE = 'PROCEDURE' -- Validación del tipo
          AND ROUTINE_SCHEMA = 'sw' -- Validación del esquema
          AND s.ROUTINE_NAME = 'USP_RESPONSABLE_REGISTER' -- Validación del nombre
)
BEGIN
    DROP PROC [sw].[USP_RESPONSABLE_REGISTER];
END;
GO
/*

*/
CREATE PROCEDURE [sw].[USP_RESPONSABLE_REGISTER]
(
    @codCia CHAR(2),
    @nombres VARCHAR(100),
    @apellidos VARCHAR(200),
    @cuRegistro VARCHAR(20)
)
WITH ENCRYPTION
AS
BEGIN
    SET NOCOUNT ON;

    DECLARE @responsableId INT;

    DECLARE @code INT,
            @message VARCHAR(300);

    SET @message = 'Registrado Satisfactoriamente.';
    SET @code = 0;

    IF EXISTS
    (
        SELECT TOP 1
               'X'
        FROM sw.RESPONSABLE p
        WHERE p.codCia = @codCia
              AND p.nombres = @nombres
              AND p.apellidos = @apellidos
			  AND COALESCE(p.eliminado,0) = 0
    )
    BEGIN
        SET @message = 'Nombres y Apellidos ingresados existentes.';
        SET @code = -1;
        GOTO Terminar;
    END;

    -- Obtener el siguiente activoId
    SELECT @responsableId = ISNULL(MAX(p.responsableId), 0) + 1
    FROM [sw].[RESPONSABLE] p
    WHERE p.codCia = @codCia;

    --BEGIN TRAN;
    BEGIN TRY

        INSERT INTO sw.RESPONSABLE
        (
            codCia,
            responsableId,
            nombres,
            apellidos,
            cuRegistro
        )
        VALUES
        (@codCia, @responsableId, @nombres, @apellidos, @cuRegistro);


    END TRY
    BEGIN CATCH
        SET @code = ERROR_NUMBER();
        SET @message = ERROR_MESSAGE();
        --OLLBACK TRAN;
        GOTO Terminar;
    END CATCH;

    --IF @@TRANCOUNT > 0
    --    COMMIT;

    Terminar:
    SELECT @code AS 'code',
           @message AS 'message';

END;
GO