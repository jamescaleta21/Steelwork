IF EXISTS
(
    SELECT TOP 1
           s.SPECIFIC_NAME
    FROM INFORMATION_SCHEMA.ROUTINES s
    WHERE s.ROUTINE_TYPE = 'PROCEDURE' -- Validación del tipo
          AND ROUTINE_SCHEMA = 'sw' -- Validación del esquema
          AND s.ROUTINE_NAME = 'USP_CATEGORIA_REGISTER' -- Validación del nombre
)
BEGIN
    DROP PROC [sw].[USP_CATEGORIA_REGISTER];
END;
GO
/*

*/
CREATE PROCEDURE [sw].[USP_CATEGORIA_REGISTER]
(
    @codCia CHAR(2),
    @descripcion VARCHAR(100),
    @cuRegistro VARCHAR(20)
)
WITH ENCRYPTION
AS
BEGIN
    SET NOCOUNT ON;

    DECLARE @categoriaId INT;

    DECLARE @code INT,
            @message VARCHAR(300);

    SET @message = 'Registrado Satisfactoriamente.';
    SET @code = 0;

    IF EXISTS
    (
        SELECT TOP 1
               'X'
        FROM sw.CATEGORIA c
        WHERE c.codCia = @codCia
              AND c.descripcion = @descripcion and isnull(c.eliminado,'') = 0
    )
    BEGIN
        SET @message = 'Descripción ingresada existente.';
        SET @code = -1;
        GOTO Terminar;
    END;

    -- Obtener el siguiente activoId
    SELECT @categoriaId = ISNULL(MAX(c.categoriaId), 0) + 1
    FROM [sw].[CATEGORIA] c
    WHERE c.codCia = @codCia;

    --BEGIN TRAN;
    BEGIN TRY

        INSERT INTO sw.CATEGORIA
        (
            codCia,
            categoriaId,
            descripcion,
            cuRegistro
        )
        VALUES
        (@codCia, @categoriaId, @descripcion, @cuRegistro);


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