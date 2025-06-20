IF EXISTS
(
    SELECT TOP 1
           s.SPECIFIC_NAME
    FROM INFORMATION_SCHEMA.ROUTINES s
    WHERE s.ROUTINE_TYPE = 'PROCEDURE' -- Validación del tipo
          AND ROUTINE_SCHEMA = 'sw' -- Validación del esquema
          AND s.ROUTINE_NAME = 'USP_CATEGORIA_UPDATE' -- Validación del nombre
)
BEGIN
    DROP PROC [sw].[USP_CATEGORIA_UPDATE];
END;
GO
/*

*/
CREATE PROCEDURE [sw].[USP_CATEGORIA_UPDATE]
(
    @codCia CHAR(2),
    @categoriaId INT,
    @descripcion VARCHAR(100),
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
        FROM sw.ACTIVO a
        WHERE a.codCia = @codCia
              AND a.descripcion = @descripcion
              AND a.categoriaId <> @categoriaId
    )
    BEGIN
        SET @message = 'Descripción ingresada existente.';
        SET @code = -1;
        GOTO Terminar;
    END;


    --BEGIN TRAN;
    BEGIN TRY


        UPDATE sw.CATEGORIA
        SET descripcion = @descripcion,
            cuRegistro = @cuRegistro,
			feRegistro = GETDATE()
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