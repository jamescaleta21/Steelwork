IF EXISTS
(
    SELECT TOP 1
           s.SPECIFIC_NAME
    FROM INFORMATION_SCHEMA.ROUTINES s
    WHERE s.ROUTINE_TYPE = 'PROCEDURE' -- Validación del tipo
          AND ROUTINE_SCHEMA = 'sw' -- Validación del esquema
          AND s.ROUTINE_NAME = 'USP_PROVEEDOR_UPDATE' -- Validación del nombre
)
BEGIN
    DROP PROC [sw].[USP_PROVEEDOR_UPDATE];
END;
GO
/*

*/
CREATE PROCEDURE [sw].[USP_PROVEEDOR_UPDATE]
(
    @codCia CHAR(2),
    @proveedorId INT,
    @razonSocial VARCHAR(100),
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
        FROM sw.PROVEEDOR p
        WHERE p.codCia = @codCia
              AND p.razonSocial = @razonSocial
              AND p.proveedorId <> @proveedorId
    )
    BEGIN
        SET @message = 'Razón Social ingresada existente.';
        SET @code = -1;
        GOTO Terminar;
    END;


    --BEGIN TRAN;
    BEGIN TRY

        UPDATE sw.PROVEEDOR
        SET razonSocial = @razonSocial,
            cuRegistro = @cuRegistro,
            feRegistro = GETDATE()
        WHERE codCia = @codCia
              AND proveedorId = @proveedorId;

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