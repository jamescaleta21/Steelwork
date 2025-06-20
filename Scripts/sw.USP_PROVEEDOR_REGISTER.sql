IF EXISTS
(
    SELECT TOP 1
           s.SPECIFIC_NAME
    FROM INFORMATION_SCHEMA.ROUTINES s
    WHERE s.ROUTINE_TYPE = 'PROCEDURE' -- Validación del tipo
          AND ROUTINE_SCHEMA = 'sw' -- Validación del esquema
          AND s.ROUTINE_NAME = 'USP_PROVEEDOR_REGISTER' -- Validación del nombre
)
BEGIN
    DROP PROC [sw].[USP_PROVEEDOR_REGISTER];
END;
GO
/*

*/
CREATE PROCEDURE [sw].[USP_PROVEEDOR_REGISTER]
(
    @codCia CHAR(2),
    @razonSocial VARCHAR(100),
    @cuRegistro VARCHAR(20)
)
WITH ENCRYPTION
AS
BEGIN
    SET NOCOUNT ON;

    DECLARE @proveedorId INT;

    DECLARE @code INT,
            @message VARCHAR(300);

    SET @message = 'Registrado Satisfactoriamente.';
    SET @code = 0;

    IF EXISTS
    (
        SELECT TOP 1
               'X'
        FROM sw.PROVEEDOR p
        WHERE p.codCia = @codCia
              AND p.razonSocial = @razonSocial
    )
    BEGIN
        SET @message = 'Razón Social ingresada existente.';
        SET @code = -1;
        GOTO Terminar;
    END;

    -- Obtener el siguiente activoId
    SELECT @proveedorId = ISNULL(MAX(p.proveedorId), 0) + 1
    FROM [sw].[PROVEEDOR] p
    WHERE p.codCia = @codCia;

    --BEGIN TRAN;
    BEGIN TRY

        INSERT INTO sw.PROVEEDOR
        (
            codCia,
            proveedorId,
            razonSocial,
            cuRegistro
        )
        VALUES
        (@codCia, @proveedorId, @razonSocial, @cuRegistro);


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