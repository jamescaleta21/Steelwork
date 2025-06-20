IF EXISTS
(
    SELECT TOP 1
           s.SPECIFIC_NAME
    FROM INFORMATION_SCHEMA.ROUTINES s
    WHERE s.ROUTINE_TYPE = 'PROCEDURE' -- Validación del tipo
          AND ROUTINE_SCHEMA = 'sw' -- Validación del esquema
          AND s.ROUTINE_NAME = 'USP_ACTIVO_REGISTER' -- Validación del nombre
)
BEGIN
    DROP PROC [sw].[USP_ACTIVO_REGISTER];
END;
GO
/*

*/
CREATE PROCEDURE [sw].[USP_ACTIVO_REGISTER]
(
    @codCia CHAR(2),
    @codigoActivo VARCHAR(20),
    @descripcion VARCHAR(100),
    @numeroSerie VARCHAR(50) = NULL,
    @categoriaId INT,
    @proveedorId INT,
    @costoInicial MONEY,
    @fechaIngreso DATE = NULL,
    @responsableId INT,
    @ubicacionId INT,
    @cuRegistro VARCHAR(20)
)
WITH ENCRYPTION
AS
BEGIN
    SET NOCOUNT ON;

    DECLARE @activoId INT;

    DECLARE @code INT,
            @message VARCHAR(300);

    SET @message = 'Registrado Satisfactoriamente.';
    SET @code = 0;

    IF LEN(RTRIM(LTRIM(@codigoActivo))) = 0
    BEGIN
        SET @message = 'Debe ingresar el Codigo.';
        SET @code = -1;
        GOTO Terminar;
    END;

    IF LEN(RTRIM(LTRIM(@descripcion))) = 0
    BEGIN
        SET @message = 'Debe ingresar la Descripcion.';
        SET @code = -2;
        GOTO Terminar;
    END;

    IF LEN(RTRIM(LTRIM(@numeroSerie))) = 0
    BEGIN
        SET @message = 'Debe ingresar el Numero de Serie.';
        SET @code = -3;
        GOTO Terminar;
    END;

    IF EXISTS
    (
        SELECT TOP 1
               'X'
        FROM sw.ACTIVO a
        WHERE a.codCia = @codCia
              AND a.codigoActivo = @codigoActivo
    )
    BEGIN
        SET @message = 'Codigo ingresado existente.';
        SET @code = -4;
        GOTO Terminar;
    END;

    IF EXISTS
    (
        SELECT TOP 1
               'X'
        FROM sw.ACTIVO a
        WHERE a.codCia = @codCia
              AND a.descripcion = @descripcion
    )
    BEGIN
        SET @message = 'Descripcion ingresada existente.';
        SET @code = -5;
        GOTO Terminar;
    END;

    -- Obtener el siguiente activoId
    SELECT @activoId = ISNULL(MAX(a.activoId), 0) + 1
    FROM [sw].[ACTIVO] a
    WHERE a.codCia = @codCia;

    BEGIN TRAN;
    BEGIN TRY

        INSERT INTO [sw].[ACTIVO]
        (
            codCia,
            activoId,
            codigoActivo,
            descripcion,
            numeroSerie,
            categoriaId,
            proveedorId,
            costoInicial,
            fechaIngreso,
            responsableId,
            ubicacionId,
            cuRegistro
        )
        VALUES
        (@codCia, @activoId, @codigoActivo, @descripcion, @numeroSerie, @categoriaId, @proveedorId, @costoInicial,
         @fechaIngreso, @responsableId, @ubicacionId, @cuRegistro);


    END TRY
    BEGIN CATCH
        SET @code = ERROR_NUMBER();
        SET @message = ERROR_MESSAGE();
        ROLLBACK TRAN;
        GOTO Terminar;
    END CATCH;

    IF @@TRANCOUNT > 0
        COMMIT;

    Terminar:
    SELECT @code AS 'code',
           @message AS 'message';

END;
GO