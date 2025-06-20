IF EXISTS
(
    SELECT TOP 1
           s.SPECIFIC_NAME
    FROM INFORMATION_SCHEMA.ROUTINES s
    WHERE s.ROUTINE_TYPE = 'PROCEDURE' -- Validación del tipo
          AND ROUTINE_SCHEMA = 'sw' -- Validación del esquema
          AND s.ROUTINE_NAME = 'USP_BAJA_REGISTER'
) -- Validación del nombre
BEGIN
    DROP PROC [sw].[USP_BAJA_REGISTER];
END;
GO
/*

*/
CREATE PROCEDURE [sw].[USP_BAJA_REGISTER]
    @codcia CHAR(2),
    @activoId INT,
    @fechaBaja DATE,
    @motivoBaja VARCHAR(200),
    @cuRegistro VARCHAR(20)
WITH ENCRYPTION
AS
BEGIN
    DECLARE @bajaId INT;
    DECLARE @ubicacionId INT;
    DECLARE @code INT,
            @message VARCHAR(300);

    SET @message = 'Registrado Satisfactoriamente.';
    SET @code = 0;

    IF @activoId = -1
    BEGIN
        SET @message = 'Debe elegir el Activo.';
        SET @code = -1;
        GOTO Terminar;
    END;


    -- Obtener el siguiente activoId
    SELECT @bajaId = ISNULL(MAX(a.bajaId), 0) + 1
    FROM [sw].[BAJA] a
    WHERE a.codCia = @codcia;

    BEGIN TRAN;
    BEGIN TRY

        SELECT TOP 1
               @ubicacionId = a.ubicacionId
        FROM sw.ACTIVO a WITH (NOLOCK)
        WHERE a.codCia = @codcia
              AND a.activoId = @activoId;

        INSERT INTO sw.BAJA
        (
            codCia,
            bajaId,
            activoId,
            fechaBaja,
            motivoBaja,
            ubicacionId,
            cuRegistro
        )
        VALUES
        (   @codcia,      -- codCia - char(2)
            @bajaId,      -- bajaId - int
            @activoId,    -- activoId - int
            @fechaBaja,   -- fechaBaja - date
            @motivoBaja,  -- motivoBaja - varchar(200)
            @ubicacionId, -- ubicacionId - int
            @cuRegistro   -- cuRegistro - varchar(20)
            );

        UPDATE sw.ACTIVO
        SET activo = 0,
            cuRegistro = @cuRegistro,
            feRegistro = GETDATE()
        WHERE codCia = @codcia
              AND activoId = @activoId;

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