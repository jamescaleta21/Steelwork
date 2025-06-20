IF EXISTS
(
    SELECT TOP 1
           s.SPECIFIC_NAME
    FROM INFORMATION_SCHEMA.ROUTINES s
    WHERE s.ROUTINE_TYPE = 'PROCEDURE' -- Validación del tipo
          AND ROUTINE_SCHEMA = 'sw' -- Validación del esquema
          AND s.ROUTINE_NAME = 'USP_MOVIMIENTO_REGISTER'
) -- Validación del nombre
BEGIN
    DROP PROC [sw].[USP_MOVIMIENTO_REGISTER];
END;
GO
/*

*/
CREATE PROCEDURE [sw].[USP_MOVIMIENTO_REGISTER]
    @codcia CHAR(2),
    @activoId INT,
    @fechaMovimiento DATE,
    @responsableOrigenId INT,
    @responsableDestinoId INT,
    @ubicacionDestinoId INT,
    @cuRegistro VARCHAR(20),
    @observacion VARCHAR(300) = NULL
WITH ENCRYPTION
AS
BEGIN
    DECLARE @movimientoId INT;
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

    IF @responsableOrigenId = -1
    BEGIN
        SET @message = 'Debe elegir el Responsable de Origen.';
        SET @code = -2;
        GOTO Terminar;
    END;

    IF @responsableDestinoId = -1
    BEGIN
        SET @message = 'Debe elegir el Responsable de Destino.';
        SET @code = -3;
        GOTO Terminar;
    END;

    IF @ubicacionDestinoId = -1
    BEGIN
        SET @message = 'Debe elegir la Ubicacion de Destino.';
        SET @code = -4;
        GOTO Terminar;
    END;

    IF @responsableDestinoId = @responsableOrigenId
    BEGIN
        SET @message = 'Responsable de Origen debe ser diferente al Responsable de Destino.';
        SET @code = -5;
        GOTO Terminar;
    END;

    IF
    (
        SELECT TOP 1
               a.ubicacionId
        FROM sw.ACTIVO a
        WHERE a.activoId = @activoId
              AND a.codCia = @codcia
    ) = @ubicacionDestinoId
    BEGIN
        SET @message = 'El activo elegido ya se encuentra en la Ubicacion de Destino.';
        SET @code = -6;
        GOTO Terminar;
    END;

    -- Obtener el siguiente activoId
    SELECT @movimientoId = ISNULL(MAX(a.movimientoId), 0) + 1
    FROM [sw].[MOVIMIENTO] a
    WHERE a.codCia = @codcia;

    BEGIN TRAN;
    BEGIN TRY

        INSERT INTO sw.MOVIMIENTO
        (
            codCia,
            movimientoId,
            activoId,
            fechaMovimiento,
            responsableIdOrigen,
            responsableIdDestino,
            ubicacionId,
            observacion,
            cuRegistro
        )
        VALUES
        (   @codcia,               -- codCia - char(2)
            @movimientoId,         -- movimientoId - int
            @activoId,             -- activoId - int
            @fechaMovimiento,      -- fechaMovimiento - date
            @responsableOrigenId,  -- responsableIdOrigen - int
            @responsableDestinoId, -- responsableIdDestino - int
            @ubicacionDestinoId,   -- ubicacionId - int
            @observacion,          -- observacion - varchar(300)
            @cuRegistro            -- cuRegistro - varchar(20)
            );

        UPDATE sw.ACTIVO
        SET responsableId = @responsableDestinoId,
            ubicacionId = @ubicacionDestinoId,
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