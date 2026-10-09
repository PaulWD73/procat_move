USE [ILDB]
GO

SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

/* Call after the move update CSV has been written, before it is published. */
CREATE OR ALTER PROCEDURE pct.ExportMoveManifestToExported
    @t_iEditsetID INT
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;

    IF EXISTS
    (
        SELECT 1
        FROM dbo.EditorialMoveManifest
        WHERE EditSetID = @t_iEditsetID
          AND MoveStatus IN ('CAPTURED', 'FAILED')
    )
        THROW 51015, 'The move manifest cannot be marked exported from its current state.', 1;

    UPDATE dbo.EditorialMoveManifest
       SET MoveStatus = 'EXPORTED',
           ExportedAt = SYSUTCDATETIME(),
           LastError = NULL
     WHERE EditSetID = @t_iEditsetID
       AND MoveStatus = 'RELEASED';
END
GO
