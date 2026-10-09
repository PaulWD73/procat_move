USE [ILDB]
GO

SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

/* Call only after the move update CSV has been safely published. */
CREATE OR ALTER PROCEDURE pct.ExportMoveManifestToCompleted
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
          AND MoveStatus <> 'EXPORTED'
          AND MoveStatus <> 'COMPLETED'
    )
        THROW 51013, 'The move manifest cannot be completed from its current state.', 1;

    UPDATE dbo.EditorialMoveManifest
       SET MoveStatus = 'COMPLETED',
           CompletedAt = SYSUTCDATETIME(),
           LastError = NULL
     WHERE EditSetID = @t_iEditsetID
       AND MoveStatus = 'EXPORTED';
END
GO