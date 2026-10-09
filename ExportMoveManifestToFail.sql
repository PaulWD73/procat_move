USE [ILDB]
GO

SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE OR ALTER PROCEDURE pct.ExportMoveManifestToFail
    @t_iEditsetID INT,
    @t_sError NVARCHAR(2000)
AS
BEGIN
    SET NOCOUNT ON;

    UPDATE dbo.EditorialMoveManifest
       SET StatusBeforeFailure = CASE WHEN MoveStatus = 'FAILED'
                                      THEN StatusBeforeFailure
                                      ELSE MoveStatus END,
           MoveStatus = 'FAILED',
           LastError = @t_sError
     WHERE EditSetID = @t_iEditsetID
       AND MoveStatus <> 'COMPLETED';
END
GO