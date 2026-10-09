USE [ILDB]
GO

SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

/*
    Run once at the very start of the scheduled release job, before capturing
    any current editsets.  Any unfinished row at this point belongs to an older
    run and must be investigated rather than silently removed.
*/
CREATE OR ALTER PROCEDURE pct.ExportMoveManifestCheck
AS
BEGIN
    SET NOCOUNT ON;

    IF EXISTS
    (
        SELECT 1
          FROM dbo.EditorialMoveManifest
         WHERE MoveStatus <> 'COMPLETED'
    )
    BEGIN
        SELECT MoveManifestID, EditSetID, CatalogueLevel, CatalogueID,
               MoveStatus, CapturedAt, ReleaseMarkedAt, ExportedAt,
               LastError, StatusBeforeFailure
          FROM dbo.EditorialMoveManifest
         WHERE MoveStatus <> 'COMPLETED'
         ORDER BY EditSetID, CatalogueLevel, CatalogueID;

        THROW 51002, 'Unfinished move-manifest entries from an earlier release require recovery.', 1;
    END;
END
GO
