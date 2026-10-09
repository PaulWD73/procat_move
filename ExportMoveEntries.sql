USE [ILDB]
GO

/****** Object:  StoredProcedure [pct].[ExportMoveEntries]    Script Date: 09/10/2026 15:48:00 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


-- =====================================================================================
-- Author:      Paul Davis / Chris Bartlett
-- Create date: 30/09/2025
-- Description: Pre-release capture of moved class and piece roots.
--
-- This procedure must run before sp_edit_release_editset, while the ES tables still
-- contain the proposed parents.  Descendant export happens after release and reads
-- the permanent manifest; it is deliberately not performed here.
-- =====================================================================================
CREATE OR ALTER   PROCEDURE [pct].[ExportMoveEntries]
    @t_iEditsetID INT
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;

    DECLARE @LockResult INT;

    BEGIN TRANSACTION;

    EXEC @LockResult = sys.sp_getapplock
        @Resource = N'PROCAT_MOVE_MANIFEST',
        @LockMode = N'Exclusive',
        @LockOwner = N'Transaction',
        @LockTimeout = 0;

    IF @LockResult < 0
    BEGIN
        ROLLBACK TRANSACTION;
        THROW 51000, 'The move manifest is already in use by another release.', 1;
    END;

    /*
       Do not duplicate or overwrite an unfinished capture for this editset.
       The scheduler performs the global stale-manifest check once, before it
       starts capturing any of the ready editsets in the current release run.
    */
    IF EXISTS
    (
        SELECT 1
        FROM dbo.EditorialMoveManifest WITH (UPDLOCK, HOLDLOCK)
        WHERE EditSetID = @t_iEditsetID
          AND MoveStatus <> 'COMPLETED'
    )
    BEGIN
        ROLLBACK TRANSACTION;
        THROW 51001, 'Unfinished entries already exist in the move manifest for this editset.', 1;
    END;

    /* Classes moved between divisions. */
    INSERT dbo.EditorialMoveManifest
        (EditSetID, CatalogueLevel, CatalogueID, MoveStatus)
    SELECT @t_iEditsetID, 3, ec.class_id, 'CAPTURED'
      FROM dbo.ES_class AS ec
      JOIN dbo.tbl_class AS c ON c.class_id = ec.class_id
     WHERE ec.edit_set_id = @t_iEditsetID
       AND COALESCE(ec.division_id, 0) <> COALESCE(c.division_id, 0);

    /* Pieces moved between a header/subheader position. */
    INSERT dbo.EditorialMoveManifest
        (EditSetID, CatalogueLevel, CatalogueID, MoveStatus)
    SELECT @t_iEditsetID, 6, ep.piece_id, 'CAPTURED'
      FROM dbo.ES_piece AS ep
      JOIN dbo.tbl_piece AS p ON p.piece_id = ep.piece_id
     WHERE ep.edit_set_id = @t_iEditsetID
       AND
       (
           COALESCE(ep.header_id, 0) <> COALESCE(p.header_id, 0)
           OR COALESCE(ep.subheader_id, 0) <> COALESCE(p.subheader_id, 0)
       );

    COMMIT TRANSACTION;

    SELECT MoveManifestID, EditSetID, CatalogueLevel, CatalogueID,
           MoveStatus, CapturedAt
      FROM dbo.EditorialMoveManifest
     WHERE EditSetID = @t_iEditsetID
       AND MoveStatus = 'CAPTURED'
     ORDER BY CatalogueLevel, CatalogueID;
END

GO

