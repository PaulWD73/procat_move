USE [ILDB]
GO

SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

-- =====================================================================================
-- Writes the post-release move update CSV into a staging directory, then
-- atomically publishes it to the Builder watched directory.
-- The trailing M on the editset token prevents Builder output-name collisions with
-- the ordinary update CSV for the same editset.
-- =====================================================================================
CREATE OR ALTER PROCEDURE pct.ExportMoveCSVData
    @t_iEditsetID INT,
    @t_Date       VARCHAR(30),
    @t_StagingRoot VARCHAR(255),
    @t_WatchRoot   VARCHAR(255),
    @t_ServerName VARCHAR(255) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;

    DECLARE @sqlCommand VARCHAR(8000);
    DECLARE @publishCommand VARCHAR(8000);
    DECLARE @stagingPath VARCHAR(500);
    DECLARE @finalPath VARCHAR(500);
    DECLARE @server VARCHAR(255) = COALESCE(NULLIF(@t_ServerName, ''), CONVERT(VARCHAR(255), SERVERPROPERTY('ServerName')));
    DECLARE @xpWasEnabled BIT;
    DECLARE @advancedWasEnabled BIT;
    DECLARE @returnCode INT;

    IF NOT EXISTS
    (
        SELECT 1
          FROM dbo.EditorialMoveManifest
         WHERE EditSetID = @t_iEditsetID
           AND MoveStatus = 'RELEASED'
    )
        THROW 51022, 'No RELEASED move-manifest entries were found for this editset.', 1;

    SET @stagingPath = @t_StagingRoot
                     + CASE WHEN RIGHT(@t_StagingRoot, 1) IN ('\', '/') THEN '' ELSE '\' END
                     + @t_Date + '_' + CONVERT(VARCHAR(20), @t_iEditsetID) + 'M_update.partial';

    SET @finalPath = @t_WatchRoot
                   + CASE WHEN RIGHT(@t_WatchRoot, 1) IN ('\', '/') THEN '' ELSE '\' END
                   + @t_Date + '_' + CONVERT(VARCHAR(20), @t_iEditsetID) + 'M_update.csv';

    SET @sqlCommand = 'bcp "EXEC ILDB.dbo.usp_editorial_export_moved_get_all_data '
                    + CONVERT(VARCHAR(20), @t_iEditsetID)
                    + '" queryout ' + QUOTENAME(@stagingPath, '"')
                    + ' -S ' + QUOTENAME(@server, '"') + ' -T -t"#%%%#" -c';

    SET @publishCommand = 'move /Y ' + QUOTENAME(@stagingPath, '"')
                        + ' ' + QUOTENAME(@finalPath, '"');

    SELECT @advancedWasEnabled = CONVERT(BIT, value_in_use)
      FROM sys.configurations
     WHERE name = 'show advanced options';
    SELECT @xpWasEnabled = CONVERT(BIT, value_in_use)
      FROM sys.configurations
     WHERE name = 'xp_cmdshell';

    BEGIN TRY
        IF @advancedWasEnabled = 0
        BEGIN
            EXEC sys.sp_configure 'show advanced options', 1;
            RECONFIGURE;
        END;

        IF @xpWasEnabled = 0
        BEGIN
            EXEC sys.sp_configure 'xp_cmdshell', 1;
            RECONFIGURE;
        END;

        EXEC @returnCode = sys.xp_cmdshell @sqlCommand, NO_OUTPUT;

        IF @returnCode <> 0
            THROW 51023, 'BCP failed while writing the move update CSV.', 1;

        EXEC @returnCode = sys.xp_cmdshell @publishCommand, NO_OUTPUT;

        IF @returnCode <> 0
            THROW 51024, 'The move CSV was written but could not be published to Builder.', 1;

        EXEC pct.ExportMoveManifestToExported @t_iEditsetID;
        EXEC pct.ExportMoveManifestToCompleted @t_iEditsetID;

        IF @xpWasEnabled = 0
        BEGIN
            EXEC sys.sp_configure 'xp_cmdshell', 0;
            RECONFIGURE;
        END;

        IF @advancedWasEnabled = 0
        BEGIN
            EXEC sys.sp_configure 'show advanced options', 0;
            RECONFIGURE;
        END;

        SELECT @finalPath AS MoveCSVFile;
    END TRY
    BEGIN CATCH
        DECLARE @error NVARCHAR(2000) = ERROR_MESSAGE();

        IF @xpWasEnabled = 0
        BEGIN TRY
            EXEC sys.sp_configure 'xp_cmdshell', 0;
            RECONFIGURE;
        END TRY
        BEGIN CATCH
            SET @error = @error + N'; additionally failed to disable xp_cmdshell: ' + ERROR_MESSAGE();
        END CATCH;

        IF @advancedWasEnabled = 0
        BEGIN TRY
            EXEC sys.sp_configure 'show advanced options', 0;
            RECONFIGURE;
        END TRY
        BEGIN CATCH
            SET @error = @error + N'; additionally failed to restore advanced options: ' + ERROR_MESSAGE();
        END CATCH;

        EXEC pct.ExportMoveManifestToFail @t_iEditsetID, @error;
        THROW;
    END CATCH;
END
GO
