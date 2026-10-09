USE [ILDB];
GO

SET ANSI_NULLS ON;
GO

SET QUOTED_IDENTIFIER ON;
GO

CREATE OR ALTER PROCEDURE [pct].[ExportCSVData]
    @t_iEditsetID        INT,
    @t_Editset_operation INT,
    @t_Date              VARCHAR(30)
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;

    DECLARE @sqlCommand VARCHAR(8000);
    DECLARE @publishCommand VARCHAR(8000);
    DECLARE @stagingRoot VARCHAR(255) = '\\na-stor01\procat\PreProd\Verity\ReleaseDataCSV\staging';
    DECLARE @watchRoot   VARCHAR(255) = '\\na-stor01\procat\PreProd\Verity\ReleaseDataCSV\data';
    DECLARE @stagingPath VARCHAR(500);
    DECLARE @finalPath   VARCHAR(500);
    DECLARE @operation   VARCHAR(10);
    DECLARE @returnCode  INT;

    SET @operation =
        CASE @t_Editset_operation
            WHEN 0 THEN 'insert'
            WHEN 1 THEN 'update'
            ELSE 'delete'
        END;

    SET @stagingPath =
        @stagingRoot + '\' + @t_Date + '_'
        + CONVERT(VARCHAR(20), @t_iEditsetID)
        + '_' + @operation + '.partial';

    SET @finalPath =
        @watchRoot + '\' + @t_Date + '_'
        + CONVERT(VARCHAR(20), @t_iEditsetID)
        + '_' + @operation + '.csv';

    SET @sqlCommand =
        'bcp "EXEC ILDB.pct.ExportGetAllData '
        + CONVERT(VARCHAR(20), @t_iEditsetID)
        + ', '
        + CONVERT(VARCHAR(10), @t_Editset_operation)
        + ', '''
        + REPLACE(@t_Date, '''', '''''')
        + '''" queryout '
        + QUOTENAME(@stagingPath, '"')
        + ' -S '
        + CONVERT(VARCHAR(255), SERVERPROPERTY('ServerName'))
        + ' -T -t"#%%%#" -c';

    SET @publishCommand =
        'move /Y '
        + QUOTENAME(@stagingPath, '"')
        + ' '
        + QUOTENAME(@finalPath, '"');

    BEGIN TRY

        PRINT 'BCP Command:';
        PRINT @sqlCommand;

        EXEC @returnCode = sys.xp_cmdshell @sqlCommand;

        PRINT 'BCP Return Code = '
            + CAST(ISNULL(@returnCode, -1) AS VARCHAR(10));

        IF @returnCode <> 0
        BEGIN
            THROW 51030,
                  'BCP failed while writing the ordinary editset CSV.',
                  1;
        END;

        PRINT 'Publish Command:';
        PRINT @publishCommand;

        EXEC @returnCode = sys.xp_cmdshell @publishCommand;

        PRINT 'MOVE Return Code = '
            + CAST(ISNULL(@returnCode, -1) AS VARCHAR(10));

        IF @returnCode <> 0
        BEGIN
            THROW 51031,
                  'The ordinary editset CSV was written but could not be published to Builder.',
                  1;
        END;

        SELECT @finalPath AS CSVFile;

    END TRY
    BEGIN CATCH

        DECLARE @ErrorMessage NVARCHAR(4000) = ERROR_MESSAGE();

        RAISERROR(
            'ExportCSVData failed. %s',
            16,
            1,
            @ErrorMessage
        );

    END CATCH;
END;
GO