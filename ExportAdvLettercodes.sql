USE [ILDB]
GO

/****** Object:  StoredProcedure [pct].[ExportAdvLettercodes]    Script Date: 07/09/2026 15:18:56 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO



-- =================================================================
-- Author:		Paul Davis / Chris Bartlett
-- Create date: 18/11/2025
-- Description:	exports lettercodes with coresponding Verity 
--				ADV collections (ADV1:A-E, ADV2:F-Z) 
-- =================================================================
CREATE OR ALTER PROCEDURE [pct].[ExportAdvLettercodes]
	
AS
BEGIN
	
	SET NOCOUNT ON;

	DECLARE @sqlCommand VARCHAR(8000)

	--TEST SQL SERVER
	DECLARE @filePath VARCHAR(255) = '\\na-stor01\procat\PreProd\Verity\ReleaseDataCSV\lettercodes_lookup.txt'

	SET @sqlCommand = 'bcp "EXEC ILDB.pct.ExportAdvLettercodes" queryout ' + QUOTENAME(@filePath, '"') + ' -S na-t-sqlc03v02\sql2 -T -c';


	/*--LIVE SQL SERVER
	DECLARE @filePath VARCHAR(255) = '\\na-stor01\procat\Prod\Verity\ReleaseDataCSV\lettercodes_lookup.txt'

	SET @sqlCommand = 'bcp "EXEC ILDB.pct.ExportAdvLettercodes" queryout ' + QUOTENAME(@filePath, '"') + ' -S na-sqlc03v01\sql1 -T -c';
	*/

	EXEC sp_configure 'show advanced options', '1'
	RECONFIGURE
-- this enables xp_cmdshell
	EXEC sp_configure 'xp_cmdshell', '1' 
	RECONFIGURE 

    -- Execute the BCP command
    EXEC xp_cmdshell @sqlCommand

	EXEC sp_configure 'show advanced options', '1'
	RECONFIGURE
-- this disables xp_cmdshell
	EXEC sp_configure 'xp_cmdshell', '0' 
	RECONFIGURE

END
GO

