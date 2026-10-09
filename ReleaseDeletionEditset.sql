USE [ILDB]
GO

/****** Object:  StoredProcedure [pct].[ReleaseDeletionEditset]    Script Date: 07/09/2026 15:25:42 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO




CREATE OR ALTER PROCEDURE [pct].[ReleaseDeletionEditset]
    @t_iEditSetID   INTEGER
AS
--===============================================================================================================
-- File name:   ReleaseDeletionEditset
-- Author:		Paul Davis / Chris Bartlett
-- Create date: 14/11/2024
-- Description: Parent stored procedure for processing a deletion edit set. Deletions are only for catalogue 
-- levels Item, subsubseries, subseries, division. Department, Series, Pieces cannot be deleted.

-- Parameters: @t_iEditSetID - ID of the edit set to process                                      
--===============================================================================================================
BEGIN
	
	DECLARE @iReadOnlyID INTEGER
    DECLARE @iLiveFlagValue INTEGER
    DECLARE @iLiveFlagValueCreated INTEGER
	SELECT @iReadOnlyID = 7
    SELECT @iLiveFlagValue = 1
    SELECT @iLiveFlagValueCreated = 2
    PRINT 'Release editset: AddAuditTrailEditSetDetails ' + CAST(@t_iEditSetID AS NVARCHAR(10))
    EXECUTE pct.AddAuditTrailEditSetDetails @t_iEditsetID, @iReadOnlyID, @iLiveFlagValue, @iLiveFlagValueCreated
    
    -- Get Items to delete and send them to delete sp    
    EXEC pct.DeleteLiveItem @t_iEditsetID
       
    -- Get SubSubSeries to delete and send them to delete sp
    EXEC pct.DeleteLiveSubHeader @t_iEditsetID
      
    -- Get SubSeries to delete and send them to delete sp
    EXEC pct.DeleteLiveHeader @t_iEditsetID
        
    -- Get Divisions to delete and send them to delete sp
    EXEC pct.DeleteLiveDivision @t_iEditsetID
  
    -- Finish by deleting the edit set details
    EXEC pct.DeleteEditset @t_iEditsetID
    
END
GO

