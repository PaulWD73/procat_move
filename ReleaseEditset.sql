USE [ILDB]
GO

/****** Object:  StoredProcedure [pct].[ReleaseEditset]    Script Date: 07/09/2026 15:26:25 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO




CREATE OR ALTER PROCEDURE [pct].[ReleaseEditset](
    @t_iEditsetID INTEGER
)
AS 
--===============================================================================================================
 --Description:
 --   Master sp for the editset release to live process
    
 -- Amendment history:

-- d1 WDP     12-Jan-1999     Added from individual files
--===============================================================================================================
BEGIN
    DECLARE @iReadOnlyID INTEGER
    DECLARE @iLiveFlagValue INTEGER
    DECLARE @iLiveFlagValueCreated INTEGER
    SELECT @iReadOnlyID = 7
    SELECT @iLiveFlagValue = 1
    SELECT @iLiveFlagValueCreated = 2
    PRINT 'Release editset:  AddAuditTrailEditSetDetails ' + CAST(@t_iEditSetID AS NVARCHAR(10))
    EXECUTE pct.AddAuditTrailEditSetDetails @t_iEditsetID, @iReadOnlyID, @iLiveFlagValue, @iLiveFlagValueCreated
    PRINT 'Release editset:  ReleaseLettercodes ' + CAST(@t_iEditSetID AS NVARCHAR(10))
    EXECUTE pct.ReleaseLettercodes @t_iEditsetID
    EXECUTE pct.ReleaseDivisions @t_iEditsetID
    EXECUTE pct.ReleaseClasses @t_iEditsetID
    EXECUTE pct.ReleaseHeaders @t_iEditsetID
    EXECUTE pct.ReleaseSubHeaders @t_iEditsetID
    EXECUTE pct.ReleasePieces @t_iEditsetID
    EXECUTE pct.ReleaseItems @t_iEditsetID
    

END
GO

