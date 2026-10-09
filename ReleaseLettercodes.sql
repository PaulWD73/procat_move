USE [ILDB]
GO

/****** Object:  StoredProcedure [pct].[ReleaseLettercodes]    Script Date: 07/09/2026 15:37:11 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO




CREATE OR ALTER PROCEDURE [pct].[ReleaseLettercodes](
    @t_iEditsetID INTEGER
)
AS 
--===============================================================================================================
-- Description:
--    This SP finds all modified or new lettercodes in an
--    edit set and updates the corresponding live tables
    
--  Amendment history:

-- d1 WDP     12-Jan-1999     Added from individual files
-- d10 DAJ    07-Mar-2000     Updated the release procedures to set the Live_Flag entry correctly
 --   d31 DAJ     25-Jul-2000     Modified the release procedures to use joins instead of cursors
--===============================================================================================================
BEGIN
    
     /* Do the actual release on a batch basis */
	 PRINT 'Release editset:   usp_editorial_release_lettercodes_from_editset ' + CAST(@t_iEditSetID AS NVARCHAR(10))
     EXECUTE pct.ReleaseLettercodesFromEditset @t_iEditsetID
     EXECUTE pct.ReleaseAllPCLinks @t_iEditsetID, 1
     PRINT 'Release editset:  UpdateLettercodesInEditsets ' + CAST(@t_iEditSetID AS NVARCHAR(10))
	 EXECUTE pct.UpdateLettercodesInEditsets @t_iEditsetID

END
GO

