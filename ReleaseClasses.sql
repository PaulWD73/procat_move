USE [ILDB]
GO

/****** Object:  StoredProcedure [pct].[ReleaseClasses]    Script Date: 07/09/2026 15:25:15 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO




CREATE OR ALTER PROCEDURE [pct].[ReleaseClasses](
    @t_iEditsetID INTEGER
)
AS 
--===============================================================================================================
-- File name:   ReleaseClasses
-- Author:		Paul Davis / Chris Bartlett
-- Create date: 14/11/2024
-- Parameters:   @t_iEditsetID - The ID of the editset whose catalogues are to be moved to audit trail
-- Description: This SP finds all modified or new classes in an
--				edit set and updates the corresponding live tables

--===============================================================================================================
BEGIN
    
     /* Do the actual release on a batch basis */
     EXECUTE pct.ReleaseClassesFromEditset @t_iEditsetID
     EXECUTE pct.ReleaseAllPCLinks @t_iEditsetID, 3
     EXECUTE pct.UpdateClassesInEditsets @t_iEditsetID

END
GO

