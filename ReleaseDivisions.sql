USE [ILDB]
GO

/****** Object:  StoredProcedure [pct].[ReleaseDivisions]    Script Date: 07/09/2026 15:25:56 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO




CREATE OR ALTER PROCEDURE [pct].[ReleaseDivisions](
    @t_iEditsetID INTEGER
)
AS 
--===============================================================================================================
-- File name:   ReleaseDivisions
-- Author:		Paul Davis / Chris Bartlett
-- Create date: 14/11/2024
-- Description: This SP finds all modified or new divisions in an
--				edit set and updates the corresponding live tables
--===============================================================================================================
BEGIN
    
     /* Do the actual release on a batch basis */
     EXECUTE pct.ReleaseDivisionsFromEditset @t_iEditsetID
     EXECUTE pct.ReleaseAllPCLinks @t_iEditsetID, 2
     EXECUTE pct.UpdateDivisionsInEditsets @t_iEditsetID

END
GO

