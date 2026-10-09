USE [ILDB]
GO

/****** Object:  StoredProcedure [pct].[ReleaseItems]    Script Date: 07/09/2026 15:27:13 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO




CREATE OR ALTER PROCEDURE [pct].[ReleaseItems](
    @t_iEditsetID INTEGER
)
AS 
--===============================================================================================================
-- File name:   ReleaseItems  
-- Author:		Paul Davis / Chris Bartlett
-- Create date: 14/11/2024
-- Description: This SP finds all modified or new items in an
--				edit set and updates the corresponding live tables
--===============================================================================================================
BEGIN
 
     /* Do the actual release on a batch basis */
     EXECUTE pct.ReleaseItemsFromEditset @t_iEditsetID;
     EXECUTE pct.ReleaseAllPCLinks @t_iEditsetID, 7;
     EXECUTE pct.UpdateItemsInEditsets @t_iEditsetID;
 
END
GO

