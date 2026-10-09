USE [ILDB]
GO

/****** Object:  StoredProcedure [pct].[ReleaseHeaders]    Script Date: 07/09/2026 15:26:47 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO




CREATE OR ALTER PROCEDURE [pct].[ReleaseHeaders](
    @t_iEditsetID INTEGER
)
AS 
--===============================================================================================================
-- File name:   ReleaseHeaders  
-- Author:		Paul Davis / Chris Bartlett
-- Create date: 14/11/2024
-- Description: This SP finds all modified or new headers in an
--				edit set and updates the corresponding live tables
--===============================================================================================================
BEGIN

     /* Do the actual release on a batch basis */
     EXECUTE pct.ReleaseHeadersFromEditset @t_iEditsetID
     EXECUTE pct.ReleaseAllPCLinks @t_iEditsetID, 4
     EXECUTE pct.UpdateHeadersInEditsets @t_iEditsetID
        
END
GO

