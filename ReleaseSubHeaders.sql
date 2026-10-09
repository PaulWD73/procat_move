USE [ILDB]
GO

/****** Object:  StoredProcedure [pct].[ReleaseSubHeaders]    Script Date: 07/09/2026 15:38:24 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO




CREATE OR ALTER PROCEDURE [pct].[ReleaseSubHeaders](
    @t_iEditsetID INTEGER
)
AS 
--===============================================================================================================
-- File name:   ReleaseSubHeaders 
-- Author:		Paul Davis / Chris Bartlett
-- Create date: 14/11/2024
-- Description: This SP finds all modified or new subheaders in an
--				edit set and updates the corresponding live tables
--===============================================================================================================
BEGIN
    
     /* Do the actual release on a batch basis */
     EXECUTE pct.ReleaseSubHeadersFromEditset @t_iEditsetID;
     EXECUTE pct.ReleaseAllPCLinks @t_iEditsetID, 5;
     EXECUTE pct.UpdateSubHeadersInEditsets @t_iEditsetID;

END
GO

