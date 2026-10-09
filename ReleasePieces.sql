USE [ILDB]
GO

/****** Object:  StoredProcedure [pct].[ReleasePieces]    Script Date: 07/09/2026 15:37:55 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO




CREATE OR ALTER PROCEDURE [pct].[ReleasePieces](
    @t_iEditsetID INTEGER
)
AS 
--===============================================================================================================
-- File name:   ReleasePieces  
-- Author:		Paul Davis / Chris Bartlett
-- Create date: 14/11/2024
-- Description: This SP finds all modified or new pieces in an
--				edit set and updates the corresponding live tables
--===============================================================================================================
BEGIN

     /* Do the actual release on a batch basis */
     EXECUTE pct.ReleasePiecesFromEditSet @t_iEditsetID;
     EXECUTE pct.ReleaseAllPCLinks @t_iEditsetID, 6;
     EXECUTE pct.UpdatePiecesInEditsets @t_iEditsetID;

END
GO

