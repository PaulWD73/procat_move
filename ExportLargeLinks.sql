USE [ILDB]
GO

/****** Object:  StoredProcedure [pct].[ExportLargeLinks]    Script Date: 07/09/2026 15:22:25 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO



-- =============================================
-- Author:        Paul Davis / Chris Bartlett
-- Create date:   14/11/2024
-- Description:   Procedure to export basic data for search indexing
-- =============================================

CREATE PROCEDURE [pct].[ExportLargeLinks]
    @t_iEditsetID AS INTEGER,
    @t_Editset_operation AS INT -- 0 = new, 1 = update, 2 = delete
AS
BEGIN
    -- Main SELECT Statement
    SELECT 
        'PE' + CONVERT(VARCHAR, Person_Reference_Id) AS LinkDataID, 
        Catalogue_ID AS CatID,
        level_no AS CatalogueLevel,
        Property_ID AS LinkTypeID
    FROM 
        ES_PC_Person t  				
    WHERE  
        t.Edit_set_ID = @t_iEditsetID
        AND Property_ID IN (12, 18, 16)

    UNION 
    SELECT 
        'CO' + CONVERT(VARCHAR, Corporate_Body_Reference_ID) AS LinkDataID,
        Catalogue_ID AS CatID,
        level_no AS CatalogueLevel,
        Property_ID AS LinkTypeID
    FROM 
        ES_PC_corp_body t 
    WHERE 
        t.Edit_set_ID = @t_iEditsetID
        AND Property_ID IN (12, 17, 32, 24, 10, 16)

    UNION 
    SELECT 
        'PL' + CONVERT(VARCHAR, Place_Reference_ID) AS LinkDataID,
        Catalogue_ID AS CatID,
        level_no AS CatalogueLevel,
        Property_ID AS LinkTypeID
    FROM 
        ES_PC_PLACE t 
    WHERE 
        t.Edit_set_ID = @t_iEditsetID 
        AND Property_ID = 19

    UNION 
    SELECT 
        'SU' + CONVERT(VARCHAR, Subject_Reference_ID) AS LinkDataID,
        Catalogue_ID AS CatID,
        level_no AS CatalogueLevel,
        Property_ID AS LinkTypeID
    FROM 
        ES_PC_SUBJECT t 
    WHERE 
        t.Edit_set_ID = @t_iEditsetID 
        AND Property_ID = 20

    ORDER BY CatalogueLevel, CatID, LinkTypeID;
END
GO

