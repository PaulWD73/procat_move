USE [ILDB]
GO

/****** Object:  StoredProcedure [pct].[ReleaseAllPCLinks]    Script Date: 07/09/2026 15:25:02 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO




CREATE OR ALTER Procedure [pct].[ReleaseAllPCLinks] @t_iEditsetID as integer, @t_iCatalogueLevel as integer
AS
--===============================================================================================================
-- File name:   ReleaseAllPCLinks 
-- Author:		Paul Davis / Chris Bartlett
-- Create date: 14/11/2024
-- Parameters:  @t_iEditsetID - ID of the edit set to release the links from
--                @t_iCatalogueLevel - The level of the source entries to release the links from           
-- Description: This SP selects the suitable new/modified entries at the specified level
--                of the catalogue from the edit set into a temp table.  It then joins
--                that list into the various link (PC_*) tables to copy the values from the
--                edit set tables into the live ones.  It has to delete the links in the live
 --               tables first, as modifications cannot be 'matched' between the two (no unique ID).
--===============================================================================================================

DECLARE @iRelatedMaterialLinkID INTEGER
SELECT @iRelatedMaterialLinkID = 36

    -- Create the table to hold the catalogue IDs
    CREATE TABLE #CatalogueEntries (
        edit_set_id     int,
        catalogue_id    int
    )
        
    -- Use the catalogue level to call the correct SP
    if @t_iCatalogueLevel = 1         -- Lettercode
    begin
        INSERT INTO #CatalogueEntries
        SELECT s.edit_set_id, s.lettercode_id as catalogue_id
            FROM ES_Lettercode_extension s
            WHERE s.edit_set_id = @t_iEditSetID
            AND s.edit_set_entry_type_id IN (1,2,4,5,6)
    end
    else if @t_iCatalogueLevel = 2    -- Division
    begin
        INSERT INTO #CatalogueEntries
        SELECT s.edit_set_id, s.division_id as catalogue_id
            FROM ES_Division_extension s
            WHERE s.edit_set_id = @t_iEditSetID
            AND s.edit_set_entry_type_id IN (1,2,4,5,6)
    end
    else if @t_iCatalogueLevel = 3    -- Class  
    begin
        INSERT INTO #CatalogueEntries
        SELECT s.edit_set_id, s.class_id as catalogue_id
            FROM ES_Class_extension s
            WHERE s.edit_set_id = @t_iEditSetID
            AND s.edit_set_entry_type_id IN (1,2,4,5,6)
    end
    else if @t_iCatalogueLevel = 4    -- Header
    begin
        INSERT INTO #CatalogueEntries
        SELECT s.edit_set_id, s.header_id as catalogue_id
            FROM ES_Header_extension s
            WHERE s.edit_set_id = @t_iEditSetID
            AND s.edit_set_entry_type_id IN (1,2,4,5,6)
    end
    else if @t_iCatalogueLevel = 5    -- Sub Header
    begin
        INSERT INTO #CatalogueEntries
        SELECT s.edit_set_id, s.subheader_id as catalogue_id
            FROM ES_SubHeader_extension s
            WHERE s.edit_set_id = @t_iEditSetID
            AND s.edit_set_entry_type_id IN (1,2,4,5,6)
    end
    else if @t_iCatalogueLevel = 6    -- Piece
    begin
        INSERT INTO #CatalogueEntries
        SELECT s.edit_set_id, s.piece_id as catalogue_id
            FROM ES_Piece_extension s
            WHERE s.edit_set_id = @t_iEditSetID
            AND s.edit_set_entry_type_id IN (1,2,4,5,6)
    end
    else if @t_iCatalogueLevel = 7    -- Item
    begin
        INSERT INTO #CatalogueEntries
        SELECT s.edit_set_id, s.item_id as catalogue_id
            FROM ES_Item_extension s
            WHERE s.edit_set_id = @t_iEditSetID
            AND s.edit_set_entry_type_id IN (1,2,4,5,6)
    end

    --Remove the existing links in the link description table
    DELETE PC_Link_Data_Element
        FROM PC_Link_Data_Element p, #CatalogueEntries e
        WHERE p.level_no = @t_iCatalogueLevel
        AND p.catalogue_id = e.catalogue_id
    
    --Add the new links in the link description table
    INSERT Into PC_Link_Data_Element (Property_ID, level_no, Catalogue_ID, Description)
    SELECT p.Property_ID, p.level_no, p.Catalogue_ID, p.Description
    FROM ES_PC_Link_Data_Element p, #CatalogueEntries e
    WHERE p.level_no = @t_iCatalogueLevel
      AND p.catalogue_id = e.catalogue_id
      AND p.edit_set_id = e.edit_set_id
    
    --Remove the existing links in the catalogue cross reference table
    DELETE FROM PC_link
        FROM PC_link p, #CatalogueEntries e
        WHERE p.level_no1 = @t_iCatalogueLevel
        AND p.catalogue_id1 = e.catalogue_id
    
    --Add the new links in the catalogue cross reference table
    INSERT Into PC_Link (level_no1, catalogue_ID1, level_no2, catalogue_ID2, Property_ID, Description)
    SELECT  level_no1, catalogue_ID1, level_no2, catalogue_ID2, Property_ID, Description
    FROM ES_PC_Link p, #CatalogueEntries e
    WHERE p.level_no1 = @t_iCatalogueLevel
      AND p.catalogue_id1 = e.catalogue_id
      AND p.edit_set_id = e.edit_set_id
    
    --Remove the existing links in pc_sfa table
    DELETE FROM PC_SFA
        FROM PC_SFA p, #CatalogueEntries e
        WHERE p.level_no = @t_iCatalogueLevel
        AND p.catalogue_id = e.catalogue_id
    
    --Add the new links  in pc_sfa table
    INSERT Into PC_SFA (SFA_Reference_ID, Catalogue_ID, level_no, Property_ID, Description)
    SELECT p.SFA_Reference_ID, p.Catalogue_ID, p.level_no, p.Property_ID, p.Description
    FROM ES_PC_SFA p, #CatalogueEntries e
    WHERE p.level_no = @t_iCatalogueLevel
      AND p.catalogue_id = e.catalogue_id
      AND p.edit_set_id = e.edit_set_id
    
    --Remove the existing links in pc_leaflet table
    DELETE FROM PC_Leaflet
        FROM PC_Leaflet p, #CatalogueEntries e
        WHERE p.level_no = @t_iCatalogueLevel
        AND p.catalogue_id = e.catalogue_id
    
    --Add the new links in pc_leaflet table
    INSERT Into PC_Leaflet (Leaflet_Reference_ID, Catalogue_ID, level_no, Property_ID, Description)
    SELECT p.Leaflet_Reference_ID, p.Catalogue_ID, p.level_no, p.Property_ID, p.Description
    FROM ES_PC_Leaflet p, #CatalogueEntries e
    WHERE p.level_no = @t_iCatalogueLevel
      AND p.catalogue_id = e.catalogue_id
      AND p.edit_set_id = e.edit_set_id
    
    --Remove the existing links
    DELETE FROM PC_corp_body
        FROM PC_corp_body p, #CatalogueEntries e
        WHERE p.level_no = @t_iCatalogueLevel
        AND p.catalogue_id = e.catalogue_id
    
    --Add the new links
    INSERT Into PC_corp_body (Catalogue_ID, Property_ID, level_no, Corporate_Body_Reference_Id, Description)
    SELECT p.Catalogue_ID, p.Property_ID, p.level_no, p.Corporate_Body_Reference_Id, p.Description
    FROM ES_PC_corp_body p, #CatalogueEntries e
    WHERE p.level_no = @t_iCatalogueLevel
      AND p.catalogue_id = e.catalogue_id
      AND p.edit_set_id = e.edit_set_id
    
    
    --Remove the existing links
    DELETE FROM PC_person
        FROM PC_person p, #CatalogueEntries e
        WHERE p.level_no = @t_iCatalogueLevel
        AND p.catalogue_id = e.catalogue_id
    
    --Add the new links
    INSERT Into  PC_person (Catalogue_ID, Property_ID, level_no,    Person_Reference_Id, Description)
    SELECT p.Catalogue_ID, p.Property_ID, p.level_no, p.Person_Reference_Id, p.Description
    FROM ES_PC_person p, #CatalogueEntries e
    WHERE p.level_no = @t_iCatalogueLevel
      AND p.catalogue_id = e.catalogue_id
      AND p.edit_set_id = e.edit_set_id
    
    --Remove the existing links
    DELETE FROM PC_place
        FROM PC_place p, #CatalogueEntries e
        WHERE p.level_no = @t_iCatalogueLevel
        AND p.catalogue_id = e.catalogue_id
    
    --Add the new links
    INSERT Into PC_place (Catalogue_ID, Property_ID, level_no, Place_Reference_Id, Description )
    SELECT p.Catalogue_ID, p.Property_ID, p.level_no, p.Place_Reference_Id, p.Description
    FROM ES_PC_place p, #CatalogueEntries e
    WHERE p.level_no = @t_iCatalogueLevel
      AND p.catalogue_id = e.catalogue_id
      AND p.edit_set_id = e.edit_set_id
    
    --Remove the existing links
    DELETE FROM PC_subject
        FROM PC_subject p, #CatalogueEntries e
        WHERE p.level_no = @t_iCatalogueLevel
        AND p.catalogue_id = e.catalogue_id

    --Add the new links
    INSERT Into PC_subject (Catalogue_ID, Property_ID, level_no, Subject_Reference_Id, Description )
    SELECT p.Catalogue_ID, p.Property_ID, p.level_no, p.Subject_Reference_Id, p.Description 
    FROM ES_PC_subject p, #CatalogueEntries e
    WHERE p.level_no = @t_iCatalogueLevel
      AND p.catalogue_id = e.catalogue_id
      AND p.edit_set_id = e.edit_set_id
    
    DROP TABLE #CatalogueEntries
    
GO

