USE [ILDB]
GO

/****** Object:  StoredProcedure [pct].[ExportLargeAppraisalInfo]    Script Date: 07/09/2026 15:21:40 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO




-- ===========================================================================
-- Author:        Paul Davis / Chris Bartlett
-- Create date:   06/12/2024
-- Description:   Procedure to export Appraisal data for search indexing
-- ===========================================================================
CREATE OR ALTER PROCEDURE [pct].[ExportLargeAppraisalInfo]
    @t_iEditsetID AS INT,
    @t_Editset_operation AS INT -- 0 = new, 1 = update, 2 = delete
AS
BEGIN
    -- Main SELECT Statement
	
    -- Lettercodes
    SELECT 
        1 AS CatalogueLevel, 
        lc.lettercode_id AS ID, 
        a.app_dest_info_text AS AppraisalInfo
    FROM 
        ES_Lettercode lc
			JOIN ES_Lettercode_Extension x ON lc.lettercode_id = x.lettercode_id
			JOIN ES_lettercode_app_dest_info a ON lc.lettercode_id = a.lettercode_id
    WHERE 
        lc.Edit_set_ID = @t_iEditsetID
        AND x.Edit_set_ID = @t_iEditsetID
        AND a.Edit_set_ID = @t_iEditsetID
          AND (
            (@t_Editset_operation = 0 AND x.edit_set_entry_type_id IN (5, 4, 1, 2)) OR
            (@t_Editset_operation = 1 AND x.edit_set_entry_type_id = 6) OR
            (@t_Editset_operation = 2 AND x.edit_set_entry_type_id = 3)
        )
    UNION

    -- Divisions
    SELECT 
        2 AS CatalogueLevel, 
        d.Division_ID AS ID, 
        a.app_dest_info_text AS AppraisalInfo
    FROM 
        ES_Division d
			JOIN ES_Division_Extension x ON d.Division_ID = x.Division_ID
			JOIN ES_Division_app_dest_info a ON d.Division_ID = a.Division_ID
    WHERE 
        d.Edit_set_ID = @t_iEditsetID
        AND x.Edit_set_ID = @t_iEditsetID
        AND a.Edit_set_ID = @t_iEditsetID
          AND (
            (@t_Editset_operation = 0 AND x.edit_set_entry_type_id IN (5, 4, 1, 2)) OR
            (@t_Editset_operation = 1 AND x.edit_set_entry_type_id = 6) OR
            (@t_Editset_operation = 2 AND x.edit_set_entry_type_id = 3)
        )
    UNION

    -- Classes
    SELECT 
        3 AS CatalogueLevel, 
        c.class_id AS ID, 
        a.app_dest_info_text AS AppraisalInfo
    FROM 
        ES_Class c
			JOIN ES_Class_Extension x ON c.class_id = x.class_id
			JOIN ES_Class_app_dest_info a ON c.class_id = a.class_id
    WHERE 
        c.Edit_set_ID = @t_iEditsetID
        AND x.Edit_set_ID = @t_iEditsetID
        AND a.Edit_set_ID = @t_iEditsetID
          AND (
            (@t_Editset_operation = 0 AND x.edit_set_entry_type_id IN (5, 4, 1, 2)) OR
            (@t_Editset_operation = 1 AND x.edit_set_entry_type_id = 6) OR
            (@t_Editset_operation = 2 AND x.edit_set_entry_type_id = 3)
        )
    UNION

    -- Headers
    SELECT 
        4 AS CatalogueLevel, 
        h.Header_ID AS ID, 
        a.app_dest_info_text AS AppraisalInfo
    FROM 
        ES_Header h
			JOIN ES_Header_Extension x ON h.Header_ID = x.Header_ID
			JOIN ES_Header_app_dest_info a ON h.Header_ID = a.Header_ID
    WHERE 
        h.Edit_set_ID = @t_iEditsetID
        AND x.Edit_set_ID = @t_iEditsetID
        AND a.Edit_set_ID = @t_iEditsetID
          AND (
            (@t_Editset_operation = 0 AND x.edit_set_entry_type_id IN (5, 4, 1, 2)) OR
            (@t_Editset_operation = 1 AND x.edit_set_entry_type_id = 6) OR
            (@t_Editset_operation = 2 AND x.edit_set_entry_type_id = 3)
        )
    UNION

    -- Subheaders
    SELECT 
        5 AS CatalogueLevel, 
        sh.SubHeader_ID AS ID, 
        a.app_dest_info_text AS AppraisalInfo
    FROM 
        ES_Subheader sh
			JOIN ES_Subheader_Extension x ON sh.SubHeader_ID = x.SubHeader_ID
			JOIN ES_Subheader_app_dest_info a ON sh.SubHeader_ID = a.SubHeader_ID
    WHERE 
        sh.Edit_set_ID = @t_iEditsetID
        AND x.Edit_set_ID = @t_iEditsetID
        AND a.Edit_set_ID = @t_iEditsetID
          AND (
            (@t_Editset_operation = 0 AND x.edit_set_entry_type_id IN (5, 4, 1, 2)) OR
            (@t_Editset_operation = 1 AND x.edit_set_entry_type_id = 6) OR
            (@t_Editset_operation = 2 AND x.edit_set_entry_type_id = 3)
        )
    UNION

    -- Pieces
    SELECT 
        6 AS CatalogueLevel, 
        p.piece_id AS ID, 
        a.app_dest_info_text AS AppraisalInfo
    FROM 
        ES_Piece p
			JOIN ES_Piece_Extension x ON p.piece_id = x.piece_id
			JOIN ES_Piece_app_dest_info a ON p.piece_id = a.piece_id
    WHERE 
        p.Edit_set_ID = @t_iEditsetID
        AND x.Edit_set_ID = @t_iEditsetID
        AND a.Edit_set_ID = @t_iEditsetID
          AND (
            (@t_Editset_operation = 0 AND x.edit_set_entry_type_id IN (5, 4, 1, 2)) OR
            (@t_Editset_operation = 1 AND x.edit_set_entry_type_id = 6) OR
            (@t_Editset_operation = 2 AND x.edit_set_entry_type_id = 3)
        )
    UNION

    -- Items
    SELECT 
        7 AS CatalogueLevel, 
        i.Item_ID AS ID, 
        a.app_dest_info_text AS AppraisalInfo
    FROM 
        ES_Item i
			JOIN ES_Item_Extension x ON i.Item_ID = x.Item_ID
			JOIN ES_Item_app_dest_info a ON i.Item_ID = a.Item_ID
    WHERE 
        i.Edit_set_ID = @t_iEditsetID
        AND x.Edit_set_ID = @t_iEditsetID
        AND a.Edit_set_ID = @t_iEditsetID
         AND (
            (@t_Editset_operation = 0 AND x.edit_set_entry_type_id IN (5, 4, 1, 2)) OR
            (@t_Editset_operation = 1 AND x.edit_set_entry_type_id = 6) OR
            (@t_Editset_operation = 2 AND x.edit_set_entry_type_id = 3)
        );
END

GO

