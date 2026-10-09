USE [ILDB]
GO

/****** Object:  StoredProcedure [pct].[ReleaseItemsFromEditset]    Script Date: 07/09/2026 15:36:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO





CREATE OR ALTER PROCEDURE [pct].[ReleaseItemsFromEditset] @t_iEditsetID AS INTEGER
AS
--===============================================================================================================
-- File name:   ReleaseItemsFromEditset 
-- Author:	    Paul Davis/ Chris Bartlett
-- Create date: 14/11/2024
-- Parameters:  @t_iEditsetID - ID of the editset to be released
-- Description: This SP finds all the new or modified entries in a particular editset and performs
--				the following actions:
--				* Update/Insert the Former References (PRO and department) into tbl_reference
--				* Copy the entries from the ES table across to the live tables
--===============================================================================================================

    DECLARE @iLevel INTEGER
    SELECT @iLevel = 7; -- Item

    DECLARE @iFormerRefTypeID INTEGER
    DECLARE @iFormerPRORefTypeID INTEGER
    DECLARE @sRef VARCHAR(255)
    SELECT @iFormerRefTypeID = (SELECT ref_type_id FROM tbl_reftype WHERE ref_type_desc='Former Reference');
    SELECT @iFormerPRORefTypeID = (SELECT ref_type_id FROM tbl_reftype WHERE ref_type_desc='PRO Former Reference');

    /* Find all modified Items in the edit set */
    SELECT * INTO #ModifiedEntry
        FROM ES_Item_extension ed
        WHERE ed.edit_set_id = @t_iEditsetID
        AND ed.edit_set_entry_type_id IN (1,2,4,5,6);
    
    /* Find all the Former References (dept) that are already in the tbl_reference table */
    SELECT e.former_reference_Department, e.Item_ID
        INTO #FormerRefDept
        FROM Item_extension p, ES_Item_extension e, #ModifiedEntry me
        WHERE p.item_id = e.item_id
        AND e.item_id = me.item_id
        AND (p.former_reference_Department <> e.former_reference_Department 
        OR (p.former_reference_Department IS NULL AND e.former_reference_Department IS NOT NULL));
    
    /* Update the entries that already exist */
    UPDATE tbl_reference
        SET
            tbl_reference.ref_char = f.former_reference_Department
        FROM
            item_extension p, #FormerRefDept f, tbl_reference t
        WHERE
            t.level_no = @iLevel
            AND p.item_id = f.item_id
            AND t.list_ref_id = p.item_id
            AND t.ref_type_id = @iFormerRefTypeID;

    /* Insert the entries that do not already exist */
    /* d49 MAP Removed #FormerRefDept from query as was preventing entries not in tbl_reference from being inserted */
    /* d50 MAP Added joins on edit_set_id */
    INSERT INTO tbl_reference (
            ref_char,
            level_no,
            list_ref_id,
            ref_type_id,
            class_id
        )
        SELECT
            ie.former_reference_Department,
            @iLevel,
            i.item_id,
            @iFormerRefTypeID,
            p.class_id
        FROM es_item i, ES_Item_extension ie, #ModifiedEntry me, es_piece p
        WHERE i.item_id = ie.item_id
        AND i.item_id NOT IN
            (SELECT list_ref_id AS item_id
                FROM tbl_reference
                WHERE level_no = @iLevel
                AND ref_type_id = @iFormerRefTypeID)
        AND i.item_id = me.item_id
        AND p.piece_id = i.piece_id
        AND ie.former_reference_Department IS NOT NULL
        AND ie.former_reference_Department <> ''
        AND i.edit_set_id = me.edit_set_id
        AND i.edit_set_id = ie.edit_set_id
        AND i.edit_set_id = p.edit_set_id; 

    /* Find all the Former References (PRO) that are already in the tbl_reference table */
    SELECT e.former_reference_pro, e.Item_ID
        INTO #FormerRefPRO
        FROM Item_extension p, ES_Item_extension e, #ModifiedEntry me
        WHERE p.item_id = e.item_id
        AND e.item_id = me.item_id
        AND (p.former_reference_pro <> e.former_reference_pro 
        OR (p.former_reference_pro IS NULL AND e.former_reference_pro IS NOT NULL));
    
    /* Update the entries that already exist */
    UPDATE tbl_reference
        SET
            tbl_reference.ref_char = f.former_reference_pro
        FROM
            item_extension p, #FormerRefPRO f, tbl_reference t
        WHERE
            t.level_no = @iLevel
            AND p.item_id = f.item_id
            AND t.list_ref_id = p.item_id
            AND t.ref_type_id = @iFormerPRORefTypeID;

    /* Insert the entries that do not already exist */
    /* d49 MAP Removed #FormerRefDept from query as was preventing entries not in tbl_reference from being inserted */
    /* d50 MAP Added joins on edit_set_id */
    INSERT INTO tbl_reference (
            ref_char,
            level_no,
            list_ref_id,
            ref_type_id,
            class_id
        )
        SELECT
            ie.former_reference_pro,
            @iLevel,
            i.item_id,
            @iFormerPRORefTypeID,
            p.class_id
        FROM es_item i, ES_Item_extension ie, #ModifiedEntry me, es_piece p
        WHERE i.item_id = ie.item_id
        AND i.item_id NOT IN
            (SELECT list_ref_id AS item_id
                FROM tbl_reference
                WHERE level_no = @iLevel
                AND ref_type_id = @iFormerPRORefTypeID)
        AND i.item_id = me.item_id
        AND p.piece_id = i.piece_id
        AND ie.former_reference_pro IS NOT NULL
        AND ie.former_reference_pro <> ''
        AND i.edit_set_id = me.edit_set_id
        AND i.edit_set_id = ie.edit_set_id
        AND i.edit_set_id = p.edit_set_id;
    
    /* Clear out any empty references that have just been put in tbl_reference */
    DELETE r
    FROM tbl_reference r
    WHERE ISNULL( r.ref_char, '' ) = ''
    AND r.list_ref_id IN (SELECT item_id FROM #FormerRefDept
		    UNION SELECT item_id FROM #FormerRefPRO);
    
    UPDATE es_item
    SET    open_date = dbo.fn_correct_opening_date(es.edit_set_type, ei.open_date, DATEADD(DAY, 1, GETDATE()), ei.closure_type,@t_iEditSetID)
    FROM   es_item ei INNER JOIN edit_set es ON ( ei.Edit_Set_ID = es.Edit_Set_ID )
    WHERE  ei.Edit_Set_ID = @t_iEditSetID;

    UPDATE dest SET 
        dest.piece_id = source.piece_id,  
        dest.item_key_order = source.item_key_order, 
        dest.item_ref = source.item_ref, 
        dest.item_title = source.item_title, 
        dest.item_scope = source.item_scope, 
        dest.date_text = source.date_text, 
        dest.closure_status = source.closure_status, 
        dest.closure_type = source.closure_type, 
        dest.closure_code = source.closure_code, 
        dest.open_date = source.open_date,
        dest.batch_id = source.batch_id, 
        dest.first_date = source.first_date, 
        dest.last_date = source.last_date, 
        dest.physcond_id = source.physcond_id, 
        dest.language_id = source.language_id
    FROM ES_Item source, tbl_Item dest, #ModifiedEntry me
    WHERE source.Edit_Set_ID = @t_iEditSetID
        AND dest.Item_ID = source.Item_ID 
        AND me.Item_ID = source.Item_ID;
    
    UPDATE dest SET 
        dest.access_condition_id = source.access_condition_id, 
        dest.accumulation_date_text = source.accumulation_date_text, 
        dest.legal_status_code = source.legal_status_code, 
        dest.map_designation = source.map_designation, 
        dest.physical_record_quantity = source.physical_record_quantity, 
        dest.physical_form_code = source.physical_form_code, 
        dest.dimensions = source.dimensions, 
        dest.restrictions_on_use = source.restrictions_on_use, 
        dest.scale_number = source.scale_number, 
        dest.Former_Reference_Department = source.Former_Reference_Department, 
        dest.Former_Reference_PRO = source.Former_Reference_PRO,
        dest.Physical_condition = source.Physical_condition
    FROM ES_Item_Extension source, Item_Extension dest, #ModifiedEntry me
    Where source.Edit_Set_ID = @t_iEditSetID
        AND dest.Item_ID = source.Item_ID 
        AND me.Item_ID = source.Item_ID;

    /*d39 Added so includes entries of type 3 (Added) as well*/
    UPDATE dest SET dest.Live_Flag = 0 
      FROM ES_Item_Extension source, Item_Extension dest
     WHERE source.Edit_Set_ID = @t_iEditSetID
       AND dest.Item_ID = source.Item_ID; 

    UPDATE dest SET 
        dest.accruals_text = source.accruals_text
     FROM ES_Item_accruals source, Item_accruals dest, #ModifiedEntry me
     Where source.Edit_Set_ID = @t_iEditSetID
        AND dest.Item_ID = source.Item_ID 
        AND me.Item_ID = source.Item_ID;
        
    UPDATE dest SET
        dest.admin_biog_hist_text = source.admin_biog_hist_text
     FROM ES_Item_admin_biog_hist source, Item_admin_biog_hist dest, #ModifiedEntry me
     WHERE source.Edit_Set_ID = @t_iEditSetID
        AND dest.Item_ID = source.Item_ID 
        AND me.Item_ID = source.Item_ID;
        
    UPDATE dest SET 
        dest.app_dest_info_text = source.app_dest_info_text
     FROM ES_Item_app_dest_info source, Item_app_dest_info dest, #ModifiedEntry me
     WHERE source.Edit_Set_ID = @t_iEditSetID
        AND dest.Item_ID = source.Item_ID 
        AND me.Item_ID = source.Item_ID;
        
    UPDATE dest SET 
        dest.arrangement_text = source.arrangement_text
     FROM ES_Item_arrangement source, Item_arrangement dest, #ModifiedEntry me
     WHERE source.Edit_Set_ID = @t_iEditSetID
        AND dest.Item_ID = source.Item_ID 
        AND me.Item_ID = source.Item_ID;
        
    UPDATE dest SET 
        dest.custodial_hist_text = source.custodial_hist_text
     FROM ES_Item_custodial_hist source, Item_custodial_hist dest, #ModifiedEntry me
     WHERE source.Edit_Set_ID = @t_iEditSetID
        AND dest.Item_ID = source.Item_ID 
        AND me.Item_ID = source.Item_ID;
        
    UPDATE dest SET 
        dest.note_text = source.note_text
     FROM ES_Item_note source, Item_note dest, #ModifiedEntry me
     WHERE source.Edit_Set_ID = @t_iEditSetID
        AND dest.Item_ID = source.Item_ID 
        AND me.Item_ID = source.Item_ID;

    UPDATE dest SET 
        dest.scope_content_text = source.scope_content_text
     FROM ES_Item_scope_content source, Item_scope_content dest, #ModifiedEntry me
     WHERE source.Edit_Set_ID = @t_iEditSetID
        AND dest.Item_ID = source.Item_ID 
        AND me.Item_ID = source.Item_ID;
        
    UPDATE dest SET 
        dest.title_text = source.title_text
     FROM ES_Item_title source, Item_title dest, #ModifiedEntry me
     WHERE source.Edit_Set_ID = @t_iEditSetID
        AND dest.Item_ID = source.Item_ID 
        AND me.Item_ID = source.Item_ID;
GO

