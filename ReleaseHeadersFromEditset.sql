USE [ILDB]
GO

/****** Object:  StoredProcedure [pct].[ReleaseHeadersFromEditset]    Script Date: 07/09/2026 15:26:59 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO




CREATE OR ALTER PROCEDURE [pct].[ReleaseHeadersFromEditset] @t_iEditsetID AS INTEGER
AS
--===============================================================================================================
-- File name:   ReleaseHeadersFromEditset  
-- Author:		Paul Davis / Chris Bartlett
-- Create date: 14/11/2024
-- Parameters:  @t_iEditsetID - ID of the editset to be released
-- Description:	This SP finds all the new or modified entries in a particular editset and performs
--				the following actions:
--				* Update/Insert the Former References (PRO and department) into tbl_reference
--				* Copy the entries from the ES table across to the live tables
--===============================================================================================================

    DECLARE @iLevel INTEGER
    SELECT @iLevel = 4 -- Header

    DECLARE @iFormerRefTypeID INTEGER
    DECLARE @iFormerPRORefTypeID INTEGER
    SELECT @iFormerRefTypeID = (SELECT ref_type_id FROM tbl_reftype WHERE ref_type_desc='Former Reference');
    SELECT @iFormerPRORefTypeID = (SELECT ref_type_id FROM tbl_reftype WHERE ref_type_desc='PRO Former Reference');

    /* Find all modified Headers in the edit set */
    SELECT * INTO #ModifiedEntry
        FROM ES_Header_extension ed
        WHERE ed.edit_set_id = @t_iEditsetID
        AND ed.edit_set_entry_type_id IN (1,2,4,5,6);
    
    /* Find all the Former References (dept) that are already in the tbl_reference table */
    SELECT e.former_reference_Department, e.Header_ID
        INTO #FormerRefDept
        FROM Header_extension p, ES_Header_extension e, #ModifiedEntry me
        WHERE p.header_id = e.header_id
        AND e.header_id = me.header_id
        AND (p.former_reference_Department <> e.former_reference_Department 
        OR (p.former_reference_Department IS NULL AND e.former_reference_Department IS NOT NULL));
    
    /* Update the entries that already exist */
    UPDATE tbl_reference
        SET
            tbl_reference.ref_char = f.former_reference_Department
        FROM
            header_extension p, #FormerRefDept f, tbl_reference t
        WHERE
            t.level_no = @iLevel
            AND p.header_id = f.header_id
            AND t.list_ref_id = p.header_id
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
            pe.former_reference_Department,
            @iLevel,
            p.header_id,
            @iFormerRefTypeID,
            p.class_id
        FROM es_header p, ES_Header_extension pe, #ModifiedEntry me
        WHERE p.header_id = pe.header_id
        AND p.header_id NOT IN
            (SELECT list_ref_id AS header_id
                FROM tbl_reference
                WHERE level_no = @iLevel
                AND ref_type_id = @iFormerRefTypeID)
         AND p.header_id = me.header_id
        AND pe.former_reference_Department IS NOT NULL
        AND pe.former_reference_Department <> ''
        AND p.edit_set_id = me.edit_set_id
        AND p.edit_set_id = pe.edit_set_id;

    /* Find all the Former References (PRO) that are already in the tbl_reference table */
    SELECT e.former_reference_pro, e.Header_ID
        INTO #FormerRefPRO
        FROM Header_extension p, ES_Header_extension e, #ModifiedEntry me
        WHERE p.header_id = e.header_id
        AND e.header_id = me.header_id
        AND (p.former_reference_pro <> e.former_reference_pro 
        OR (p.former_reference_pro IS NULL AND e.former_reference_pro IS NOT NULL));
    
    /* Update the entries that already exist */
    UPDATE tbl_reference
        SET
            tbl_reference.ref_char = f.former_reference_pro
        FROM
            header_extension p, #FormerRefPRO f, tbl_reference t
        WHERE
            t.level_no = @iLevel
            AND p.header_id = f.header_id
            AND t.list_ref_id = p.header_id
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
            pe.former_reference_pro,
            @iLevel,
            p.header_id,
            @iFormerPRORefTypeID,
            p.class_id
        FROM es_header p, ES_Header_extension pe, #ModifiedEntry me
        WHERE p.header_id = pe.header_id
        AND p.header_id NOT IN
            (SELECT list_ref_id AS header_id
                FROM tbl_reference
                WHERE level_no = @iLevel
                AND ref_type_id = @iFormerPRORefTypeID)
        AND p.header_id = me.header_id
        AND pe.former_reference_pro IS NOT NULL
        AND pe.former_reference_pro <> ''
        AND p.edit_set_id = me.edit_set_id
        AND p.edit_set_id = pe.edit_set_id;
    
    /* Clear out any empty references that have just been put in tbl_reference */ 
    DELETE r
    FROM tbl_reference r
    WHERE ISNULL( r.ref_char, '' ) = ''
    AND r.list_ref_id IN (SELECT header_id FROM #FormerRefDept
		    UNION SELECT header_id FROM #FormerRefPRO);
    
     UPDATE dest Set 
        dest.class_hdr_no = source.class_hdr_no, 
        dest.header_title = source.header_title, 
        dest.header_scope = source.header_scope, 
        dest.date_text = source.date_text, 
        dest.batch_id = source.batch_id
     FROM ES_Header source, tbl_Header dest, #ModifiedEntry me
     WHERE source.Edit_Set_ID = @t_iEditSetID
        AND dest.Header_ID = source.Header_ID 
        AND me.Header_ID = source.Header_ID;
    
    UPDATE dest Set 
        dest.access_condition_id = source.access_condition_id, 
        dest.accumulation_date_text = source.accumulation_date_text, 
        dest.covering_from_date = source.covering_from_date, 
        dest.covering_to_date = source.covering_to_date, 
        dest.language_id = source.language_id, 
        dest.legal_status_code = source.legal_status_code, 
        dest.map_designation = source.map_designation, 
        dest.physical_condition = source.physical_condition, 
        dest.physical_record_quantity = source.physical_record_quantity, 
        dest.physical_form_code = source.physical_form_code, 
        dest.dimensions = source.dimensions , 
        dest.restrictions_on_use = source.restrictions_on_use, 
        dest.scale_number = source.scale_number, 
        dest.Former_Reference_Department = source.Former_Reference_Department, 
        dest.Former_Reference_PRO = source.Former_Reference_PRO, 
        dest.Record_Opening_Date = source.Record_Opening_Date
     FROM ES_Header_Extension source, Header_Extension dest, #ModifiedEntry me
     WHERE source.Edit_Set_ID = @t_iEditSetID
        AND dest.Header_ID = source.Header_ID 
        AND me.Header_ID = source.Header_ID;

    /*d39 Added so includes entries of type 3 (Added) as well*/
    UPDATE dest SET dest.Live_Flag = 0 
      FROM ES_Header_Extension source, Header_Extension dest
     WHERE source.Edit_Set_ID = @t_iEditSetID
       AND dest.Header_ID = source.Header_ID;

    UPDATE dest SET 
        dest.accruals_text = source.accruals_text
     FROM ES_Header_accruals source, Header_accruals dest, #ModifiedEntry me
     WHERE source.Edit_Set_ID = @t_iEditSetID
        AND dest.Header_ID = source.Header_ID 
        AND me.Header_ID = source.Header_ID;
        
    UPDATE dest SET
        dest.admin_biog_hist_text = source.admin_biog_hist_text
     FROM ES_Header_admin_biog_hist source, Header_admin_biog_hist dest, #ModifiedEntry me
     WHERE source.Edit_Set_ID = @t_iEditSetID
        AND dest.Header_ID = source.Header_ID 
        AND me.Header_ID = source.Header_ID;
        
    UPDATE dest SET 
        dest.app_dest_info_text = source.app_dest_info_text
     FROM ES_Header_app_dest_info source, Header_app_dest_info dest, #ModifiedEntry me
     WHERE source.Edit_Set_ID = @t_iEditSetID
        AND dest.Header_ID = source.Header_ID 
        AND me.Header_ID = source.Header_ID;
        
    UPDATE dest SET 
        dest.arrangement_text = source.arrangement_text
     FROM ES_Header_arrangement source, Header_arrangement dest, #ModifiedEntry me
     WHERE source.Edit_Set_ID = @t_iEditSetID
        AND dest.Header_ID = source.Header_ID 
        AND me.Header_ID = source.Header_ID;
        
    UPDATE dest SET 
        dest.custodial_hist_text = source.custodial_hist_text
     FROM ES_Header_custodial_hist source, Header_custodial_hist dest, #ModifiedEntry me
     WHERE source.Edit_Set_ID = @t_iEditSetID
        AND dest.Header_ID = source.Header_ID 
        AND me.Header_ID = source.Header_ID;
        
    UPDATE dest SET 
        dest.note_text = source.note_text
     FROM ES_Header_note source, Header_note dest, #ModifiedEntry me
     WHERE source.Edit_Set_ID = @t_iEditSetID
        AND dest.Header_ID = source.Header_ID 
        AND me.Header_ID = source.Header_ID;

    UPDATE dest SET 
        dest.scope_content_text = source.scope_content_text
     FROM ES_Header_scope_content source, Header_scope_content dest, #ModifiedEntry me
     WHERE source.Edit_Set_ID = @t_iEditSetID
        AND dest.Header_ID = source.Header_ID 
        AND me.Header_ID = source.Header_ID;
        
    UPDATE dest SET 
        dest.title_text = source.title_text
     FROM ES_Header_title source, Header_title dest, #ModifiedEntry me
     WHERE source.Edit_Set_ID = @t_iEditSetID
        AND dest.Header_ID = source.Header_ID 
        AND me.Header_ID = source.Header_ID;

GO

