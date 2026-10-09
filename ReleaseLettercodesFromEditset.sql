USE [ILDB]
GO

/****** Object:  StoredProcedure [pct].[ReleaseLettercodesFromEditset]    Script Date: 07/09/2026 15:37:25 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO




CREATE OR ALTER Procedure [pct].[ReleaseLettercodesFromEditset] @t_iEditsetID as integer
AS
--===============================================================================================================
-- File name:   ReleaseLettercodesFromEditset 
-- Author:		Paul Davis / Chris Bartlett
-- Create date: 14/11/2024
-- Parameters:  @t_iEditsetID - ID of the editset to be released
-- Description: This SP finds all the new or modified entries in a particular editset and performs
--				the following actions:
--				* Update/Insert the Former References (PRO and department) into tbl_reference
--				* Copy the entries from the ES table across to the live tables
--===============================================================================================================

    DECLARE @iLevel INTEGER
    SELECT @iLevel = 1 -- Lettercode

    DECLARE @iFormerRefTypeID INTEGER
    DECLARE @iFormerPRORefTypeID INTEGER
    SELECT @iFormerRefTypeID = (SELECT ref_type_id FROM tbl_reftype WHERE ref_type_desc='Former Reference')
    SELECT @iFormerPRORefTypeID = (SELECT ref_type_id FROM tbl_reftype WHERE ref_type_desc='PRO Former Reference')

    /* Find all modified Lettercodes in the edit set */
    SELECT * INTO #ModifiedEntry
        FROM ES_Lettercode_extension ed
        WHERE ed.edit_set_id = @t_iEditsetID
        AND ed.edit_set_entry_type_id IN (1,2,4,5,6)
    
    /* Find all the Former References (dept) that are already in the tbl_reference table */
    SELECT e.former_reference_Department, e.Lettercode_ID
        INTO #FormerRefDept
        FROM Lettercode_extension p, ES_Lettercode_extension e, #ModifiedEntry me
        WHERE p.lettercode_id = e.lettercode_id
        AND e.lettercode_id = me.lettercode_id
        AND (p.former_reference_Department <> e.former_reference_Department 
        OR (p.former_reference_Department IS NULL AND e.former_reference_Department IS NOT NULL))
    
    /* Update the entries that already exist */
    UPDATE tbl_reference
        SET
            tbl_reference.ref_char = f.former_reference_Department
        FROM
            lettercode_extension p, #FormerRefDept f, tbl_reference t
        WHERE
            t.level_no = @iLevel
            AND p.lettercode_id = f.lettercode_id
            AND t.list_ref_id = p.lettercode_id
            AND t.ref_type_id = @iFormerRefTypeID

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
            p.lettercode_id,
            @iFormerRefTypeID,
            NULL
        FROM es_lettercode p, ES_Lettercode_extension pe, #ModifiedEntry me
        WHERE p.lettercode_id = pe.lettercode_id
        AND p.lettercode_id NOT IN
            (SELECT list_ref_id as lettercode_id
                FROM tbl_reference
                WHERE level_no = @iLevel
                AND ref_type_id = @iFormerRefTypeID)
        AND p.lettercode_id = me.lettercode_id
        AND pe.former_reference_Department IS NOT NULL
        AND pe.former_reference_Department <> ''
        AND p.edit_set_id = me.edit_set_id
        AND p.edit_set_id = pe.edit_set_id

    /* Find all the Former References (PRO) that are already in the tbl_reference table */
    SELECT e.former_reference_pro, e.Lettercode_ID
        INTO #FormerRefPRO
        FROM Lettercode_extension p, ES_Lettercode_extension e, #ModifiedEntry me
        WHERE p.lettercode_id = e.lettercode_id
        AND e.lettercode_id = me.lettercode_id
        AND (p.former_reference_pro <> e.former_reference_pro 
        OR (p.former_reference_pro IS NULL AND e.former_reference_pro IS NOT NULL))
    
    /* Update the entries that already exist */
    UPDATE tbl_reference
        SET
            tbl_reference.ref_char = f.former_reference_pro
        FROM
            lettercode_extension p, #FormerRefPRO f, tbl_reference t
        WHERE
            t.level_no = @iLevel
            AND p.lettercode_id = f.lettercode_id
            AND t.list_ref_id = p.lettercode_id
            AND t.ref_type_id = @iFormerPRORefTypeID

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
            p.lettercode_id,
            @iFormerPRORefTypeID,
            NULL
        FROM es_lettercode p, ES_Lettercode_extension pe, #ModifiedEntry me
        WHERE p.lettercode_id = pe.lettercode_id
        AND p.lettercode_id NOT IN
            (SELECT list_ref_id as lettercode_id
                FROM tbl_reference
                WHERE level_no = @iLevel
                AND ref_type_id = @iFormerPRORefTypeID)
        AND p.lettercode_id = me.lettercode_id
        AND pe.former_reference_pro IS NOT NULL
        AND pe.former_reference_pro <> ''
        AND p.edit_set_id = me.edit_set_id
        AND p.edit_set_id = pe.edit_set_id


    /* Clear out any empty references that have just been put in tbl_reference */
    DELETE r
    FROM tbl_reference r
    WHERE isnull( r.ref_char, '' ) = ''
    AND r.list_ref_id IN (SELECT lettercode_id FROM #FormerRefDept
		    UNION SELECT lettercode_id FROM #FormerRefPRO)
    
    Update dest Set 
        dest.letter_code = source.letter_code, 
        dest.lettercode_title = source.lettercode_title, 
        dest.closure_status = source.closure_status
    From ES_Lettercode source, tbl_Lettercode dest, #ModifiedEntry me
    Where source.Edit_Set_ID = @t_iEditSetID
        AND dest.Lettercode_ID = source.Lettercode_ID 
        AND me.Lettercode_ID = source.Lettercode_ID
    
    Update dest Set 
        dest.access_condition_id = source.access_condition_id,
        dest.accumulation_date_text = source.accumulation_date_text, 
        dest.covering_date_text = source.covering_date_text, 
        dest.covering_from_date = source.covering_from_date, 
        dest.covering_to_date = source.covering_to_date, 
        dest.language_id = source.language_id,
        dest.legal_status_code = source.legal_status_code, 
        dest.map_designation = source.map_designation, 
        dest.physical_condition = source.physical_condition, 
        dest.physical_record_quantity = source.physical_record_quantity, 
        dest.physical_form_code = source.physical_form_code, 
        dest.dimensions = source.dimensions,
        dest.restrictions_on_use = source.restrictions_on_use, 
        dest.scale_number = source.scale_number, 
        dest.Former_Reference_Department = source.Former_Reference_Department, 
        dest.Former_Reference_PRO = source.Former_Reference_PRO, 
        dest.Record_Opening_Date = source.Record_Opening_Date
     From ES_Lettercode_Extension source, Lettercode_Extension dest, #ModifiedEntry me
     Where source.Edit_Set_ID = @t_iEditSetID
        AND dest.Lettercode_ID = source.Lettercode_ID 
        AND me.Lettercode_ID = source.Lettercode_ID

    /*d39 Added so includes entries of type 3 (Added) as well*/
    UPDATE dest SET dest.Live_Flag = 0 
      FROM ES_Lettercode_Extension source, Lettercode_Extension dest
     WHERE source.Edit_Set_ID = @t_iEditSetID
       AND dest.Lettercode_ID = source.Lettercode_ID 

    UPDATE dest SET 
        dest.accruals_text = source.accruals_text
    From ES_Lettercode_accruals source, Lettercode_accruals dest, #ModifiedEntry me
    Where source.Edit_Set_ID = @t_iEditSetID
        AND dest.Lettercode_ID = source.Lettercode_ID 
        AND me.Lettercode_ID = source.Lettercode_ID
        
    UPDATE dest SET
        dest.admin_biog_hist_text = source.admin_biog_hist_text
    From ES_Lettercode_admin_biog_hist source, Lettercode_admin_biog_hist dest, #ModifiedEntry me
    Where source.Edit_Set_ID = @t_iEditSetID
        AND dest.Lettercode_ID = source.Lettercode_ID 
        AND me.Lettercode_ID = source.Lettercode_ID
        
    UPDATE dest SET 
        dest.app_dest_info_text = source.app_dest_info_text
    From ES_Lettercode_app_dest_info source, Lettercode_app_dest_info dest, #ModifiedEntry me
    Where source.Edit_Set_ID = @t_iEditSetID
        AND dest.Lettercode_ID = source.Lettercode_ID 
        AND me.Lettercode_ID = source.Lettercode_ID
        
    UPDATE dest SET 
        dest.arrangement_text = source.arrangement_text
    From ES_Lettercode_arrangement source, Lettercode_arrangement dest, #ModifiedEntry me
    Where source.Edit_Set_ID = @t_iEditSetID
        AND dest.Lettercode_ID = source.Lettercode_ID 
        AND me.Lettercode_ID = source.Lettercode_ID
        
    UPDATE dest SET 
        dest.custodial_hist_text = source.custodial_hist_text
    From ES_Lettercode_custodial_hist source, Lettercode_custodial_hist dest, #ModifiedEntry me
    Where source.Edit_Set_ID = @t_iEditSetID
        AND dest.Lettercode_ID = source.Lettercode_ID 
        AND me.Lettercode_ID = source.Lettercode_ID
        
    UPDATE dest SET 
        dest.note_text = source.note_text
    From ES_Lettercode_note source, Lettercode_note dest, #ModifiedEntry me
    Where source.Edit_Set_ID = @t_iEditSetID
        AND dest.Lettercode_ID = source.Lettercode_ID 
        AND me.Lettercode_ID = source.Lettercode_ID

    UPDATE dest SET 
        dest.scope_content_text = source.scope_content_text
    From ES_Lettercode_scope_content source, Lettercode_scope_content dest, #ModifiedEntry me
    Where source.Edit_Set_ID = @t_iEditSetID
        AND dest.Lettercode_ID = source.Lettercode_ID 
        AND me.Lettercode_ID = source.Lettercode_ID
        
    UPDATE dest SET 
        dest.title_text = source.title_text
    From ES_Lettercode_title source, Lettercode_title dest, #ModifiedEntry me
    Where source.Edit_Set_ID = @t_iEditSetID
        AND dest.Lettercode_ID = source.Lettercode_ID 
        AND me.Lettercode_ID = source.Lettercode_ID

GO

