USE [ILDB]
GO

/****** Object:  StoredProcedure [pct].[FindReadyEditsets]    Script Date: 07/09/2026 15:24:42 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO




CREATE OR ALTER PROCEDURE [pct].[FindReadyEditsets] AS 
--===============================================================================================
-- File name:   usp_editorial_find_ready_editsets
-- Author:		Paul Davis / Chris Bartlett
-- Create date: 14/11/2024
-- Description: This SP finds all edit sets that are ready for release to live.
--===============================================================================================
BEGIN
    
    /*Get the catalogue flag that indicates that pieces and items are not DORIS-ready*/
    DECLARE @iDORISNotReadyFlag INTEGER
    SELECT @iDORISNotReadyFlag = 
        (SELECT flag_value FROM tbl_flags WHERE flag_name='FLAG_DORIS_NOT_READY')
    
    /*Clean down the table of not yet ready edit sets that are in editorial release*/
    DELETE NotReadyEditSet
    
    
    /* find all edit sets whose latest stage is pre-release */
    SELECT es.edit_set_id, es.Edit_Set_Type, es.name INTO #ReleasedEditsets
    FROM 
    edit_set es JOIN edit_set_history esh 
    ON es.edit_set_id = esh.edit_set_id
    WHERE esh.Transtition_date 
        = (SELECT MAX(Transtition_date) FROM edit_set_history h WHERE h.edit_set_id = esh.edit_set_id)
    AND esh.edit_set_to_stage = 4
    
    
    /*find any edit sets that have pieces that are not DORIS-ready
      by performing a bit-wise AND of the record_status with the DORIS
      flag value
    */
    INSERT INTO NotReadyEditSet(
        edit_set_id,
        edit_set_name,
        name,
        catalogue_id,
        reason
    )
    SELECT 
        x.edit_set_id,
        x.name,
        p.piece_scope,
        p.piece_id,
        'Piece not DORIS ready'

    FROM tbl_piece p, ES_piece ep,  #ReleasedEditsets x
    WHERE p.piece_id = ep.piece_id
    AND x.edit_set_id = ep.edit_set_id
    AND (p.record_status & @iDORISNotReadyFlag) = @iDORISNotReadyFlag
    ORDER by x.edit_set_id

    /*find any edit sets that have items that are not DORIS-ready
      by performing a bit-wise AND of the record_status with the DORIS
      flag value
    */
    INSERT INTO NotReadyEditSet(
        edit_set_id,
        edit_set_name,
        name,
        catalogue_id,
        reason
    )
    SELECT 
        x.edit_set_id,
        x.name,
        i.item_scope,
        i.item_id,
        'Item not DORIS ready'
    
    FROM tbl_item i, ES_item ei, #ReleasedEditsets x
    WHERE i.item_id = ei.item_id
    AND x.edit_set_id = ei.edit_set_id
    AND (i.record_status & @iDORISNotReadyFlag) = @iDORISNotReadyFlag
    ORDER by x.edit_set_id
    
    
    /* find all edit sets that have entries linked to non-approved
       Subject AF terms
    */
    INSERT INTO NotReadyEditSet(
        edit_set_id,
        edit_set_name,
        name,
        catalogue_id,
        reason
    )
    SELECT 
        x.edit_set_id,
        x.name,
        s.Subject_Term_Text,
        s.subject_reference_id,
        'Subject AF term not approved'

    FROM subject s, ES_PC_subject eps, #ReleasedEditsets x
    WHERE s.subject_reference_id = eps.subject_reference_id
    AND x.edit_set_id = eps.edit_set_id
    AND s.authority_status IN (3,4)
    ORDER by x.edit_set_id

    /* find all edit sets that have entries linked to non-approved
       Person AF terms
    */
    INSERT INTO NotReadyEditSet(
        edit_set_id,
        edit_set_name,
        name,
        catalogue_id,
        reason
    )
    SELECT 
        x.edit_set_id,
        x.name,
        (p.Forename_Text + ' ' + p.Surname_Text),
        p.person_reference_id,
        'Person AF term not approved'
        
    FROM person p, ES_PC_person epp, #ReleasedEditsets x
    WHERE p.person_reference_id = epp.person_reference_id
    AND x.edit_set_id = epp.edit_set_id
    AND p.authority_status IN (3,4)
    ORDER by x.edit_set_id

    /* find all edit sets that have entries linked to non-approved
       Place AF terms
    */
    INSERT INTO NotReadyEditSet(
        edit_set_id,
        edit_set_name,
        name,
        catalogue_id,
        reason
    )
    SELECT 
        x.edit_set_id,
        x.name,
        p.Place_Name_Text,
        p.place_reference_id,
        'Place AF term not approved'

    FROM place p, ES_PC_place epp, #ReleasedEditsets x
    WHERE p.place_reference_id = epp.place_reference_id
    AND x.edit_set_id = epp.edit_set_id
    AND p.authority_status IN (3,4)
    ORDER by x.edit_set_id

    /* find all edit sets that have entries linked to non-approved
       Corporate name AF terms
    */
    INSERT INTO NotReadyEditSet(
        edit_set_id,
        edit_set_name,
        name,
        catalogue_id,
        reason
    )
    SELECT 
        x.edit_set_id,
        x.name,
        c.Corporate_Body_Name_Text,
        c.corporate_body_reference_id,
        'Corporate name AF term not approved'

    FROM corporate_body c, ES_PC_corp_body epc, #ReleasedEditsets x
    WHERE c.corporate_body_reference_id = epc.corporate_body_reference_id
    AND x.edit_set_id = epc.edit_set_id
    AND c.authority_status IN (3,4)
    ORDER by x.edit_set_id, x.name, c.Corporate_Body_Name_Text
    
    /* find any new lettercodes that have not yet been assigned Verity collections
       to be indexed in */
    INSERT INTO NotReadyEditSet(
        edit_set_id,
        edit_set_name,
        name,
        catalogue_id,
        reason
    )
    SELECT 
        x.edit_set_id,
        x.name,
        l.letter_code,
        l.lettercode_id,
        ('System administrator has not assigned Verity collections for lettercode: ' + l.letter_code)

    FROM ES_lettercode l JOIN #ReleasedEditsets x ON l.edit_set_id = x.edit_set_id
        LEFT OUTER JOIN lcode_verity_collection lvc ON l.lettercode_id = lvc.lettercode_id
        WHERE lvc.lettercode_id IS NULL
    GROUP by x.edit_set_id, x.name, l.letter_code, l.lettercode_id
    ORDER BY l.letter_code
    
    /* delete edit sets from the set of those in editorial release
       that are not ready yet
    */
    DELETE FROM #ReleasedEditsets
    WHERE edit_set_id IN (SELECT DISTINCT edit_set_id FROM NotReadyEditSet)
        
    SELECT DISTINCT edit_set_id, edit_set_type, name FROM #ReleasedEditsets
    
END
GO

