USE [ILDB]
GO

/****** Object:  StoredProcedure [pct].[ExportLettercode]    Script Date: 07/09/2026 15:23:34 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO





-- =============================================
-- Author:        Paul Davis / Chris Bartlett
-- Create date:   14/11/2024
-- Description:   Procedure to export basic data for search indexing
-- =============================================
CREATE OR ALTER PROCEDURE [pct].[ExportLettercode]
    @t_iEditsetID AS INT,
    @t_Editset_operation AS INT -- 0 = new, 1 = update, 2 = delete
AS
BEGIN
    SELECT 
		DISTINCT(lc.lettercode_id) AS ID,
        1 AS CatalogueLevel,  
        lc.letter_code AS LetterCodeRef, 
        NULL AS DivisionRef, 
        NULL AS ClassRef, 
        NULL AS SubClassRef, 
		NULL AS HeaderRef,
        NULL AS SubheaderRef, 
        NULL AS PieceRef, 
        NULL AS ItemRef, 
        NULL AS PieceKeyOrder, 
        NULL AS ItemKeyOrder, 
        x.Covering_From_Date AS FirstDate, 
        x.Covering_To_Date AS LastDate, 
        x.covering_date_text AS CoveringDates, 
        x.map_designation AS MapDesig, 
        x.Former_Reference_Department AS FormerRef, 
        x.Former_Reference_PRO AS FormerPRORef, 
        x.scale_number AS MapScale,
        x.restrictions_on_use AS Restrictions,
        x.accumulation_date_text AS AccumDates,
        x.record_opening_date AS OpeningDate,
        x.legal_status_code AS LegalStatus,
        x.physical_form_code AS PhyDescForm,
        x.physical_record_quantity AS Quantity,
        x.dimensions AS Dimensions,
        x.access_condition_id AS AccessCond,
        x.language_id AS Language,
        x.physical_condition AS PhysCond,
        NULL AS ClosureCode,
        NULL AS ClosureType,
        lc.closure_status AS ClosureStat
    FROM 
        ES_Lettercode lc
			JOIN ES_Lettercode_Extension x ON lc.lettercode_id = x.lettercode_id
    WHERE 
        lc.Edit_set_ID = @t_iEditsetID
        AND x.Edit_set_ID = @t_iEditsetID
          AND (
            (@t_Editset_operation = 0 AND x.edit_set_entry_type_id IN (5, 4, 1, 2)) OR
            (@t_Editset_operation = 1 AND x.edit_set_entry_type_id = 6) OR
            (@t_Editset_operation = 2 AND x.edit_set_entry_type_id = 3)
        );
END

GO

