USE [ILDB]
GO

/****** Object:  StoredProcedure [pct].[ExportDivision]    Script Date: 07/09/2026 15:20:01 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO



-- =============================================
-- Author:        Paul Davis / Chris Bartlett
-- Create date:   14/11/2024
-- Description:   Procedure to export basic data for search indexing
-- =============================================
CREATE OR ALTER PROCEDURE [pct].[ExportDivision]
    @t_iEditsetID AS INT,
    @t_Editset_operation AS INT -- 0 = new, 1 = update, 2 = delete
AS
BEGIN
    SELECT 
	    DISTINCT(d.Division_ID) AS ID, 
        2 AS CatalogueLevel, 
        lc.letter_code AS LetterCodeRef, 
        d.Division_No AS DivisionRef, 
        NULL AS ClassRef, 
        NULL AS SubClassRef, 
		NULL As HeaderRef,
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
        d.closure_status AS ClosureStat
		
    FROM ES_lettercode lc
    JOIN ES_Division d ON d.lettercode_id = lc.lettercode_id
    JOIN ES_division_extension x ON d.Division_ID = x.Division_ID
    WHERE lc.Edit_set_ID = @t_iEditsetID
      AND d.Edit_set_ID = @t_iEditsetID
      AND x.Edit_set_ID = @t_iEditsetID
      AND (
        (@t_Editset_operation = 0 AND x.edit_set_entry_type_id IN (5, 4, 1, 2)) OR
        (@t_Editset_operation = 1 AND x.edit_set_entry_type_id = 6) OR
        (@t_Editset_operation = 2 AND x.edit_set_entry_type_id = 3)
    );
END

GO

