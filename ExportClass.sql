USE [ILDB]
GO

/****** Object:  StoredProcedure [pct].[ExportClass]    Script Date: 07/09/2026 15:19:23 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO



-- ========================================================================
-- Author:        Paul Davis / Chris Bartlett
-- Create date:   06/12/2024
-- Description:   Procedure to export class data for search indexing
-- ========================================================================

CREATE OR ALTER PROCEDURE [pct].[ExportClass]
    @t_iEditsetID AS INT,
    @t_Editset_operation AS INT, -- 0 = new, 1 = update, 2 = delete
    @t_eCatStructure AS INT -- 0 = CatPlain, 1 = CATWithDiv
AS
BEGIN  
    SELECT 
	DISTINCT(c.class_id) AS ID,
        3 AS CatalogueLevel, 
         
        lc.letter_code AS LetterCodeRef,
        CASE 
            WHEN @t_eCatStructure = 0 THEN NULL
            WHEN @t_eCatStructure = 1 THEN d.Division_No
        END AS DivisionRef,
        c.class_no AS ClassRef,
        c.subclass_no AS SubClassRef,
		NULL AS HeaderRef,
        NULL AS SubheaderRef,
        NULL AS PieceRef,
        NULL AS ItemRef,
        NULL AS PieceKeyOrder,
        NULL AS ItemKeyOrder,
        c.First_Date AS FirstDate,
        c.Last_Date AS LastDate,
        c.Date_text AS CoveringDates,
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
        c.closure_status AS ClosureStat
		
    FROM 
        ES_lettercode lc
			JOIN ES_class c ON c.lettercode_id = lc.lettercode_id
			JOIN ES_class_extension x ON c.class_id = x.class_id
			LEFT JOIN ES_division d ON c.Division_ID = d.Division_ID
    WHERE 
        c.Edit_set_ID = @t_iEditsetID
        AND lc.Edit_set_ID = @t_iEditsetID
        AND x.Edit_set_ID = @t_iEditsetID
        AND (
            (@t_eCatStructure = 0 AND (c.Division_ID IS NULL OR c.Division_ID = 0)) OR
            (@t_eCatStructure = 1 AND c.Division_ID IS NOT NULL AND c.Division_ID <> 0)
        )
        AND (
            (@t_Editset_operation = 0 AND x.edit_set_entry_type_id IN (5, 4, 1, 2)) OR
            (@t_Editset_operation = 1 AND x.edit_set_entry_type_id = 6) OR
            (@t_Editset_operation = 2 AND x.edit_set_entry_type_id = 3)
        );
END

GO

