USE [ILDB]
GO

/****** Object:  StoredProcedure [pct].[ExportHeader]    Script Date: 07/09/2026 15:20:31 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO





-- =======================================================================
-- Author:        Paul Davis / Chris Bartlett
-- Create date:   06/12/2024
-- Description:   Procedure to export header data for search indexing
-- =======================================================================
CREATE OR ALTER PROCEDURE [pct].[ExportHeader]
    @t_iEditsetID AS INT,
    @t_Editset_operation AS INT, -- 0 = new, 1 = update, 2 = delete
    @t_eCatStructure AS INT -- 0 = CatPlain, 1 = CATWithDiv
AS
BEGIN
    -- Main SELECT Statement
    SELECT 
		DISTINCT(h.Header_ID) AS ID,
        4 AS CatalogueLevel, 
        lc.letter_code AS LetterCodeRef,
        CASE 
            WHEN @t_eCatStructure = 0 THEN NULL
            WHEN @t_eCatStructure = 1 THEN d.Division_No
        END AS DivisionRef,
        c.class_no AS ClassRef,
        c.subclass_no AS SubClassRef,
		h.class_hdr_no AS HeaderRef,
        NULL AS SubheaderRef,
        NULL AS PieceRef,
        NULL AS ItemRef,
        NULL AS PieceKeyOrder,
        NULL AS ItemKeyOrder,
        x.Covering_From_Date AS FirstDate,
        x.Covering_To_Date AS LastDate,
        h.Date_text AS CoveringDates,
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
        NULL AS  ClosureCode,
        NULL AS  ClosureType,
        NULL AS  ClosureStat
		
    FROM 
        ES_Lettercode lc
			JOIN ES_Class c ON c.lettercode_id = lc.lettercode_id
			JOIN ES_Header h ON h.class_id = c.class_id
			JOIN ES_Header_Extension x ON h.Header_ID = x.Header_ID
			LEFT JOIN ES_Division d ON @t_eCatStructure = 1 AND c.Division_ID = d.Division_ID
    WHERE 
        lc.Edit_set_ID = @t_iEditsetID
        AND c.Edit_set_ID = @t_iEditsetID
        AND h.Edit_set_ID = @t_iEditsetID
        AND x.Edit_set_ID = @t_iEditsetID
        AND (@t_eCatStructure = 0 AND (c.Division_ID IS NULL OR c.Division_ID = 0)
             OR @t_eCatStructure = 1 AND c.Division_ID IS NOT NULL AND c.Division_ID <> 0)
        AND (
        (@t_Editset_operation = 0 AND x.edit_set_entry_type_id IN (5, 4, 1, 2)) OR
        (@t_Editset_operation = 1 AND x.edit_set_entry_type_id = 6) OR
        (@t_Editset_operation = 2 AND x.edit_set_entry_type_id = 3)
    );
END

GO

