USE [ILDB]
GO

/****** Object:  StoredProcedure [pct].[ExportItems]    Script Date: 07/09/2026 15:20:46 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO



-- ========================================================================
-- Author:        Paul Davis / Chris Bartlett
-- Create date:   06/12/2024
-- Description:   Procedure to export items data for search indexing
-- ========================================================================
CREATE OR ALTER PROCEDURE [pct].[ExportItems]
    @t_iEditsetID AS INT,
    @t_Editset_operation AS INT, -- 0 = new, 1 = update, 2 = delete
    @t_eCatStructure AS INT -- CATPlain=0, CATWithDiv=1, CATWithHeader=2, CATWithSubheader=3, CATWithHeaderAndDiv=4, CATWithSubheaderAndDiv=5
AS
BEGIN
    -- Set NOCOUNT ON to prevent extra result sets
    SET NOCOUNT ON;

    -- Main SELECT Statement
    SELECT
		DISTINCT(i.item_id) AS ID,
        7 AS CatalogueLevel,
        lc.letter_code AS LetterCodeRef,

        -- Conditional Columns
        CASE 
            WHEN @t_eCatStructure IN (0, 2, 3) THEN NULL
            WHEN @t_eCatStructure IN (1, 4, 5) THEN d.Division_No
        END AS DivisionRef,
        
        c.class_no AS ClassRef,
        c.subclass_no AS SubclassRef,

        CASE 
            WHEN @t_eCatStructure IN (0, 1) THEN NULL
            WHEN @t_eCatStructure IN (2, 3, 4, 5) THEN h.class_hdr_no
        END AS HeaderRef,
        
        CASE 
            WHEN @t_eCatStructure IN (0, 1, 2, 4) THEN NULL
            WHEN @t_eCatStructure IN (3, 5) THEN sh.subheader_no
        END AS SubheaderRef,

		p.piece_ref AS PieceRef,
        i.item_ref AS ItemRef,
		p.piece_key_order AS PieceKeyOrder,
        i.item_key_order AS ItemKeyOrder,
        i.First_Date AS FirstDate,
        i.Last_Date AS LastDate,
        i.Date_text AS CoveringDates,
        x.map_designation AS MapDesig,
        x.Former_Reference_Department AS FormerRef,
        x.Former_Reference_PRO AS FormerPRORef,
        x.scale_number AS MapScale,
        x.restrictions_on_use  AS Restrictions,
        x.accumulation_date_text  AS AccumDates,
        i.open_date AS OpeningDate,
        x.legal_status_code  AS LegalStatus,
        x.physical_form_code AS PhyDescForm,
        x.physical_record_quantity AS Quantity,
        x.dimensions  AS Dimensions,
        x.access_condition_id  AS AccessCond,
        i.language_id  AS Language,
        x.physical_condition  AS PhysCond,
        i.closure_code  AS ClosureCode,
        i.closure_type  AS ClosureType,
        i.closure_status  AS ClosureStat
		
    FROM
        ES_Lettercode lc
			JOIN ES_Class c ON c.lettercode_id = lc.lettercode_id
			JOIN ES_Piece p ON p.class_id = c.class_id
			JOIN ES_item i ON i.piece_id = p.piece_id
			JOIN ES_item_Extension x ON x.item_id = i.item_id

        -- Conditional Joins
			LEFT JOIN ES_Division d ON @t_eCatStructure IN (1, 4, 5) AND c.Division_ID = d.Division_ID
			LEFT JOIN ES_Header h ON @t_eCatStructure IN (2, 3, 4, 5) AND p.Header_ID = h.Header_ID
			LEFT JOIN ES_Subheader sh ON @t_eCatStructure IN (3, 5) AND p.SubHeader_ID = sh.SubHeader_ID
    WHERE
        -- Edit Set Filters
        lc.Edit_set_ID = @t_iEditsetID
        AND c.Edit_set_ID = @t_iEditsetID
        AND p.Edit_set_ID = @t_iEditsetID
		AND i.Edit_set_ID = @t_iEditsetID
        AND x.Edit_set_ID = @t_iEditsetID
        
        -- Additional Filters for Joined Tables
        AND (@t_eCatStructure NOT IN (1, 4, 5) OR d.Edit_set_ID = @t_iEditsetID)
        AND (@t_eCatStructure NOT IN (2, 3, 4, 5) OR h.Edit_set_ID = @t_iEditsetID)
        AND (@t_eCatStructure NOT IN (3, 5) OR sh.Edit_set_ID = @t_iEditsetID)
        
        -- Mode Filters
        AND (
        (@t_Editset_operation = 0 AND x.edit_set_entry_type_id IN (5, 4, 1, 2)) OR
        (@t_Editset_operation = 1 AND x.edit_set_entry_type_id = 6) OR
        (@t_Editset_operation = 2 AND x.edit_set_entry_type_id = 3)
    );

    SET NOCOUNT OFF;
END;

GO

