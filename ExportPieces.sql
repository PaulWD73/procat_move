USE [ILDB]
GO

/****** Object:  StoredProcedure [pct].[ExportPieces]    Script Date: 07/09/2026 15:24:07 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO




-- ======================================================================
-- Author:        Paul Davis / Chris Bartlett
-- Create date:   06/12/2024
-- Description:   Procedure to export pieces data for search indexing
-- ======================================================================
CREATE OR ALTER PROCEDURE [pct].[ExportPieces]
    @t_iEditsetID AS INT,
    @t_Editset_operation AS INT, -- 0 = new, 1 = update, 2 = delete
    @t_eCatStructure AS INT -- 0 = CatPlain, 1 = CATWithDiv
AS
BEGIN
    -- Main SELECT Statement
    SELECT
		DISTINCT(p.piece_id) AS ID,
        6 AS CatalogueLevel,
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
        NULL AS ItemRef,
        p.piece_key_order AS PieceKeyOrder,
        NULL AS ItemKeyOrder,
        p.First_Date AS FirstDate,
        p.Last_Date AS LastDate,
        p.Date_text AS CoveringDates,
        x.map_designation AS MapDesig,
        x.Former_Reference_Department AS FormerRef,
        x.Former_Reference_PRO AS FormerPRORef,
        x.scale_number AS MapScale,
        x.restrictions_on_use AS Restrictions,
        x.accumulation_date_text AS AccumDates,
        p.open_date AS OpeningDate,
        x.legal_status_code AS LegalStatus,
        x.physical_form_code AS PhyDescForm,
        x.physical_record_quantity AS Quantity,
        x.dimensions AS Dimensions,
        x.access_condition_id AS AccessCond,
        p.language_id AS Language,
        x.physical_condition AS PhysCond,
        p.closure_code AS ClosureCode,
        p.closure_type AS ClosureType,
        p.closure_status AS ClosureStat
		
		
    FROM
        ES_Lettercode lc
			JOIN ES_Class c ON c.lettercode_id = lc.lettercode_id
			JOIN ES_Piece p ON p.class_id = c.class_id
			JOIN ES_Piece_Extension x ON x.piece_id = p.piece_id

        -- Conditional Joins
			LEFT JOIN ES_Division d ON @t_eCatStructure IN (1, 4, 5) AND c.Division_ID = d.Division_ID
			LEFT JOIN ES_Header h ON @t_eCatStructure IN (2, 3, 4, 5) AND p.Header_ID = h.Header_ID
			LEFT JOIN ES_Subheader sh ON @t_eCatStructure IN (3, 5) AND p.SubHeader_ID = sh.SubHeader_ID
    WHERE
        -- Edit Set Filters
        lc.Edit_set_ID = @t_iEditsetID
        AND c.Edit_set_ID = @t_iEditsetID
        AND p.Edit_set_ID = @t_iEditsetID
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

END;

GO

