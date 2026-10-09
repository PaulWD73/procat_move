USE [ILDB]
GO

SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

-- =====================================================================================
-- Author:      Paul Davis / Chris Bartlett
-- Description: Post-release export of descendants affected by catalogue moves.
--
-- A moved class causes live descendants at levels 4-7 to be refreshed.
-- A moved piece causes its live level-7 items to be refreshed.
-- The moved root itself is already present in the ordinary editset update CSV.
--
-- Output is the same 42-column contract consumed by DahuVerityBuilder.
-- =====================================================================================
CREATE OR ALTER PROCEDURE pct.ExportMoveGetAllData
    @t_iEditsetID INT
WITH EXECUTE AS OWNER
AS
BEGIN
    SET NOCOUNT ON;

    IF EXISTS
    (
        SELECT 1
          FROM dbo.EditorialMoveManifest
         WHERE EditSetID = @t_iEditsetID
           AND MoveStatus IN ('CAPTURED', 'FAILED')
    )
        THROW 51020, 'Move descendants cannot be exported before the release has completed.', 1;

    IF NOT EXISTS
    (
        SELECT 1
          FROM dbo.EditorialMoveManifest
         WHERE EditSetID = @t_iEditsetID
           AND MoveStatus IN ('RELEASED', 'EXPORTED')
    )
        THROW 51021, 'No released move-manifest entries were found for this editset.', 1;

    CREATE TABLE #Targets
    (
        CatalogueLevel TINYINT NOT NULL,
        CatalogueID    INT     NOT NULL,
        CONSTRAINT PK_MoveExportTargets PRIMARY KEY (CatalogueLevel, CatalogueID)
    );

    /* Level 4 headers below moved classes. */
    INSERT #Targets (CatalogueLevel, CatalogueID)
    SELECT DISTINCT 4, h.header_id
      FROM dbo.EditorialMoveManifest AS mm
      JOIN dbo.tbl_header AS h ON h.class_id = mm.CatalogueID
     WHERE mm.EditSetID = @t_iEditsetID
       AND mm.CatalogueLevel = 3
       AND mm.MoveStatus IN ('RELEASED', 'EXPORTED');

    /* Level 5 subheaders below moved classes. */
    INSERT #Targets (CatalogueLevel, CatalogueID)
    SELECT DISTINCT 5, sh.subheader_id
      FROM dbo.EditorialMoveManifest AS mm
      JOIN dbo.tbl_header AS h ON h.class_id = mm.CatalogueID
      JOIN dbo.tbl_subheader AS sh ON sh.header_id = h.header_id
     WHERE mm.EditSetID = @t_iEditsetID
       AND mm.CatalogueLevel = 3
       AND mm.MoveStatus IN ('RELEASED', 'EXPORTED');

    /* Level 6 pieces below moved classes. */
    INSERT #Targets (CatalogueLevel, CatalogueID)
    SELECT DISTINCT 6, p.piece_id
      FROM dbo.EditorialMoveManifest AS mm
      JOIN dbo.tbl_piece AS p ON p.class_id = mm.CatalogueID
     WHERE mm.EditSetID = @t_iEditsetID
       AND mm.CatalogueLevel = 3
       AND mm.MoveStatus IN ('RELEASED', 'EXPORTED');

    /* Level 7 items below moved classes and moved pieces. */
    INSERT #Targets (CatalogueLevel, CatalogueID)
    SELECT DISTINCT 7, i.item_id
      FROM dbo.EditorialMoveManifest AS mm
      JOIN dbo.tbl_piece AS p
        ON mm.CatalogueLevel = 3
       AND p.class_id = mm.CatalogueID
      JOIN dbo.tbl_item AS i ON i.piece_id = p.piece_id
     WHERE mm.EditSetID = @t_iEditsetID
       AND mm.MoveStatus IN ('RELEASED', 'EXPORTED')
       AND NOT EXISTS
           (SELECT 1 FROM #Targets AS t WHERE t.CatalogueLevel = 7 AND t.CatalogueID = i.item_id);

    INSERT #Targets (CatalogueLevel, CatalogueID)
    SELECT DISTINCT 7, i.item_id
      FROM dbo.EditorialMoveManifest AS mm
      JOIN dbo.tbl_item AS i ON i.piece_id = mm.CatalogueID
     WHERE mm.EditSetID = @t_iEditsetID
       AND mm.CatalogueLevel = 6
       AND mm.MoveStatus IN ('RELEASED', 'EXPORTED')
       AND NOT EXISTS
           (SELECT 1 FROM #Targets AS t WHERE t.CatalogueLevel = 7 AND t.CatalogueID = i.item_id);

    /*
       Read the post-release hierarchy exclusively from live tables.  LEFT JOINs for
       optional text rows deliberately produce blank CSV fields rather than dropping
       an otherwise valid catalogue record.
    */
    SELECT d.*
      INTO #MoveData
      FROM
      (
        /* Headers (level 4). */
        SELECT
            h.header_id AS ID, CAST(4 AS INT) AS CatalogueLevel,
            lc.letter_code AS LetterCodeRef,
            CASE WHEN ISNULL(c.division_id, 0) = 0 THEN NULL ELSE dv.division_no END AS DivisionRef,
            c.class_no AS ClassRef, c.subclass_no AS SubClassRef,
            h.class_hdr_no AS HeaderRef, CAST(NULL AS INT) AS SubheaderRef,
            CAST(NULL AS VARCHAR(40)) AS PieceRef, CAST(NULL AS VARCHAR(40)) AS ItemRef,
            CAST(NULL AS INT) AS PieceKeyOrder, CAST(NULL AS INT) AS ItemKeyOrder,
            x.covering_from_date AS FirstDate, x.covering_to_date AS LastDate,
            h.date_text AS CoveringDates, x.map_designation AS MapDesig,
            x.Former_Reference_Department AS FormerRef, x.Former_Reference_PRO AS FormerPRORef,
            x.scale_number AS MapScale, x.restrictions_on_use AS Restrictions,
            x.accumulation_date_text AS AccumDates, x.record_opening_date AS OpeningDate,
            x.legal_status_code AS LegalStatus, x.physical_form_code AS PhyDescForm,
            x.physical_record_quantity AS Quantity, x.dimensions AS Dimensions,
            x.access_condition_id AS AccessCond, x.language_id AS [Language],
            x.physical_condition AS PhysCond,
            CAST(NULL AS INT) AS ClosureCode, CAST(NULL AS CHAR(1)) AS ClosureType,
            CAST(NULL AS CHAR(1)) AS ClosureStat,
            sc.scope_content_text AS ScopeContent, ti.title_text AS TitleContent,
            ah.admin_biog_hist_text AS AdminHistory, ac.accruals_text AS Accruals,
            ap.app_dest_info_text AS AppraisalInfo, ar.arrangement_text AS Arrangement,
            ch.custodial_hist_text AS CustHist, nt.note_text AS NoteText
          FROM #Targets AS target
          JOIN dbo.tbl_header AS h ON target.CatalogueLevel = 4 AND h.header_id = target.CatalogueID
          JOIN dbo.header_extension AS x ON x.header_id = h.header_id AND x.live_flag IN (0, 1)
          JOIN dbo.tbl_class AS c ON c.class_id = h.class_id
          JOIN dbo.tbl_lettercode AS lc ON lc.lettercode_id = c.lettercode_id
          LEFT JOIN dbo.tbl_division AS dv ON dv.division_id = c.division_id AND ISNULL(c.division_id, 0) <> 0
          LEFT JOIN dbo.header_scope_content AS sc ON sc.header_id = h.header_id
          LEFT JOIN dbo.header_title AS ti ON ti.header_id = h.header_id
          LEFT JOIN dbo.header_admin_biog_hist AS ah ON ah.header_id = h.header_id
          LEFT JOIN dbo.header_accruals AS ac ON ac.header_id = h.header_id
          LEFT JOIN dbo.header_app_dest_info AS ap ON ap.header_id = h.header_id
          LEFT JOIN dbo.header_arrangement AS ar ON ar.header_id = h.header_id
          LEFT JOIN dbo.header_custodial_hist AS ch ON ch.header_id = h.header_id
          LEFT JOIN dbo.header_note AS nt ON nt.header_id = h.header_id

        UNION ALL

        /* Subheaders (level 5). */
        SELECT
            sh.subheader_id, 5, lc.letter_code,
            CASE WHEN ISNULL(c.division_id, 0) = 0 THEN NULL ELSE dv.division_no END,
            c.class_no, c.subclass_no, h.class_hdr_no, sh.subheader_no,
            NULL, NULL, NULL, NULL,
            x.covering_from_date, x.covering_to_date, sh.date_text,
            x.map_designation, x.Former_Reference_Department, x.Former_Reference_PRO,
            x.scale_number, x.restrictions_on_use, x.accumulation_date_text,
            x.record_opening_date, x.legal_status_code, x.physical_form_code,
            x.physical_record_quantity, x.dimensions, x.access_condition_id,
            x.language_id, x.physical_condition, NULL, NULL, NULL,
            sc.scope_content_text, ti.title_text, ah.admin_biog_hist_text,
            ac.accruals_text, ap.app_dest_info_text, ar.arrangement_text,
            ch.custodial_hist_text, nt.note_text
          FROM #Targets AS target
          JOIN dbo.tbl_subheader AS sh ON target.CatalogueLevel = 5 AND sh.subheader_id = target.CatalogueID
          JOIN dbo.subheader_extension AS x ON x.subheader_id = sh.subheader_id AND x.live_flag IN (0, 1)
          JOIN dbo.tbl_header AS h ON h.header_id = sh.header_id
          JOIN dbo.tbl_class AS c ON c.class_id = h.class_id
          JOIN dbo.tbl_lettercode AS lc ON lc.lettercode_id = c.lettercode_id
          LEFT JOIN dbo.tbl_division AS dv ON dv.division_id = c.division_id AND ISNULL(c.division_id, 0) <> 0
          LEFT JOIN dbo.subheader_scope_content AS sc ON sc.subheader_id = sh.subheader_id
          LEFT JOIN dbo.subheader_title AS ti ON ti.subheader_id = sh.subheader_id
          LEFT JOIN dbo.subheader_admin_biog_hist AS ah ON ah.subheader_id = sh.subheader_id
          LEFT JOIN dbo.subheader_accruals AS ac ON ac.subheader_id = sh.subheader_id
          LEFT JOIN dbo.subheader_app_dest_info AS ap ON ap.subheader_id = sh.subheader_id
          LEFT JOIN dbo.subheader_arrangement AS ar ON ar.subheader_id = sh.subheader_id
          LEFT JOIN dbo.subheader_custodial_hist AS ch ON ch.subheader_id = sh.subheader_id
          LEFT JOIN dbo.subheader_note AS nt ON nt.subheader_id = sh.subheader_id

        UNION ALL

        /* Pieces (level 6). */
        SELECT
            p.piece_id, 6, lc.letter_code,
            CASE WHEN ISNULL(c.division_id, 0) = 0 THEN NULL ELSE dv.division_no END,
            c.class_no, c.subclass_no,
            CASE WHEN ISNULL(p.header_id, 0) = 0 THEN NULL ELSE h.class_hdr_no END,
            CASE WHEN ISNULL(p.subheader_id, 0) = 0 THEN NULL ELSE sh.subheader_no END,
            p.piece_ref, NULL, p.piece_key_order, NULL,
            p.first_date, p.last_date, p.date_text,
            x.map_designation, x.Former_Reference_Department, x.Former_Reference_PRO,
            x.scale_number, x.restrictions_on_use, x.accumulation_date_text,
            p.open_date, x.legal_status_code, x.physical_form_code,
            x.physical_record_quantity, x.dimensions, x.access_condition_id,
            p.language_id, x.physical_condition,
            p.closure_code, p.closure_type, p.closure_status,
            sc.scope_content_text, ti.title_text, ah.admin_biog_hist_text,
            ac.accruals_text, ap.app_dest_info_text, ar.arrangement_text,
            ch.custodial_hist_text, nt.note_text
          FROM #Targets AS target
          JOIN dbo.tbl_piece AS p ON target.CatalogueLevel = 6 AND p.piece_id = target.CatalogueID
          JOIN dbo.piece_extension AS x ON x.piece_id = p.piece_id AND x.live_flag IN (0, 1)
          JOIN dbo.tbl_class AS c ON c.class_id = p.class_id
          JOIN dbo.tbl_lettercode AS lc ON lc.lettercode_id = c.lettercode_id
          LEFT JOIN dbo.tbl_division AS dv ON dv.division_id = c.division_id AND ISNULL(c.division_id, 0) <> 0
          LEFT JOIN dbo.tbl_header AS h ON h.header_id = p.header_id AND ISNULL(p.header_id, 0) <> 0
          LEFT JOIN dbo.tbl_subheader AS sh ON sh.subheader_id = p.subheader_id AND ISNULL(p.subheader_id, 0) <> 0
          LEFT JOIN dbo.piece_scope_content AS sc ON sc.piece_id = p.piece_id
          LEFT JOIN dbo.piece_title AS ti ON ti.piece_id = p.piece_id
          LEFT JOIN dbo.piece_admin_biog_hist AS ah ON ah.piece_id = p.piece_id
          LEFT JOIN dbo.piece_accruals AS ac ON ac.piece_id = p.piece_id
          LEFT JOIN dbo.piece_app_dest_info AS ap ON ap.piece_id = p.piece_id
          LEFT JOIN dbo.piece_arrangement AS ar ON ar.piece_id = p.piece_id
          LEFT JOIN dbo.piece_custodial_hist AS ch ON ch.piece_id = p.piece_id
          LEFT JOIN dbo.piece_note AS nt ON nt.piece_id = p.piece_id

        UNION ALL

        /* Items (level 7). */
        SELECT
            i.item_id, 7, lc.letter_code,
            CASE WHEN ISNULL(c.division_id, 0) = 0 THEN NULL ELSE dv.division_no END,
            c.class_no, c.subclass_no,
            CASE WHEN ISNULL(p.header_id, 0) = 0 THEN NULL ELSE h.class_hdr_no END,
            CASE WHEN ISNULL(p.subheader_id, 0) = 0 THEN NULL ELSE sh.subheader_no END,
            p.piece_ref, i.item_ref, p.piece_key_order, i.item_key_order,
            i.first_date, i.last_date, i.date_text,
            x.map_designation, x.Former_Reference_Department, x.Former_Reference_PRO,
            x.scale_number, x.restrictions_on_use, x.accumulation_date_text,
            i.open_date, x.legal_status_code, x.physical_form_code,
            x.physical_record_quantity, x.dimensions, x.access_condition_id,
            i.language_id, x.physical_condition,
            i.closure_code, i.closure_type, i.closure_status,
            sc.scope_content_text, ti.title_text, ah.admin_biog_hist_text,
            ac.accruals_text, ap.app_dest_info_text, ar.arrangement_text,
            ch.custodial_hist_text, nt.note_text
          FROM #Targets AS target
          JOIN dbo.tbl_item AS i ON target.CatalogueLevel = 7 AND i.item_id = target.CatalogueID
          JOIN dbo.item_extension AS x ON x.item_id = i.item_id AND x.live_flag IN (0, 1)
          JOIN dbo.tbl_piece AS p ON p.piece_id = i.piece_id
          JOIN dbo.tbl_class AS c ON c.class_id = p.class_id
          JOIN dbo.tbl_lettercode AS lc ON lc.lettercode_id = c.lettercode_id
          LEFT JOIN dbo.tbl_division AS dv ON dv.division_id = c.division_id AND ISNULL(c.division_id, 0) <> 0
          LEFT JOIN dbo.tbl_header AS h ON h.header_id = p.header_id AND ISNULL(p.header_id, 0) <> 0
          LEFT JOIN dbo.tbl_subheader AS sh ON sh.subheader_id = p.subheader_id AND ISNULL(p.subheader_id, 0) <> 0
          LEFT JOIN dbo.item_scope_content AS sc ON sc.item_id = i.item_id
          LEFT JOIN dbo.item_title AS ti ON ti.item_id = i.item_id
          LEFT JOIN dbo.item_admin_biog_hist AS ah ON ah.item_id = i.item_id
          LEFT JOIN dbo.item_accruals AS ac ON ac.item_id = i.item_id
          LEFT JOIN dbo.item_app_dest_info AS ap ON ap.item_id = i.item_id
          LEFT JOIN dbo.item_arrangement AS ar ON ar.item_id = i.item_id
          LEFT JOIN dbo.item_custodial_hist AS ch ON ch.item_id = i.item_id
          LEFT JOIN dbo.item_note AS nt ON nt.item_id = i.item_id
      ) AS d;

    CREATE TABLE #LinkData
    (
        LinkDataID    VARCHAR(50) NOT NULL,
        ID            INT         NOT NULL,
        CatalogueLevel INT        NOT NULL,
        LinkTypeID    INT         NOT NULL
    );

    INSERT #LinkData (LinkDataID, ID, CatalogueLevel, LinkTypeID)
    SELECT 'PE' + CONVERT(VARCHAR(40), p.Person_Reference_Id), p.Catalogue_ID, p.level_no, p.Property_ID
      FROM dbo.PC_Person AS p
      JOIN #Targets AS t ON t.CatalogueLevel = p.level_no AND t.CatalogueID = p.Catalogue_ID
     WHERE p.Property_ID IN (12, 18, 16)
    UNION ALL
    SELECT 'CO' + CONVERT(VARCHAR(40), c.Corporate_Body_Reference_ID), c.Catalogue_ID, c.level_no, c.Property_ID
      FROM dbo.PC_Corp_Body AS c
      JOIN #Targets AS t ON t.CatalogueLevel = c.level_no AND t.CatalogueID = c.Catalogue_ID
     WHERE c.Property_ID IN (12, 17, 32, 24, 10, 16)
    UNION ALL
    SELECT 'PL' + CONVERT(VARCHAR(40), p.Place_Reference_ID), p.Catalogue_ID, p.level_no, p.Property_ID
      FROM dbo.PC_Place AS p
      JOIN #Targets AS t ON t.CatalogueLevel = p.level_no AND t.CatalogueID = p.Catalogue_ID
     WHERE p.Property_ID = 19
    UNION ALL
    SELECT 'SU' + CONVERT(VARCHAR(40), s.Subject_Reference_ID), s.Catalogue_ID, s.level_no, s.Property_ID
      FROM dbo.PC_Subject AS s
      JOIN #Targets AS t ON t.CatalogueLevel = s.level_no AND t.CatalogueID = s.Catalogue_ID
     WHERE s.Property_ID = 20;

    /* Header row followed by Builder-compatible data rows. */
    SELECT 'ID', 'CatalogueLevel', 'LetterCodeRef', 'DivisionRef', 'ClassRef',
           'SubClassRef', 'HeaderRef', 'SubheaderRef', 'PieceRef', 'ItemRef',
           'PieceKeyOrder', 'ItemKeyOrder', 'FirstDate', 'LastDate', 'CoveringDates',
           'MapDesig', 'FormerRef', 'FormerPRORef', 'MapScale', 'Restrictions',
           'AccumDates', 'OpeningDate', 'LegalStatus', 'PhyDescForm', 'Quantity',
           'Dimensions', 'AccessCond', 'Language', 'PhysCond', 'ClosureCode',
           'ClosureType', 'ClosureStat', 'ScopeContent', 'TitleContent',
           'AdminHistory', 'Accruals', 'AppraisalInfo', 'Arrangement', 'CustHist',
           'NoteText', 'LinkDataID', 'LinkTypeID'
    UNION ALL
    SELECT
        NULLIF(CONVERT(VARCHAR(MAX), m.ID), ''),
        NULLIF(CONVERT(VARCHAR(MAX), m.CatalogueLevel), ''),
        NULLIF(CONVERT(VARCHAR(MAX), m.LetterCodeRef), ''),
        NULLIF(CONVERT(VARCHAR(MAX), m.DivisionRef), ''),
        NULLIF(CONVERT(VARCHAR(MAX), m.ClassRef), ''),
        NULLIF(CONVERT(VARCHAR(MAX), m.SubClassRef), ''),
        NULLIF(CONVERT(VARCHAR(MAX), m.HeaderRef), ''),
        NULLIF(CONVERT(VARCHAR(MAX), m.SubheaderRef), ''),
        NULLIF(CONVERT(VARCHAR(MAX), m.PieceRef), ''),
        NULLIF(CONVERT(VARCHAR(MAX), m.ItemRef), ''),
        NULLIF(CONVERT(VARCHAR(MAX), m.PieceKeyOrder), ''),
        NULLIF(CONVERT(VARCHAR(MAX), m.ItemKeyOrder), ''),
        NULLIF(CONVERT(VARCHAR(MAX), m.FirstDate), ''),
        NULLIF(CONVERT(VARCHAR(MAX), m.LastDate), ''),
        NULLIF(CONVERT(VARCHAR(MAX), m.CoveringDates), ''),
        NULLIF(CONVERT(VARCHAR(MAX), m.MapDesig), ''),
        NULLIF(CONVERT(VARCHAR(MAX), m.FormerRef), ''),
        NULLIF(CONVERT(VARCHAR(MAX), m.FormerPRORef), ''),
        NULLIF(CONVERT(VARCHAR(MAX), m.MapScale), ''),
        NULLIF(CONVERT(VARCHAR(MAX), m.Restrictions), ''),
        NULLIF(CONVERT(VARCHAR(MAX), m.AccumDates), ''),
        NULLIF(CONVERT(VARCHAR(MAX), m.OpeningDate), ''),
        NULLIF(CONVERT(VARCHAR(MAX), m.LegalStatus), ''),
        NULLIF(CONVERT(VARCHAR(MAX), m.PhyDescForm), ''),
        NULLIF(CONVERT(VARCHAR(MAX), m.Quantity), ''),
        NULLIF(CONVERT(VARCHAR(MAX), m.Dimensions), ''),
        NULLIF(CONVERT(VARCHAR(MAX), m.AccessCond), ''),
        NULLIF(CONVERT(VARCHAR(MAX), m.[Language]), ''),
        NULLIF(CONVERT(VARCHAR(MAX), m.PhysCond), ''),
        NULLIF(CONVERT(VARCHAR(MAX), m.ClosureCode), ''),
        NULLIF(CONVERT(VARCHAR(MAX), m.ClosureType), ''),
        NULLIF(CONVERT(VARCHAR(MAX), m.ClosureStat), ''),
        NULLIF(REPLACE(REPLACE(CONVERT(VARCHAR(MAX), m.ScopeContent), CHAR(13), ''), CHAR(10), ''), ''),
        NULLIF(REPLACE(REPLACE(CONVERT(VARCHAR(MAX), m.TitleContent), CHAR(13), ''), CHAR(10), ''), ''),
        NULLIF(REPLACE(REPLACE(CONVERT(VARCHAR(MAX), m.AdminHistory), CHAR(13), ''), CHAR(10), ''), ''),
        NULLIF(REPLACE(REPLACE(CONVERT(VARCHAR(MAX), m.Accruals), CHAR(13), ''), CHAR(10), ''), ''),
        NULLIF(REPLACE(REPLACE(CONVERT(VARCHAR(MAX), m.AppraisalInfo), CHAR(13), ''), CHAR(10), ''), ''),
        NULLIF(REPLACE(REPLACE(CONVERT(VARCHAR(MAX), m.Arrangement), CHAR(13), ''), CHAR(10), ''), ''),
        NULLIF(REPLACE(REPLACE(CONVERT(VARCHAR(MAX), m.CustHist), CHAR(13), ''), CHAR(10), ''), ''),
        NULLIF(REPLACE(REPLACE(CONVERT(VARCHAR(MAX), m.NoteText), CHAR(13), ''), CHAR(10), ''), ''),
        NULLIF(CONVERT(VARCHAR(MAX), l.LinkDataID), ''),
        NULLIF(CONVERT(VARCHAR(MAX), l.LinkTypeID), '')
      FROM #MoveData AS m
      LEFT JOIN #LinkData AS l
        ON l.CatalogueLevel = m.CatalogueLevel
       AND l.ID = m.ID;
END
GO
