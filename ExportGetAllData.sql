USE [ILDB]
GO

/****** Object:  StoredProcedure [pct].[ExportGetAllData]    Script Date: 07/09/2026 15:20:15 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


-- ==================================================================================
-- Author:		Paul Davis/ Chris Bartlett
-- Create date: 06/12/2024
-- Description:	Procedure to start PROCAT Export all data for search indexing	
-- ==================================================================================

CREATE OR ALTER PROCEDURE [pct].[ExportGetAllData]
	@t_iEditsetID AS INT,
	@t_Editset_operation INT, -- 0=new, 1=update, 2=delete
	--@t_CollectionType INT, -- 0=CollAdvanced, 1=CollClassUp
	@t_Date AS VARCHAR(30) --todo: create it upstream for the editset - deals with crossing midnight boundaries when writing editsets?)
WITH EXECUTE AS OWNER 

AS	
BEGIN
SET NOCOUNT ON;
-- TODO: need to figure out data types from schema. Refer to individual SPs for table names/column names
	DECLARE @MainData TABLE (
	    ID INT, 
		CatalogueLevel INT,
	    LetterCodeRef VARCHAR(4), 
	    DivisionRef INT, 
	    ClassRef INT, 
	    SubClassRef INT, 
		HeaderRef INT,
	    SubheaderRef INT, 
	    PieceRef VARCHAR(40), 
	    ItemRef VARCHAR(40), 
	    PieceKeyOrder INT, 
	    ItemKeyOrder INT, 
	    FirstDate INT, 
	    LastDate INT, 
	    CoveringDates VARCHAR(255), 
	    MapDesig VARCHAR(100), 
	    FormerRef VARCHAR(255), 
	    FormerPRORef VARCHAR(255),
		MapScale INT,
		Restrictions VARCHAR(500),
		AccumDates VARCHAR(255),
		OpeningDate DATETIME,
		LegalStatus SMALLINT,
		PhyDescForm SMALLINT,
		Quantity INT,
		Dimensions VARCHAR(20),
		AccessCond SMALLINT,
		Language SMALLINT,
		PhysCond VARCHAR(255),
		ClosureCode INT,
		ClosureType CHAR(1),
		ClosureStat CHAR(1)
	);
	
	DECLARE @ScopeData TABLE (
		CatalogueLevel INT,
        ID INT, 
        ScopeContent VARCHAR(8000)
	)
	DECLARE @TitleData TABLE (
		CatalogueLevel INT,
        ID INT, 
        TitleContent VARCHAR(8000)
	)
	DECLARE @AdminData TABLE (
		CatalogueLevel INT,
        ID INT, 
        AdminHistory VARCHAR(8000)
	)
	DECLARE @AccrualData TABLE (
		CatalogueLevel INT,
        ID INT, 
        Accruals VARCHAR(8000)
	)
	DECLARE @AppInfoData TABLE (
		CatalogueLevel INT,
        ID INT, 
        AppraisalInfo VARCHAR(8000)
	)
	DECLARE @ArrangementData TABLE (
		CatalogueLevel INT,
        ID INT, 
        Arrangement VARCHAR(8000)
	)
	DECLARE @CustHisttData TABLE (
		CatalogueLevel INT,
        ID INT, 
        CustHist VARCHAR(8000)
	)
	
	DECLARE @NoteData TABLE (
		CatalogueLevel INT,
        ID INT, 
        NoteText VARCHAR(8000)
	)
	
	DECLARE @LinkData TABLE(	
		LinkDataID VARCHAR(20),
        ID INT,
		CatalogueLevel INT,
        LinkTypeID VARCHAR(20)
	)
	
	-- Insert from each sub-query
	INSERT INTO @MainData EXEC pct.ExportLettercode @t_iEditsetID, @t_Editset_operation

	INSERT INTO @MainData EXEC pct.ExportDivision @t_iEditsetID, @t_Editset_operation
	
	INSERT INTO @MainData EXEC pct.ExportClass  @t_iEditsetID, @t_Editset_operation, 0
	INSERT INTO @MainData EXEC pct.ExportClass  @t_iEditsetID, @t_Editset_operation, 1

	INSERT INTO @MainData EXEC pct.ExportHeader  @t_iEditsetID, @t_Editset_operation, 0
	INSERT INTO @MainData EXEC pct.ExportHeader  @t_iEditsetID, @t_Editset_operation, 1

	INSERT INTO @MainData EXEC pct.ExportSubheader  @t_iEditsetID, @t_Editset_operation, 0
	INSERT INTO @MainData EXEC pct.ExportSubheader  @t_iEditsetID, @t_Editset_operation, 1

	INSERT INTO @MainData EXEC pct.ExportPieces @t_iEditsetID, @t_Editset_operation, 0
	INSERT INTO @MainData EXEC pct.ExportPieces @t_iEditsetID, @t_Editset_operation, 1
	INSERT INTO @MainData EXEC pct.ExportPieces @t_iEditsetID, @t_Editset_operation, 2
	INSERT INTO @MainData EXEC pct.ExportPieces @t_iEditsetID, @t_Editset_operation, 3
	INSERT INTO @MainData EXEC pct.ExportPieces @t_iEditsetID, @t_Editset_operation, 4
	INSERT INTO @MainData EXEC pct.ExportPieces @t_iEditsetID, @t_Editset_operation, 5

	INSERT INTO @MainData EXEC pct.ExportItems @t_iEditsetID, @t_Editset_operation, 0
	INSERT INTO @MainData EXEC pct.ExportItems @t_iEditsetID, @t_Editset_operation, 1
	INSERT INTO @MainData EXEC pct.ExportItems @t_iEditsetID, @t_Editset_operation, 2
	INSERT INTO @MainData EXEC pct.ExportItems @t_iEditsetID, @t_Editset_operation, 3
	INSERT INTO @MainData EXEC pct.ExportItems @t_iEditsetID, @t_Editset_operation, 4
	INSERT INTO @MainData EXEC pct.ExportItems @t_iEditsetID, @t_Editset_operation, 5

	INSERT INTO @ScopeData EXEC pct.ExportLargeScopeContent @t_iEditsetID, @t_Editset_operation
	INSERT INTO @TitleData EXEC pct.ExportLargeTitle @t_iEditsetID, @t_Editset_operation
	INSERT INTO @AdminData EXEC pct.ExportLargeAdminHistory @t_iEditsetID, @t_Editset_operation
	INSERT INTO @AccrualData EXEC pct.ExportLargeAccruals @t_iEditsetID, @t_Editset_operation
	INSERT INTO @AppInfoData EXEC pct.ExportLargeAppraisalInfo @t_iEditsetID, @t_Editset_operation
	INSERT INTO @ArrangementData EXEC pct.ExportLargeArrangement @t_iEditsetID, @t_Editset_operation
	INSERT INTO @CustHisttData EXEC pct.ExportLargeCustodialHistory @t_iEditsetID, @t_Editset_operation
	INSERT INTO @NoteData EXEC pct.ExportLargeNote @t_iEditsetID, @t_Editset_operation

	INSERT INTO @LinkData EXEC pct.ExportLargeLinks @t_iEditsetID, @t_Editset_operation


	SELECT m.ID, m.CatalogueLevel, m.LetterCodeRef, m.DivisionRef, m.ClassRef, m.SubClassRef, m.HeaderRef, m.SubheaderRef, m.PieceRef, m.ItemRef, m.PieceKeyOrder, 
			m.ItemKeyOrder, m.FirstDate, m.LastDate, m.CoveringDates, m.MapDesig, m.FormerRef, m.FormerPRORef, m.MapScale, m.Restrictions, m.AccumDates, m.OpeningDate,
			m.LegalStatus, m.PhyDescForm, m.Quantity, m.Dimensions, m.AccessCond, m.Language, m.PhysCond, m.ClosureCode, m.ClosureType, m.ClosureStat, s.ScopeContent, t.TitleContent,
			ah.AdminHistory, ac.Accruals, ap.AppraisalInfo, ar.Arrangement, c.CustHist, n.NoteText, l.LinkDataID, l.LinkTypeID 
	INTO #TempTable
	FROM (SELECT ID, CatalogueLevel, LetterCodeRef, DivisionRef, ClassRef, SubClassRef, HeaderRef, SubheaderRef, PieceRef, ItemRef, PieceKeyOrder, 
					ItemKeyOrder, FirstDate, LastDate, CoveringDates, MapDesig, FormerRef, FormerPRORef, MapScale, Restrictions, AccumDates, OpeningDate,
						LegalStatus, PhyDescForm, Quantity, Dimensions, AccessCond, Language, PhysCond, ClosureCode, ClosureType, ClosureStat 
		  FROM @MainData) m
				JOIN (SELECT CatalogueLevel, ID, ScopeContent FROM @ScopeData) s
					ON m.CatalogueLevel = s.CatalogueLevel AND m.ID = s.ID
				JOIN (SELECT CatalogueLevel, ID, TitleContent FROM @TitleData) t
					ON m.CatalogueLevel = t.CatalogueLevel AND m.ID = t.ID
				JOIN (SELECT CatalogueLevel, ID, AdminHistory FROM @AdminData) ah
					ON m.CatalogueLevel = ah.CatalogueLevel AND m.ID = ah.ID
				JOIN (SELECT CatalogueLevel, ID, Accruals FROM @AccrualData) ac
					ON m.CatalogueLevel = ac.CatalogueLevel AND m.ID = ac.ID
				JOIN (SELECT CatalogueLevel, ID, AppraisalInfo FROM @AppInfoData) ap
					ON m.CatalogueLevel = ap.CatalogueLevel AND m.ID = ap.ID
				JOIN (SELECT CatalogueLevel, ID, Arrangement FROM @ArrangementData) ar
					ON m.CatalogueLevel = ar.CatalogueLevel AND m.ID = ar.ID
				JOIN (SELECT CatalogueLevel, ID, CustHist FROM @CustHisttData) c
					ON m.CatalogueLevel = c.CatalogueLevel AND m.ID = c.ID
				JOIN (SELECT CatalogueLevel, ID, NoteText FROM @NoteData) n
					ON m.CatalogueLevel = n.CatalogueLevel AND m.ID = n.ID
				LEFT OUTER JOIN (SELECT CatalogueLevel, ID, LinkTypeID, LinkDataID FROM @LinkData) l
					ON m.CatalogueLevel = l.CatalogueLevel AND m.ID = l.ID


	SELECT 'ID', 'CatalogueLevel','LetterCodeRef', 'DivisionRef', 'ClassRef', 'SubClassRef', 'HeaderRef', 'SubheaderRef' , 'PieceRef' , 'ItemRef' , 'PieceKeyOrder' , 
			'ItemKeyOrder' , 'FirstDate' , 'LastDate' , 'CoveringDates', 'MapDesig' , 'FormerRef', 'FormerPRORef' ,'MapScale' ,'Restrictions' ,'AccumDates','OpeningDate',
			'LegalStatus','PhyDescForm' ,' Quantity','Dimensions','AccessCond','Language','PhysCond','ClosureCode' ,'ClosureType','ClosureStat', 'ScopeContent', 'TitleContent',
			'AdminHistory', 'Accruals', 'AppraisalInfo', 'Arrangement', 'CustHist', 'NoteText', 'LinkDataID','LinkTypeID'
	UNION ALL
	SELECT	NULLIF(CAST(ID AS VARCHAR(MAX)), ''), 
			NULLIF(CAST(CatalogueLevel AS VARCHAR(MAX)), ''), 
			NULLIF(CAST(LetterCodeRef AS VARCHAR(MAX)), ''), 
			NULLIF(CAST(DivisionRef AS VARCHAR(MAX)), ''),  
			NULLIF(CAST(ClassRef AS VARCHAR(MAX)), ''), 
			NULLIF(CAST(SubClassRef AS VARCHAR(MAX)), ''), 
			NULLIF(CAST(HeaderRef AS VARCHAR(MAX)), ''),
			NULLIF(CAST(SubheaderRef AS VARCHAR(MAX)), ''), 
			NULLIF(CAST(PieceRef AS VARCHAR(MAX)), ''), 
			NULLIF(CAST(ItemRef AS VARCHAR(MAX)), ''), 
			NULLIF(CAST(PieceKeyOrder AS VARCHAR(MAX)), ''), 
			NULLIF(CAST(ItemKeyOrder AS VARCHAR(MAX)), ''), 
			NULLIF(CAST(FirstDate AS VARCHAR(MAX)), ''), 
			NULLIF(CAST(LastDate AS VARCHAR(MAX)), ''), 
			NULLIF(CAST(CoveringDates AS VARCHAR(MAX)), ''), 
			NULLIF(CAST(MapDesig AS VARCHAR(MAX)), ''), 
			NULLIF(CAST(FormerRef AS VARCHAR(MAX)), ''), 
			NULLIF(CAST(FormerPRORef AS VARCHAR(MAX)), ''), 
			NULLIF(CAST(MapScale AS VARCHAR(MAX)), ''), 
			NULLIF(CAST(Restrictions AS VARCHAR(MAX)), ''), 
			NULLIF(CAST(AccumDates AS VARCHAR(MAX)), ''), 
			NULLIF(CAST(OpeningDate AS VARCHAR(MAX)), ''), 
			NULLIF(CAST(LegalStatus AS VARCHAR(MAX)), ''), 
			NULLIF(CAST(PhyDescForm AS VARCHAR(MAX)), ''), 
			NULLIF(CAST(Quantity AS VARCHAR(MAX)), ''), 
			NULLIF(CAST(Dimensions AS VARCHAR(MAX)), ''), 
			NULLIF(CAST(AccessCond AS VARCHAR(MAX)), ''), 
			NULLIF(CAST(Language AS VARCHAR(MAX)), ''), 
			NULLIF(CAST(PhysCond AS VARCHAR(MAX)), ''), 
			NULLIF(CAST(ClosureCode AS VARCHAR(MAX)), ''), 
			NULLIF(CAST(ClosureType AS VARCHAR(MAX)), ''),
			NULLIF(CAST(ClosureStat AS VARCHAR(MAX)), ''), 
			NULLIF(REPLACE(REPLACE(CAST(ScopeContent AS VARCHAR(MAX)),CHAR(13), ''), CHAR(10), ''), ''), 
			NULLIF(REPLACE(REPLACE(CAST(TitleContent AS VARCHAR(MAX)),CHAR(13), ''), CHAR(10), ''), ''),
			NULLIF(REPLACE(REPLACE(CAST(AdminHistory AS VARCHAR(MAX)),CHAR(13), ''), CHAR(10), ''), ''), 
			NULLIF(REPLACE(REPLACE(CAST(Accruals AS VARCHAR(MAX)),CHAR(13), ''), CHAR(10), ''), ''), 
			NULLIF(REPLACE(REPLACE(CAST(AppraisalInfo AS VARCHAR(MAX)),CHAR(13), ''), CHAR(10), ''), ''), 
			NULLIF(REPLACE(REPLACE(CAST(Arrangement AS VARCHAR(MAX)),CHAR(13), ''), CHAR(10), ''), ''), 
			NULLIF(REPLACE(REPLACE(CAST(CustHist AS VARCHAR(MAX)),CHAR(13), ''), CHAR(10), ''), ''), 
			NULLIF(REPLACE(REPLACE(CAST(NoteText AS VARCHAR(MAX)),CHAR(13), ''), CHAR(10), ''), ''),
			NULLIF(CAST(LinkDataID AS VARCHAR(MAX)), ''), 
			NULLIF(CAST(LinkTypeID AS VARCHAR(MAX)), '')
			
	FROM #TempTable

END;
GO

