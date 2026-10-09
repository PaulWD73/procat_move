VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BOBasicCatalogueEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    BOBasicCatalogueEntry.cls
' System:       PROCAT, PRO
' Copyright:    (C) Quidnunc Limited
'
' Description:  Business object to handle the creation, updation and deletion of catalogue entries
'
' Amendment history:
'   d1   ndk      26Jul1999      Created
'                                CreateNewCatalogueFromScratch, CreateNewCatalogueByCopying
'                                were moved from BoEditSet and changed their names
'   d2   ndk      17Aug1999      Moved AddCataloguesIntoEditSet from BOEditSet
'   d3   ndk      08Oct1999      Major changes, Added audit trail functions, restructured catalogue related functions
'                                Added more fields and validations
'   d4   ndk      20Oct1999      Added scope and content property to be used while creating piece and item
'                                Updating live_flag
'   d5   ksk      21Oct1999      Added the method AddCatalogueEntryWithParentsIntoEditset and changed
'                                AddCataloguesIntoEditSet to use the new private method to add the
'                                catalogue along with parents with appropriate status into edit set
'   d6   ksk      22Oct1999      Added parameter to set the live flag for the added entry
'                                AddCatalogueEntryWithParentsIntoEditset
'   d7   ndk      23Oct1999      While creating new editset entries the method AddCatalogueEntryWithParentsIntoEditset is
'                                called to add the parents. The actual insertion is done in the live tables and after calling this
'                                method it will move the entry and its parents(if required) into the editset
'                                Copying link items while create by copying, saving link items while create from scratch
'   d8   ndk      26Oct1999      Minor change in InsertRowsByCopying method
'   d9   ndk      27Oct1999      Validation for levels of the copied and created catalogues when created by copying
'   d10  ksk      02Nov1999      Added methods the check for uniqueness for editset entry and the parent uniqueness
'   d11  ndk      03Nov1999      Removed load method, all fields for level copied when created by copying
'                                Added method SqlForExtensionTableForCreateByCopy
'   d13  ndk      10Nov1999      Added method SpecificCatalogueLevels to return required catalogues for use in list box
'                                Copying other fields (not used in procat) in the main table when creating by
'                                copying
'                                Using global constants for detault values , removed local constants
'   d14  ndk      18Nov1999      Using quotes method to insert certain text fields
'   d15  ksk      18Nov1999      Added a method similar to Add catalogue entry but this time for import
'                                ImportCataloguesIntoAnEditSet
'   d13  ndk      24Nov1999      Access condition is now updated when a editset entry is created from scratch or lettercode
'                                using default value of parents value
'   d14  ksk      13Dec1999      Made the duplicate add (adding a catalogue into an edit set as editatble
'                                when it already exists as one in another edit set) as a global constant string
'   d15  ndk      16Dec1999      Access condition was being returned twice in the sql when creating by copying, now corrected
'   d16  ndk      22Dec1999      Changed database error handling
'   d17  ndk      29Dec1999      Rechecking releasing of variables, especially if an error is deliberately raised
'   d18  wdp      09Feb2000      Removed reference to record_open_date in SQLForExtensionTableCreateByCopy
'   d19  tmj      21Feb2000      Changed error message in AddCatalogueIntoEditset
'   d20  wdp      24Feb2000      Added sorted title for lettercodes in CreateFromScratch
'   d21  spb      29Feb2000      Removed reference to BOEditorialPermissions
'   d22  WDP      02Mar2000      Fixed problems using InsertAndReturnIdentity in CreateFromScratch
'                                for pieces and items(doesnt work due to ILDB triggers)
'   d23  WDP      06Mar2000      Fixed defect in Piece/item key order generation in SetDataToBeGenerated
'   d24  tmj      06Mar2000      Enhanced the checks to see if an object has been moved
'   d25  tmj      10Mar2000      Fixed defect in GetClassID
'   d26  tmj      21Mar2000      Fixed CheckNotMoved
'   d27  tmj      22Mar2000      Fixed Check Piece not moved
'   d28  tmj      24Mar2000      Changed checking for moving to pass the test if attempting to add to the same editset
'   d29  daj      27Mar2000      Added extra error handling to cope with the error thrown by stored procedures when
'                                attempting to add a non-live parent to an edit set
'   d30  daj      20Jul2000      Added public wrapper around AddCatalogueEntryWithParentsIntoEditset to add a single entry
'   d31  alm      18Sep2001      Supports the generation and saving of start / end dates when creating new catalogue entries
'   d32  alm      25Oct2001      Added checking for when adding entries to a delete editset, or selecting a parent that may
'                                be in a delete editset. Added the functions CheckCatalogueIsReferenced and
'                                CheckCatalogueParentMarkedForDeletion.
'   d33  alm      29Oct2001      Altered sql for item to detect blank as well
'                                as null references
'   d34  alm      30Oct2001      Shows a different message for an item than for other levels when
'                                it is determined that it is referenced
'   d35  tomh     18Mar2002      Have made the title get stored in the tbl_tables when added to an edit set
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'private framework objects
Private m_oFWDBHelper As FWDBHelper

'private business objects
Private m_oBOCatalogueElementMatrix  As BOCatalogueElementMatrix

'Private property variables
Private m_vCommonEditSetID As Variant
Private m_vCommonCatalogueID As Variant
Private m_vCommonCatalogueLevelID As Variant
Private m_vCommonTitleText As Variant
Private m_vCommonScopeAndContent As Variant
Private m_vCommonLegalStatusID As Variant
Private m_vCommonLanguageID As Variant
Private m_vCommonCoveringDateText As Variant
Private m_vCommonStartDate As Variant 'd31
Private m_vCommonEndDate As Variant 'd31
Private m_vCommonParentCatalogueID As Variant
Private m_vCommonParentCatalogueLevelID As Variant
Private m_vCommonReference As Variant
Private m_colCreatorNamesPerson As Collection
Private m_colCreatorNamesCorporateBody As Collection
Private m_colPlaceOfDeposit As Collection

'default values
Private m_vCommonAccessConditionID As Variant

'Other private variables
Private m_lParentLetterCodeID As Long  '(lettercode id for class under division)
Private m_lParentClassID As Long   'class_id for piece under header and subheader
Private m_lParentHeaderID As Long      'header_id for piece under sub header
Private m_lNewPieceKeyOrder As Long
Private m_lNewItemKeyOrder As Long

'enums
Private Enum CatalogueCreateType
    ectCreateFromScratch
    ectCreateByCopy
End Enum

'Property Methods

Private Sub Initialise()
    'framwork and bo
    Set m_oBOCatalogueElementMatrix = New BOCatalogueElementMatrix
    Set m_oFWDBHelper = New FWDBHelper
    
    'collection properties
    Set m_colCreatorNamesPerson = New Collection
    Set m_colCreatorNamesCorporateBody = New Collection
    Set m_colPlaceOfDeposit = New Collection
End Sub
Private Sub Terminate()
    'framwork and bo
    Set m_oBOCatalogueElementMatrix = Nothing
    Set m_oFWDBHelper = Nothing

    'collection properties
    Set m_colCreatorNamesPerson = Nothing
    Set m_colCreatorNamesCorporateBody = Nothing
    Set m_colPlaceOfDeposit = Nothing
End Sub
Private Sub Class_Initialize()
    Initialise
End Sub
Private Sub Class_Terminate()
    Terminate
End Sub
Public Property Get Name()
    Name = "BOBasicCatalogueEntry"
End Property

Public Property Let EditsetID(ByVal t_vEditSetID As Variant)
    m_vCommonEditSetID = t_vEditSetID
End Property
Public Property Get EditsetID() As Variant
    EditsetID = m_vCommonEditSetID
End Property
Public Property Let CatalogueID(ByVal t_vCatalogueID As Variant)
    m_vCommonCatalogueID = t_vCatalogueID
End Property
Public Property Get CatalogueID() As Variant
    CatalogueID = m_vCommonCatalogueID
End Property
Public Property Let CatalogueLevelID(ByVal t_vCatalogueLevelD As Variant)
    m_vCommonCatalogueLevelID = t_vCatalogueLevelD
End Property
Public Property Get CatalogueLevelID() As Variant
    CatalogueLevelID = m_vCommonCatalogueLevelID
End Property
Public Property Let TitleText(ByVal t_vTitleText As Variant)
    m_vCommonTitleText = t_vTitleText
End Property
Public Property Get TitleText() As Variant
    TitleText = m_vCommonTitleText
End Property
Public Property Let ScopeAndContent(ByVal t_vScopeAndContent As Variant)
    m_vCommonScopeAndContent = t_vScopeAndContent
End Property
Public Property Get ScopeAndContent() As Variant
    ScopeAndContent = m_vCommonScopeAndContent
End Property
Public Property Let LegalStatusID(ByVal t_vLegalStatusID As Variant)
    m_vCommonLegalStatusID = t_vLegalStatusID
End Property
Public Property Get LegalStatusID() As Variant
    LegalStatusID = m_vCommonLegalStatusID
End Property
Public Property Let LanguageID(ByVal t_vLanguageID As Variant)
    m_vCommonLanguageID = t_vLanguageID
End Property
Public Property Get LanguageID() As Variant
    LanguageID = m_vCommonLanguageID
End Property
Public Property Let CoveringDateText(ByVal t_vCoveringDateText As Variant)
    m_vCommonCoveringDateText = t_vCoveringDateText
End Property
Public Property Let StartDate(ByVal t_vStartDate As Variant) 'd31
    m_vCommonStartDate = t_vStartDate
End Property
Public Property Let EndDate(ByVal t_vEndDate As Variant) 'd31
    m_vCommonEndDate = t_vEndDate
End Property
Public Property Get CoveringDateText() As Variant
    CoveringDateText = m_vCommonCoveringDateText
End Property
Public Property Get StartDate() As Variant 'd31
    StartDate = m_vCommonStartDate
End Property
Public Property Get EndDate() As Variant 'd31
    EndDate = m_vCommonEndDate
End Property
Public Property Let ParentCatalogueID(ByVal t_vParentCatalogueID As Variant)
    m_vCommonParentCatalogueID = t_vParentCatalogueID
End Property
Public Property Let ParentCatalogueLevelID(ByVal t_vParentCatalogueLevelID As Variant)
    m_vCommonParentCatalogueLevelID = t_vParentCatalogueLevelID
End Property
Public Property Let Reference(ByVal t_vReference As Variant)
    m_vCommonReference = t_vReference
End Property
Public Property Let CreatorNamesPerson(ByVal t_colCreatorNamesPerson As Collection)
    Set m_colCreatorNamesPerson = t_colCreatorNamesPerson
End Property
Public Property Let CreatorNamesCorporateBody(ByVal t_colCreatorNamesCorporateBody As Collection)
    Set m_colCreatorNamesCorporateBody = t_colCreatorNamesCorporateBody
End Property
Public Property Let PlaceOfDeposit(ByVal t_colPlaceOfDeposit As Collection)
    Set m_colPlaceOfDeposit = t_colPlaceOfDeposit
End Property
Public Property Let AccessConditionID(ByVal t_vAccessConditionID As Variant)
    m_vCommonAccessConditionID = t_vAccessConditionID
End Property
Public Property Get AccessConditionID() As Variant
    AccessConditionID = m_vCommonAccessConditionID
End Property


'Public Methods

Public Sub CreateFromScratch()
' Description:
'
' Creates a new catalogue entry from scratch using the properties set by the view object
'
' Amendment history:
'
' d1   ndk      26Jul1999      Copied from BOEditSet (Fully rewritten)
' d7   ndk      23Oct1999      While creating new editset entries the method AddCatalogueEntryWithParentsIntoEditset is
'                              called to add the parents. The actual insertion is done in the live tables and after calling this
'                              method it will move the entry and its parents(if required) into the editset
' d16  ndk      22Dec1999      Changed database error handling
' d20  wdp      24Feb2000      Added sorted title for lettercodes in CreateFromScratch
' d22  WDP      02Mar2000      Fixed problems using InsertAndReturnIdentity in CreateFromScratch
'                                for pieces and items(doesnt work due to ILDB triggers)
'
    
    On Error GoTo ErrorHandler
    
    Dim sSource As String, sSql As String
    Dim lNewCatalogueID As Long, eCommonCatalogueLevelID As ECATALOGUELEVEL, sIDColoumnName As String

    sSource = Me.Name & ".CreateFromScratch()"
    
    'get the enum equivalent
    eCommonCatalogueLevelID = GetCatalogueLevelEnum(CInt(m_vCommonCatalogueLevelID))

    'Getnerate piece key order, item key order and Real parent (for class and piece)
    SetDataToBeGenerated
    
    ' Insert Row in Main table and get the value of the id coloumn (live)
    sSql = SqlForMainTable( _
            CatalogueCreateType.ectCreateFromScratch, _
            eCommonCatalogueLevelID, _
            m_vCommonParentCatalogueLevelID, _
            CLng(m_vCommonParentCatalogueID))

    'For pieces and items InsertAndReturnIdentity will return the wrong
    'ID (due to ILDB insert triggers creating records in ILDBAudit)
    'so we must retrieve the identity seperately
    
    If eCommonCatalogueLevelID < ctPiece Then
        lNewCatalogueID = _
            m_oFWDBHelper.InsertAndReturnIdentity(sSql)
    Else
        Dim lKeyOrder As Long
        Dim lProperParentID As Long
        Select Case eCommonCatalogueLevelID
            Case ctPiece
                lKeyOrder = m_lNewPieceKeyOrder
                lProperParentID = m_lParentClassID
                'For pieces that are direct children of classes
                'this will not be set so we need to set it
                'from the parent ID
                If lProperParentID = 0 Then
                    lProperParentID = CLng(m_vCommonParentCatalogueID)
                End If
                
            Case ctItem
                lKeyOrder = m_lNewItemKeyOrder
                lProperParentID = CLng(m_vCommonParentCatalogueID)
            Case Else
                RaiseError sSource, PEGeneralInternalError, "Invalid catalogue level."
        End Select
        
        m_oFWDBHelper.Execute (sSql)
        lNewCatalogueID = GetPitemIDFromParentIDAndKeyOrder(eCommonCatalogueLevelID, lProperParentID, lKeyOrder)
    End If
    'The variable to hold the coloumn names of catalogueid and editsetid
    sIDColoumnName = GetColumnName(ECATALOGUECOLUMNTYPE.cctIDColumn, eCommonCatalogueLevelID)
    
    'Add a row to extension table
    m_oFWDBHelper.Execute _
        SqlForExtensionTable(ectCreateFromScratch, lNewCatalogueID, eCommonCatalogueLevelID)    ' Live
    
    'Add rows to accruals table
    InsertRows ECATALOGUETABLETYPE.cttAccrualsTable, eCommonCatalogueLevelID, sIDColoumnName, lNewCatalogueID
    
    'Add rows to Admin history table
    InsertRows ECATALOGUETABLETYPE.cttAdminHistoryTable, eCommonCatalogueLevelID, sIDColoumnName, lNewCatalogueID
    
    'Add rows to App Destruction Info  table
    InsertRows ECATALOGUETABLETYPE.cttAppDestructionTable, eCommonCatalogueLevelID, sIDColoumnName, lNewCatalogueID
    
    'Add rows to Arrangement table
    InsertRows ECATALOGUETABLETYPE.cttArrangementTable, eCommonCatalogueLevelID, sIDColoumnName, lNewCatalogueID
    
    'Add rows to Custodial History table
    InsertRows ECATALOGUETABLETYPE.cttCustodialHistoryTable, eCommonCatalogueLevelID, sIDColoumnName, lNewCatalogueID
    
    'Add rows to note table
    InsertRows ECATALOGUETABLETYPE.cttNoteTable, eCommonCatalogueLevelID, sIDColoumnName, lNewCatalogueID
    
    If eCommonCatalogueLevelID = ECATALOGUELEVEL.ctPiece Or eCommonCatalogueLevelID = ECATALOGUELEVEL.ctItem Then
        'Add rows to scope and content table (with data
        InsertRows ECATALOGUETABLETYPE.cttScopeAndContentTable, eCommonCatalogueLevelID, sIDColoumnName, lNewCatalogueID, _
            gk_sDBCOLScopeAndContent, QuoteValueWithQuotes(CStr(ScopeAndContent))

        'Add rows to Title table (empty row)
        InsertRows ECATALOGUETABLETYPE.cttTitleTable, eCommonCatalogueLevelID, sIDColoumnName, lNewCatalogueID
    Else    'Other levels
        
        'Add rows to scope and content table (empty row)
        InsertRows ECATALOGUETABLETYPE.cttScopeAndContentTable, eCommonCatalogueLevelID, sIDColoumnName, lNewCatalogueID
        
        'Add rows to Title table (with data)
        InsertRows ECATALOGUETABLETYPE.cttTitleTable, eCommonCatalogueLevelID, sIDColoumnName, lNewCatalogueID, _
            gk_sDBCOLTitle, QuoteValueWithQuotes(CStr(TitleText))
    End If
    
    If CatalogueLevelID = ECATALOGUELEVEL.ctLetterCode Then         'Creating a letter code
        'Update the es_lettercode_sortedtitle
        m_oFWDBHelper.Execute _
            " INSERT " & gk_sDBTABLELetterCodeSortedTitle & "(" & _
            gk_sDBCOLLetterCodeID & ", " & gk_sDBCOLSortTitle & ")" & _
            " VALUES ( " & lNewCatalogueID & _
            ", " & QuoteValueWithQuotes(GetSortedTitle(m_vCommonTitleText)) & ")"       'Live

    End If
    
    'Copy the link items to reader
    
    'creator name person
    AddLinkItemsToReader ELinkTermType.elCreatorNamesPerson, m_colCreatorNamesPerson, lNewCatalogueID, eCommonCatalogueLevelID
    
    'creator name corp body
    AddLinkItemsToReader ELinkTermType.elCreatorNamesCorporateBody, m_colCreatorNamesCorporateBody, lNewCatalogueID, eCommonCatalogueLevelID
    
    'creator name place of deposit
    AddLinkItemsToReader ELinkTermType.elPlaceOfDeposit, m_colPlaceOfDeposit, lNewCatalogueID, eCommonCatalogueLevelID
    
    'Now transfer the main entry from live to editorial and its parents (as read only if they do not already exist)
    AddCatalogueEntryWithParentsIntoEditset _
        CLng(m_vCommonEditSetID), lNewCatalogueID, eCommonCatalogueLevelID, _
        ECATALOGUEENTRYSTATUS.cesPROCATCreated, ELIVEFLAGSTATUS.lfsNotLive, ECATALOGUEENTRYSTATUS.cesReadeOnly
    
    Exit Sub
    
ErrorHandler:
    Select Case Err.Number
        Case EPROCATINTERNALERROR.PEDBSQLServerErrorUniqueConstraintViolation, _
            EPROCATINTERNALERROR.PEDBSQLServerErrorUniqueIndexViolation
        
            RaiseError sSource, EPROCATINTERNALERROR.PEDBGeneralUniqueConstraintViolation, _
            "The reference you entered for the catalogue is not unique"
    End Select
                
    'handled here if none of the above
    ConvertError sSource
End Sub

Public Sub CreateByCopying( _
            ByVal t_iSourceCatalogueLevel As Integer, _
            ByVal t_lSourceCatalogueID As Long)
            
' Description:
'
' Creates a new lettercode by copying details of an entry from the live table
' The mandatory fields are shown to the user for modification
'
' Amendment history:
'
' d1   ndk      26Jul1999      Copied from BOEditSet (Fully rewritten)
' d7   ndk      23Oct1999      While creating new editset entries the method AddCatalogueEntryWithParentsIntoEditset is
'                              called to add the parents. The actual insertion is done in the live tables and after calling this
'                              method it will move the entry and its parents(if required) into the editset
' d9   ndk      27Oct1999      Validation for levels of the copied and created catalogues when created by copying
' d11  ndk      03Nov1999      All fields for level copied when created by copying
' d16  ndk      22Dec1999      Changed database error handling
' d23   wdp     06Mar2000      Altered code to get id for new pieces and items
'

    On Error GoTo ErrorHandler
    Dim sSource As String, sSql As String
    Dim lNewCatalogueID As Long
    Dim eSourceCatalogueLevel As ECATALOGUELEVEL, eTargetCatalogueLevel As ECATALOGUELEVEL
    Dim sIDColoumnNameForSource As String, sIDColoumnNameForTarget As String
 
    sSource = Me.Name & ".CreateByCopying"
    
    'get the enum equivalent
    eTargetCatalogueLevel = GetCatalogueLevelEnum(CInt(m_vCommonCatalogueLevelID))
    eSourceCatalogueLevel = GetCatalogueLevelEnum(t_iSourceCatalogueLevel)
        
    'Validations
    If Not eSourceCatalogueLevel = ECATALOGUELEVEL.ctPiece Then
        'If you are not copying from an piece then both the levels should be same
        If eSourceCatalogueLevel <> eTargetCatalogueLevel Then
            RaiseError sSource, PEInvalidParentLevelForCreateByCopying, _
                "The catalogue you have choosen can not be the parent for the catalogue you are trying to create"
        End If
    Else
        'if you are copying from a piece
        'the target catalogue level should be an piece or item
        If Not eTargetCatalogueLevel = ECATALOGUELEVEL.ctItem And _
            Not eTargetCatalogueLevel = ECATALOGUELEVEL.ctPiece Then
            
            RaiseError sSource, PEInvalidParentLevelForCreateByCopying, _
                "The catalogue you have choosen can not be the parent for the catalogue you are trying to create"
        End If
    End If
    
    'primary key column names
    sIDColoumnNameForTarget = GetColumnName(ECATALOGUECOLUMNTYPE.cctIDColumn, eTargetCatalogueLevel)
    sIDColoumnNameForSource = GetColumnName(ECATALOGUECOLUMNTYPE.cctIDColumn, eSourceCatalogueLevel)
    
    'Getnerate piece key order, item key order and Real parent (for class and piece)
    SetDataToBeGenerated
    
    ' Insert Row in Main table and get the value of the id coloumn (live)
    
    sSql = SqlForMainTable( _
                CatalogueCreateType.ectCreateByCopy, _
                eTargetCatalogueLevel, _
                m_vCommonParentCatalogueLevelID, _
                CLng(m_vCommonParentCatalogueID), _
                t_iSourceCatalogueLevel, _
                t_lSourceCatalogueID)

    'For pieces and items InsertAndReturnIdentity will return the wrong
    'ID (due to ILDB insert triggers creating records in ILDBAudit)
    'so we must retrieve the identity seperately
    
    If eTargetCatalogueLevel < ctPiece Then
        lNewCatalogueID = _
            m_oFWDBHelper.InsertAndReturnIdentity(sSql)
    Else
        Dim lKeyOrder As Long
        Dim lProperParentID As Long
        Select Case eTargetCatalogueLevel
            Case ctPiece
                lKeyOrder = m_lNewPieceKeyOrder
                lProperParentID = m_lParentClassID
                'For pieces that are direct children of classes
                'this will not be set so we need to set it
                'from the parent ID
                If lProperParentID = 0 Then
                    lProperParentID = CLng(m_vCommonParentCatalogueID)
                End If
                
            Case ctItem
                lKeyOrder = m_lNewItemKeyOrder
                lProperParentID = CLng(m_vCommonParentCatalogueID)
            Case Else
                RaiseError sSource, PEGeneralInternalError, "Invalid catalogue level."
        End Select
        
        m_oFWDBHelper.Execute (sSql)
        lNewCatalogueID = GetPitemIDFromParentIDAndKeyOrder(eTargetCatalogueLevel, lProperParentID, lKeyOrder)
    End If
    
        'Add rows to extension table
    sSql = SqlForExtensionTable( _
            ectCreateByCopy, lNewCatalogueID, eTargetCatalogueLevel, _
            eSourceCatalogueLevel, t_lSourceCatalogueID)

    m_oFWDBHelper.Execute sSql
    
    'Add rows to accruals table
    InsertRowsByCopying _
        ECATALOGUETABLETYPE.cttAccrualsTable, eTargetCatalogueLevel, eSourceCatalogueLevel, _
        sIDColoumnNameForTarget, sIDColoumnNameForSource, lNewCatalogueID, t_lSourceCatalogueID, _
        gk_sDBCOLAccruals, gk_sDBCOLAccruals
    
    'Add rows to Admin history table
    InsertRowsByCopying _
        ECATALOGUETABLETYPE.cttAdminHistoryTable, eTargetCatalogueLevel, eSourceCatalogueLevel, _
        sIDColoumnNameForTarget, sIDColoumnNameForSource, lNewCatalogueID, t_lSourceCatalogueID, _
        gk_sDBCOLAdministrativeHistory, gk_sDBCOLAdministrativeHistory
    
    'Add rows to App Destruction Info table
    InsertRowsByCopying _
        ECATALOGUETABLETYPE.cttAppDestructionTable, eTargetCatalogueLevel, eSourceCatalogueLevel, _
        sIDColoumnNameForTarget, sIDColoumnNameForSource, lNewCatalogueID, t_lSourceCatalogueID, _
        gk_sDBCOLAppraisalDestructionInformation, gk_sDBCOLAppraisalDestructionInformation
    
    'Add rows to Arrangement table
    InsertRowsByCopying _
        ECATALOGUETABLETYPE.cttArrangementTable, eTargetCatalogueLevel, eSourceCatalogueLevel, _
        sIDColoumnNameForTarget, sIDColoumnNameForSource, lNewCatalogueID, t_lSourceCatalogueID, _
        gk_sDBCOLArrangement, gk_sDBCOLArrangement
    
    'Add rows to Custodial History table
    InsertRowsByCopying _
        ECATALOGUETABLETYPE.cttCustodialHistoryTable, eTargetCatalogueLevel, eSourceCatalogueLevel, _
        sIDColoumnNameForTarget, sIDColoumnNameForSource, lNewCatalogueID, t_lSourceCatalogueID, _
        gk_sDBCOLCustodialHistory, gk_sDBCOLCustodialHistory

    'Add rows to note table
    InsertRowsByCopying _
        ECATALOGUETABLETYPE.cttNoteTable, eTargetCatalogueLevel, eSourceCatalogueLevel, _
        sIDColoumnNameForTarget, sIDColoumnNameForSource, lNewCatalogueID, t_lSourceCatalogueID, _
        gk_sDBCOLNoteText, gk_sDBCOLNoteText

    If eTargetCatalogueLevel = ECATALOGUELEVEL.ctPiece Or eTargetCatalogueLevel = ECATALOGUELEVEL.ctItem Then
        'Add rows to scope and content table (for piece and item add the scope and content typed by the user)
        InsertRowsByCopying _
            ECATALOGUETABLETYPE.cttScopeAndContentTable, eTargetCatalogueLevel, eSourceCatalogueLevel, _
            sIDColoumnNameForTarget, sIDColoumnNameForSource, lNewCatalogueID, t_lSourceCatalogueID, _
            gk_sDBCOLScopeAndContent, QuoteValueWithQuotes(CStr(ScopeAndContent))
        
        'Add rows to Title table
        InsertRowsByCopying _
            ECATALOGUETABLETYPE.cttTitleTable, eTargetCatalogueLevel, eSourceCatalogueLevel, _
            sIDColoumnNameForTarget, sIDColoumnNameForSource, lNewCatalogueID, t_lSourceCatalogueID, _
            gk_sDBCOLTitle, gk_sDBCOLTitle
    Else
            'Add rows to scope and content table (for other levels add the title typed by user)
        InsertRowsByCopying _
            ECATALOGUETABLETYPE.cttScopeAndContentTable, eTargetCatalogueLevel, eSourceCatalogueLevel, _
            sIDColoumnNameForTarget, sIDColoumnNameForSource, lNewCatalogueID, t_lSourceCatalogueID, _
            gk_sDBCOLScopeAndContent, gk_sDBCOLScopeAndContent
        
        'Add rows to Title table
        InsertRowsByCopying _
            ECATALOGUETABLETYPE.cttTitleTable, eTargetCatalogueLevel, eSourceCatalogueLevel, _
            sIDColoumnNameForTarget, sIDColoumnNameForSource, lNewCatalogueID, t_lSourceCatalogueID, _
            gk_sDBCOLTitle, QuoteValueWithQuotes(CStr(TitleText))
    End If
    
    'Copy the link items into reader
    CopyAllLinksToReader _
        t_lSourceCatalogueID, eSourceCatalogueLevel, lNewCatalogueID, eTargetCatalogueLevel
    
    'creator name person
    AddLinkItemsToReader ELinkTermType.elCreatorNamesPerson, m_colCreatorNamesPerson, lNewCatalogueID, eTargetCatalogueLevel
    
    'creator name corp body
    AddLinkItemsToReader ELinkTermType.elCreatorNamesCorporateBody, m_colCreatorNamesCorporateBody, lNewCatalogueID, eTargetCatalogueLevel
    
    'creator name place of deposit
    AddLinkItemsToReader ELinkTermType.elPlaceOfDeposit, m_colPlaceOfDeposit, lNewCatalogueID, eTargetCatalogueLevel
    
    'Now transfer the main entry from live to editorial and its parents (as read only if they do not already exist)
    AddCatalogueEntryWithParentsIntoEditset _
        CLng(m_vCommonEditSetID), lNewCatalogueID, eTargetCatalogueLevel, _
        ECATALOGUEENTRYSTATUS.cesPROCATCreated, ELIVEFLAGSTATUS.lfsNotLive, ECATALOGUEENTRYSTATUS.cesReadeOnly
    
    Exit Sub
    
ErrorHandler:
    Select Case Err.Number
        Case EPROCATINTERNALERROR.PEDBSQLServerErrorUniqueConstraintViolation, _
        EPROCATINTERNALERROR.PEDBSQLServerErrorUniqueIndexViolation
        
            RaiseError sSource, EPROCATINTERNALERROR.PEDBGeneralUniqueConstraintViolation, _
            "The reference you entered for the catalogue is not unique"
    End Select
    
    'handled here if none of the above
    ConvertError sSource

End Sub

Public Function ChildCatalogueLevels(ByVal t_iCatalogueLevelID As Integer) As FWRecordList
' Description:
'
' Return the child catalogues for a given catalogue
'
' Amendment history:
'
' d1   ndk      26Jul1999      Copied from BOEditSet

On Error GoTo ErrorHandler

    Dim sSource As String
    Dim sSql As String

    sSource = Me.Name & ".ChildCatalogueLevels"
    
    'Construct Sql to search the database
    'construct the sql
    sSql = "sp_Edit_GetChildCatalogueLevels " & _
        " @t_iCatalogueLevelId = " & t_iCatalogueLevelID & _
        ", @t_iLevelIDLetterCode = " & GetCatalogueLevelCode(ctLetterCode) & _
        ", @t_sLeveNameLetterCode = " & Quotes(GetCatalogueLevelDescription(GetCatalogueLevelCode(ctLetterCode))) & _
        ", @t_iLevelDDivision = " & GetCatalogueLevelCode(ctDivision) & _
        ", @t_sLevelNameDivision = " & Quotes(GetCatalogueLevelDescription(GetCatalogueLevelCode(ctDivision))) & _
        ", @t_iLevelIDClass = " & GetCatalogueLevelCode(ctClass) & _
        ", @t_sLevelNameClass = " & Quotes(GetCatalogueLevelDescription(GetCatalogueLevelCode(ctClass))) & _
        ", @t_iLevelIDHeader = " & GetCatalogueLevelCode(ctHeaders) & _
        ", @t_sLevelNameHeader = " & Quotes(GetCatalogueLevelDescription(GetCatalogueLevelCode(ctHeaders))) & _
        ", @t_iLevelIDSubHeader = " & GetCatalogueLevelCode(ctSubHeaders) & _
        ", @t_sLevelNameSubHeader = " & Quotes(GetCatalogueLevelDescription(GetCatalogueLevelCode(ctSubHeaders))) & _
        ", @t_iLevelIDPiece = " & GetCatalogueLevelCode(ctPiece) & _
        ", @t_sLevelNamePiece = " & Quotes(GetCatalogueLevelDescription(GetCatalogueLevelCode(ctPiece))) & _
        ", @t_iLevelIDItem = " & GetCatalogueLevelCode(ctItem) & _
        ", @t_sLevelNameItem = " & Quotes(GetCatalogueLevelDescription(GetCatalogueLevelCode(ctItem)))
        
    'Execute the sp
    Set ChildCatalogueLevels = m_oFWDBHelper.FindBySql(sSql)
    
    Exit Function
    
ErrorHandler:
    ConvertError sSource
End Function

Public Function SpecificCatalogueLevels( _
     t_arLevels() As Integer) As FWRecordList
    
' Description:
'
' Return required catalogues as a recordlist for use in list box
'
' Amendment history:
'
'   d13  ndk      10Nov1999     Created

On Error GoTo ErrorHandler

    Dim sSource As String
    Dim sSql As String, sLevelsRequired As String, iCounter As Integer
        
    sSource = Me.Name & ".SpecificCatalogueLevels"
    
    'List of levels required by the user
    For iCounter = LBound(t_arLevels) To UBound(t_arLevels)
        If sLevelsRequired = "" Then
            sLevelsRequired = t_arLevels(iCounter)
        Else
            sLevelsRequired = sLevelsRequired & "," & t_arLevels(iCounter)
        End If
    Next
    
    'Construct Sql to search the database
    sSql = _
        " SELECT " & _
            gk_sDBCOLCatalogueLevel & "," & _
            gk_sDBCOLLevelDesc & _
        " FROM " & gk_sDBTABLELevels & _
        " WHERE " & gk_sDBCOLCatalogueLevel & _
            " IN ( " & sLevelsRequired & " )"
            
    'get the recordlist
    Set SpecificCatalogueLevels = m_oFWDBHelper.FindBySql(sSql)
    
    Exit Function
ErrorHandler:
    ConvertError sSource
End Function




Private Function SqlForMainTable( _
            ByVal t_eCatalogueCreateType As CatalogueCreateType, _
            ByVal t_eTargetCatalogueLevelID As ECATALOGUELEVEL, _
            ByVal t_eParentCatalogueLevelID As ECATALOGUELEVEL, _
            ByVal t_lParentCatalogueID As Long, _
            Optional ByVal t_eSourceCatalogueLevelID As ECATALOGUELEVEL, _
            Optional ByVal t_lSourceCatalogueID As Long) As String
                    
' Description:
'
' Returns the insert sql for the maintable eg es_lettercode, es_class
'
' Amendment history:
'
' d1   ndk      26Jul1999      Created
' d13  ndk      10Nov1999      Copying other fields (not used in procat) in the main table when creating by
'                              copying
    
    On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim sSqlColoumns As String, sSqlAdditionalColoumns As String
    Dim sSqlValues As String, sSqlAdditionalValues As String
    Dim sFieldsFromTargetTable As String, sFieldsFromSourceTable As String
    
    sSource = Me.Name & ".SqlForMainTable"
        
    'sql To insert a row in the main table
    
    'coloumn names
    If Not t_eTargetCatalogueLevelID = ECATALOGUELEVEL.ctLetterCode Then
        sSqlColoumns = ParentCatalogueColumnName(t_eTargetCatalogueLevelID, t_eParentCatalogueLevelID)
    End If
    
    'Get coloumns specicific to the catalogue level
    sSqlAdditionalColoumns = SqlForSpecificColoumnsInMainTable(t_eTargetCatalogueLevelID)
    
    If Not sSqlAdditionalColoumns = "" Then  'There are some additional coloumns
        If sSqlColoumns = "" Then
            sSqlColoumns = sSqlAdditionalColoumns
        Else
            sSqlColoumns = sSqlColoumns & "," & sSqlAdditionalColoumns
        End If
    End If
    
    'coloumn values
    If Not t_eTargetCatalogueLevelID = ECATALOGUELEVEL.ctLetterCode Then
        sSqlValues = t_lParentCatalogueID
    End If
    
    'Get values specicific to the catalogue level
    sSqlAdditionalValues = SqlForSpecificValuesInMainTable(t_eTargetCatalogueLevelID)
    
    If Not sSqlAdditionalValues = "" Then
        If sSqlValues = "" Then
            sSqlValues = sSqlAdditionalValues
        Else
            sSqlValues = sSqlValues & "," & sSqlAdditionalValues
        End If
    End If
    
    'Create by copying
    If t_eCatalogueCreateType = CatalogueCreateType.ectCreateByCopy Then
        
        'Fields for target table
        sFieldsFromTargetTable = SqlForMainTableForCreateByCopy(t_eTargetCatalogueLevelID)
        
        'Fields for source table
        sFieldsFromSourceTable = SqlForMainTableForCreateByCopy(t_eSourceCatalogueLevelID)
        
        If Not sFieldsFromTargetTable = "" Then   'There are some additional values
            sSqlColoumns = sSqlColoumns & ", " & sFieldsFromTargetTable
        End If
        
        If Not sFieldsFromSourceTable = "" Then  'There are some additional values
            sSqlValues = sSqlValues & ", " & sFieldsFromSourceTable
        End If


        SqlForMainTable = _
            " INSERT " & GetTableName(ECATALOGUETABLETYPE.cttMainTable, ESectionType.SECReader, t_eTargetCatalogueLevelID) & " (" & _
                    sSqlColoumns & ")" & _
            " SELECT " & sSqlValues & _
            " FROM " & GetTableName(ECATALOGUETABLETYPE.cttMainTable, ESectionType.SECReader, t_eSourceCatalogueLevelID) & _
            " WHERE " & GetColumnName(ECATALOGUECOLUMNTYPE.cctIDColumn, t_eSourceCatalogueLevelID) & _
                "=" & t_lSourceCatalogueID
    Else
                
        SqlForMainTable = _
            "INSERT " & GetTableName(ECATALOGUETABLETYPE.cttMainTable, ESectionType.SECReader, t_eTargetCatalogueLevelID) & " (" & _
                sSqlColoumns & ")" & _
            " VALUES (" & _
                sSqlValues & ")"
    End If
    
    Exit Function
ErrorHandler:
    ConvertError sSource
End Function

Private Function SqlForSpecificColoumnsInMainTable(ByVal t_eCatalogueLevelID As ECATALOGUELEVEL) As String
' Description:
'
' Returns the coloumns in which the data is to be inserted in the main table
'
' Amendment history:
'
' d1   ndk      26Jul1999      Created
' d31  alm      18Sep2001      Inserts columns for start and end dates
' d35  Tomh     18Mar2002      Have put it columns for title
    
    On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim StartDateCol As String
    Dim EndDateCol As String
    
    sSource = Me.Name & ".SqlForSpecificColoumnsInMainTable"
    
    ' d31 Get start and end date column names
    StartDateCol = GetColumnName(cctStartDateColumn, t_eCatalogueLevelID)
    EndDateCol = GetColumnName(cctEndDateColumn, t_eCatalogueLevelID)

    'Additional coloumns specific to tables
    Select Case t_eCatalogueLevelID
        Case ECATALOGUELEVEL.ctLetterCode

                SqlForSpecificColoumnsInMainTable = gk_sDBCOLLetterCodeCode
                
                ' d35 tomh Title column
                SqlForSpecificColoumnsInMainTable = SqlForSpecificColoumnsInMainTable & ", " & gk_sDBTableLettercodeTitle
        
        Case ECATALOGUELEVEL.ctDivision
                
                SqlForSpecificColoumnsInMainTable = gk_sDBCOLDivisionNo
                
                ' d35 tomh Title column
                SqlForSpecificColoumnsInMainTable = SqlForSpecificColoumnsInMainTable & ", " & gk_sDBTableDivisionTitle
                
        Case ECATALOGUELEVEL.ctClass
                'Class no
                SqlForSpecificColoumnsInMainTable = gk_sDBCOLClassNumber
                
                If GetCatalogueLevelEnum(CInt(m_vCommonParentCatalogueLevelID)) = ECATALOGUELEVEL.ctDivision Then
                    'Creating class under division - set the lettercode
                    
                    SqlForSpecificColoumnsInMainTable = SqlForSpecificColoumnsInMainTable & ", " & _
                        gk_sDBCOLLetterCodeID
                End If
                
                'Date related coloumns
                SqlForSpecificColoumnsInMainTable = SqlForSpecificColoumnsInMainTable & ", " & _
                        gk_sDBCOLDateText
                        
                'd31 Start date
                SqlForSpecificColoumnsInMainTable = SqlForSpecificColoumnsInMainTable & ", " & _
                StartDateCol
                
                'd31 End date
                SqlForSpecificColoumnsInMainTable = SqlForSpecificColoumnsInMainTable & ", " & _
                EndDateCol
                
                ' d35 tomh Title column
                SqlForSpecificColoumnsInMainTable = SqlForSpecificColoumnsInMainTable & ", " & gk_sDBTableClassTitle
            
            
        Case ECATALOGUELEVEL.ctHeaders
              
                SqlForSpecificColoumnsInMainTable = gk_sDBCOLClassHeaderNo
                
                'Date related coloumns
                SqlForSpecificColoumnsInMainTable = SqlForSpecificColoumnsInMainTable & ", " & gk_sDBCOLDateText
                
                ' d35 tomh Title column
                SqlForSpecificColoumnsInMainTable = SqlForSpecificColoumnsInMainTable & ", " & gk_sDBTableHeaderTitle
            
        Case ECATALOGUELEVEL.ctSubHeaders

                SqlForSpecificColoumnsInMainTable = gk_sDBCOLSubHeaderNo
                
                'Date related coloumns
                SqlForSpecificColoumnsInMainTable = SqlForSpecificColoumnsInMainTable & ", " & gk_sDBCOLDateText
                
                ' d35 tomh Title column
                SqlForSpecificColoumnsInMainTable = SqlForSpecificColoumnsInMainTable & ", " & gk_sDBTableSubHeaderTitle
            
            
        Case ECATALOGUELEVEL.ctPiece
                
                'Piece key order and reference
                SqlForSpecificColoumnsInMainTable = gk_sDBColPiecekeyOrder & ", " & gk_sDBCOLPieceRef
                
                If GetCatalogueLevelEnum(CInt(m_vCommonParentCatalogueLevelID)) = ECATALOGUELEVEL.ctHeaders Then
                    'Creating piece under header - set the class_id
                    SqlForSpecificColoumnsInMainTable = SqlForSpecificColoumnsInMainTable & ", " & _
                        gk_sDBCOLClassID
                ElseIf GetCatalogueLevelEnum(CInt(m_vCommonParentCatalogueLevelID)) = ECATALOGUELEVEL.ctSubHeaders Then
                    'Creating piece under sub header - set the class_id and headerid
                    SqlForSpecificColoumnsInMainTable = SqlForSpecificColoumnsInMainTable & ", " & _
                    gk_sDBCOLClassID & "," & gk_sDBCOLHeaderID
                End If
                
                'Columns populated with constants
                SqlForSpecificColoumnsInMainTable = SqlForSpecificColoumnsInMainTable & "," & _
                    gk_sDBCOLClosureStatus & "," & _
                    gk_sDBCOLClosureType & "," & _
                    gk_sDBCOLClosureCode & "," & _
                    gk_sDBCOLRecordStatus
                
                'Date related coloumns
                SqlForSpecificColoumnsInMainTable = _
                        SqlForSpecificColoumnsInMainTable & ", " & _
                        gk_sDBCOLDateText & ", " & _
                        gk_sDBCOLLanguageID
                
                'd31 Start date
                SqlForSpecificColoumnsInMainTable = SqlForSpecificColoumnsInMainTable & ", " & _
                StartDateCol
                
                'd31 End date
                SqlForSpecificColoumnsInMainTable = SqlForSpecificColoumnsInMainTable & ", " & _
                EndDateCol
                            
            
        Case ECATALOGUELEVEL.ctItem
                'Date related coloumns
                SqlForSpecificColoumnsInMainTable = _
                        gk_sDBCOLDateText & ", " & _
                        gk_sDBCOLLanguageID
                            
                'Columns populated with constants
                SqlForSpecificColoumnsInMainTable = SqlForSpecificColoumnsInMainTable & "," & _
                    gk_sDBCOLClosureStatus & "," & _
                    gk_sDBCOLClosureType & "," & _
                    gk_sDBCOLClosureCode & "," & _
                    gk_sDBCOLRecordStatus
                            
                'Item key order and itemref
                SqlForSpecificColoumnsInMainTable = _
                        SqlForSpecificColoumnsInMainTable & ", " & gk_sDBCOLItemKeyOrder & ", " & gk_sDBCOLItemRef
                        
                'd31 Start date
                SqlForSpecificColoumnsInMainTable = SqlForSpecificColoumnsInMainTable & ", " & _
                StartDateCol
                
                'd31 End date
                SqlForSpecificColoumnsInMainTable = SqlForSpecificColoumnsInMainTable & ", " & _
                EndDateCol


    End Select
    
    Exit Function
    
ErrorHandler:
    ConvertError sSource
End Function
        
                        
Private Function SqlForSpecificValuesInMainTable( _
                    ByVal t_eCatalogueLevelID As ECATALOGUELEVEL _
                    ) As String
    
' Description:
'
' Returns the values to be inserted in the main table
'
' Amendment history:
'
' d1   ndk      26Jul1999      Created
' d31  alm      18Sep2001      Inserts values for start and end dates
' d35  Tomh     18Mar2002      This inserts values for the title
    
    On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim sStrippedTitle As String
    
    sSource = Me.Name & ".SqlForSpecificValuesInMainTable"
    
    'Additional values specific to tables
    Select Case t_eCatalogueLevelID
        Case ECATALOGUELEVEL.ctLetterCode
                
                SqlForSpecificValuesInMainTable = QuoteValueWithQuotes(CStr(m_vCommonReference))
        
                ' tomh d35 (added this bit)
                ' Need to format the title string for tbl_header table
                sStrippedTitle = RemoveTagsFromString(TitleText)
                sStrippedTitle = SpecialCharacterDecode(sStrippedTitle)
                
                ' want to check title not too long
                If Len(sStrippedTitle) > 255 Then
                    sStrippedTitle = Left(sStrippedTitle, 255)
                End If

                ' Title column
                SqlForSpecificValuesInMainTable = _
                        SqlForSpecificValuesInMainTable & ", " & QuoteValueWithQuotes(sStrippedTitle)
        
        Case ECATALOGUELEVEL.ctDivision
        
                SqlForSpecificValuesInMainTable = CLng(m_vCommonReference)
                
                ' tomh d35 (added this bit)
                ' Need to format the title string for tbl_header table
                sStrippedTitle = RemoveTagsFromString(TitleText)
                sStrippedTitle = SpecialCharacterDecode(sStrippedTitle)
                
                ' want to check title not too long
                If Len(sStrippedTitle) > 255 Then
                    sStrippedTitle = Left(sStrippedTitle, 255)
                End If

                ' Title column
                SqlForSpecificValuesInMainTable = _
                        SqlForSpecificValuesInMainTable & ", " & QuoteValueWithQuotes(sStrippedTitle)
                
        Case ECATALOGUELEVEL.ctClass
                
                'Class no
                SqlForSpecificValuesInMainTable = CLng(m_vCommonReference)
                
                If GetCatalogueLevelEnum(CInt(m_vCommonParentCatalogueLevelID)) = ECATALOGUELEVEL.ctDivision Then
                    'Creating class under division - set the lettercode
                    
                    SqlForSpecificValuesInMainTable = SqlForSpecificValuesInMainTable & ", " & _
                        m_lParentLetterCodeID
                End If
                           
                'Date related coloumns
                SqlForSpecificValuesInMainTable = _
                        SqlForSpecificValuesInMainTable & ", " & _
                        QuoteValueWithQuotes(CStr(CoveringDateText))
                        
                SqlForSpecificValuesInMainTable = _
                        SqlForSpecificValuesInMainTable & ", " & _
                        CStr(StartDate)
                        
                SqlForSpecificValuesInMainTable = _
                        SqlForSpecificValuesInMainTable & ", " & _
                        CStr(EndDate)
                
                ' tomh d35 (added this bit)
                ' Need to format the title string for tbl_header table
                sStrippedTitle = RemoveTagsFromString(TitleText)
                sStrippedTitle = SpecialCharacterDecode(sStrippedTitle)
                
                ' want to check title not too long
                If Len(sStrippedTitle) > 255 Then
                    sStrippedTitle = Left(sStrippedTitle, 255)
                End If

                ' Title column
                SqlForSpecificValuesInMainTable = _
                        SqlForSpecificValuesInMainTable & ", " & QuoteValueWithQuotes(sStrippedTitle)
                        
        Case ECATALOGUELEVEL.ctHeaders
              
                SqlForSpecificValuesInMainTable = CLng(m_vCommonReference)
                
                'Date related coloumns
                SqlForSpecificValuesInMainTable = _
                        SqlForSpecificValuesInMainTable & ", " & QuoteValueWithQuotes(CoveringDateText)
                
                ' tomh d35 (added this bit)
                ' Need to format the title string for tbl_header table
                sStrippedTitle = RemoveTagsFromString(TitleText)
                sStrippedTitle = SpecialCharacterDecode(sStrippedTitle)
                
                ' want to check title not too long
                If Len(sStrippedTitle) > 255 Then
                    sStrippedTitle = Left(sStrippedTitle, 255)
                End If

                ' Title column
                SqlForSpecificValuesInMainTable = _
                        SqlForSpecificValuesInMainTable & ", " & QuoteValueWithQuotes(sStrippedTitle)
                
                
        Case ECATALOGUELEVEL.ctSubHeaders
                
                SqlForSpecificValuesInMainTable = CLng(m_vCommonReference)
                
                'Date related coloumns
          
                SqlForSpecificValuesInMainTable = _
                        SqlForSpecificValuesInMainTable & ", " & QuoteValueWithQuotes(CoveringDateText)
                
                ' tomh d35 (added this bit)
                ' Need to format the title string for tbl_header table
                sStrippedTitle = RemoveTagsFromString(TitleText)
                sStrippedTitle = SpecialCharacterDecode(sStrippedTitle)
                
                ' want to check title not too long
                If Len(sStrippedTitle) > 255 Then
                    sStrippedTitle = Left(sStrippedTitle, 255)
                End If

                ' Title column
                SqlForSpecificValuesInMainTable = _
                        SqlForSpecificValuesInMainTable & ", " & QuoteValueWithQuotes(sStrippedTitle)

                
        Case ECATALOGUELEVEL.ctPiece
                
                'Piece key order and reference
                SqlForSpecificValuesInMainTable = _
                            m_lNewPieceKeyOrder & ", " & _
                            QuoteValueWithQuotes(CStr(m_vCommonReference))
                
                If GetCatalogueLevelEnum(CInt(m_vCommonParentCatalogueLevelID)) = ECATALOGUELEVEL.ctHeaders Then
                    'Creating piece under header - set the class_id
                    SqlForSpecificValuesInMainTable = SqlForSpecificValuesInMainTable & ", " & _
                        m_lParentClassID
                ElseIf GetCatalogueLevelEnum(CInt(m_vCommonParentCatalogueLevelID)) = ECATALOGUELEVEL.ctSubHeaders Then
                    'Creating piece under sub header - set the class_id and headerid
                    SqlForSpecificValuesInMainTable = SqlForSpecificValuesInMainTable & ", " & _
                    m_lParentClassID & "," & m_lParentHeaderID
                        
                End If
                
                'Columns populated with constants
                SqlForSpecificValuesInMainTable = SqlForSpecificValuesInMainTable & "," & _
                    Quotes(gk_sDefaultClosureStatus) & "," & _
                    Quotes(gk_sDefaultClosureType) & "," & _
                    gk_lDefaultClosureCode & "," & _
                    gk_lDefaultRecordStatus
                
                'Date related coloumns
                SqlForSpecificValuesInMainTable = _
                        SqlForSpecificValuesInMainTable & ", " & _
                        QuoteValueWithQuotes(CoveringDateText) & ", " & _
                        IIf(VariantStringAsString(LanguageID) = "", "Null", LanguageID)
                        
                SqlForSpecificValuesInMainTable = _
                        SqlForSpecificValuesInMainTable & ", " & _
                        CStr(StartDate)
                        
                SqlForSpecificValuesInMainTable = _
                        SqlForSpecificValuesInMainTable & ", " & _
                        CStr(EndDate)
                
                    
        Case ECATALOGUELEVEL.ctItem
                'Date related coloumns
                SqlForSpecificValuesInMainTable = _
                        QuoteValueWithQuotes(CoveringDateText) & ", " & _
                        IIf(VariantStringAsString(LanguageID) = "", "Null", LanguageID)
                        
                'Columns populated with constants
                SqlForSpecificValuesInMainTable = SqlForSpecificValuesInMainTable & "," & _
                    Quotes(gk_sDefaultClosureStatus) & "," & _
                    Quotes(gk_sDefaultClosureType) & "," & _
                    gk_lDefaultClosureCode & "," & _
                    gk_lDefaultRecordStatus
                
                'Item key order and reference
                SqlForSpecificValuesInMainTable = _
                    SqlForSpecificValuesInMainTable & ", " & _
                    m_lNewItemKeyOrder & ", " & _
                    QuoteValueWithQuotes(CStr(m_vCommonReference))
                    
                SqlForSpecificValuesInMainTable = _
                        SqlForSpecificValuesInMainTable & ", " & _
                        CStr(StartDate)
                        
                SqlForSpecificValuesInMainTable = _
                        SqlForSpecificValuesInMainTable & ", " & _
                        CStr(EndDate)

    End Select
    
    Exit Function
    
ErrorHandler:
    ConvertError sSource

End Function
                        
Private Function SqlForExtensionTable( _
                            ByVal t_eCatalogueCreateType As CatalogueCreateType, _
                            ByVal t_lNewCatalogueID As Long, _
                            ByVal t_eTargetCatalogueLevelID As ECATALOGUELEVEL, _
                            Optional ByVal t_eSourceCatalogueLevelID As ECATALOGUELEVEL, _
                            Optional ByVal t_lSourceCatalogueID As Long) As String
                    
' Description:
'
' Returns the insert sql for the extension eg es_lettercode_extension, es_class_extension
'
' Amendment history:
'
' d1   ndk      26Jul1999      Created
' d4   ndk      20Oct1999      Updating live_flag
' d11  ndk      03Nov1999      Adds all fields for the level when create by copying

    On Error GoTo ErrorHandler
    
    Dim sSqlColoumns As String, sSqlAdditionalColoumns As String, sSqlValues As String, sSqlAdditionalValues As String
    Dim sSource As String, sCommonFieldsToBeCopied As String
    
    sSource = Me.Name & ".SqlForExtensionTable"
    
    'common coloumns in the extension table
    sSqlColoumns = _
        GetColumnName(ECATALOGUECOLUMNTYPE.cctIDColumn, t_eTargetCatalogueLevelID) & "," & _
        gk_sDBCOLLiveFlag
    
    'Get the additional coloumns
    sSqlAdditionalColoumns = SqlForSpecificColoumnsInExtensionTable(t_eTargetCatalogueLevelID)
    
    If Not sSqlAdditionalColoumns = "" Then  'There are some additional coloumns
        sSqlColoumns = sSqlColoumns & ", " & sSqlAdditionalColoumns
    End If
    
    'Common values in the extension table
    sSqlValues = t_lNewCatalogueID & "," & _
        ELIVEFLAGSTATUS.lfsNotLive
    
    'Get the additional coloumns
    sSqlAdditionalValues = SqlForSpecificValuesInExtensionTable(t_eTargetCatalogueLevelID)
    
    If Not sSqlAdditionalValues = "" Then  'There are some additional values
        sSqlValues = sSqlValues & ", " & sSqlAdditionalValues
    End If

    'Coloumns and values
    If t_eCatalogueCreateType = ectCreateByCopy Then
        'All columns for that catalogue
        sCommonFieldsToBeCopied = SqlForExtensionTableForCreateByCopy(t_eTargetCatalogueLevelID, t_eSourceCatalogueLevelID)
        
        If Not sCommonFieldsToBeCopied = "" Then  'There are some additional values
            sSqlValues = sSqlValues & ", " & sCommonFieldsToBeCopied
        End If

        If Not sCommonFieldsToBeCopied = "" Then  'There are some additional values
            sSqlColoumns = sSqlColoumns & ", " & sCommonFieldsToBeCopied
        End If

        'creating catalogue by copy
        SqlForExtensionTable = _
            " INSERT " & GetTableName(ECATALOGUETABLETYPE.cttExtensionTable, ESectionType.SECReader, t_eTargetCatalogueLevelID) & " (" & _
                    sSqlColoumns & ")" & _
            " SELECT " & sSqlValues & _
            " FROM " & GetTableName(ECATALOGUETABLETYPE.cttExtensionTable, ESectionType.SECReader, t_eSourceCatalogueLevelID) & _
            " WHERE " & GetColumnName(ECATALOGUECOLUMNTYPE.cctIDColumn, t_eSourceCatalogueLevelID) & _
                "=" & t_lSourceCatalogueID
    Else
        'creating catalogue from scratch
        SqlForExtensionTable = _
            " INSERT " & GetTableName(ECATALOGUETABLETYPE.cttExtensionTable, ESectionType.SECReader, t_eTargetCatalogueLevelID) & " (" & _
                    sSqlColoumns & ")" & _
            " VALUES  (" & sSqlValues & ")"
    End If
        
    Exit Function
    
ErrorHandler:
    ConvertError sSource
End Function
                    
Private Function SqlForSpecificColoumnsInExtensionTable( _
                    ByVal t_eCatalogueLevelID As ECATALOGUELEVEL _
                    ) As String
     
' Description:  Returns the coloumns in which the data is to be inserted in the ext table
'
' Amendment history:
'   d1   ndk      26Jul1999      Created
'   d13  ndk      24Nov1999      Including access condition now
'   d31  alm      18Sep2001      Inserts columns for start and end dates
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    On Error GoTo ErrorHandler
    
     
    Dim sSource As String
    Dim StartDateCol As String
    Dim EndDateCol As String
    
    sSource = Me.Name & ".SqlForSpecificColoumnsInExtensionTable"
    
    ' d31 Get start and end date column names
    StartDateCol = GetColumnName(cctStartDateColumn, t_eCatalogueLevelID)
    EndDateCol = GetColumnName(cctEndDateColumn, t_eCatalogueLevelID)
         
     'Additional coloumns specific to tables
    Select Case t_eCatalogueLevelID
        Case ECATALOGUELEVEL.ctLetterCode, ECATALOGUELEVEL.ctDivision
                'For letter code and division
                SqlForSpecificColoumnsInExtensionTable = _
                        gk_sDBCOLCoveringDate & ", " & _
                        StartDateCol & ", " & _
                        EndDateCol & ", " & _
                        gk_sDBCOLLanguageID & ", " & _
                        gk_sDBCOLLegalStatusID & "," & _
                        gk_sDBCOLAccessConditionID

        Case ECATALOGUELEVEL.ctClass
                'For class
                SqlForSpecificColoumnsInExtensionTable = _
                        gk_sDBCOLLanguageID & ", " & _
                        gk_sDBCOLLegalStatusID & "," & _
                        gk_sDBCOLAccessConditionID

        Case ECATALOGUELEVEL.ctHeaders, ECATALOGUELEVEL.ctSubHeaders
                'For header and sub header
                SqlForSpecificColoumnsInExtensionTable = _
                        gk_sDBCOLLanguageID & ", " & _
                        gk_sDBCOLLegalStatusID & "," & _
                        gk_sDBCOLAccessConditionID & ", " & _
                        StartDateCol & ", " & _
                        EndDateCol

       Case ECATALOGUELEVEL.ctPiece, ECATALOGUELEVEL.ctItem
                'For piece
                SqlForSpecificColoumnsInExtensionTable = gk_sDBCOLLegalStatusID
    End Select
    
    Exit Function
    
ErrorHandler:
    ConvertError sSource
End Function
Private Function SqlForSpecificValuesInExtensionTable( _
                    ByVal t_eCatalogueLevelID As ECATALOGUELEVEL _
                    ) As String

' Description:  Returns the values to be inserted in the ext table
'
' Amendment history:
'   d1   ndk      26Jul1999      Created
'   d13  ndk      24Nov1999      Including access condition now
'   d31  alm      18Sep2001      Inserts values for start and end dates
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    On Error GoTo ErrorHandler
    
    Dim sSource As String
    
    sSource = Me.Name & ".SqlForSpecificValuesInExtensionTable"

    Select Case t_eCatalogueLevelID
        Case ECATALOGUELEVEL.ctLetterCode, ECATALOGUELEVEL.ctDivision
                'For letter code and division

               SqlForSpecificValuesInExtensionTable = _
                        QuoteValueWithQuotes(CoveringDateText) & ", " & _
                        CStr(StartDate) & ", " & _
                        CStr(EndDate) & ", " & _
                        IIf(VariantStringAsString(LanguageID) = "", "Null", LanguageID) & ", " & _
                        IIf(VariantStringAsString(LegalStatusID) = "", "Null", LegalStatusID) & ", " & _
                        IIf(VariantStringAsString(AccessConditionID) = "", "Null", AccessConditionID)
        Case ECATALOGUELEVEL.ctClass
                'For class
               SqlForSpecificValuesInExtensionTable = _
                        IIf(VariantStringAsString(LanguageID) = "", "Null", LanguageID) & ", " & _
                        IIf(VariantStringAsString(LegalStatusID) = "", "Null", LegalStatusID) & ", " & _
                        IIf(VariantStringAsString(AccessConditionID) = "", "Null", AccessConditionID)
        Case ECATALOGUELEVEL.ctHeaders, ECATALOGUELEVEL.ctSubHeaders
                'For header and sub header
               SqlForSpecificValuesInExtensionTable = _
                        IIf(VariantStringAsString(LanguageID) = "", "Null", LanguageID) & ", " & _
                        IIf(VariantStringAsString(LegalStatusID) = "", "Null", LegalStatusID) & ", " & _
                        IIf(VariantStringAsString(AccessConditionID) = "", "Null", AccessConditionID) & ", " & _
                        CStr(StartDate) & ", " & _
                        CStr(EndDate)

       Case ECATALOGUELEVEL.ctPiece, ECATALOGUELEVEL.ctItem
                'For piece
                SqlForSpecificValuesInExtensionTable = _
                        IIf(VariantStringAsString(LegalStatusID) = "", "Null", LegalStatusID)
    End Select
    
    Exit Function
    
ErrorHandler:
    ConvertError sSource
End Function


Private Sub InsertRows( _
        ByVal t_eTableType As ECATALOGUETABLETYPE, _
        ByVal t_eCatalogueLevelID As ECATALOGUELEVEL, _
        ByVal t_sIDColoumnName As String, _
        ByVal t_lNewCatalogueID As Long, _
        Optional ByVal t_sFieldNames As String, _
        Optional ByVal t_sFieldContents As String)
' Description:  Inserts rows in the live and editorial table while creating entry by scratch
'
' Amendment history:
'   d1   ndk      26Jul1999      Created
'   d7   ndk      23Oct1999      does not any more insert in the editorial tables
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    On Error GoTo ErrorHandler
    
    Dim sSql As String, sSource As String
    
    sSource = Me.Name & ".InsertRows"
    
    'For the live table
    sSql = _
        "INSERT " & GetTableName(t_eTableType, ESectionType.SECReader, t_eCatalogueLevelID) & "(" & _
            t_sIDColoumnName & _
            IIf(t_sFieldNames = "", "", "," & t_sFieldNames) & ")" & _
        "VALUES (" & _
            t_lNewCatalogueID & _
            IIf(t_sFieldContents = "", "", "," & t_sFieldContents) & ")"

    m_oFWDBHelper.Execute sSql
        
    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub
        

Private Function InsertRowsByCopying( _
    ByVal t_eTableType As ECATALOGUETABLETYPE, _
    ByVal t_eTargetCatalogueLevelID As ECATALOGUELEVEL, _
    ByVal t_eSourceCatalogueLevelID As ECATALOGUELEVEL, _
    ByVal t_sIDColoumnNameForTarget As String, _
    ByVal t_sIDColoumnNameForSource As String, _
    ByVal t_lNewCatalogueID As Long, _
    ByVal t_lSourceCatalogueID As Long, _
    Optional ByVal t_sFieldNames As String, _
    Optional ByVal t_sFieldContents As String)

' Description:  Inserts rows in the live and editorial table while creating entry by copying
'
' Amendment history:
'   d1   ndk      26Jul1999      Created
'   d7   ndk      23Oct1999      does not any more insert in the editorial tables
'   d8   ndk      26Oct1999      Minor change
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    On Error GoTo ErrorHandler
    
    Dim sSql As String, sSource As String
    
    sSource = Me.Name & ".InsertRowsByCopying"
    'For the Live table
    sSql = _
        " INSERT " & GetTableName(t_eTableType, ESectionType.SECReader, t_eTargetCatalogueLevelID) & "(" & _
            t_sIDColoumnNameForTarget & _
            IIf(t_sFieldNames = "", "", "," & t_sFieldNames) & ")" & _
        " SELECT " & _
            t_lNewCatalogueID & _
            IIf(t_sFieldContents = "", "", "," & t_sFieldContents) & _
        " FROM " & GetTableName(t_eTableType, ESectionType.SECReader, t_eSourceCatalogueLevelID) & _
        " WHERE " & t_sIDColoumnNameForSource & "=" & t_lSourceCatalogueID

    m_oFWDBHelper.Execute sSql

    Exit Function
    
ErrorHandler:
    ConvertError sSource
End Function
        

Public Function ImportCataloguesIntoAnEditSet(t_lEditSetID As Long, _
                                            t_ofwRecordList As FWRecordList, _
                                            t_eCatalogueLevel As ECATALOGUELEVEL)
'Imports the list of catalogues passed as a fwrecord list into an editset
' All the catalogues in the FWRecordList should be of the same level
'   and the ID column should have the name in the constant gk_sDBCOLCatalogueID
'
'   d15  ksk      18Nov1999      Added a method similar to Add catalogue entry but this time for import
'                                ImportCataloguesIntoAnEditSet
'   d29  daj      27Mar2000      Added extra error handling to cope with the error thrown by stored procedures when
'                                attempting to add a non-live parent to an edit set
'
On Error GoTo ErrHandler
    Dim i As Integer
    Dim sSource As String, sMessage As String
    Dim lCatalogueID As Long
    
    sSource = Me.Name & ".ImportCataloguesIntoAnEditSet"
    
    
    'Generate and execute the SQL using the parameter
    While Not t_ofwRecordList.IsAtEnd
    
        lCatalogueID = t_ofwRecordList.GetFieldValue(gk_sDBCOLCatalogueID)
        
        sMessage = "Importing " & _
                    " Catalogue ID - " & lCatalogueID & _
                    " Catalogue Level - " & GetCatalogueLevelDescription( _
                                    GetCatalogueLevelCode(t_eCatalogueLevel))
        LogMessage sMessage, EProfile.pComplete
         
        If Not CheckIfCatalogueExistsAsEditable(lCatalogueID, t_eCatalogueLevel) Then
            AddCatalogueEntryWithParentsIntoEditset t_lEditSetID, lCatalogueID, _
                        t_eCatalogueLevel, ECATALOGUEENTRYSTATUS.cesImported, _
                        ELIVEFLAGSTATUS.lfsNotLive, ECATALOGUEENTRYSTATUS.cesReadeOnly
        Else
            RaiseError sSource, EPROCATINTERNALERROR.PECatalogue, _
                gk_sErrorDescriptionAddDuplicateEditableCatalogue & _
                " Catalogue ID - " & lCatalogueID & _
                " Catalogue Level - " & GetCatalogueLevelDescription( _
                                    GetCatalogueLevelCode(t_eCatalogueLevel))
        End If
        
        
        t_ofwRecordList.MoveNext
    Wend
    
    Exit Function
    
ErrHandler:
    ' d29 - Added extra handling to cope with adding a non-live parent to editset
    If InStr(Err.Description, gk_sErrorDescriptionNonLiveParent) > 0 Then
        ' Remove the SQL rubbish from the start
        Err.Description = Mid(Err.Description, InStr(Err.Description, gk_sErrorDescriptionNonLiveParent))
        Err.Number = EPROCATINTERNALERROR.PECannotAddNonLiveParent
    End If
    ConvertError sSource
End Function

Public Function AddCataloguesIntoEditSet(t_lEditSetID As Long, t_lCatalogueID() As Long, _
                        t_eCatalogueLevel() As ECATALOGUELEVEL)
'Add the passed list of catalogues into the editset
'   This method Skips adding catalogues into an editset which are already present
'   in the editset but selected agin by the user to add.
'   d2   ndk      17Aug1999      Moved AddCataloguesIntoEditSet from BOEditSet (Created by KSK)
'   d5   ksk      21Oct1999      Changed to reflect the new private method to add catalogue with parents
'   d19  tmj      21Feb2000      Changed the error message when entry already in an editset to display
'                                catalogue reference rather than catalogue ID
'   d23  tmj      03Mar2000      Modified to enable handling of case when parent has been moved
'   d29  daj      27Mar2000      Added extra error handling to cope with the error thrown by stored procedures when
'                                attempting to add a non-live parent to an edit set
'   d32  alm      25Oct2001      Checks if the editset is a delete editset and if it is then
'                                makes sure the entry being added is valid
'   d34  alm      30Oct2001      Shows a different message for an item than for other levels when
'                                it is determined that it is referenced
On Error GoTo ErrHandler
    Dim i              As Integer
    Dim sSource        As String
    Dim sMessage       As String
    Dim bDeleteEditSet As Boolean
    Dim oBOEditSet     As BOEditSet
    
    sSource = Me.Name & ".AddCataloguesIntoEditSet"
    
    'Make sure that the number of catalogueID elements match Catalogue levels
    If UBound(t_lCatalogueID) <> UBound(t_eCatalogueLevel) Then
        RaiseError sSource, peEditSet, "Catalogue ID and Catalogue Level List do not match"
    End If
    
    'd32 - Get the type of the edit set
    Set oBOEditSet = New BOEditSet
    oBOEditSet.EditsetID = t_lEditSetID
    oBOEditSet.Load
    
    'd32 Determine if the editset is a delete edit set
    '(4 is hardcoded as the edit set types do not match the values
    'in the enumeration ECreateEditSetUsage)
    If oBOEditSet.EditSetTypeID = 4 Then
        bDeleteEditSet = True
    Else
        bDeleteEditSet = False
    End If
    
    Set oBOEditSet = Nothing
    
    'Moved error chrecking to here so it is separated.
    'This is better as either all or none of the entries are added, giving the user more
    'certainty
    For i = 1 To UBound(t_lCatalogueID)
        If CheckIfCatalogueExistsAsEditable(t_lCatalogueID(i), t_eCatalogueLevel(i)) Then
            RaiseError sSource, EPROCATINTERNALERROR.PECatalogue, _
                gk_sErrorDescriptionAddDuplicateEditableCatalogue & _
                " Catalogue Ref - " & GenerateCatalogueReferenceFromID(t_lCatalogueID(i), t_eCatalogueLevel(i)) & _
                " Catalogue Level - " & GetCatalogueLevelDescription( _
                                    GetCatalogueLevelCode(t_eCatalogueLevel(i)))
        End If
        If Not CheckWhetherParentNotMoved(t_lCatalogueID(i), t_eCatalogueLevel(i), t_lEditSetID) Then
            RaiseError sSource, EPROCATINTERNALERROR.PECatalogue, _
                gk_sErrorDescriptionParentMoved & _
                " Catalogue Ref - " & GenerateCatalogueReferenceFromID(t_lCatalogueID(i), t_eCatalogueLevel(i)) & _
                " Catalogue Level - " & GetCatalogueLevelDescription( _
                                    GetCatalogueLevelCode(t_eCatalogueLevel(i)))
        End If
        
        'd32 - If a delete edit set check that the entries are non-referenced items
        If bDeleteEditSet = True Then
            If CheckCatalogueIsReferenced(t_lCatalogueID(i), t_eCatalogueLevel(i)) Then
                
                ' d34 The error depends on the catalogue entry
                If t_eCatalogueLevel(i) <> ctItem Then
                    RaiseError sSource, EPROCATINTERNALERROR.PECatalogue, _
                        gk_sErrorDeleteCatalogueHasReferences & _
                        " (Catalogue Ref - " & GenerateCatalogueReferenceFromID(t_lCatalogueID(i), t_eCatalogueLevel(i)) & _
                        ", Catalogue Level - " & GetCatalogueLevelDescription( _
                                            GetCatalogueLevelCode(t_eCatalogueLevel(i))) & ")<br>"
                Else
                    RaiseError sSource, EPROCATINTERNALERROR.PECatalogue, _
                        gk_sErrorDeleteCatalogueItemHasReferences & _
                        " (Catalogue Ref - " & GenerateCatalogueReferenceFromID(t_lCatalogueID(i), t_eCatalogueLevel(i)) & _
                        ", Catalogue Level - " & GetCatalogueLevelDescription( _
                                            GetCatalogueLevelCode(t_eCatalogueLevel(i))) & ")<br>"
                End If
            End If
        End If
    Next i
    
    'Generate and execute the SQL using the parameter
    '   d10  ksk      02Nov1999 Created
    For i = 1 To UBound(t_lCatalogueID)
        AddCatalogueEntryWithParentsIntoEditset t_lEditSetID, t_lCatalogueID(i), _
                    t_eCatalogueLevel(i), ECATALOGUEENTRYSTATUS.cesAddedFromLive, _
                    ELIVEFLAGSTATUS.lfsLiveAndAddedToEditorial, ECATALOGUEENTRYSTATUS.cesReadeOnly
        
        sMessage = "Importing " & _
                    " Catalogue ID - " & t_lCatalogueID(i) & _
                    ", Catalogue Level - " & GetCatalogueLevelDescription( _
                                        GetCatalogueLevelCode(t_eCatalogueLevel(i)))
        LogMessage sMessage, EProfile.pComplete
    
    
    Next i
    
    Exit Function
    
ErrHandler:
    ' d29 - Added extra handling to cope with adding a non-live parent to editset
    If InStr(Err.Description, gk_sErrorDescriptionNonLiveParent) > 0 Then
        ' Remove the SQL rubbish from the start
        Err.Description = Mid(Err.Description, InStr(Err.Description, gk_sErrorDescriptionNonLiveParent))
        Err.Number = EPROCATINTERNALERROR.PECannotAddNonLiveParent
    End If
    ConvertError sSource
End Function


Public Function AddSingleCatalogueEntryIntoEditSet(t_lEditSetID As Long, t_lCatalogueID As Long, _
                        t_eCatalogueLevel As ECATALOGUELEVEL)
' Add a single catalogue entry into the editset, checking for editable and moved, etc. first
'
'   d30  daj      20Jul2000      Created
On Error GoTo ErrHandler
    
    Dim sSource As String, sMessage As String
    
    sSource = Me.Name & ".AddSingleCatalogueEntryIntoEditSet"
    
    ' Check that the values passed in are sensible
    If (t_lEditSetID = 0) Or (t_lCatalogueID = 0) Or (t_eCatalogueLevel = ctNone) Then
        RaiseError sSource, EPROCATINTERNALERROR.PECatalogue, "Must supply edit set ID, catalogue ID and level"
    End If
    
    ' Check to see if the entry is editable or has moved
    If CheckIfCatalogueExistsAsEditable(t_lCatalogueID, t_eCatalogueLevel) Then
        RaiseError sSource, EPROCATINTERNALERROR.PECatalogue, _
            gk_sErrorDescriptionAddDuplicateEditableCatalogue & _
            " Catalogue Ref - " & GenerateCatalogueReferenceFromID(t_lCatalogueID, t_eCatalogueLevel) & _
            " Catalogue Level - " & GetCatalogueLevelDescription( _
                                GetCatalogueLevelCode(t_eCatalogueLevel))
    End If
    
    If Not CheckWhetherParentNotMoved(t_lCatalogueID, t_eCatalogueLevel, t_lEditSetID) Then
        RaiseError sSource, EPROCATINTERNALERROR.PECatalogue, _
            gk_sErrorDescriptionParentMoved & _
            " Catalogue Ref - " & GenerateCatalogueReferenceFromID(t_lCatalogueID, t_eCatalogueLevel) & _
            " Catalogue Level - " & GetCatalogueLevelDescription( _
                                GetCatalogueLevelCode(t_eCatalogueLevel))
    End If
    
    ' If we get this far, it's OK to add it into the database
    sMessage = "Importing " & _
                " Catalogue ID - " & t_lCatalogueID & _
                " Catalogue Level - " & GetCatalogueLevelDescription( _
                                    GetCatalogueLevelCode(t_eCatalogueLevel))
    LogMessage sMessage, EProfile.pComplete
    
    AddCatalogueEntryWithParentsIntoEditset t_lEditSetID, t_lCatalogueID, _
                t_eCatalogueLevel, ECATALOGUEENTRYSTATUS.cesAddedFromLive, _
                ELIVEFLAGSTATUS.lfsLiveAndAddedToEditorial, ECATALOGUEENTRYSTATUS.cesReadeOnly
    
    Exit Function
    
ErrHandler:
    If InStr(Err.Description, gk_sErrorDescriptionNonLiveParent) > 0 Then
        ' Remove the SQL rubbish from the start
        Err.Description = Mid(Err.Description, InStr(Err.Description, gk_sErrorDescriptionNonLiveParent))
        Err.Number = EPROCATINTERNALERROR.PECannotAddNonLiveParent
    End If
    ConvertError sSource
End Function

Private Function AddCatalogueEntryWithParentsIntoEditset(t_lEditSetID As Long, _
                        t_lCatalogueID As Long, _
                        t_eCatalogueLevel As ECATALOGUELEVEL, _
                        t_eAddedCatalogueEntryStatusToSet As ECATALOGUEENTRYSTATUS, _
                        t_eAddedCatalogueEntryLiveStatusToSet As ELIVEFLAGSTATUS, _
                        Optional t_eParentCatalogueEntryStatus As _
                            ECATALOGUEENTRYSTATUS = ECATALOGUEENTRYSTATUS.cesReadeOnly)
' This private method takes in the edit set (t_lEditSetID) into which a catalogue entry
' (t_lCatalogueID, t_eCatalogueLevel) with the status set to t_eAddedCatalogueEntryStatusToSet
' needs to be added along with the list of its parents with status of t_eParentCatalogueEntryStatus.
' The default value of the t_eParentCatalogueEntryStatus is to set the parents as readonly entries.
'   d5   ksk      21Oct1999      This method to adds the catalogue along with parents with
'                                   appropriate status into edit set
'   d29  daj      27Mar2000      Added extra error handling to cope with the error thrown by stored procedures when
'                                attempting to add a non-live parent to an edit set
'
On Error GoTo ErrHandler
    Dim sSql As String
    Dim sSource As String
    
    sSource = Me.Name & ".AddCatalogueEntryWithParentsIntoEditset"
    

    sSql = "Execute sp_Edit_AddCatalogueEntryDetailsWithParents " & _
                    " @t_iEditsetID = " & CStr(t_lEditSetID) & _
                    ", @t_iCatalogueID = " & CStr(t_lCatalogueID) & _
                    ", @t_iCatalogueLevel =" & _
                        GetCatalogueLevelCode(t_eCatalogueLevel) & _
                    ", @t_iEditsetEntryTypeValueToSet = " & _
                            CStr(t_eAddedCatalogueEntryStatusToSet) & _
                    ", @t_iLiveFlagValueToSet = " & _
                            CStr(t_eAddedCatalogueEntryLiveStatusToSet) & _
                    ", @t_iReadOnlyEditsetEntryType = " & _
                            CStr(t_eParentCatalogueEntryStatus)
                
    m_oFWDBHelper.Execute sSql
    Exit Function
    
ErrHandler:
    ' d29 - Added handling of adding non-live entry to edit set
    If InStr(Err.Description, gk_sErrorDescriptionNonLiveParent) > 0 Then
        ' Remove the SQL rubbish from the start
        Err.Description = Mid(Err.Description, InStr(Err.Description, gk_sErrorDescriptionNonLiveParent))
        Err.Number = EPROCATINTERNALERROR.PECannotAddNonLiveParent
    End If
    ConvertError sSource
End Function

Private Function SetDataToBeGenerated()
'
'Description
'Generates values for the variables
'm_iNewPieceKeyOrder    -Piece key order while inserting pieces
'm_iNewItemKeyOrder     -Item key order while inserting pieces
'm_lParentLetterCodeID  -lettercode id for class under division
'm_lParentClassID       -class_id for piece under header and subheader
'm_lParentHeaderID      -header_id for piece under sub header
'
' Amendment history:
'
'   d3   ndk      08Oct1999     Created
'   d23  WDP      06Mar2000      Fixed defect in Piece/item key
'                                order generation in SetDataToBeGenerated
    On Error GoTo ErrorHandler

    Dim sSource As String, sSql As String
    Dim oFWRecordlist As FWRecordList
    
    sSource = Me.Name & ".SetDataToBeGenerated"
    
    'If a new class is being created under a division
    If GetCatalogueLevelEnum(CInt(m_vCommonCatalogueLevelID)) = ECATALOGUELEVEL.ctClass And _
        GetCatalogueLevelEnum(CInt(m_vCommonParentCatalogueLevelID)) = ECATALOGUELEVEL.ctDivision Then
        
        'Get the parent lettercode of the division
        
        'Construct the sql
        sSql = _
            " Select " & gk_sDBCOLLetterCodeID & _
            " From " & gk_sDBTABLEDivision & _
            " Where " & gk_sDBCOLDivisionID & " = " & CLng(m_vCommonParentCatalogueID)
            
        'Run the query
        Set oFWRecordlist = m_oFWDBHelper.FindBySql(sSql)
        
        'Get the lettercode id
        m_lParentLetterCodeID = oFWRecordlist.GetFieldValue(gk_sDBCOLLetterCodeID)
        
    'Piece under header
    ElseIf GetCatalogueLevelEnum(CInt(m_vCommonCatalogueLevelID)) = ECATALOGUELEVEL.ctPiece And _
        GetCatalogueLevelEnum(CInt(m_vCommonParentCatalogueLevelID)) = ECATALOGUELEVEL.ctHeaders Then
        
        'Get the class id for the piece
        
        'Construct the sql
        sSql = _
            " Select " & gk_sDBCOLClassID & _
            " From " & gk_sDBTABLEHeader & _
            " Where " & gk_sDBCOLHeaderID & " = " & CLng(m_vCommonParentCatalogueID)
            
        'Run the query
        Set oFWRecordlist = m_oFWDBHelper.FindBySql(sSql)
        
        'Get the class id
        m_lParentClassID = oFWRecordlist.GetFieldValue(gk_sDBCOLClassID)
        
    'Piece under sub header (get headerid and classid)
    ElseIf GetCatalogueLevelEnum(CInt(m_vCommonCatalogueLevelID)) = ECATALOGUELEVEL.ctPiece And _
        GetCatalogueLevelEnum(CInt(m_vCommonParentCatalogueLevelID)) = ECATALOGUELEVEL.ctSubHeaders Then
        
        'Get the header id for the piece
        'Construct the sql
        sSql = _
            " Select " & gk_sDBCOLHeaderID & _
            " From " & gk_sDBTABLESubHeader & _
            " Where " & gk_sDBCOLSubHeaderID & " = " & CLng(m_vCommonParentCatalogueID)
            
        'Run the query
        Set oFWRecordlist = m_oFWDBHelper.FindBySql(sSql)
        
        'Get the header id for the piece
        m_lParentHeaderID = oFWRecordlist.GetFieldValue(gk_sDBCOLHeaderID)
        
        
        'Get the class id for the piece
        'Construct the sql (Join subheader and header as there is no class_id in subheader)
        sSql = _
            " Select " & gk_sDBCOLClassID & _
            " From " & gk_sDBTABLESubHeader & "," & gk_sDBTABLEHeader & _
            " Where " & gk_sDBCOLSubHeaderID & " = " & CLng(m_vCommonParentCatalogueID) & " AND " & _
                        gk_sDBTABLEHeader & "." & gk_sDBCOLHeaderID & "=" & _
                            gk_sDBTABLESubHeader & "." & gk_sDBCOLHeaderID

            
        'Run the query
        Set oFWRecordlist = m_oFWDBHelper.FindBySql(sSql)
        
        'Get the class id
        m_lParentClassID = oFWRecordlist.GetFieldValue(gk_sDBCOLClassID)
            
    End If

    'piece key order
    If GetCatalogueLevelEnum(CInt(m_vCommonCatalogueLevelID)) = ECATALOGUELEVEL.ctPiece Then
        'Creating a piece, get the piece key order
        
        Dim lClassID As Long
        If Not CInt(m_vCommonParentCatalogueLevelID) = ctClass Then
            lClassID = m_lParentClassID
        Else
            lClassID = CLng(m_vCommonParentCatalogueID)
        End If
        'Construct the sql
        sSql = "Select Max(" & gk_sDBColPiecekeyOrder & ") from " & gk_sDBTABLEPiece & _
                " WHERE " & gk_sDBCOLClassID & "=" & lClassID
        
        'Run the query
        Set oFWRecordlist = m_oFWDBHelper.FindBySql(sSql)
        
        If oFWRecordlist.RecordCount > 0 Then
            'New piece key order
            m_lNewPieceKeyOrder = gk_iIncrementForPieceKeyOrder + _
                IIf(IsNull(oFWRecordlist.GetFieldValue(0)), 0, oFWRecordlist.GetFieldValue(0))
        Else
            m_lNewPieceKeyOrder = 0
        End If
        
    End If
    
    'item key order
    If GetCatalogueLevelEnum(CInt(m_vCommonCatalogueLevelID)) = ECATALOGUELEVEL.ctItem Then
        'Creating a item, get the itemkey order
        
        'Construct the sql
        sSql = "Select Max(" & gk_sDBCOLItemKeyOrder & ") from " & gk_sDBTABLEItem & _
                " WHERE " & gk_sDBCOLPieceID & "=" & CLng(m_vCommonParentCatalogueID)
        
        'Run the query
        Set oFWRecordlist = m_oFWDBHelper.FindBySql(sSql)
        
        'New item key order
        m_lNewItemKeyOrder = gk_iIncrementForItemKeyOrder + _
            IIf(IsNull(oFWRecordlist.GetFieldValue(0)), 0, oFWRecordlist.GetFieldValue(0))
    End If
    
    Set oFWRecordlist = Nothing
    
    Exit Function
ErrorHandler:
    ConvertError sSource
End Function

Public Function AuditTrailForCatalogue( _
    ByVal t_lCatalogueID As Long, _
    ByVal t_eCatalogueLevelID As ECATALOGUELEVEL) As FWRecordList
    
'   Returns the list of last two editsets where the catalogue is present
'   These editsets are the audit_trail editsets
'
' Amendment history:
'
'   d3   ndk      08Oct1999     Created

    On Error GoTo ErrorHandler

    Dim sSource As String, sSql As String
    
    sSource = Me.Name & ".AuditTrailForCatalogue"
    
    'Construct the sql
    sSql = "sp_Edit_GetAuditTrailForCatalogue " & _
            "@t_iCatalogueID = " & t_lCatalogueID & "," & _
            "@t_iCatalogueLevel = " & GetCatalogueLevelCode(t_eCatalogueLevelID)
    
    Set AuditTrailForCatalogue = m_oFWDBHelper.FindBySql(sSql)
    
    Exit Function
    
ErrorHandler:
    ConvertError sSource
End Function


Private Sub AddLinkItemsToReader( _
    t_eLinkTermType As ELinkTermType, _
    t_colLinkItems As Collection, _
    t_lCatalogueID As Long, _
    t_eCatalogueLevelID As ECATALOGUELEVEL)
'
'Description: Update the reader tables with creator names and place of deposit
'             This is used when create by copying and create from scratch
'
' Amendment history:
'   d7   ndk      23Oct1999      Created
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String, sSql As String
    Dim sMasterTable As String, sLinksTable As String
    Dim sIDColumnName As String, eProperty As ECATALOGUEPROPERTIES
    Dim recLinkItemRecord  As Variant
    
    sSource = Me.Name & ".AddLinkItemsToReader"

    'case structure to get the following details
    'Links table - table to extract link item details from
    'Master table - master table for the link items (af files)
    'ID column name - id column for the link and master tables
    'Property - property id of the link item
    
    Select Case t_eLinkTermType
        Case ELinkTermType.elCreatorNamesCorporateBody
            sLinksTable = GetTableName(cttCorporateBody, ESectionType.SECReader)
            sMasterTable = gk_sDBTABLECorporateBody
            sIDColumnName = gk_sDBCOLCorporateBodyReferenceID
            eProperty = ECATALOGUEPROPERTIES.cpCreatorNames
        Case ELinkTermType.elCreatorNamesPerson
            sLinksTable = GetTableName(cttPerson, ESectionType.SECReader)
            sMasterTable = gk_sDBTABLEPerson
            sIDColumnName = gk_sDBCOLPersonReferenceID
            eProperty = ECATALOGUEPROPERTIES.cpCreatorNames
        Case ELinkTermType.elPlaceOfDeposit
            sLinksTable = GetTableName(cttCorporateBody, ESectionType.SECReader)
            sMasterTable = gk_sDBTABLECorporateBody
            sIDColumnName = gk_sDBCOLCorporateBodyReferenceID
            eProperty = ECATALOGUEPROPERTIES.cpPlaceOfDeposit
    End Select
    
    'To add
    For Each recLinkItemRecord In t_colLinkItems
        'Columns
        sSql = _
            " INSERT " & sLinksTable & "(" & _
            sIDColumnName & "," & _
            gk_sDBCOLCatalogueID & "," & _
            gk_sDBCOLCatalogueLevel & "," & _
            gk_sDBCOLPropertyID & ")"
    
        'Values
        sSql = _
            sSql & " VALUES ( " & _
            recLinkItemRecord.lID & "," & _
            t_lCatalogueID & "," & _
            GetCatalogueLevelCode(t_eCatalogueLevelID) & "," & _
            eProperty & ")"
        
        'Execute the sql
        m_oFWDBHelper.Execute sSql
    
    Next
    
    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub


Private Sub CopyAllLinksToReader( _
    ByVal t_lSourceCatalogueID As Long, _
    ByVal t_eSourceCatalogueLevel As ECATALOGUELEVEL, _
    ByVal t_lTargetCatalogueID As Long, _
    ByVal t_eTargetCatalogueLevel As ECATALOGUELEVEL)
'
'   Copies the link items from the source catalogue to target catalogue in reader
'
' Amendment history:
'
'   d7   ndk      23Oct1999      Created

    On Error GoTo ErrorHandler

    Dim sSource As String, sSql As String
    
    sSource = Me.Name & ".CopyAllLinksToReader"
    
    'create the query
    sSql = _
        "sp_Edit_CopyAllLinksToReaderForCreateByCopying " & _
            "@t_iSourceCatalogueID = " & t_lSourceCatalogueID & "," & _
            "@t_iSourceCatalogueLevel = " & GetCatalogueLevelCode(t_eSourceCatalogueLevel) & "," & _
            "@t_iTargetCatalogueID = " & t_lTargetCatalogueID & "," & _
            "@t_iTargetCatalogueLevel = " & GetCatalogueLevelCode(t_eTargetCatalogueLevel) & "," & _
            "@t_iPropertyIDForCreatorNames = " & ECATALOGUEPROPERTIES.cpCreatorNames & "," & _
            "@t_iPropertyIDForPlaceOfDeposit = " & ECATALOGUEPROPERTIES.cpPlaceOfDeposit
    'run the sp
    m_oFWDBHelper.Execute sSql
            
    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub


    

  
Private Function CheckIfCatalogueExistsAsEditable(t_lCatalogueID As Long, _
                                        t_eCatalogueLevel As ECATALOGUELEVEL) As Boolean
'Checks in all editsets the passed catalogue entry to be unique
'   d10  ksk      02Nov1999 Created
'
On Error GoTo ErrorHandler

Dim sSource As String, sSql As String
Dim sTablename As String
Dim sIDColumnName As String
Dim ofwResult As FWRecordList

    sSource = Me.Name & ".CheckIfCatalogueExistsAsEditable"
    
    CheckIfCatalogueExistsAsEditable = False
    
    sTablename = GetTableName(ECATALOGUETABLETYPE.cttExtensionTable, _
                    ESectionType.SECeditorial, t_eCatalogueLevel)
    sIDColumnName = GetColumnName(ECATALOGUECOLUMNTYPE.cctIDColumn, t_eCatalogueLevel)
    sSql = "SELECT * FROM " & sTablename & " WHERE " & sIDColumnName & " = " & CStr(t_lCatalogueID)
    sSql = sSql & " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & ECATALOGUEENTRYSTATUS.cesReadeOnly
    
    Set ofwResult = m_oFWDBHelper.FindBySql(sSql)
    If Not ofwResult Is Nothing Then
        If ofwResult.RecordCount > 0 Then
            CheckIfCatalogueExistsAsEditable = True
        End If
    End If
    
    Exit Function
    
ErrorHandler:
    ConvertError sSource
End Function

Private Function CheckCatalogueIsReferenced(t_lCatalogueID As Long, _
                                            t_eCatalogueLevel As ECATALOGUELEVEL) As Boolean
'Description:
'  Checks that the catalogue entry does not have publically visable references
'
' Amendment history:
'   d32  alm      24Oct2001     Created
'   d33  alm      29Oct2001     Altered sql for item to detect blank as well
'                               as null references
'
On Error GoTo ErrorHandler

    ' Declare locals
    Dim sSource   As String
    Dim sSql      As String
    Dim ofwResult As FWRecordList
    
    ' Define sSource
    sSource = Me.Name & ".CheckCatalogueIsNotReferenced"

    ' Initiate
    CheckCatalogueIsReferenced = False

    ' Determine which level we are at and if the entry has any children either in
    ' the live tables or another editset
    Select Case t_eCatalogueLevel
    
        Case ctDivision:
        
            ' Create the sql that will determine if this division has children
            ' (series)
            sSql = " IF EXISTS (SELECT TOP 1 " & gk_sDBCOLClassID & _
                   " FROM " & gk_sDBTABLEClass & _
                   " WHERE " & gk_sDBCOLDivisionID & " = " & t_lCatalogueID & ")" & _
                   " SELECT 1 ELSE " & _
                   " IF EXISTS (SELECT TOP 1 " & gk_sDBCOLClassID & _
                   " FROM " & gk_sDBTABLEEditsetClass & _
                   " WHERE " & gk_sDBCOLDivisionID & " = " & t_lCatalogueID & ")" & _
                   " SELECT 1 ELSE SELECT 0"
        
        Case ctHeaders:
        
            ' Create the sql that will determine if this sub-series has children
            ' (sub-sub-series or piece)
            sSql = " IF EXISTS (SELECT TOP 1 " & gk_sDBCOLSubHeaderID & _
                   " FROM " & gk_sDBTABLESubHeader & _
                   " WHERE " & gk_sDBCOLHeaderID & " = " & t_lCatalogueID & ")" & _
                   " SELECT 1 ELSE " & _
                   " IF EXISTS (SELECT TOP 1 " & gk_sDBCOLSubHeaderID & _
                   " FROM " & gk_sDBTABLEEditsetSubheader & _
                   " WHERE " & gk_sDBCOLHeaderID & " = " & t_lCatalogueID & ")" & _
                   " SELECT 1 ELSE " & _
                   " IF EXISTS (SELECT TOP 1 " & gk_sDBCOLPieceID & _
                   " FROM " & gk_sDBTABLEPiece & _
                   " WHERE " & gk_sDBCOLHeaderID & " = " & t_lCatalogueID & ")" & _
                   " SELECT 1 ELSE " & _
                   " IF EXISTS (SELECT TOP 1 " & gk_sDBCOLPieceID & _
                   " FROM " & gk_sDBTABLEEditsetPiece & _
                   " WHERE " & gk_sDBCOLHeaderID & " = " & t_lCatalogueID & ")" & _
                   " SELECT 1 ELSE SELECT 0"

        Case ctSubHeaders:
        
            ' Create the sql that will determine if this sub-sub-series has
            ' children (piece)
            sSql = " IF EXISTS (SELECT TOP 1 " & gk_sDBCOLPieceID & _
                   " FROM " & gk_sDBTABLEPiece & _
                   " WHERE " & gk_sDBCOLSubHeaderID & " = " & t_lCatalogueID & ")" & _
                   " SELECT 1 ELSE " & _
                   " IF EXISTS (SELECT TOP 1 " & gk_sDBCOLPieceID & _
                   " FROM " & gk_sDBTABLEEditsetPiece & _
                   " WHERE " & gk_sDBCOLSubHeaderID & " = " & t_lCatalogueID & ")" & _
                   " SELECT 1 ELSE SELECT 0"
    
        Case ctItem:
        
            ' Create the sql that will determine if this item does not have a
            ' null reference
            ' d33 - altered sql to detect blank as well as null references
            sSql = " IF EXISTS (SELECT " & gk_sDBCOLItemID & _
                   " FROM " & gk_sDBTABLEItem & _
                   " WHERE " & gk_sDBCOLItemID & " = " & t_lCatalogueID & _
                   " AND NOT ( " & gk_sDBCOLItemRef & " IS NULL OR " & _
                   gk_sDBCOLItemRef & " = '')) " & _
                   " SELECT 1 ELSE " & _
                   " IF EXISTS (SELECT " & gk_sDBCOLItemID & _
                   " FROM " & gk_sDBTABLEEditsetItem & _
                   " WHERE " & gk_sDBCOLItemID & " = " & t_lCatalogueID & _
                   " AND NOT ( " & gk_sDBCOLItemRef & " IS NULL OR " & _
                   gk_sDBCOLItemRef & " = '')) " & _
                   " SELECT 1 ELSE SELECT 0"
        
        Case Else:
        
            ' All other levels must have publically visable references
            RaiseError sSource, EPROCATINTERNALERROR.PECatalogue, _
                gk_sErrorDeleteCatalogueIsWrongLevel & _
                " (Catalogue Ref - " & GenerateCatalogueReferenceFromID(t_lCatalogueID, t_eCatalogueLevel) & _
                ", Catalogue Level - " & GetCatalogueLevelDescription( _
                                    GetCatalogueLevelCode(t_eCatalogueLevel)) & ")<br>"
            Exit Function
    
    End Select
    
    ' Determine if the entry has children
    If Len(sSql) > 0 Then
        Set ofwResult = m_oFWDBHelper.FindBySql(sSql)
        If Not ofwResult Is Nothing Then
            If ofwResult.RecordCount > 0 Then
                If ofwResult.GetData(0) = 1 Then
                    CheckCatalogueIsReferenced = True
                Else
                    CheckCatalogueIsReferenced = False
                End If
            Else
            End If
        End If
    End If
        
    Exit Function
    
ErrorHandler:
    ConvertError sSource
End Function
  
Public Function CheckCatalogueParentMarkedForDeletion( _
                 t_lCatalogueID As Long, _
                 t_eCatalogueLevel As ECATALOGUELEVEL) As Boolean
'Description:
'  Checks that the catalogue entry does not have a parent that is in a
'  delete editset
'
' Amendment history:
'   d32  alm      24Oct2001     Created
'
On Error GoTo ErrorHandler

    ' Declare locals
    Dim sSource    As String
    Dim sSql       As String
    Dim sTable1    As String
    Dim sTable2    As String
    Dim ofwResult  As FWRecordList
    
    ' Define sSource
    sSource = Me.Name & ".CheckCatalogueParentMarkedForDeletion"

    ' Initiate
    CheckCatalogueParentMarkedForDeletion = False

    ' Determine which level we are at and if the entry is marked for deletion
    Select Case t_eCatalogueLevel
    
        Case ctDivision:
        
            ' Create the sql that will determine if this parent is marked for deletion
            sTable1 = " " & gk_sDBTABLEEditsetDivision & "."
            sTable2 = " " & gk_sDBTABLEEditset & "."
            sSql = " IF EXISTS (SELECT TOP 1 " & gk_sDBCOLDivisionID & _
                   " FROM " & gk_sDBTABLEEditsetDivision & _
                   " JOIN " & gk_sDBTABLEEditset & " ON " & _
                   sTable2 & gk_sDBCOLEditSetID & " = " & sTable1 & gk_sDBCOLEditSetID & _
                   " WHERE " & sTable1 & gk_sDBCOLDivisionID & " = " & t_lCatalogueID & _
                   " AND " & sTable2 & gk_sDBCOLEditSetType & " = 4)" & _
                   " SELECT 1 ELSE SELECT 0"
        
        Case ctHeaders:
        
            ' Create the sql that will determine if this parent is marked for deletion
            sTable1 = " " & gk_sDBTABLEEditsetHeader & "."
            sTable2 = " " & gk_sDBTABLEEditset & "."
            sSql = " IF EXISTS (SELECT TOP 1 " & gk_sDBCOLHeaderID & _
                   " FROM " & gk_sDBTABLEEditsetHeader & _
                   " JOIN " & gk_sDBTABLEEditset & " ON " & _
                   sTable2 & gk_sDBCOLEditSetID & " = " & sTable1 & gk_sDBCOLEditSetID & _
                   " WHERE " & sTable1 & gk_sDBCOLHeaderID & " = " & t_lCatalogueID & _
                   " AND " & sTable2 & gk_sDBCOLEditSetType & " = 4)" & _
                   " SELECT 1 ELSE SELECT 0"

        Case ctSubHeaders:
        
            ' Create the sql that will determine if this parent is marked for deletion
            sTable1 = " " & gk_sDBTABLEEditsetSubheader & "."
            sTable2 = " " & gk_sDBTABLEEditset & "."
            sSql = " IF EXISTS (SELECT TOP 1 " & gk_sDBCOLSubHeaderID & _
                   " FROM " & gk_sDBTABLEEditsetSubheader & _
                   " JOIN " & gk_sDBTABLEEditset & " ON " & _
                   sTable2 & gk_sDBCOLEditSetID & " = " & sTable1 & gk_sDBCOLEditSetID & _
                   " WHERE " & sTable1 & gk_sDBCOLSubHeaderID & " = " & t_lCatalogueID & _
                   " AND " & sTable2 & gk_sDBCOLEditSetType & " = 4)" & _
                   " SELECT 1 ELSE SELECT 0"
    
    End Select
    
    ' Determine if the entry has children
    If Len(sSql) > 0 Then
        Set ofwResult = m_oFWDBHelper.FindBySql(sSql)
        If Not ofwResult Is Nothing Then
            If ofwResult.RecordCount > 0 Then
                If ofwResult.GetData(0) = 1 Then
                    CheckCatalogueParentMarkedForDeletion = True
                Else
                    CheckCatalogueParentMarkedForDeletion = False
                End If
            Else
            End If
        End If
    End If
        
    Exit Function
    
ErrorHandler:
    ConvertError sSource
End Function
  
Private Function SqlForExtensionTableForCreateByCopy( _
                    ByVal t_eTargetCatalogueLevelID As ECATALOGUELEVEL, _
                    ByVal t_eSourceCatalogueLevelID As ECATALOGUELEVEL _
                    ) As String
     
' Description: This methods returns the columns list specific for that level
'              which are copied from the copied catalogue to the new catalogue, the user will
'              not be seeing or modifying these (unlike dates, languages etc) columns
'
'              Only those columns are returned only if the properties are common for both the source and target catalogues
'              This  feature is helpful when creating an item from piece
'
'              These columns are present in the extension table
'
' Amendment history:
'   d11  ndk      03Nov1999      Created
'   d13  ndk      10Nov1999      Not copying closure type,closure code,Closure status here
'   d15  ndk      16Dec1999      Access condition was being returned twice in the sql when creating by copying, now corrected
'
    On Error GoTo ErrorHandler
     
    Dim sSource As String, sSql As String
    
    sSource = Me.Name & ".SqlForSpecificColoumnsInExtensionTable"

    sSql = ""
    
    'accumulation dates
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eTargetCatalogueLevelID, cpAccumulationDates) And _
        m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eSourceCatalogueLevelID, cpAccumulationDates) Then
        
        sSql = IIf(sSql = "", gk_sDBCOLAccumulationdates, sSql & "," & gk_sDBCOLAccumulationdates)
    End If
    
    'Former reference department
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eTargetCatalogueLevelID, cpFormerDepartmentReference) And _
        m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eSourceCatalogueLevelID, cpFormerDepartmentReference) Then
        
        sSql = IIf(sSql = "", gk_sDBCOLFormerDepartmentReference, sSql & "," & gk_sDBCOLFormerDepartmentReference)
    End If
    
    'Former reference PRO
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eTargetCatalogueLevelID, cpFormerPROReference) And _
        m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eSourceCatalogueLevelID, cpFormerPROReference) Then
        
        sSql = IIf(sSql = "", gk_sDBCOLFormerPROReference, sSql & "," & gk_sDBCOLFormerPROReference)
    End If
    
    'Map designation
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eTargetCatalogueLevelID, cpMapDesignation) And _
        m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eSourceCatalogueLevelID, cpMapDesignation) Then
        
        sSql = IIf(sSql = "", gk_sDBCOLMapDesignation, sSql & "," & gk_sDBCOLMapDesignation)
    End If
    
    'Map scale number
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eTargetCatalogueLevelID, cpMapscaleNumber) And _
        m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eSourceCatalogueLevelID, cpMapscaleNumber) Then
        
        sSql = IIf(sSql = "", gk_sDBCOLMapScaleNumber, sSql & "," & gk_sDBCOLMapScaleNumber)
    End If
    
    'Physical condition
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eTargetCatalogueLevelID, cpPhysicalCondition) And _
        m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eSourceCatalogueLevelID, cpPhysicalCondition) Then
        
        sSql = IIf(sSql = "", gk_sDBCOLPhysicalCondition, sSql & "," & gk_sDBCOLPhysicalCondition)
    End If
    
    'Physical description - dimensions
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eTargetCatalogueLevelID, cpPhysicalDescription_Dimension) And _
        m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eSourceCatalogueLevelID, cpPhysicalDescription_Dimension) Then
        
        sSql = IIf(sSql = "", gk_sDBCOLPhysicalDescriptionDimension, sSql & "," & gk_sDBCOLPhysicalDescriptionDimension)
    End If
    
    'Physical description - record quantity
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eTargetCatalogueLevelID, cpPhysicalDescription_RecordQuantity) And _
        m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eSourceCatalogueLevelID, cpPhysicalDescription_RecordQuantity) Then
        
        sSql = IIf(sSql = "", gk_sDBCOLPhysicalDescriptionRecordQuantity, sSql & "," & gk_sDBCOLPhysicalDescriptionRecordQuantity)
    End If
    
    'Physical description - form code
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eTargetCatalogueLevelID, cpPhysicalDescription_Form) And _
        m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eSourceCatalogueLevelID, cpPhysicalDescription_Form) Then
        
        sSql = IIf(sSql = "", gk_sDBCOLPhysicalDescriptionFormID, sSql & "," & gk_sDBCOLPhysicalDescriptionFormID)
    End If
    
    'Defunct field (for pieces and items only - they have open_date instead)
    'record opening date
    'If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eTargetCatalogueLevelID, cpRecordOpeningDate) And _
    '    m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eSourceCatalogueLevelID, cpRecordOpeningDate) Then
        
    '    sSql = IIf(sSql = "", gk_sDBCOLRecordOpeningDate, sSql & "," & gk_sDBCOLRecordOpeningDate)
    'End If
     
    SqlForExtensionTableForCreateByCopy = sSql
    
    Exit Function
    
ErrorHandler:
    ConvertError sSource
End Function

Private Function SqlForMainTableForCreateByCopy( _
                    ByVal t_eCatalogueLevelID As ECATALOGUELEVEL) As String
     
' Description: This methods returns the columns list specific for that level
'              This method is called once for getting the columns in the source
'              table and second time to get the columns for the target table
'
'              These are not the properties of the table and in procat these can
'              not be seen or modified currently (unlike dates, languages,dimensions etc)
'
'              These columns are present in the main table
'              For all the levels other than piece and item
'              the data can be copied from to the same level
'
' Amendment history:
'   d13  ndk      10Nov1999      Created
    
    On Error GoTo ErrorHandler
     
    Dim sSource As String, sSql As String
    
    sSource = Me.Name & ".SqlForMainTableForCreateByCopy"

    sSql = ""
    
    'division - closure_status
    'class - closure status,batch_id,intro_note,subclass_no,p_deposit
    'header - batch_id,header_scope,
    'subheder - batch_id,subheader_scope,
    'piece - physcond_id,batch_id,open_date,piece_scope
    'item -physcond_id,batch_id,open_date,item_scope
    
    'lettercode is not copied
    Select Case t_eCatalogueLevelID
        'division
        Case ECATALOGUELEVEL.ctDivision
            sSql = gk_sDBCOLClosureStatus
    
        'class
        Case ECATALOGUELEVEL.ctClass
            sSql = gk_sDBCOLClosureStatus & "," & _
                   gk_sDBCOLBatchID & "," & _
                   gk_sDBCOLIntroNote & "," & _
                   gk_sDBCOLSubClassNumber & "," & _
                   gk_sDBCOLPDeposit
                   
        'header
        Case ECATALOGUELEVEL.ctHeaders
            sSql = gk_sDBCOLBatchID & "," & _
                    gk_sDBCOLHeaderScope
        
        'subheader
        Case ECATALOGUELEVEL.ctSubHeaders
            sSql = gk_sDBCOLBatchID & "," & _
                    gk_sDBCOLSubHeaderScope
            
        'piece
        Case ECATALOGUELEVEL.ctPiece
            sSql = gk_sDBCOLPhysicCondID & "," & _
                    gk_sDBCOLBatchID & "," & _
                    gk_sDBCOLOpenDate & "," & _
                    gk_sDBCOLPieceScope
            
        'item
        Case ECATALOGUELEVEL.ctItem
            sSql = gk_sDBCOLPhysicCondID & "," & _
                    gk_sDBCOLBatchID & "," & _
                    gk_sDBCOLOpenDate & "," & _
                    gk_sDBCOLItemScope
    End Select
    
    'access condition
    SqlForMainTableForCreateByCopy = sSql
    
    Exit Function
    
ErrorHandler:
    ConvertError sSource
End Function



Function GetClassID(t_lCatalogueID As Long, t_eCatalogueLevel As ECATALOGUELEVEL) As Long
' Description: This function returns the Class ID of a catalogue entry below class in the
'              hierarchy
'
' Amendment history:
'   d24  tmj      03Mar2000      Created
'   d25  tmj      10Mar2000      Fixed defect in SQL statement to select class for an item
    
    Dim sSource As String
    Dim sSql As String
    Dim oFWRecordlist As FWRecordList
    Dim vClassID As Variant
    
    On Error GoTo ErrorHandler
    
    sSource = Me.Name & ".GetClassID"

    
    sSql = ""
    
    Select Case t_eCatalogueLevel
        Case ECATALOGUELEVEL.ctHeaders
            sSql = "SELECT " & gk_sDBCOLClassID & " FROM " & gk_sDBTABLEHeader & _
                   " WHERE " & gk_sDBCOLHeaderID & " = " & t_lCatalogueID
        Case ECATALOGUELEVEL.ctSubHeaders
            sSql = "SELECT th." & gk_sDBCOLClassID & " FROM " & gk_sDBTABLEHeader & " AS th, " & _
                   gk_sDBTABLESubHeader & " AS ts" & _
                   " WHERE ts." & gk_sDBCOLHeaderID & " = th." & gk_sDBCOLHeaderID & " AND " & _
                   "ts." & gk_sDBCOLSubHeaderID & " = " & t_lCatalogueID
        Case ECATALOGUELEVEL.ctPiece
            sSql = "SELECT " & gk_sDBCOLClassID & " FROM " & gk_sDBTABLEPiece & _
                   " WHERE " & gk_sDBCOLPieceID & " = " & t_lCatalogueID
        Case ECATALOGUELEVEL.ctItem
            sSql = "SELECT tp." & gk_sDBCOLClassID & " FROM " & gk_sDBTABLEPiece & " AS tp, " & _
                   gk_sDBTABLEItem & " AS ti" & _
                   " WHERE tp." & gk_sDBCOLPieceID & " = ti." & gk_sDBCOLPieceID & " AND " & _
                   "ti." & gk_sDBCOLItemID & " = " & t_lCatalogueID
        Case Else
            RaiseError sSource, PEGeneralInternalError, "Invalid Catalogue Error"
    End Select
        
    Set oFWRecordlist = m_oFWDBHelper.FindBySql(sSql)
    vClassID = oFWRecordlist.GetFieldValue(gk_sDBCOLClassID)
    GetClassID = CLng(vClassID)
        
    Exit Function
    
ErrorHandler:
    ConvertError sSource
End Function

Function GetItemPieceID(t_lItemID As Long) As Long
' Description: This function returns the Piece ID of an item below the piece in the
'              hierarchy
'
' Amendment history:
'   d24  tmj      03Mar2000      Created
    
    Dim sSource As String
    Dim sSql As String
    Dim oFWRecordlist As FWRecordList
    Dim vPieceID As Variant
    
    On Error GoTo ErrorHandler
    
    sSource = Me.Name & ".GetItemPieceID"
    
    sSql = "SELECT " & gk_sDBCOLPieceID & " FROM " & gk_sDBTABLEItem & _
           " WHERE " & gk_sDBCOLItemID & " = " & t_lItemID
        
    Set oFWRecordlist = m_oFWDBHelper.FindBySql(sSql)
    vPieceID = oFWRecordlist.GetFieldValue(gk_sDBCOLPieceID)
    GetItemPieceID = CLng(vPieceID)
        
    Exit Function
    
ErrorHandler:
    ConvertError sSource
End Function

Function CheckWhetherParentNotMoved(ByVal t_lCatalogueID As Long, _
                                        ByVal t_eCatalogueLevel As ECATALOGUELEVEL, ByVal t_lEditSetID) As Boolean
' Description: This function returns true if the parents of the object have not been moved
'              and false if they have. It exploits the fact that only pieces and classes can
'              be moved as their references will not change
'
' Amendment history:
'   d24  tmj      03Mar2000      Created
'   d28  tmj      24Mar2000      Passed EditsetID parameter

    Dim lClassID As Long
    Dim lPieceID As Long
    Dim sSource As String

    On Error GoTo ErrorHandler

    sSource = Me.Name & ".CheckWhetherParentNotMoved"
    
    
    If ((t_eCatalogueLevel = ctLetterCode) Or (t_eCatalogueLevel = ctDivision)) Then
        CheckWhetherParentNotMoved = True
        Exit Function
    ElseIf t_eCatalogueLevel = ctClass Then
        lClassID = t_lCatalogueID
    Else
        lClassID = GetClassID(t_lCatalogueID, t_eCatalogueLevel)
    End If

    CheckWhetherParentNotMoved = CheckClassNotMoved(lClassID, t_lEditSetID)

    If t_eCatalogueLevel = ctItem Then
        lPieceID = GetItemPieceID(t_lCatalogueID)
        CheckWhetherParentNotMoved = CheckPieceNotMoved(lPieceID, t_lEditSetID)
        Exit Function
    End If

    Exit Function

ErrorHandler:
    ConvertError sSource
End Function




Function CheckNotMoved(ByVal t_lCatalogueID As Long, _
                        ByVal t_eCatalogueLevel As ECATALOGUELEVEL, ByVal t_lEditSetID As Long) As Boolean
' Description: This function returns true if the object or it's parents have not been moved
'              and false if they have. It exploits the fact that only pieces and classes can
'              be moved as their references will not change
'
' Amendment history:
'   d24  tmj      03Mar2000      Created
'   d26  tmj      21Mar2000      Fixed so deals with case parents properly

    Dim bRes As Boolean
    Dim lClassID As Long
    Dim lPieceID As Long
    Dim sSql As String
    Dim oFWRecordlist As FWRecordList
    Dim sSource As String

    On Error GoTo ErrorHandler

    sSource = Me.Name & ".CheckNotMoved"

    bRes = True

    Select Case t_eCatalogueLevel
        Case ctLetterCode, ctDivision
            bRes = True
        Case ctClass
            bRes = CheckClassNotMoved(t_lCatalogueID, t_lEditSetID)
        Case ctHeaders, ctSubHeaders
            lClassID = GetClassID(t_lCatalogueID, t_eCatalogueLevel)
            bRes = CheckClassNotMoved(lClassID, t_lEditSetID)
        Case ctPiece
            lClassID = GetClassID(t_lCatalogueID, t_eCatalogueLevel)
            bRes = CheckClassNotMoved(lClassID, t_lEditSetID)
            If bRes Then
                bRes = CheckPieceNotMoved(t_lCatalogueID, t_lEditSetID)
            End If
        Case ctItem
            lClassID = GetClassID(t_lCatalogueID, t_eCatalogueLevel)
            bRes = CheckClassNotMoved(lClassID, t_lEditSetID)
            If bRes Then
                lPieceID = GetItemPieceID(t_lCatalogueID)
                bRes = CheckPieceNotMoved(lPieceID, t_lEditSetID)
            End If
    End Select

    CheckNotMoved = bRes

    Exit Function

ErrorHandler:
    ConvertError sSource
End Function

Private Function CheckClassNotMoved(t_lClassID As Long, ByVal t_lEditSetID As Long) As Boolean
' Description: This function returns true if the class specified by the id has not been moved
' and false otherwise. It returns true if the piece is not in an editset
'
' Amendment history:
'   d24  tmj      06Mar2000      Created
'   d28  tmj      24Mar2000      Enhanced to pass if adding to same editset

    Dim bRes As Boolean
    Dim sSql As String
    Dim oFWRecordlist As FWRecordList
    Dim sSource As String

    Const ksNoClass As String = "NoClass"

    On Error GoTo ErrorHandler

    bRes = True

    sSource = Me.Name & ".CheckClassNotMoved"
   
    If CheckIfCatalogueExistsAsEditable(t_lClassID, ctClass) Then
        sSql = "SELECT * FROM " & gk_sDBTABLEEditsetClass & " As tesc, " & _
               gk_sDBTABLEClass & " AS tc WHERE " & _
               "tc." & gk_sDBCOLDivisionID & " <> " & "tesc." & gk_sDBCOLDivisionID & " AND " & _
               "tc." & gk_sDBCOLClassID & " = " & "tesc." & gk_sDBCOLClassID & " AND " & _
               "tc." & gk_sDBCOLClassID & " = " & t_lClassID & " AND " & _
               "tesc." & gk_sDBCOLEditSetID & " <> " & t_lEditSetID
                   
        Set oFWRecordlist = m_oFWDBHelper.FindBySql(sSql)
        
        'If there are no recordsets there is a difference between the editset and live versions
        'so the parent has been moved
        If oFWRecordlist.RecordCount > 0 Then bRes = False
    End If
    
    CheckClassNotMoved = bRes

    Exit Function

ErrorHandler:
    ConvertError sSource
End Function

Private Function CheckPieceNotMoved(t_lPieceID As Long, ByVal t_lEditSetID As Long) As Boolean
' Description: This function returns true if the piece specified by the id has not been moved
' relative to it's parent class and false otherwise. It returns true if the piece is not in
' an editset
'
' Amendment history:
'   d24  tmj      06Mar2000      Created
'   d27  tmj      22Mar2000      Fixed
'   d28  tmj      24Mar2000      Enhanced to pass if adding to same editset

    Dim bRes As Boolean
    Dim lClassID As Long
    Dim sSql As String
    Dim oFWRecordlist As FWRecordList
    Dim sSource As String

    Const ksNoPiece As String = "NoPiece"

    On Error GoTo ErrorHandler

    bRes = True

    sSource = Me.Name & ".CheckClassNotMoved"
   
    If CheckIfCatalogueExistsAsEditable(t_lPieceID, ctPiece) Then
        sSql = "SELECT * FROM " & gk_sDBTABLEEditsetPiece & _
               " As tesp, " & gk_sDBTABLEPiece & " AS tp WHERE " & _
               "(((tp." & gk_sDBCOLHeaderID & " <> " & "tesp." & gk_sDBCOLHeaderID & ") OR " & _
                 "(tp." & gk_sDBCOLHeaderID & " IS NULL AND " & "tesp." & gk_sDBCOLHeaderID & " IS NOT NULL) OR " & _
                 "(tp." & gk_sDBCOLHeaderID & " IS NOT NULL AND " & "tesp." & gk_sDBCOLHeaderID & " IS NULL))" & _
               " OR " & _
               "((tp." & gk_sDBCOLSubHeaderID & " <> " & "tesp." & gk_sDBCOLSubHeaderID & ") OR " & _
                "(tp." & gk_sDBCOLSubHeaderID & " IS NULL AND " & "tesp." & gk_sDBCOLSubHeaderID & " IS NOT NULL) OR " & _
                "(tp." & gk_sDBCOLSubHeaderID & " IS NOT NULL AND " & "tesp." & gk_sDBCOLSubHeaderID & " IS NULL)))" & _
               " AND " & "tp." & gk_sDBCOLPieceID & " = " & "tesp." & gk_sDBCOLPieceID & _
               " AND " & "tp." & gk_sDBCOLPieceID & " = " & t_lPieceID & _
               " AND " & "tesp." & gk_sDBCOLEditSetID & " <> " & t_lEditSetID

                   
        Set oFWRecordlist = m_oFWDBHelper.FindBySql(sSql)
        
        'If there are no recordsets there is a difference between the editset and live versions
        'so the parent has been moved
        If oFWRecordlist.RecordCount > 0 Then bRes = False
    End If
    
    CheckPieceNotMoved = bRes

    Exit Function

ErrorHandler:
    ConvertError sSource
End Function


