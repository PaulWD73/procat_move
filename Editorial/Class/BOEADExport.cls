VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BOEADExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    BOEADExport.cls
' System:       PRO, PROCAT
' Copyright:    (C) Quidnunc Limited
'
' Description:  This class contains the methods and properties needed to
'    import data from the database
'
'    This business object is used to create a parsed catalogue entry based the DB ID passed
'
'
' Amendment history:
'   Vers    Auth            Date             Comment
'   d1      A.Farrow        25-Jul-2000      Created from BOEADImport
'   d2      D Johnston      03-Aug-2000      Added IIFs in CreateParsedCollection to cope with Null descriptions on link terms
'   d3      A.Farrow        04-Aug-2000      Added code to extract leaflet and sfa data to the
'                                            Publication Note(published) parsed cat entry string.
'                                            and the Unpublihed finding aid string.
'   d4      A.Farrow        09-Aug-2000      Changed ExtractFromIndexTerms to take into account level
'   d5      A.Farrow        11-Aug-2000      Changed main table extraction to get Language id for pieces/items
'                                            Altered slightly the code to get the SFA and Leaflet links.
'   d6      D Johnston      14-Aug-2000      Tidied up for release
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' ObjectContext object and ASP objects can be accessed (if necessary)
Private m_objContext As ObjectContext


'The current state of catalogues (ID) are stored in the local variables

Private m_lLetterCodeID As Long
Private m_lDivisionID As Long
Private m_lClassID As Long
Private m_lHeaderID As Long
Private m_lSubHeaderID As Long
Private m_lPieceID As Long
Private m_lItemID As Long
Private m_lBatchID As Long
'Values that are generated prior to the Extraction of piece
Private m_lPiecekeyOrder As Long
Private m_lItemkeyOrder As Long

Private m_oboParsedCatalogueEntry As ParsedCatalogueEntry
Private m_oFWDBHelper As FWDBHelper


Private Const m_sMaxOrderKeyValueName As String = "MaxOrderKey"


'
' Properties
'


''''''''''''''''' Implementation ''''''''''''''''''''''

'
' Private functions
'


'
' Class initialisation
'

Private Sub Initialise()
    
    m_lLetterCodeID = 0
    m_lDivisionID = 0
    m_lClassID = 0
    m_lHeaderID = 0
    m_lSubHeaderID = 0
    m_lPieceID = 0
    m_lItemID = 0
    m_lBatchID = 0
    m_lPiecekeyOrder = 0
    m_lItemkeyOrder = 0
    
    Set m_objContext = GetObjectContext()
    
    Set m_oboParsedCatalogueEntry = New ParsedCatalogueEntry
    Set m_oFWDBHelper = New FWDBHelper
    
End Sub

Private Sub Terminate()
    Set m_objContext = Nothing
    Set m_oboParsedCatalogueEntry = Nothing
    Set m_oFWDBHelper = Nothing
End Sub

Private Sub Class_Initialize()
    Initialise
End Sub

Private Sub Class_Terminate()
    Terminate
End Sub

'Properties
Public Property Set ParsedCatalogueEntry(t_objParsedCatalogueEntry As ParsedCatalogueEntry)
    Set m_oboParsedCatalogueEntry = t_objParsedCatalogueEntry
End Property

Public Property Get ParsedCatalogueEntry() As ParsedCatalogueEntry
    Set ParsedCatalogueEntry = m_oboParsedCatalogueEntry
End Property

Public Property Let BatchID(t_lBatchID As Long)
    m_lBatchID = t_lBatchID
End Property

Public Property Get BatchID() As Long
    BatchID = m_lBatchID
End Property

'Methods

Public Function Retrieve(t_iCatalogueLevel As Integer, t_lID As Long, t_lBatchID As Long) As Boolean
' Method to retrieve an entire catalogue entry and it's details into a ParsedCatalogueEntry
'
'
'   d1      A.Farrow        25-Jul-2000      Created from BOEADImport

On Error GoTo ErrHandler

    Dim sSQL As String
    Dim eCurrentLevel As ECATALOGUELEVEL
    Dim lID As Long
    Dim colTemp As Collection
    Dim sSource As String
    Dim ofwResult As FWRecordList
    Dim sSetParsedObj As String
    Dim sBatchRef As String
    
    ' Link for the related/separated material
    Dim rsmLink As Variant

    sSource = "BOEADExport.Retrieve"
    
    ' Set default value
    Retrieve = False
   
    
    'Initailise the parsed catalogue entry
    m_oboParsedCatalogueEntry.Initialise
    
    ' Get level number enum
    eCurrentLevel = GetCatalogueLevelEnum(t_iCatalogueLevel)
    
    'Get the catalogue reference to display
    sBatchRef = GetRefFromDBID(eCurrentLevel, t_lID)
    sBatchRef = sBatchRef & " ( " & GenerateExportReference(GetEntireRefFromDBID(eCurrentLevel, t_lID)) & " )"
    
    'Update the member varaiables with the ID value
    UpdateLocalVariable eCurrentLevel, t_lID
    
    'Extract from the Main table
    MainTableExtraction eCurrentLevel
    LogParseMessage "", "Extracted from the main table for " & sBatchRef, pInfo, t_lBatchID
    
    'Extract from the Extenstion table
    ExtensionTableExtraction eCurrentLevel
    LogParseMessage "", "Extracted from the extension table for " & sBatchRef, pInfo, t_lBatchID
    
    'Get SFA links for unpublished links and published links
    GetSFALinks eCurrentLevel, ECATALOGUEPROPERTIES.cpUnpublishedFindingAids
    GetSFALinks eCurrentLevel, ECATALOGUEPROPERTIES.cpPublicationNote
    LogParseMessage "", "Extracted from the SFA links for " & sBatchRef, pInfo, t_lBatchID
    
    'Get the leaflets in a similar fashion
    GetLeafletLinks eCurrentLevel, ECATALOGUEPROPERTIES.cpUnpublishedFindingAids
    GetLeafletLinks eCurrentLevel, ECATALOGUEPROPERTIES.cpPublicationNote
    LogParseMessage "", "Extracted from the leaflet links for " & sBatchRef, pInfo, t_lBatchID
    
    'LogMessage "Inserted into the extenstion table", EProfile.pComplete
    
    'Extract from the accruals table
    sSQL = SingleFieldTableExtractionQuery(eCurrentLevel, ECATALOGUETABLETYPE.cttAccrualsTable)
    
    Set ofwResult = m_oFWDBHelper.FindBySql(sSQL)
    
    sSetParsedObj = SetParsedEntry(ofwResult, ECATALOGUETABLETYPE.cttAccrualsTable)
    
    LogMessage "Extracted from the accruals table", EProfile.pComplete
    LogParseMessage "", "Extracted from the accruals table for " & sBatchRef, pInfo, t_lBatchID
    
    'Extract from the Admin history table
    sSQL = SingleFieldTableExtractionQuery(eCurrentLevel, ECATALOGUETABLETYPE.cttAdminHistoryTable)
    
    Set ofwResult = m_oFWDBHelper.FindBySql(sSQL)
    sSetParsedObj = SetParsedEntry(ofwResult, ECATALOGUETABLETYPE.cttAdminHistoryTable)
    
    LogMessage "Exported from the admin history table", EProfile.pComplete
    LogParseMessage "", "Extracted from the admin info table for " & sBatchRef, pInfo, t_lBatchID
    
    'Extract from the appraisal info table
    sSQL = SingleFieldTableExtractionQuery(eCurrentLevel, ECATALOGUETABLETYPE.cttAppDestructionTable)
    
    Set ofwResult = m_oFWDBHelper.FindBySql(sSQL)
    sSetParsedObj = SetParsedEntry(ofwResult, ECATALOGUETABLETYPE.cttAppDestructionTable)
     
    LogMessage "Exported from the appraisal info table", EProfile.pComplete
    LogParseMessage "", "Extracted from the appraisal table for " & sBatchRef, pInfo, t_lBatchID
    'Extract from the  arrangement table
    sSQL = SingleFieldTableExtractionQuery(eCurrentLevel, ECATALOGUETABLETYPE.cttArrangementTable)
    
    Set ofwResult = m_oFWDBHelper.FindBySql(sSQL)
    sSetParsedObj = SetParsedEntry(ofwResult, ECATALOGUETABLETYPE.cttArrangementTable)

    
    LogMessage "Exported from the arrangement table", EProfile.pComplete
    LogParseMessage "", "Extracted from the arrangement table for " & sBatchRef, pInfo, t_lBatchID
    
    'Extract from the custodial history table
    sSQL = SingleFieldTableExtractionQuery(eCurrentLevel, ECATALOGUETABLETYPE.cttCustodialHistoryTable)
    
    Set ofwResult = m_oFWDBHelper.FindBySql(sSQL)
    sSetParsedObj = SetParsedEntry(ofwResult, ECATALOGUETABLETYPE.cttCustodialHistoryTable)

    LogMessage "Exported from the custodial history table", EProfile.pComplete
    LogParseMessage "", "Extracted from the custodial history table for " & sBatchRef, pInfo, t_lBatchID
    
    'Extract from the note table
    sSQL = SingleFieldTableExtractionQuery(eCurrentLevel, ECATALOGUETABLETYPE.cttNoteTable)
    
    Set ofwResult = m_oFWDBHelper.FindBySql(sSQL)
    sSetParsedObj = SetParsedEntry(ofwResult, ECATALOGUETABLETYPE.cttNoteTable)

    
    LogMessage "Extracted from the note table", EProfile.pComplete
    
    'Extract from the scope and content table
    sSQL = SingleFieldTableExtractionQuery(eCurrentLevel, ECATALOGUETABLETYPE.cttScopeAndContentTable)
    
    Set ofwResult = m_oFWDBHelper.FindBySql(sSQL)
    sSetParsedObj = SetParsedEntry(ofwResult, ECATALOGUETABLETYPE.cttScopeAndContentTable)
    
    LogMessage "Extracted from the scope and content table", EProfile.pComplete
    LogParseMessage "", "Extracted from the custodial history table for " & sBatchRef, pInfo, t_lBatchID
    
    
    'Extract from the title table
    sSQL = SingleFieldTableExtractionQuery(eCurrentLevel, ECATALOGUETABLETYPE.cttTitleTable)
    
    Set ofwResult = m_oFWDBHelper.FindBySql(sSQL)
    sSetParsedObj = SetParsedEntry(ofwResult, ECATALOGUETABLETYPE.cttTitleTable)

    LogMessage "Extracted from the title table", EProfile.pComplete
    LogParseMessage "", "Extracted from the title table for " & sBatchRef, pInfo, t_lBatchID
    
    'Extraction of data for Publication note field
    sSQL = ExtractFromLinkDataElement(eCurrentLevel, ECATALOGUEPROPERTIES.cpPublicationNote)
    
    Set ofwResult = m_oFWDBHelper.FindBySql(sSQL)
    sSetParsedObj = SetParsedLinkEntry(ofwResult, ECATALOGUEPROPERTIES.cpPublicationNote)
    
    LogMessage "Extracted the Link data elements table for publication note ", EProfile.pComplete
    LogParseMessage "", "Extracted the Link data elements table for publication note, for " & sBatchRef, pInfo, t_lBatchID
    
    'Extraction of data for unpublished finding aids field
    sSQL = ExtractFromLinkDataElement(eCurrentLevel, ECATALOGUEPROPERTIES.cpUnpublishedFindingAids)
    
    Set ofwResult = m_oFWDBHelper.FindBySql(sSQL)
    sSetParsedObj = SetParsedLinkEntry(ofwResult, ECATALOGUEPROPERTIES.cpUnpublishedFindingAids)
    
    LogMessage "Extracted the Link data elements table for unpublished finding aids ", EProfile.pComplete
    LogParseMessage "", "Extracted the Link data elements table for unpublished finding aids for " & sBatchRef, pInfo, t_lBatchID
    
    
    'Extracting the data for related materials field
    LogMessage "Extracting the link data elements table for related material hanging text", EProfile.pComplete
    
    sSQL = ExtractFromLinkDataElement(eCurrentLevel, ECATALOGUEPROPERTIES.cpRelatedMaterials)
    
    Set ofwResult = m_oFWDBHelper.FindBySql(sSQL)
    sSetParsedObj = SetParsedLinkEntry(ofwResult, ECATALOGUEPROPERTIES.cpRelatedMaterials)

    '************************** STILL TO GET THE ACTUAL LINKS BACK ************************
    ' d11 - Add the actual links
    'LogMessage "Inserting into the link table for related material links", EProfile.pComplete
    'Dim colRelatedLinks As Collection
    'Set colRelatedLinks = m_oboParsedCatalogueEntry.RelatedMaterialLinks
    'For Each rsmLink In colRelatedLinks
       sSQL = ExtractFromLink(eCurrentLevel, ECATALOGUEPROPERTIES.cpRelatedMaterials)
        Set ofwResult = m_oFWDBHelper.FindBySql(sSQL)
        sSetParsedObj = CreateParsedCollection(ofwResult, ECATALOGUEPROPERTIES.cpRelatedMaterials, ECATALOGUETABLETYPE.cttLink)
    'Next
    
    
    'Extraction of data for seperate materials field
    LogMessage "Inserting into the link data elements table for separate material hanging text", EProfile.pComplete
    sSQL = ExtractFromLinkDataElement(eCurrentLevel, ECATALOGUEPROPERTIES.cpSeperatedMaterials)
    LogParseMessage "", "Extracting the link data elements table for separate material hanging text, for " & sBatchRef, pInfo, t_lBatchID
    
    Set ofwResult = m_oFWDBHelper.FindBySql(sSQL)
    sSetParsedObj = SetParsedLinkEntry(ofwResult, ECATALOGUEPROPERTIES.cpSeperatedMaterials)
    
    
    
    '************************* GET THE ACTUAL LINKS BACK ***********************
  
        sSQL = ExtractFromLink(eCurrentLevel, ECATALOGUEPROPERTIES.cpSeperatedMaterials)
        Set ofwResult = m_oFWDBHelper.FindBySql(sSQL)
        sSetParsedObj = CreateParsedCollection(ofwResult, ECATALOGUEPROPERTIES.cpSeperatedMaterials, ECATALOGUETABLETYPE.cttLink)

    
    
    'Insersion of data for Copies infomation TEXT ONLY field
    sSQL = ExtractFromLinkDataElement(eCurrentLevel, ECATALOGUEPROPERTIES.cpCopiesInformation)
    
    Set ofwResult = m_oFWDBHelper.FindBySql(sSQL)
    sSetParsedObj = SetParsedLinkEntry(ofwResult, ECATALOGUEPROPERTIES.cpCopiesInformation)
    
    LogMessage "Extracted the Link data elements table for copies info ", EProfile.pComplete
    LogParseMessage "", "Extracted the Link data elements table for copies info, for " & sBatchRef, pInfo, t_lBatchID
    
    'Insersion of data for immediate source of accq TEXT ONLY field
    sSQL = ExtractFromLinkDataElement(eCurrentLevel, ECATALOGUEPROPERTIES.cpImmediateSourceOfAcquisition)
    
    Set ofwResult = m_oFWDBHelper.FindBySql(sSQL)
    sSetParsedObj = SetParsedLinkEntry(ofwResult, ECATALOGUEPROPERTIES.cpImmediateSourceOfAcquisition)
    
    LogMessage "Extracted from the Link data elements table for immediate source of accq ", EProfile.pComplete
    LogParseMessage "", "Extracted from the Link data elements table for immediate source of accq, for " & sBatchRef, pInfo, t_lBatchID
    
    'Copies Info Corporate body ID description Export
    SaveIDDescriptionInCollecation eCurrentLevel, _
                            ECATALOGUETABLETYPE.cttCorporateBody, _
                            ECATALOGUEPROPERTIES.cpCopiesInformation
    
    LogMessage "Extracting from the the pc corporate body table for copies info", EProfile.pComplete
    LogParseMessage "", "Extracting from the the pc corporate body table for copies info, for " & sBatchRef, pInfo, t_lBatchID
    
    'Immediate Source Of acquisition Corporate body ID description Export
    SaveIDDescriptionInCollecation eCurrentLevel, _
                            ECATALOGUETABLETYPE.cttCorporateBody, _
                            ECATALOGUEPROPERTIES.cpImmediateSourceOfAcquisition

    
    LogMessage "Extracting from the pc corporate body table for Immediate source of accq", EProfile.pComplete
    LogParseMessage "", "Extracting from the pc corporate body table for Immediate source of accq, for " & sBatchRef, pInfo, t_lBatchID
     
    'Immediate Source Of acquisition Person ID description Export
    SaveIDDescriptionInCollecation eCurrentLevel, _
                            ECATALOGUETABLETYPE.cttPerson, _
                            ECATALOGUEPROPERTIES.cpImmediateSourceOfAcquisition
    
    LogMessage "Extracting from the pc person table for Immediate source of accq", EProfile.pComplete
    LogParseMessage "", "Extracting from the pc person table for Immediate source of accq, for " & sBatchRef, pInfo, t_lBatchID
     
    'Location of Originals Corporate body ID description Export
    SaveIDDescriptionInCollecation eCurrentLevel, _
                            ECATALOGUETABLETYPE.cttCorporateBody, _
                            ECATALOGUEPROPERTIES.cpLocationOfOriginals
    
    LogMessage "Extracting from the pc corporate body table for Location of Originals", EProfile.pComplete
    LogParseMessage "", "Extracting from the pc corporate body table for Location of Originals, for " & sBatchRef, pInfo, t_lBatchID
     
    'Creator names Export
    SaveIDDescriptionInCollecation eCurrentLevel, _
                            ECATALOGUETABLETYPE.cttCorporateBody, _
                            ECATALOGUEPROPERTIES.cpCreatorNames
    
    LogMessage "Extracting from the pc corporate body table for creator names", EProfile.pComplete
    LogParseMessage "", "Extracting from the pc corporate body table for creator names, for " & sBatchRef, pInfo, t_lBatchID
     
    'Creator names Export
    SaveIDDescriptionInCollecation eCurrentLevel, _
                            ECATALOGUETABLETYPE.cttPerson, _
                            ECATALOGUEPROPERTIES.cpCreatorNames
    
    ' d13 - Place of deposit links
    LogMessage "Extracting the from pc corporate body table for place of deposit", EProfile.pComplete
    LogParseMessage "", "Extracting the from pc corporate body table for place of deposit, for " & sBatchRef, pInfo, t_lBatchID
     
    SaveIDDescriptionInCollecation eCurrentLevel, _
                            ECATALOGUETABLETYPE.cttCorporateBody, _
                            ECATALOGUEPROPERTIES.cpPlaceOfDeposit
    
    LogMessage "Extracting into from the pc person table for creator names", EProfile.pComplete
    LogParseMessage "", "Extracting into from the pc person table for creator names, for " & sBatchRef, pInfo, t_lBatchID
    'Corporate Export
    SaveIDDescriptionInCollecation eCurrentLevel, _
                            ECATALOGUETABLETYPE.cttCorporateBody, _
                            ECATALOGUEPROPERTIES.cpIndexTermCorporateName
    
    LogMessage "Extracting from the pc corporate body table for Index term corporate body", EProfile.pComplete
     LogParseMessage "", "Extracting from the pc corporate body table for Index term corporate body, for   " & sBatchRef, pInfo, t_lBatchID
    'Person Export
    SaveIDDescriptionInCollecation eCurrentLevel, _
                            ECATALOGUETABLETYPE.cttPerson, _
                            ECATALOGUEPROPERTIES.cpIndexTermPersonName
    
    LogMessage "Extracting from the pc person table for Index term person", EProfile.pComplete
    LogParseMessage "", "Extracting from the pc person table for Index term person, for " & sBatchRef, pInfo, t_lBatchID
    'Place Export
    SaveIDDescriptionInCollecation eCurrentLevel, _
                            ECATALOGUETABLETYPE.cttPlace, _
                            ECATALOGUEPROPERTIES.cpIndexTermPlaceName
    
    LogMessage "Extracting from the pc place table for Index term Place", EProfile.pComplete
    LogParseMessage "", "Extracting from the pc place table for Index term Place, for " & sBatchRef, pInfo, t_lBatchID
    'Subject Export
    SaveIDDescriptionInCollecation eCurrentLevel, _
                            ECATALOGUETABLETYPE.cttSubject, _
                            ECATALOGUEPROPERTIES.cpIndexTermSubject
                            
    LogMessage "Extracting from the pc place table for Index term Subject", EProfile.pComplete
    LogParseMessage "", "Extracted the pc subject table for Index term Subject, for " & sBatchRef, pInfo, t_lBatchID
    ' Return success
    Retrieve = True
    
    Exit Function
    
ErrHandler:
    LogLoadInfo m_lBatchID, Err.Description, , , pFatalError
    LogMessage Err.Description, pSystemErrors
    ConvertError sSource
End Function

Private Sub SaveIDDescriptionInCollecation(t_eCatalogueLevel As ECATALOGUELEVEL, _
                            t_eTableType As ECATALOGUETABLETYPE, _
                            t_eCatalogueProperty As ECATALOGUEPROPERTIES)
'Method to get a back a record set and put into collection
'
'   Vers    Auth        Date        Comment
'   d12 DAJ     16Mar2000   Fixed SaveIDDescription to convert to string rather than a number
'   d15 DAJ     19Jun2000   Removed extra quoting around index term descriptions
On Error GoTo ErrHandle
Dim typPair As GenericLinkRecord
Dim sSQL As String
Dim sSource As String
Dim i As Integer
Dim ofwResult As New FWRecordList

    sSource = "BOEADExport.SaveIDDescriptionInCollecation"
                            
        
        sSQL = ExtractFromIndexTerms(t_eTableType, _
                        t_eCatalogueLevel, t_eCatalogueProperty)
        
        Set ofwResult = m_oFWDBHelper.FindBySql(sSQL)
        Dim d As String
        CreateParsedCollection ofwResult, t_eCatalogueProperty, t_eTableType
        
        'Set the object to empty to stop any problems with transactions.
        Set ofwResult = Nothing
    
    Exit Sub
ErrHandle:
    ConvertError sSource

End Sub




Private Sub SetKeyOrderMemberVariables(t_eLevel As ECATALOGUELEVEL)
'This method sets the value of the piece key order or item key order
'   member variables
'   d8  DAJ     08Mar2000   Fixed SetKeyOrderMemberVariables to get correct key order for item
'   d14 WDP     27Mar2000   Removed Embedded SQL using Max and replaced with sp call in SetKeyOrderMemberVariables
On Error GoTo ErrHandler

Dim sSQL As String
Dim oFWRecordList As FWRecordList
Dim sSource As String

    sSource = "BOEADImport.SetKeyOrderMemberVariables"
    
    Select Case t_eLevel
    Case ECATALOGUELEVEL.ctPiece
    
        sSQL = gk_sDBSPGetMaxPieceKeyOrder & _
            QuoteValueWithQuotes(m_oboParsedCatalogueEntry.LetterCode) & ", " & _
            m_oboParsedCatalogueEntry.ClassNumber
            
        Set oFWRecordList = m_oFWDBHelper.FindBySql(sSQL)
        
        If IsNull(oFWRecordList.GetFieldValue(m_sMaxOrderKeyValueName)) Then
            m_lPiecekeyOrder = gk_iIncrementForPieceKeyOrder
        Else
            m_lPiecekeyOrder = oFWRecordList.GetFieldValue(m_sMaxOrderKeyValueName) + _
                                gk_iIncrementForPieceKeyOrder
        End If
        
    Case ECATALOGUELEVEL.ctItem
        
        'Cannot use embedded SQL that contains aggregate function (max etc.)
        'as this causes errors in connection pooling. use an SP instead.
        
        sSQL = gk_sDBSPGetMaxItemKeyOrder & _
            QuoteValueWithQuotes(m_oboParsedCatalogueEntry.LetterCode) & ", " & _
            m_oboParsedCatalogueEntry.ClassNumber & ", " & _
            QuoteValueWithQuotes(m_oboParsedCatalogueEntry.PieceRef)
            
        Set oFWRecordList = m_oFWDBHelper.FindBySql(sSQL)
        
        If IsNull(oFWRecordList.GetFieldValue(m_sMaxOrderKeyValueName)) Then
            m_lItemkeyOrder = gk_iIncrementForItemKeyOrder
        Else
            m_lItemkeyOrder = oFWRecordList.GetFieldValue(m_sMaxOrderKeyValueName) + _
                                gk_iIncrementForItemKeyOrder
        End If
    End Select

    Exit Sub
ErrHandler:
    ConvertError sSource
End Sub


Private Sub UpdateLocalVariable(t_eLevel As ECATALOGUELEVEL, t_lID As Long)
'Updates the specified levels member variable to the ID value
    
    Select Case t_eLevel
    
    Case ECATALOGUELEVEL.ctLetterCode
            
        m_lLetterCodeID = t_lID
            
    Case ECATALOGUELEVEL.ctDivision
        
        m_lDivisionID = t_lID
        
    Case ECATALOGUELEVEL.ctClass
    
        m_lClassID = t_lID
        
    Case ECATALOGUELEVEL.ctHeaders
    
        m_lHeaderID = t_lID
        
    Case ECATALOGUELEVEL.ctSubheaders
        
        m_lSubHeaderID = t_lID
        
    Case ECATALOGUELEVEL.ctPiece
        
        m_lPieceID = t_lID
        
    Case ECATALOGUELEVEL.ctItem
        
        m_lItemID = t_lID
        
    End Select

End Sub


Private Sub ClearValuesBasedOnLevels()
'This method clears values based on the current level.
' The use of this is to make sure that the values available at
' any point is of the right
Dim i As Integer

    m_lPiecekeyOrder = 0
    m_lItemkeyOrder = 0
    For i = m_oboParsedCatalogueEntry.Level To gk_iNumberOfLevels
        
        Select Case GetCatalogueLevelEnum(i)
        Case ECATALOGUELEVEL.ctLetterCode
            
            m_lLetterCodeID = 0
            m_lDivisionID = 0
            m_lClassID = 0
            m_lHeaderID = 0
            m_lSubHeaderID = 0
            m_lPieceID = 0
            m_lItemID = 0
            
        Case ECATALOGUELEVEL.ctDivision
        
            m_lDivisionID = 0
            m_lClassID = 0
            m_lHeaderID = 0
            m_lSubHeaderID = 0
            m_lPieceID = 0
            m_lItemID = 0
            
        Case ECATALOGUELEVEL.ctClass
            
            m_lClassID = 0
            m_lHeaderID = 0
            m_lSubHeaderID = 0
            m_lPieceID = 0
            m_lItemID = 0
            
        Case ECATALOGUELEVEL.ctHeaders
            
            m_lHeaderID = 0
            m_lSubHeaderID = 0
            m_lPieceID = 0
            m_lItemID = 0
            
        Case ECATALOGUELEVEL.ctSubheaders
            
            m_lSubHeaderID = 0
            m_lPieceID = 0
            m_lItemID = 0
            
        Case ECATALOGUELEVEL.ctPiece
            
            m_lPieceID = 0
            m_lItemID = 0
            
            
        Case ECATALOGUELEVEL.ctItem
            
            m_lItemID = 0
            
        End Select
    
    Next i
    
    'If the Division number is 0 then the division is not present for a class
    If m_oboParsedCatalogueEntry.DivisionNumber = 0 Then
        m_lDivisionID = 0
    End If

    'If the header number is 0 then the subheader number will also be zero
    '  and the piece is a parent of class
    If m_oboParsedCatalogueEntry.HeaderNumber = 0 Then
        m_lHeaderID = 0
        m_lSubHeaderID = 0
    End If
    
    'If the subheader number is 0 then a subheader is not present for a class
    If m_oboParsedCatalogueEntry.SubHeaderNumber = 0 Then
        m_lSubHeaderID = 0
    End If

End Sub

Private Function CheckValues()
On Error GoTo ErrHandler

Dim sSource As String

    sSource = "BOEADImport.CheckValues"

    
    If m_oboParsedCatalogueEntry Is Nothing Then
        RaiseError sSource, EPROCATINTERNALERROR.PEImport, _
            "Need to set the property ParsedCatalogueEntry"
    End If
    
    If m_lBatchID <= 0 Then
        RaiseError sSource, EPROCATINTERNALERROR.PEImport, _
            "Batch ID property needs to be set"
    End If

    Exit Function
ErrHandler:
    ConvertError sSource
End Function

Private Function FindNonLiveParents(t_eLevel As ECATALOGUELEVEL) As String
' Returns the SQL required to find any non-live parents of the current entry
'
'   d14 DAJ     27Mar2000   Added checking to ensure that new entries do not have non-live parents

    Dim sSQL As String
    Dim sParent1MainTablename As String, sParent2MainTablename As String, sParent3MainTablename As String
    Dim sParent1ExtTablename As String, sParent2ExtTablename As String, sParent3ExtTablename As String
    Dim sParent1IDColumnName As String, sParent2IDColumnName As String, sParent3IDColumnName As String
    
    ' Lettercodes have no parents; just return
    If t_eLevel = ctLetterCode Then
        sSQL = ""
        Exit Function
    End If
    
    
    Select Case t_eLevel
    
    Case ECATALOGUELEVEL.ctDivision
            
            ' Lettercode parents
            ' Both batch ID and live_flag are in the extension table
            sParent1ExtTablename = GetTableName(ECATALOGUETABLETYPE.cttExtensionTable, _
                            ESectionType.SECReader, ctLetterCode)
                            
            sParent1IDColumnName = GetColumnName(ECATALOGUECOLUMNTYPE.cctIDColumn, ctLetterCode)
            
            sSQL = "SELECT " & sParent1IDColumnName & " FROM " & sParent1ExtTablename & " (nolock)" & _
                " WHERE " & sParent1IDColumnName & " = " & m_lLetterCodeID & _
                " AND " & gk_sDBCOLBatchID & " <> " & m_lBatchID & _
                " AND " & gk_sDBCOLLiveFlag & " = " & ELIVEFLAGSTATUS.lfsNotLive
                

    Case ECATALOGUELEVEL.ctClass
            ' Lettercode parents
            ' Both batch ID and live_flag are in the extension table
            sParent1ExtTablename = GetTableName(ECATALOGUETABLETYPE.cttExtensionTable, _
                            ESectionType.SECReader, ctLetterCode)
                            
            sParent1IDColumnName = GetColumnName(ECATALOGUECOLUMNTYPE.cctIDColumn, ctLetterCode)
            
            ' Division parents
            ' Batch ID -> main table, live_flag -> extension table
            sParent2MainTablename = GetTableName(ECATALOGUETABLETYPE.cttMainTable, _
                            ESectionType.SECReader, ctDivision)
            sParent2ExtTablename = GetTableName(ECATALOGUETABLETYPE.cttExtensionTable, _
                            ESectionType.SECReader, ctDivision)
                            
            sParent2IDColumnName = GetColumnName(ECATALOGUECOLUMNTYPE.cctIDColumn, ctDivision)
            
            ' SQL
            sSQL = "SELECT " & sParent1IDColumnName & " AS " & sParent1IDColumnName & "," & _
                " Null AS " & sParent2IDColumnName & " " & _
                " FROM " & sParent1ExtTablename & " (nolock)" & _
                " WHERE (" & sParent1IDColumnName & " = " & m_lLetterCodeID & _
                " AND " & gk_sDBCOLBatchID & " <> " & m_lBatchID & _
                " AND " & gk_sDBCOLLiveFlag & " = " & ELIVEFLAGSTATUS.lfsNotLive & ")"
            sSQL = sSQL & " UNION " & _
                " SELECT Null AS " & sParent1IDColumnName & ", " & _
                " DM." & sParent2IDColumnName & " AS " & sParent2IDColumnName & " " & _
                " FROM " & sParent2MainTablename & " DM (nolock), " & _
                sParent2ExtTablename & " DE (nolock)" & _
                " WHERE (DE." & sParent2IDColumnName & " = " & m_lDivisionID & _
                " AND DM." & sParent2IDColumnName & " = " & m_lDivisionID & _
                " AND DM." & gk_sDBCOLBatchID & " <> " & m_lBatchID & _
                " AND DE." & gk_sDBCOLLiveFlag & " = " & ELIVEFLAGSTATUS.lfsNotLive & ")"

    Case ECATALOGUELEVEL.ctHeaders
    
            ' Class parents
            ' Batch ID -> main table, live_flag -> extension table
            sParent1MainTablename = GetTableName(ECATALOGUETABLETYPE.cttMainTable, _
                            ESectionType.SECReader, ctClass)
            sParent1ExtTablename = GetTableName(ECATALOGUETABLETYPE.cttExtensionTable, _
                            ESectionType.SECReader, ctClass)
                            
            sParent1IDColumnName = GetColumnName(ECATALOGUECOLUMNTYPE.cctIDColumn, ctClass)
            
            sSQL = "SELECT CM." & sParent1IDColumnName & _
                " FROM " & sParent1MainTablename & " CM (nolock), " & _
                sParent1ExtTablename & " CE (nolock)" & _
                " WHERE CM." & sParent1IDColumnName & " = " & m_lClassID & _
                " AND CE." & sParent1IDColumnName & " = " & m_lClassID & _
                " AND CM." & gk_sDBCOLBatchID & " <> " & m_lBatchID & _
                " AND CE." & gk_sDBCOLLiveFlag & " = " & ELIVEFLAGSTATUS.lfsNotLive
    
    Case ECATALOGUELEVEL.ctSubheaders
    
            ' Class parents
            ' Batch ID -> main table, live_flag -> extension table
            sParent1MainTablename = GetTableName(ECATALOGUETABLETYPE.cttMainTable, _
                            ESectionType.SECReader, ctClass)
            sParent1ExtTablename = GetTableName(ECATALOGUETABLETYPE.cttExtensionTable, _
                            ESectionType.SECReader, ctClass)
                            
            sParent1IDColumnName = GetColumnName(ECATALOGUECOLUMNTYPE.cctIDColumn, ctClass)
            
            ' Header parents
            ' Batch ID -> main table, live_flag -> extension table
            sParent2MainTablename = GetTableName(ECATALOGUETABLETYPE.cttMainTable, _
                            ESectionType.SECReader, ctHeaders)
            sParent2ExtTablename = GetTableName(ECATALOGUETABLETYPE.cttExtensionTable, _
                            ESectionType.SECReader, ctHeaders)
                            
            sParent2IDColumnName = GetColumnName(ECATALOGUECOLUMNTYPE.cctIDColumn, ctHeaders)
            
            ' SQL
            sSQL = "SELECT CM." & sParent1IDColumnName & " AS " & sParent1IDColumnName & "," & _
                " Null AS " & sParent2IDColumnName & " " & _
                " FROM " & sParent1MainTablename & " CM (nolock), " & _
                sParent1ExtTablename & " CE (nolock)" & _
                " WHERE (CM." & sParent1IDColumnName & " = " & m_lClassID & _
                " AND CE." & sParent1IDColumnName & " = " & m_lClassID & _
                " AND CM." & gk_sDBCOLBatchID & " <> " & m_lBatchID & _
                " AND CE." & gk_sDBCOLLiveFlag & " = " & ELIVEFLAGSTATUS.lfsNotLive & ")"
            sSQL = sSQL & " UNION " & _
                " SELECT Null AS " & sParent1IDColumnName & ", " & _
                " HM." & sParent2IDColumnName & " AS " & sParent2IDColumnName & " " & _
                " FROM " & sParent2MainTablename & " HM (nolock), " & _
                sParent2ExtTablename & " HE (nolock)" & _
                " WHERE (HM." & sParent2IDColumnName & " = " & m_lHeaderID & _
                " AND HE." & sParent2IDColumnName & " = " & m_lHeaderID & _
                " AND HM." & gk_sDBCOLBatchID & " <> " & m_lBatchID & _
                " AND HE." & gk_sDBCOLLiveFlag & " = " & ELIVEFLAGSTATUS.lfsNotLive & ")"
    
    Case ECATALOGUELEVEL.ctPiece
    
            ' Class parents
            ' Batch ID -> main table, live_flag -> extension table
            sParent1MainTablename = GetTableName(ECATALOGUETABLETYPE.cttMainTable, _
                            ESectionType.SECReader, ctClass)
            sParent1ExtTablename = GetTableName(ECATALOGUETABLETYPE.cttExtensionTable, _
                            ESectionType.SECReader, ctClass)
                            
            sParent1IDColumnName = GetColumnName(ECATALOGUECOLUMNTYPE.cctIDColumn, ctClass)
            
            ' Header parents
            ' Batch ID -> main table, live_flag -> extension table
            sParent2MainTablename = GetTableName(ECATALOGUETABLETYPE.cttMainTable, _
                            ESectionType.SECReader, ctHeaders)
            sParent2ExtTablename = GetTableName(ECATALOGUETABLETYPE.cttExtensionTable, _
                            ESectionType.SECReader, ctHeaders)
                            
            sParent2IDColumnName = GetColumnName(ECATALOGUECOLUMNTYPE.cctIDColumn, ctHeaders)
            
            ' SubHeader parents
            ' Batch ID -> main table, live_flag -> extension table
            sParent3MainTablename = GetTableName(ECATALOGUETABLETYPE.cttMainTable, _
                            ESectionType.SECReader, ctSubheaders)
            sParent3ExtTablename = GetTableName(ECATALOGUETABLETYPE.cttExtensionTable, _
                            ESectionType.SECReader, ctSubheaders)
                            
            sParent3IDColumnName = GetColumnName(ECATALOGUECOLUMNTYPE.cctIDColumn, ctSubheaders)
            
            ' SQL
            sSQL = "SELECT CM." & sParent1IDColumnName & " AS " & sParent1IDColumnName & "," & _
                " Null AS " & sParent2IDColumnName & ", " & _
                " Null AS " & sParent3IDColumnName & " " & _
                " FROM " & sParent1MainTablename & " CM (nolock), " & _
                sParent1ExtTablename & " CE (nolock)" & _
                " WHERE (CM." & sParent1IDColumnName & " = " & m_lClassID & _
                " AND CE." & sParent1IDColumnName & " = " & m_lClassID & _
                " AND CM." & gk_sDBCOLBatchID & " <> " & m_lBatchID & _
                " AND CE." & gk_sDBCOLLiveFlag & " = " & ELIVEFLAGSTATUS.lfsNotLive & ")"
                
            sSQL = sSQL & " UNION " & _
                " SELECT Null AS " & sParent1IDColumnName & ", " & _
                " HM." & sParent2IDColumnName & " AS " & sParent2IDColumnName & ", " & _
                " Null AS " & sParent3IDColumnName & " " & _
                " FROM " & sParent2MainTablename & " HM (nolock), " & _
                sParent2ExtTablename & " HE (nolock)" & _
                " WHERE (HM." & sParent2IDColumnName & " = " & m_lHeaderID & _
                " AND HE." & sParent2IDColumnName & " = " & m_lHeaderID & _
                " AND HM." & gk_sDBCOLBatchID & " <> " & m_lBatchID & _
                " AND HE." & gk_sDBCOLLiveFlag & " = " & ELIVEFLAGSTATUS.lfsNotLive & ")"
                
            sSQL = sSQL & " UNION " & _
                " SELECT Null AS " & sParent1IDColumnName & ", " & _
                " Null AS " & sParent2IDColumnName & ", " & _
                " SHM." & sParent3IDColumnName & " AS " & sParent3IDColumnName & " " & _
                " FROM " & sParent3MainTablename & " SHM (nolock), " & _
                sParent3ExtTablename & " SHE (nolock)" & _
                " WHERE (SHM." & sParent3IDColumnName & " = " & m_lHeaderID & _
                " AND SHE." & sParent3IDColumnName & " = " & m_lHeaderID & _
                " AND SHM." & gk_sDBCOLBatchID & " <> " & m_lBatchID & _
                " AND SHE." & gk_sDBCOLLiveFlag & " = " & ELIVEFLAGSTATUS.lfsNotLive & ")"
    
    Case ECATALOGUELEVEL.ctItem
            ' Piece parents
            ' Batch ID -> main table, live_flag -> extension table
            sParent1MainTablename = GetTableName(ECATALOGUETABLETYPE.cttMainTable, _
                            ESectionType.SECReader, ctPiece)
            sParent1ExtTablename = GetTableName(ECATALOGUETABLETYPE.cttExtensionTable, _
                            ESectionType.SECReader, ctPiece)
                            
            sParent1IDColumnName = GetColumnName(ECATALOGUECOLUMNTYPE.cctIDColumn, ctPiece)
            
            sSQL = "SELECT PM." & sParent1IDColumnName & _
                " FROM " & sParent1MainTablename & " PM (nolock), " & _
                sParent1ExtTablename & " PE (nolock)" & _
                " WHERE PM." & sParent1IDColumnName & " = " & m_lPieceID & _
                " AND PE." & sParent1IDColumnName & " = " & m_lPieceID & _
                " AND PM." & gk_sDBCOLBatchID & " <> " & m_lBatchID & _
                " AND PE." & gk_sDBCOLLiveFlag & " = " & ELIVEFLAGSTATUS.lfsNotLive
    
    End Select

    FindNonLiveParents = sSQL
    
End Function


Private Function MainTableExtraction(t_eLevel As ECATALOGUELEVEL) As String
'Returns the SQL for extracting from the main table


    Dim sSQL As String
    Dim sTablename As String
    Dim sIDColumnName As String
    Dim ofwResult As New FWRecordList
    
    sTablename = GetTableName(ECATALOGUETABLETYPE.cttMainTable, _
                    ESectionType.SECReader, t_eLevel)
                    
    sIDColumnName = GetColumnName(ECATALOGUECOLUMNTYPE.cctIDColumn, t_eLevel)
    
    Select Case t_eLevel
    
    Case ECATALOGUELEVEL.ctLetterCode
            
            sSQL = "SELECT " & gk_sDBCOLLetterCodeCode & " FROM " & sTablename & " (nolock)"

            sSQL = sSQL & " WHERE " & sIDColumnName & " = " & GetCatalogueID(t_eLevel)
            
            Set ofwResult = m_oFWDBHelper.FindBySql(sSQL)
            
            ' If we have some results store it in the parsed catalogue entry
            If Not ofwResult Is Nothing Then
                If ofwResult.RecordCount > 0 Then
                    m_oboParsedCatalogueEntry.LetterCode = ofwResult.GetFieldValue(gk_sDBCOLLetterCodeCode)
                End If
            End If
                
            
    
    Case ECATALOGUELEVEL.ctDivision
    
            sSQL = "SELECT " & gk_sDBCOLDivisionNo & " FROM " & sTablename & " (nolock)"

            sSQL = sSQL & " WHERE " & sIDColumnName & " = " & GetCatalogueID(t_eLevel)
                       
            Set ofwResult = m_oFWDBHelper.FindBySql(sSQL)
                 
             ' If we have some results store it in the parsed catalogue entry
            If Not ofwResult Is Nothing Then
                If ofwResult.RecordCount > 0 Then
                    m_oboParsedCatalogueEntry.DivisionNumber = ofwResult.GetFieldValue(gk_sDBCOLDivisionNo)
                End If
            End If

    Case ECATALOGUELEVEL.ctClass
    
            sSQL = "SELECT " & gk_sDBCOLClassNumber & _
                "," & gk_sDBCOLDateText & _
                " FROM " & sTablename & " (nolock)"

            sSQL = sSQL & " WHERE " & sIDColumnName & " = " & GetCatalogueID(t_eLevel)
                       
            Set ofwResult = m_oFWDBHelper.FindBySql(sSQL)
             ' If we have some results store it in the parsed catalogue entry
            If Not ofwResult Is Nothing Then
                If ofwResult.RecordCount > 0 Then
                    m_oboParsedCatalogueEntry.ClassNumber = ofwResult.GetFieldValue(gk_sDBCOLClassNumber)
                    m_oboParsedCatalogueEntry.DateText = IIf(IsNull(ofwResult.GetFieldValue(gk_sDBCOLDateText)), "", ofwResult.GetFieldValue(gk_sDBCOLDateText))
                End If
            End If
    
    Case ECATALOGUELEVEL.ctHeaders
    
            sSQL = "SELECT " & gk_sDBCOLClassHeaderNo & _
                "," & gk_sDBCOLDateText & _
                " FROM " & sTablename & " (nolock)"

            sSQL = sSQL & " WHERE " & sIDColumnName & " = " & GetCatalogueID(t_eLevel)
                       
            Set ofwResult = m_oFWDBHelper.FindBySql(sSQL)
             ' If we have some results store it in the parsed catalogue entry
            If Not ofwResult Is Nothing Then
                If ofwResult.RecordCount > 0 Then
                    m_oboParsedCatalogueEntry.HeaderNumber = ofwResult.GetFieldValue(gk_sDBCOLClassHeaderNo)
                    m_oboParsedCatalogueEntry.DateText = IIf(IsNull(ofwResult.GetFieldValue(gk_sDBCOLDateText)), "", ofwResult.GetFieldValue(gk_sDBCOLDateText))
                End If
            End If
    
    Case ECATALOGUELEVEL.ctSubheaders
    
            sSQL = "SELECT " & gk_sDBCOLSubHeaderNo & _
                "," & gk_sDBCOLDateText & _
                " FROM " & sTablename & " (nolock)"

            sSQL = sSQL & " WHERE " & sIDColumnName & " = " & GetCatalogueID(t_eLevel)
                       
            Set ofwResult = m_oFWDBHelper.FindBySql(sSQL)
             ' If we have some results store it in the parsed catalogue entry
            If Not ofwResult Is Nothing Then
                If ofwResult.RecordCount > 0 Then
                    m_oboParsedCatalogueEntry.SubClassNumber = ofwResult.GetFieldValue(gk_sDBCOLSubHeaderNo)
                    m_oboParsedCatalogueEntry.DateText = IIf(IsNull(ofwResult.GetFieldValue(gk_sDBCOLDateText)), "", ofwResult.GetFieldValue(gk_sDBCOLDateText))
                End If
            End If
    
    Case ECATALOGUELEVEL.ctPiece
    
            sSQL = "SELECT " & gk_sDBCOLPieceRef & _
                "," & gk_sDBCOLDateText & _
                "," & gk_sDBCOLLanguageID & _
                " FROM " & sTablename & " (nolock)"

            sSQL = sSQL & " WHERE " & sIDColumnName & " = " & GetCatalogueID(t_eLevel)
                       
            Set ofwResult = m_oFWDBHelper.FindBySql(sSQL)
             ' If we have some results store it in the parsed catalogue entry
            If Not ofwResult Is Nothing Then
                If ofwResult.RecordCount > 0 Then
                    m_oboParsedCatalogueEntry.PieceRef = ofwResult.GetFieldValue(gk_sDBCOLPieceRef)
                    m_oboParsedCatalogueEntry.DateText = IIf(IsNull(ofwResult.GetFieldValue(gk_sDBCOLDateText)), "", ofwResult.GetFieldValue(gk_sDBCOLDateText))
                    m_oboParsedCatalogueEntry.LanguageID = IIf(IsNull(ofwResult.GetFieldValue(gk_sDBCOLLanguageID)), 0, ofwResult.GetFieldValue(gk_sDBCOLLanguageID))
                End If
            End If
    
    Case ECATALOGUELEVEL.ctItem
    'For Item key reference if an empty string is passed then convert to null
             sSQL = "SELECT " & gk_sDBCOLItemRef & _
                "," & gk_sDBCOLDateText & _
                "," & gk_sDBCOLLanguageID & _
                " FROM " & sTablename & " (nolock)"

            sSQL = sSQL & " WHERE " & sIDColumnName & " = " & GetCatalogueID(t_eLevel)
                       
            Set ofwResult = m_oFWDBHelper.FindBySql(sSQL)
             ' If we have some results store it in the parsed catalogue entry
            If Not ofwResult Is Nothing Then
                If ofwResult.RecordCount > 0 Then
                    m_oboParsedCatalogueEntry.ItemRef = _
                    IIf(IsNull(ofwResult.GetFieldValue(gk_sDBCOLItemRef)), "", ofwResult.GetFieldValue(gk_sDBCOLItemRef))
                    m_oboParsedCatalogueEntry.DateText = IIf(IsNull(ofwResult.GetFieldValue(gk_sDBCOLDateText)), "", ofwResult.GetFieldValue(gk_sDBCOLDateText))
                    m_oboParsedCatalogueEntry.LanguageID = IIf(IsNull(ofwResult.GetFieldValue(gk_sDBCOLLanguageID)), 0, ofwResult.GetFieldValue(gk_sDBCOLLanguageID))
                End If
            End If
    
    End Select


End Function



Private Function SingleFieldTableExtractionQuery(t_eLevel As ECATALOGUELEVEL, _
                                                t_eTableType As ECATALOGUETABLETYPE) As String
'Returns select string for scope and content table
'
    Dim sSQL As String
    Dim sTablename As String
    Dim sIDColumnName As String
    
    sTablename = GetTableName(t_eTableType, _
                    ESectionType.SECReader, t_eLevel)
                    
    sIDColumnName = GetColumnName(ECATALOGUECOLUMNTYPE.cctIDColumn, t_eLevel)

    sSQL = "SELECT " & GetSingleFieldTableFieldName(t_eTableType) & " FROM " & sTablename & " (nolock)"

    sSQL = sSQL & " WHERE " & _
                    sIDColumnName & " = " & GetCatalogueID(t_eLevel)

    SingleFieldTableExtractionQuery = sSQL

End Function


Private Function ExtractFromLinkDataElement(t_eLevel As ECATALOGUELEVEL, _
                            t_eProperty As ECATALOGUEPROPERTIES) As String
   
   Dim sSQL As String
    
    sSQL = "SELECT " & gk_sDBCOLDescription & _
                " FROM " & gk_sDBTABLEPCLinkDataElement & " (nolock)"

    sSQL = sSQL & " WHERE " & gk_sDBCOLCatalogueID & " = " & GetCatalogueID(t_eLevel) & _
        " AND " & gk_sDBCOLCatalogueLevel & " = " & t_eLevel & _
        " AND " & t_eProperty & " = " & gk_sDBCOLPropertyID

    ExtractFromLinkDataElement = sSQL

End Function

Private Function ExtractFromLink(t_eLevel As ECATALOGUELEVEL, _
                            t_eProperty As ECATALOGUEPROPERTIES) As String
'   d11 DAJ     15Mar2000   Added ExtractFromLink to add the Related/Separated material links into the database
'                           (copied from ExtractFromLinkDataElement)
    
    Dim sSQL As String
    
    ' Catalogue entry 1 is the source, 2 is the destination.  But ...
    ' for related materials the link is bi-directional; we must put in
    ' another link going back the way...
    
    ' For both types of link, insert the 'outbound' link
    sSQL = "SELECT " & gk_sDBCOLCatalogueID2 & "," & _
        gk_sDBCOLCatalogueLevel2 & "," & _
        gk_sDBCOLDescription & " FROM " & gk_sDBTABLEPCLink & " (nolock)"

    sSQL = sSQL & " WHERE " & _
        gk_sDBCOLCatalogueLevel1 & " = " & _
        t_eLevel & " AND " & _
        gk_sDBCOLCatalogueID1 & " = " & _
        GetCatalogueID(t_eLevel) & " AND " & _
        gk_sDBCOLPropertyID & " = " & _
        t_eProperty
    
    ExtractFromLink = sSQL
    
End Function

Private Function ExtractFromIndexTerms(t_eTableType As ECATALOGUETABLETYPE, _
                            t_eLevel As ECATALOGUELEVEL, _
                            t_eProperty As ECATALOGUEPROPERTIES) As String
   
    Dim sTablename As String
    Dim sSQL  As String
    
    sTablename = GetTableName(t_eTableType, _
                    ESectionType.SECReader, t_eLevel)
                    
    sSQL = "SELECT " & GetReferenceColumnNameInLinkTable(t_eTableType) & "," _
                & gk_sDBCOLDescription & _
                " FROM " & sTablename & " (nolock)"
            

    sSQL = sSQL & " Where " & gk_sDBCOLCatalogueID & " = " & GetCatalogueID(t_eLevel) & _
                        " AND " & gk_sDBCOLPropertyID & " = " & t_eProperty & _
                        " AND " & gk_sDBCOLCatalogueLevel & " = " & t_eLevel

    ExtractFromIndexTerms = sSQL

End Function



Private Function GetSingleFieldTableFieldName(t_eTableType As ECATALOGUETABLETYPE) As String
'Returns the column name for the single column table names
    
    Select Case t_eTableType
    Case ECATALOGUETABLETYPE.cttAccrualsTable
        
        GetSingleFieldTableFieldName = gk_sDBCOLAccruals
    
    Case ECATALOGUETABLETYPE.cttAdminHistoryTable
        
        GetSingleFieldTableFieldName = gk_sDBCOLAdministrativeHistory
    
    Case ECATALOGUETABLETYPE.cttAppDestructionTable
        
        GetSingleFieldTableFieldName = gk_sDBCOLAppraisalDestructionInformation
    
    Case ECATALOGUETABLETYPE.cttArrangementTable
        
        GetSingleFieldTableFieldName = gk_sDBCOLArrangement
    
    Case ECATALOGUETABLETYPE.cttCustodialHistoryTable
        
        GetSingleFieldTableFieldName = gk_sDBCOLCustodialHistory
    
    Case ECATALOGUETABLETYPE.cttNoteTable
        
        GetSingleFieldTableFieldName = gk_sDBCOLNoteText
    
    Case ECATALOGUETABLETYPE.cttScopeAndContentTable
        
        GetSingleFieldTableFieldName = gk_sDBCOLScopeAndContent
    
    Case ECATALOGUETABLETYPE.cttTitleTable
        
        GetSingleFieldTableFieldName = gk_sDBCOLTitle
    
    End Select
    
End Function

Private Function GetReferenceColumnNameInLinkTable(t_eTableType As ECATALOGUETABLETYPE) As String
    
    Select Case t_eTableType
    
    Case ECATALOGUETABLETYPE.cttCorporateBody
    
        GetReferenceColumnNameInLinkTable = gk_sDBCOLCorpBodyReferenceID
        
    Case ECATALOGUETABLETYPE.cttPerson
    
        GetReferenceColumnNameInLinkTable = gk_sDBCOLPersonReferenceID
        
    Case ECATALOGUETABLETYPE.cttPlace
    
        GetReferenceColumnNameInLinkTable = gk_sDBCOLPlaceReferenceID
        
    Case ECATALOGUETABLETYPE.cttSFA
    
        GetReferenceColumnNameInLinkTable = gk_sDBCOLSFAReferenceID
        
    Case ECATALOGUETABLETYPE.cttSubject
    
        GetReferenceColumnNameInLinkTable = gk_sDBCOLSubjectReferenceID
        
    End Select
End Function

Private Function GetSingleFieldTableFieldValue(t_eTableType As ECATALOGUETABLETYPE) As Variant
'Returns the value from the parsed catalogue entry object for the single column table names

    Select Case t_eTableType
    Case ECATALOGUETABLETYPE.cttAccrualsTable
        
        GetSingleFieldTableFieldValue = m_oboParsedCatalogueEntry.Accruals
        
    Case ECATALOGUETABLETYPE.cttAdminHistoryTable
        
        GetSingleFieldTableFieldValue = m_oboParsedCatalogueEntry.History
        
    Case ECATALOGUETABLETYPE.cttAppDestructionTable
        
        GetSingleFieldTableFieldValue = m_oboParsedCatalogueEntry.Appraisal
        
    Case ECATALOGUETABLETYPE.cttArrangementTable
        
        GetSingleFieldTableFieldValue = m_oboParsedCatalogueEntry.Arrangement
        
    Case ECATALOGUETABLETYPE.cttCustodialHistoryTable
        
        GetSingleFieldTableFieldValue = m_oboParsedCatalogueEntry.CustHist
        
    Case ECATALOGUETABLETYPE.cttNoteTable
        
        GetSingleFieldTableFieldValue = m_oboParsedCatalogueEntry.Note
        
    Case ECATALOGUETABLETYPE.cttScopeAndContentTable
        
        GetSingleFieldTableFieldValue = m_oboParsedCatalogueEntry.ScopeContent
        
    Case ECATALOGUETABLETYPE.cttTitleTable
        
        GetSingleFieldTableFieldValue = m_oboParsedCatalogueEntry.Title
        
    End Select
    
End Function




Private Function GetCatalogueID(t_eLevel As ECATALOGUELEVEL) As Long
' Returns the Id (in the member variables) corresponding to the passed level
    
    Select Case t_eLevel
    Case ECATALOGUELEVEL.ctLetterCode
        
        GetCatalogueID = m_lLetterCodeID
        
    Case ECATALOGUELEVEL.ctDivision
    
        GetCatalogueID = m_lDivisionID
        
    Case ECATALOGUELEVEL.ctClass
    
        GetCatalogueID = m_lClassID
        
    Case ECATALOGUELEVEL.ctHeaders
        
        GetCatalogueID = m_lHeaderID
        
    Case ECATALOGUELEVEL.ctSubheaders
        
        GetCatalogueID = m_lSubHeaderID
        
    Case ECATALOGUELEVEL.ctPiece
    
        GetCatalogueID = m_lPieceID
        
    Case ECATALOGUELEVEL.ctItem
    
        GetCatalogueID = m_lItemID
        
    End Select

End Function


Private Function GetLinkTableDescriptionFieldValue(t_eProperty As ECATALOGUEPROPERTIES) As Variant
'Return the description
    
    
    Select Case t_eProperty
    
    Case ECATALOGUEPROPERTIES.cpPublicationNote
        
        GetLinkTableDescriptionFieldValue = m_oboParsedCatalogueEntry.PublicationNote
        
    Case ECATALOGUEPROPERTIES.cpUnpublishedFindingAids
        
        GetLinkTableDescriptionFieldValue = m_oboParsedCatalogueEntry.UnpublishedFindingAids
        
    Case ECATALOGUEPROPERTIES.cpRelatedMaterials
        
        GetLinkTableDescriptionFieldValue = m_oboParsedCatalogueEntry.RelatedMaterialSingleText
        
    Case ECATALOGUEPROPERTIES.cpSeperatedMaterials
        
        GetLinkTableDescriptionFieldValue = m_oboParsedCatalogueEntry.SeparatedMaterialSingleText
        
    Case ECATALOGUEPROPERTIES.cpCopiesInformation
        
        GetLinkTableDescriptionFieldValue = m_oboParsedCatalogueEntry.CopiesSingleText
        
    Case ECATALOGUEPROPERTIES.cpImmediateSourceOfAcquisition
        
        GetLinkTableDescriptionFieldValue = m_oboParsedCatalogueEntry.SourceAcquisitionSingleText
        
    End Select
    

End Function


Private Function NullifZero(t_lValue As Long) As Variant
'Function returns null (string) is the value is zero else returns
'       the value
    If t_lValue = 0 Then
        NullifZero = "Null"
    Else
        NullifZero = t_lValue
    End If
End Function

Private Function NullifEmptyString(t_sValue As String) As Variant
'Function returns null (string) is the value is zero else returns
'       the value
    If t_sValue = "" Then
        NullifEmptyString = Null
    Else
        NullifEmptyString = t_sValue
    End If
End Function



Private Function IsSubClassNumberCorrect() As Boolean
' Description:
'   Searches for series entries which match this one in terms of
'   series sub-number; if entries already in the database have them,
'   then this entry must also have one.  If entries in the database
'   do not have them, then there cannot be another series entry with
'   that series number.  For example:
'       In database: XXX 4
'           Entry = XXX 4/1 - Invalid
'           Entry = XXX 4 - Invalid - but let pass to get ID at a later stage
'       In database: XXX 4/1
'           Entry = XXX 4/2 - Valid
'           Entry = XXX 4 - Invalid
'
' Amendment History:
'   d7  DAJ     08Mar2000   Created
'   d10 DAJ     13Mar2000   Changed to allow the InsertAndReturnIdentity function
'                           to get the ID of the class if it exists - this is required
'                           if it is a parent.  Basically, only check for classes with
'                           the subnumber as NULL if the subnumber is non-zero
    
    ' Declarations
    Dim ofwResult As FWRecordList
    Dim sSQL As String
    
    ' Set default return value
    IsSubClassNumberCorrect = True
    
    ' d10 - We only check if the subnumber is non-zero; it is zero,
    ' then any duplication will be caught later so the ID can be
    ' retrieved
    If m_oboParsedCatalogueEntry.SubClassNumber = 0 Then
        Exit Function
    End If
    
    ' Create the SQL to find the record
    
    ' We want to find entries with the same class number all the time...
    sSQL = "SELECT * FROM " & gk_sDBTABLEClass & " (nolock)" & _
        " WHERE " & gk_sDBCOLClassNumber & "=" & m_oboParsedCatalogueEntry.ClassNumber & _
        " AND " & gk_sDBCOLLetterCodeID & " = " & m_lLetterCodeID
        
    ' ... and we want to find entries with same class number
    ' AND a null in subnumber
    sSQL = sSQL & " AND " & gk_sDBCOLSubClassNumber & " IS NULL"
    
    ' Execute the SQL
    Set ofwResult = m_oFWDBHelper.FindBySql(sSQL)
    
    ' If records are found, then the current entry is invalid
    If ofwResult.RecordCount > 0 Then
        IsSubClassNumberCorrect = False
    End If

End Function


Private Sub SetCatalogueEntryRecordIDInMemberVariables(t_eLevel As ECATALOGUELEVEL)
'Method to set the member varable for the current level with the ID value as the
' insert failed because the record was already present
'
    Dim ofwResult As FWRecordList
    Dim sIDColumnName As String
    
    sIDColumnName = GetColumnName(ECATALOGUECOLUMNTYPE.cctIDColumn, t_eLevel)
    
    Set ofwResult = m_oFWDBHelper.FindBySql(GetSingleRecordSelectSQL(t_eLevel))
    
    If ofwResult.RecordCount > 0 Then
    
        Select Case t_eLevel
        Case ECATALOGUELEVEL.ctLetterCode
            
            m_lLetterCodeID = ofwResult.GetFieldValue(sIDColumnName)
            
        Case ECATALOGUELEVEL.ctDivision
        
            m_lDivisionID = ofwResult.GetFieldValue(sIDColumnName)
            
        Case ECATALOGUELEVEL.ctClass
            
            m_lClassID = ofwResult.GetFieldValue(sIDColumnName)
            
        Case ECATALOGUELEVEL.ctHeaders
            
            m_lHeaderID = ofwResult.GetFieldValue(sIDColumnName)
            
        Case ECATALOGUELEVEL.ctSubheaders
            
            m_lSubHeaderID = ofwResult.GetFieldValue(sIDColumnName)
            
        Case ECATALOGUELEVEL.ctPiece
            
            m_lPieceID = ofwResult.GetFieldValue(sIDColumnName)
            
        Case ECATALOGUELEVEL.ctItem
            
            m_lItemID = ofwResult.GetFieldValue(sIDColumnName)
            
        End Select
        
    End If

End Sub



Private Function GetSingleRecordSelectSQL(t_eLevel As ECATALOGUELEVEL) As String
'Returns a SQL for selecting the ID for the record present in the level
' Since Item ref can be null this can lead to a query with multiple rows
'
'   d7  DAJ     08Mar2000   Added subclass number handling
    Dim sSQL As String
    
    Select Case t_eLevel
    
    Case ECATALOGUELEVEL.ctLetterCode
    
        sSQL = "SELECT " & gk_sDBCOLLetterCodeID & " FROM " & _
                gk_sDBTABLELCode & " (nolock) WHERE " & _
                gk_sDBCOLLetterCodeCode & " = " & QuoteValueWithQuotes(m_oboParsedCatalogueEntry.LetterCode)
    
    Case ECATALOGUELEVEL.ctDivision
    
        sSQL = "SELECT " & gk_sDBTABLEDivision & "." & gk_sDBCOLDivisionID & " FROM " & _
                gk_sDBTABLELCode & " (nolock), " & _
                gk_sDBTABLEDivision & " (nolock) WHERE " & _
                gk_sDBTABLELCode & "." & gk_sDBCOLLetterCodeCode & " = " & QuoteValueWithQuotes(m_oboParsedCatalogueEntry.LetterCode) & _
                " AND " & gk_sDBTABLELCode & "." & gk_sDBCOLLetterCodeID & " = " & gk_sDBTABLEDivision & "." & gk_sDBCOLLetterCodeID & _
                " AND " & gk_sDBTABLEDivision & "." & gk_sDBCOLDivisionNo & " = " & CStr(m_oboParsedCatalogueEntry.DivisionNumber)
    
    Case ECATALOGUELEVEL.ctClass
        
        sSQL = "SELECT " & gk_sDBTABLEClass & "." & gk_sDBCOLClassID & " FROM " & _
                gk_sDBTABLELCode & " (nolock), " & _
                gk_sDBTABLEClass & " (nolock) WHERE " & _
                gk_sDBTABLELCode & "." & gk_sDBCOLLetterCodeCode & " = " & QuoteValueWithQuotes(m_oboParsedCatalogueEntry.LetterCode) & _
                " AND " & gk_sDBTABLELCode & "." & gk_sDBCOLLetterCodeID & " = " & gk_sDBTABLEClass & "." & gk_sDBCOLLetterCodeID & _
                " AND " & gk_sDBTABLEClass & "." & gk_sDBCOLClassNumber & " = " & CStr(m_oboParsedCatalogueEntry.ClassNumber)
        
        ' d7 - Added check for subclass number
        If m_oboParsedCatalogueEntry.SubClassNumber <> 0 Then
            sSQL = sSQL & _
                " AND " & gk_sDBTABLEClass & "." & gk_sDBCOLSubClassNumber & " = " & m_oboParsedCatalogueEntry.SubClassNumber
        End If
        
    Case ECATALOGUELEVEL.ctHeaders
    
        sSQL = "SELECT " & gk_sDBTABLEHeader & "." & gk_sDBCOLHeaderID & " FROM " & _
                gk_sDBTABLEHeader & " (nolock), " & _
                gk_sDBTABLELCode & " (nolock), " & _
                gk_sDBTABLEClass & " (nolock) WHERE " & _
                gk_sDBTABLELCode & "." & gk_sDBCOLLetterCodeCode & " = " & QuoteValueWithQuotes(m_oboParsedCatalogueEntry.LetterCode) & _
                " AND " & gk_sDBTABLELCode & "." & gk_sDBCOLLetterCodeID & " = " & gk_sDBTABLEClass & "." & gk_sDBCOLLetterCodeID & _
                " AND " & gk_sDBTABLEClass & "." & gk_sDBCOLClassNumber & " = " & CStr(m_oboParsedCatalogueEntry.ClassNumber) & _
                " AND " & gk_sDBTABLEClass & "." & gk_sDBCOLClassID & " = " & gk_sDBTABLEHeader & "." & gk_sDBCOLClassID & _
                " AND " & gk_sDBTABLEHeader & "." & gk_sDBCOLClassHeaderNo & " = " & CStr(m_oboParsedCatalogueEntry.HeaderNumber)
    
    Case ECATALOGUELEVEL.ctSubheaders
        
        sSQL = "SELECT " & gk_sDBTABLESubHeader & "." & gk_sDBCOLSubHeaderID & " FROM " & _
                gk_sDBTABLESubHeader & " (nolock), " & _
                gk_sDBTABLEHeader & " (nolock), " & _
                gk_sDBTABLELCode & " (nolock), " & _
                gk_sDBTABLEClass & " (nolock) WHERE " & _
                gk_sDBTABLELCode & "." & gk_sDBCOLLetterCodeCode & " = " & QuoteValueWithQuotes(m_oboParsedCatalogueEntry.LetterCode) & _
                " AND " & gk_sDBTABLELCode & "." & gk_sDBCOLLetterCodeID & " = " & gk_sDBTABLEClass & "." & gk_sDBCOLLetterCodeID & _
                " AND " & gk_sDBTABLEClass & "." & gk_sDBCOLClassNumber & " = " & CStr(m_oboParsedCatalogueEntry.ClassNumber) & _
                " AND " & gk_sDBTABLEClass & "." & gk_sDBCOLClassID & " = " & gk_sDBTABLEHeader & "." & gk_sDBCOLClassID & _
                " AND " & gk_sDBTABLEHeader & "." & gk_sDBCOLClassHeaderNo & " = " & CStr(m_oboParsedCatalogueEntry.HeaderNumber) & _
                " AND " & gk_sDBTABLEHeader & "." & gk_sDBCOLHeaderID & " = " & gk_sDBTABLESubHeader & "." & gk_sDBCOLHeaderID & _
                " AND " & gk_sDBTABLESubHeader & "." & gk_sDBCOLSubHeaderNo & " = " & CStr(m_oboParsedCatalogueEntry.SubHeaderNumber)
    
    Case ECATALOGUELEVEL.ctPiece
    
        sSQL = "SELECT " & gk_sDBTABLEPiece & "." & gk_sDBCOLPieceID & " FROM " & _
                gk_sDBTABLEPiece & " (nolock), " & _
                gk_sDBTABLELCode & " (nolock), " & _
                gk_sDBTABLEClass & " (nolock) WHERE " & _
                gk_sDBTABLELCode & "." & gk_sDBCOLLetterCodeCode & " = " & QuoteValueWithQuotes(m_oboParsedCatalogueEntry.LetterCode) & _
                " AND " & gk_sDBTABLELCode & "." & gk_sDBCOLLetterCodeID & " = " & gk_sDBTABLEClass & "." & gk_sDBCOLLetterCodeID & _
                " AND " & gk_sDBTABLEClass & "." & gk_sDBCOLClassNumber & " = " & CStr(m_oboParsedCatalogueEntry.ClassNumber) & _
                " AND " & gk_sDBTABLEClass & "." & gk_sDBCOLClassID & " = " & gk_sDBTABLEPiece & "." & gk_sDBCOLClassID & _
                " AND " & gk_sDBTABLEPiece & "." & gk_sDBCOLPieceRef & " = " & QuoteValueWithQuotes(m_oboParsedCatalogueEntry.PieceRef)
    
    Case ECATALOGUELEVEL.ctItem
    
        sSQL = "SELECT " & gk_sDBTABLEItem & "." & gk_sDBCOLItemID & " FROM " & _
                gk_sDBTABLEItem & " (nolock), " & _
                gk_sDBTABLEPiece & " (nolock), " & _
                gk_sDBTABLELCode & " (nolock), " & _
                gk_sDBTABLEClass & " (nolock) WHERE " & _
                gk_sDBTABLELCode & "." & gk_sDBCOLLetterCodeCode & " = " & QuoteValueWithQuotes(m_oboParsedCatalogueEntry.LetterCode) & _
                " AND " & gk_sDBTABLELCode & "." & gk_sDBCOLLetterCodeID & " = " & gk_sDBTABLEClass & "." & gk_sDBCOLLetterCodeID & _
                " AND " & gk_sDBTABLEClass & "." & gk_sDBCOLClassNumber & " = " & CStr(m_oboParsedCatalogueEntry.ClassNumber) & _
                " AND " & gk_sDBTABLEClass & "." & gk_sDBCOLClassID & " = " & gk_sDBTABLEPiece & "." & gk_sDBCOLClassID & _
                " AND " & gk_sDBTABLEPiece & "." & gk_sDBCOLPieceRef & " = " & QuoteValueWithQuotes(m_oboParsedCatalogueEntry.PieceRef) & _
                " AND " & gk_sDBTABLEPiece & "." & gk_sDBCOLPieceID & " = " & gk_sDBTABLEItem & "." & gk_sDBCOLPieceID & _
                " AND " & gk_sDBTABLEItem & "." & gk_sDBCOLItemRef & " = " & QuoteValueWithQuotes(m_oboParsedCatalogueEntry.ItemRef)

    End Select
    
    GetSingleRecordSelectSQL = sSQL
End Function


Private Function SetParsedEntry(t_ofwResult As FWRecordList, t_eTableType As ECATALOGUETABLETYPE)
' Description: Sets the values of the member variable parsed catalogue
'
' Amendment History
'   Vers    Auth            Date            Comment
'   d1      A.Farrow        25-Jul-2000     Created
'

'Make sure that result contains something
    If Not t_ofwResult Is Nothing Then
        If t_ofwResult.RecordCount > 0 Then
                'Select the variable to set
            Select Case t_eTableType
                Case ECATALOGUETABLETYPE.cttAccrualsTable
        
                    m_oboParsedCatalogueEntry.Accruals = IIf(IsNull(t_ofwResult.GetFieldValue(GetSingleFieldTableFieldName(ECATALOGUETABLETYPE.cttAccrualsTable))), "", t_ofwResult.GetFieldValue(GetSingleFieldTableFieldName(ECATALOGUETABLETYPE.cttAccrualsTable)))
        
                Case ECATALOGUETABLETYPE.cttAdminHistoryTable
        
                    m_oboParsedCatalogueEntry.History = IIf(IsNull(t_ofwResult.GetFieldValue(GetSingleFieldTableFieldName(ECATALOGUETABLETYPE.cttAdminHistoryTable))), "", t_ofwResult.GetFieldValue(GetSingleFieldTableFieldName(ECATALOGUETABLETYPE.cttAdminHistoryTable)))
        
                Case ECATALOGUETABLETYPE.cttAppDestructionTable
        
                    m_oboParsedCatalogueEntry.Appraisal = IIf(IsNull(t_ofwResult.GetFieldValue(GetSingleFieldTableFieldName(ECATALOGUETABLETYPE.cttAppDestructionTable))), "", t_ofwResult.GetFieldValue(GetSingleFieldTableFieldName(ECATALOGUETABLETYPE.cttAppDestructionTable)))
        
                Case ECATALOGUETABLETYPE.cttArrangementTable
        
                    m_oboParsedCatalogueEntry.Arrangement = IIf(IsNull(t_ofwResult.GetFieldValue(GetSingleFieldTableFieldName(ECATALOGUETABLETYPE.cttArrangementTable))), "", t_ofwResult.GetFieldValue(GetSingleFieldTableFieldName(ECATALOGUETABLETYPE.cttArrangementTable)))
        
                Case ECATALOGUETABLETYPE.cttCustodialHistoryTable
        
                    m_oboParsedCatalogueEntry.CustHist = IIf(IsNull(t_ofwResult.GetFieldValue(GetSingleFieldTableFieldName(ECATALOGUETABLETYPE.cttCustodialHistoryTable))), "", t_ofwResult.GetFieldValue(GetSingleFieldTableFieldName(ECATALOGUETABLETYPE.cttCustodialHistoryTable)))
        
                Case ECATALOGUETABLETYPE.cttNoteTable
        
                    m_oboParsedCatalogueEntry.Note = IIf(IsNull(t_ofwResult.GetFieldValue(GetSingleFieldTableFieldName(ECATALOGUETABLETYPE.cttNoteTable))), "", t_ofwResult.GetFieldValue(GetSingleFieldTableFieldName(ECATALOGUETABLETYPE.cttNoteTable)))
        
                Case ECATALOGUETABLETYPE.cttScopeAndContentTable
        
                    m_oboParsedCatalogueEntry.ScopeContent = IIf(IsNull(t_ofwResult.GetFieldValue(GetSingleFieldTableFieldName(ECATALOGUETABLETYPE.cttScopeAndContentTable))), "", t_ofwResult.GetFieldValue(GetSingleFieldTableFieldName(ECATALOGUETABLETYPE.cttScopeAndContentTable)))
        
                Case ECATALOGUETABLETYPE.cttTitleTable
        
                    m_oboParsedCatalogueEntry.Title = IIf(IsNull(t_ofwResult.GetFieldValue(GetSingleFieldTableFieldName(ECATALOGUETABLETYPE.cttTitleTable))), "", t_ofwResult.GetFieldValue(GetSingleFieldTableFieldName(ECATALOGUETABLETYPE.cttTitleTable)))
        
            End Select
        End If
    End If
            
End Function

Private Function SetParsedLinkEntry(t_ofwResult As FWRecordList, t_eProperty As ECATALOGUEPROPERTIES)
' Description: Sets the values of the member variable parsed catalogue
'
' Amendment History
'   Vers    Auth            Date            Comment
'   d1      A.Farrow        25-Jul-2000     Created
'
Dim sFieldValue As String

        If Not t_ofwResult Is Nothing Then
            If t_ofwResult.RecordCount > 0 Then
                'Select the variable to set
                Select Case t_eProperty
    
                    Case ECATALOGUEPROPERTIES.cpPublicationNote
        
                        m_oboParsedCatalogueEntry.PublicationNote = _
                            t_ofwResult.GetFieldValue(gk_sDBCOLDescription)
        
                    Case ECATALOGUEPROPERTIES.cpUnpublishedFindingAids
        
                        m_oboParsedCatalogueEntry.UnpublishedFindingAids = _
                            t_ofwResult.GetFieldValue(gk_sDBCOLDescription)

                    Case ECATALOGUEPROPERTIES.cpRelatedMaterials
        
                        m_oboParsedCatalogueEntry.RelatedMaterialSingleText = _
                            t_ofwResult.GetFieldValue(gk_sDBCOLDescription)


                    Case ECATALOGUEPROPERTIES.cpSeperatedMaterials
        
                        m_oboParsedCatalogueEntry.SeparatedMaterialSingleText = _
                            t_ofwResult.GetFieldValue(gk_sDBCOLDescription)

        
                    Case ECATALOGUEPROPERTIES.cpCopiesInformation
        
                        m_oboParsedCatalogueEntry.CopiesSingleText = _
                            t_ofwResult.GetFieldValue(gk_sDBCOLDescription)
        
                    Case ECATALOGUEPROPERTIES.cpImmediateSourceOfAcquisition
        
                        m_oboParsedCatalogueEntry.SourceAcquisitionSingleText = _
                            t_ofwResult.GetFieldValue(gk_sDBCOLDescription)

        
                End Select
            End If
        End If
End Function

Private Function CreateParsedCollection(ofwResult As FWRecordList, t_eProperty As ECATALOGUEPROPERTIES, t_eTable As ECATALOGUETABLETYPE)
' Description: Recussively add entries to the collections in the parsed catalogue
'
' Amendment History
'   Vers    Auth            Date            Comment
'   d1      A.Farrow        26-Jul-2000     Created
'   d2      D Johnston      03-Aug-2000     Added IIFs in CreateParsedCollection to cope with Null descriptions on link terms
'
 Dim CorpBodyID As Long


 'Now we have the results back we can store them in an array
    If Not ofwResult Is Nothing Then
        If ofwResult.RecordCount > 0 Then
            While Not ofwResult.IsAtEnd
                Select Case t_eProperty
                     
                    '  We construct the pasrsed catalogue entry with ref IDs and desc
                    Case ECATALOGUEPROPERTIES.cpCopiesInformation
                        m_oboParsedCatalogueEntry.AddCopy ofwResult.GetFieldValue(GetReferenceColumnNameInLinkTable(t_eTable)), gk_sDBCOLDescription
            
                    Case ECATALOGUEPROPERTIES.cpImmediateSourceOfAcquisition
                        If t_eTable = ECATALOGUETABLETYPE.cttCorporateBody Then
                            m_oboParsedCatalogueEntry.AddSourceAcquisitionCorpBody ofwResult.GetFieldValue(GetReferenceColumnNameInLinkTable(t_eTable)), _
                                IIf(IsNull(ofwResult.GetFieldValue(gk_sDBCOLDescription)), "", ofwResult.GetFieldValue(gk_sDBCOLDescription))
                        Else
                            m_oboParsedCatalogueEntry.AddSourceAcquisitionPerson ofwResult.GetFieldValue(GetReferenceColumnNameInLinkTable(t_eTable)), _
                                IIf(IsNull(ofwResult.GetFieldValue(gk_sDBCOLDescription)), "", ofwResult.GetFieldValue(gk_sDBCOLDescription))
                        End If
                    
                    Case ECATALOGUEPROPERTIES.cpLocationOfOriginals
                        m_oboParsedCatalogueEntry.AddLocationOriginal ofwResult.GetFieldValue(GetReferenceColumnNameInLinkTable(t_eTable)), _
                            IIf(IsNull(ofwResult.GetFieldValue(gk_sDBCOLDescription)), "", ofwResult.GetFieldValue(gk_sDBCOLDescription))
                        
                    
                    ' The following properties do not need a description
                    Case ECATALOGUEPROPERTIES.cpCreatorNames
                        If t_eTable = ECATALOGUETABLETYPE.cttCorporateBody Then
                            CorpBodyID = CLng(ofwResult.GetFieldValue(GetReferenceColumnNameInLinkTable(t_eTable)))
                            m_oboParsedCatalogueEntry.AddCreatorCorpBody CorpBodyID
                        Else
                            m_oboParsedCatalogueEntry.AddCreatorPerson ofwResult.GetFieldValue(GetReferenceColumnNameInLinkTable(t_eTable))
                        End If
                        
                    Case ECATALOGUEPROPERTIES.cpPlaceOfDeposit
                        m_oboParsedCatalogueEntry.AddPlaceDeposit ofwResult.GetFieldValue(GetReferenceColumnNameInLinkTable(t_eTable))
                        
                    Case ECATALOGUEPROPERTIES.cpIndexTermCorporateName
                        m_oboParsedCatalogueEntry.AddCorpBodyIndex ofwResult.GetFieldValue(GetReferenceColumnNameInLinkTable(t_eTable))
                        
                    Case ECATALOGUEPROPERTIES.cpIndexTermPersonName
                        m_oboParsedCatalogueEntry.AddPersonIndex ofwResult.GetFieldValue(GetReferenceColumnNameInLinkTable(t_eTable))
                        
                    Case ECATALOGUEPROPERTIES.cpIndexTermSubject
                        m_oboParsedCatalogueEntry.AddSubjectIndex ofwResult.GetFieldValue(GetReferenceColumnNameInLinkTable(t_eTable))
                        
                    Case ECATALOGUEPROPERTIES.cpIndexTermPlaceName
                        m_oboParsedCatalogueEntry.AddPlaceIndex ofwResult.GetFieldValue(GetReferenceColumnNameInLinkTable(t_eTable))
                    
                    Case ECATALOGUEPROPERTIES.cpRelatedMaterials
                        m_oboParsedCatalogueEntry.AddRelatedMaterialLink ofwResult.GetFieldValue(gk_sDBCOLCatalogueID2), ofwResult.GetFieldValue("level_no2") _
                                , IIf(IsNull(ofwResult.GetFieldValue(gk_sDBCOLDescription)), "", ofwResult.GetFieldValue(gk_sDBCOLDescription))
                                
                    Case ECATALOGUEPROPERTIES.cpSeperatedMaterials
                         m_oboParsedCatalogueEntry.AddSeparatedMaterialLink ofwResult.GetFieldValue(gk_sDBCOLCatalogueID2), ofwResult.GetFieldValue("level_no2") _
                                , IIf(IsNull(ofwResult.GetFieldValue(gk_sDBCOLDescription)), "", ofwResult.GetFieldValue(gk_sDBCOLDescription))
                End Select
                ofwResult.MoveNext
            Wend
        End If
    End If
End Function

Private Function ExtensionTableExtraction(t_eLevel As ECATALOGUELEVEL)
' Description: Pull out all the values from the extension table
'
' Amendment History
'   Vers    Auth            Date            Comment
'   d1      A.Farrow        25-Jul-2000     Created
'
Dim sSQL As String
    Dim sTablename As String
    Dim sIDColumnName As String
    Dim ofwResult As New FWRecordList
    
    sTablename = GetTableName(ECATALOGUETABLETYPE.cttExtensionTable, _
                    ESectionType.SECReader, t_eLevel)
                    
    sIDColumnName = GetColumnName(ECATALOGUECOLUMNTYPE.cctIDColumn, t_eLevel)
            
            sSQL = "SELECT * FROM " & sTablename

            sSQL = sSQL & " WHERE " & sIDColumnName & " = " & GetCatalogueID(t_eLevel)
            
            
            Set ofwResult = m_oFWDBHelper.FindBySql(sSQL)
            If Not ofwResult Is Nothing Then
                If ofwResult.RecordCount > 0 Then
                        
                        'Set All the values
                        m_oboParsedCatalogueEntry.AccessConditionID = IIf(IsNull(ofwResult.GetFieldValue(gk_sDBCOLAccessConditionID)), 0, ofwResult.GetFieldValue(gk_sDBCOLAccessConditionID))
                        If Not (t_eLevel = ECATALOGUELEVEL.ctPiece Or t_eLevel = ECATALOGUELEVEL.ctItem) Then
                            m_oboParsedCatalogueEntry.LanguageID = IIf(IsNull(ofwResult.GetFieldValue(gk_sDBCOLLanguageID)), 0, ofwResult.GetFieldValue(gk_sDBCOLLanguageID))
                        End If
                        
                        m_oboParsedCatalogueEntry.AccumulationDates = IIf(IsNull(ofwResult.GetFieldValue(gk_sDBCOLAccumulationdates)), "", ofwResult.GetFieldValue(gk_sDBCOLAccumulationdates))
                        
                        ' We only get the Covering dates from the main table
                        If t_eLevel = ECATALOGUELEVEL.ctLetterCode Or t_eLevel = ECATALOGUELEVEL.ctDivision Then
                            m_oboParsedCatalogueEntry.DateText = IIf(IsNull(ofwResult.GetFieldValue(gk_sDBCOLCoveringDates)), "", ofwResult.GetFieldValue(gk_sDBCOLCoveringDates))
                        End If
                        m_oboParsedCatalogueEntry.LegalStatusID = IIf(IsNull(ofwResult.GetFieldValue(gk_sDBCOLLegalStatusID)), 0, ofwResult.GetFieldValue(gk_sDBCOLLegalStatusID))
                        m_oboParsedCatalogueEntry.MapDesignation = IIf(IsNull(ofwResult.GetFieldValue(gk_sDBCOLMapDesignation)), "", ofwResult.GetFieldValue(gk_sDBCOLMapDesignation))
                        m_oboParsedCatalogueEntry.PhysdescCondition = IIf(IsNull(ofwResult.GetFieldValue(gk_sDBCOLPhysicalCondition)), "", ofwResult.GetFieldValue(gk_sDBCOLPhysicalCondition))
                        m_oboParsedCatalogueEntry.PhysdescFormID = IIf(IsNull(ofwResult.GetFieldValue(gk_sDBCOLPhysicalDescriptionFormID)), 0, ofwResult.GetFieldValue(gk_sDBCOLPhysicalDescriptionFormID))
                        m_oboParsedCatalogueEntry.PhysdescDimensions = IIf(IsNull(ofwResult.GetFieldValue(gk_sDBCOLPhysicalDescriptionDimension)), "", ofwResult.GetFieldValue(gk_sDBCOLPhysicalDescriptionDimension))
                        m_oboParsedCatalogueEntry.RestrictionsOnUse = IIf(IsNull(ofwResult.GetFieldValue(gk_sDBCOLRestrictionsOnUse)), "", ofwResult.GetFieldValue(gk_sDBCOLRestrictionsOnUse))
                        m_oboParsedCatalogueEntry.ReferenceDepartment = IIf(IsNull(ofwResult.GetFieldValue(gk_sDBCOLFormerDepartmentReference)), "", ofwResult.GetFieldValue(gk_sDBCOLFormerDepartmentReference))
                        m_oboParsedCatalogueEntry.ReferencePRO = IIf(IsNull(ofwResult.GetFieldValue(gk_sDBCOLFormerPROReference)), "", ofwResult.GetFieldValue(gk_sDBCOLFormerPROReference))

                        If Not IsNull(ofwResult.GetFieldValue(gk_sDBCOLPhysicalDescriptionRecordQuantity)) Then
                            m_oboParsedCatalogueEntry.PhysdescExtent = ofwResult.GetFieldValue(gk_sDBCOLPhysicalDescriptionRecordQuantity)
                            m_oboParsedCatalogueEntry.PhysdescExtentNull = False
                        Else
                            m_oboParsedCatalogueEntry.PhysdescExtentNull = True
                        End If
                        
                        If Not IsNull(ofwResult.GetFieldValue(gk_sDBCOLMapScaleNumber)) Then
                            m_oboParsedCatalogueEntry.ScaleNumber = ofwResult.GetFieldValue(gk_sDBCOLMapScaleNumber)
                            m_oboParsedCatalogueEntry.ScaleNumberNull = False
                        Else
                            m_oboParsedCatalogueEntry.ScaleNumberNull = True
                        End If
    
                End If
            End If
End Function

Private Function GetSFALinks(t_eLevel As ECATALOGUELEVEL, t_eProperty As ECATALOGUEPROPERTIES)
' Description: Get any unpublished and published finding aids for SFAs
'
' Amendment History
'   Vers    Auth            Date            Comment
'   d1      A.Farrow        04-Aug-2000     Created
'
    Dim sDesc As String
    Dim lReferenceID As Long
    Dim ofwResult As FWRecordList
    Dim ofwLinkedTableResult As FWRecordList
    Dim sSQL As String
    Dim linkRecord As GenericLinkRecord
    ' Get the sfa unpublished aids first
    
    'Get the link
    sSQL = "SELECT " & gk_sDBCOLDescription & "," _
                & gk_sDBCOLSFAReferenceID & _
                " FROM " & gk_sDBTABLEPCSFA & " (nolock)"
            
    sSQL = sSQL & " Where " & gk_sDBCOLCatalogueID & " = " & GetCatalogueID(t_eLevel) & _
                        " AND " & gk_sDBCOLCatalogueLevel & " = " & t_eLevel & _
                        " AND " & gk_sDBCOLPropertyID & " = " & t_eProperty
                        
     Set ofwResult = m_oFWDBHelper.FindBySql(sSQL)
        
        If Not ofwResult Is Nothing Then
            If ofwResult.RecordCount > 0 Then
                While Not ofwResult.IsAtEnd
                    sDesc = ofwResult.GetFieldValue(gk_sDBCOLDescription) & ","
                    lReferenceID = ofwResult.GetFieldValue(gk_sDBCOLSFAReferenceID)
                    
                    'Get the text from the linked table
                    sSQL = "SELECT " & gk_sDBCOLSFATitleText & _
                        " FROM " & gk_sDBTABLESFA & " (nolock)"
                        
                     sSQL = sSQL & " Where " & gk_sDBCOLSFAReferenceID & " = " & lReferenceID
                     
                     Set ofwLinkedTableResult = m_oFWDBHelper.FindBySql(sSQL)
                     
                     If Not ofwLinkedTableResult Is Nothing Then
                        If ofwLinkedTableResult.RecordCount > 0 Then
                            sDesc = sDesc & ofwLinkedTableResult.GetFieldValue(gk_sDBCOLSFATitleText)
                            If t_eProperty = cpUnpublishedFindingAids Then
                                m_oboParsedCatalogueEntry.AddUnpublishedFindingAidSFALinks lReferenceID, sDesc
                            Else
                                m_oboParsedCatalogueEntry.AddPublicationFindingAidSFALinks lReferenceID, sDesc
                            End If
                        End If
                     End If
                     ofwResult.MoveNext
                Wend
            End If
        End If
        
        linkRecord.sOptionalDescription = sDesc
        
                                           
End Function

Private Function GetLeafletLinks(t_eLevel As ECATALOGUELEVEL, t_eProperty As ECATALOGUEPROPERTIES)
' Description: Get any unpublished and published finding aids for leaflets
'
' Amendment History
'   Vers    Auth            Date            Comment
'   d1      A.Farrow        04-Aug-2000     Created
'
    Dim sDesc As String
    Dim lReferenceID As Long
    Dim ofwResult As FWRecordList
    Dim ofwLinkedTableResult As FWRecordList
    Dim sSQL As String
    ' Get the sfa unpublished aids first
    
    'Get the link
    sSQL = "SELECT " & gk_sDBCOLDescription & "," _
                & gk_sDBCOLLeafletReferenceID & _
                " FROM " & gk_sDBTABLEPCLeaflet & " (nolock)"
            
    sSQL = sSQL & " Where " & gk_sDBCOLCatalogueID & " = " & GetCatalogueID(t_eLevel) & _
                        " AND " & gk_sDBCOLCatalogueLevel & " = " & t_eLevel & _
                        " AND " & gk_sDBCOLPropertyID & " = " & t_eProperty
                        
    Set ofwResult = m_oFWDBHelper.FindBySql(sSQL)
        
    If Not ofwResult Is Nothing Then
        If ofwResult.RecordCount > 0 Then
            While Not ofwResult.IsAtEnd
                sDesc = ofwResult.GetFieldValue(gk_sDBCOLDescription) & ","
                lReferenceID = ofwResult.GetFieldValue(gk_sDBCOLLeafletReferenceID)
                
                'Get the text from the linked table
                sSQL = "SELECT " & gk_sDBCOLLeafletTitleText _
                    & " FROM " & gk_sDBTABLELeaflet & " (nolock)"
                    
                 sSQL = sSQL & " Where " & gk_sDBCOLLeafletReferenceID & " = " & lReferenceID
                 
                 Set ofwLinkedTableResult = m_oFWDBHelper.FindBySql(sSQL)
                 
                 If Not ofwLinkedTableResult Is Nothing Then
                    If ofwLinkedTableResult.RecordCount > 0 Then
                        sDesc = sDesc & ofwLinkedTableResult.GetFieldValue(gk_sDBCOLLeafletTitleText)
                        If t_eProperty = cpUnpublishedFindingAids Then
                            m_oboParsedCatalogueEntry.AddUnpublishedFindingAidLeafletLinks lReferenceID, sDesc
                        Else
                            m_oboParsedCatalogueEntry.AddPublicationFindingAidLeafletLinks lReferenceID, sDesc
                        End If
                    End If
                 End If
                 ofwResult.MoveNext
            Wend
        End If
    End If
                                           
End Function
Private Function GenerateExportReference(t_cCatRef As CatReferenceParts) As String
'   Description :  Generates the export reference.
'
'   Vers    Auth        Date            Comment
'   d1      A.Farrow    07-Aug-2000     Created
'
    
    Dim sRef As String
    
    If Not t_cCatRef.sLettercodeRef = "" Then
        sRef = sRef & t_cCatRef.sLettercodeRef
    End If
    If Not t_cCatRef.sDivisionRef = "" Then
        sRef = sRef & "-" & "d" & t_cCatRef.sDivisionRef
    End If
    If Not t_cCatRef.sClassRef = "" Then
        sRef = sRef & "-" & "c" & t_cCatRef.sClassRef
    End If
    If Not t_cCatRef.sSubclassRef = "" Then
        sRef = sRef & "-" & t_cCatRef.sSubclassRef
    End If
    If Not t_cCatRef.sHeaderRef = "" Then
        sRef = sRef & "-" & "H" & t_cCatRef.sHeaderRef
    End If
    If Not t_cCatRef.sSubheaderRef = "" Then
        sRef = sRef & "-" & "h" & t_cCatRef.sSubheaderRef
    End If
    If Not t_cCatRef.sPieceRef = "" Then
        sRef = sRef & "-" & "p" & t_cCatRef.sPieceRef
    End If
    If Not t_cCatRef.sItemRef = "" Then
        sRef = sRef & "-" & "i" & t_cCatRef.sItemRef
    End If
    
    GenerateExportReference = sRef

End Function

