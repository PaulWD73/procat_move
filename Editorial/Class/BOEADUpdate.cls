VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BOEADUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    BOEADUpdate.cls
' System:       PRO, PROCAT
' Copyright:    (C) Quidnunc Limited
'
' Description:  This class contains the methods and properties needed to
'    update catalogue entries in the database from EAD files.
'
'    This business object assumes that catalogues are loaded in the proper
'    order at all times. That is the call to Add is made only from the root
'    to the leaf ie from Lettercode to item level successively.
'    The business object preserves states for this purpose.
'
'    Converts all lookup 0 values to Null and only Item reference string empty
'    value to null. All other fields are inserted without any data change.
'    Few fields have a null field in the parsed catalogue entry as property which is used
'           m_oboParsedCatalogueEntry.PhysdescExtentNull and
'           m_oboParsedCatalogueEntry.ScaleNumberNull
'
'    Since for the Item level there are no Uniqueness constraint it needs to be
'    implemented in a select query. This query will ignore if the value is null.
'    Null item references are always assumed to be non unique.
'
' Amendment history:
'   Version Name            Date            Comments
'   d1      D Johnston      19-Jul-2000     Created from BOEADImport
'   d2      D Johnston      04-Sep-2000     Added level checking to *HasChanged functions, and tidied up
'   d3      D Johnston      05-Oct-2000     Removed the code to delete reciprocal Related Material links, as
'                                           the release code no longer puts them in
'   d4      A Manic         28-Feb-2001     Updates start / end dates
'   d5      A Manic         26-Nov-2001     GetSingleFieldTableFieldValue modified to insert spaces between adjacent paragraphs marks
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' ObjectContext object and ASP objects can be accessed (if necessary)
Private m_objContext As ObjectContext


'The current state of catalogues (ID) are stored in the local variables

Private m_lBatchID As Long
Private m_lCatalogueID As Long
Private m_lEditsetID As Long

Private m_oboParsedCatalogueEntry As ParsedCatalogueEntry
Private m_oFWDBHelper As FWDBHelper
Private m_oBOCatalogueDBHelper As BOCatalogueDBHelper

Private Const m_sMaxOrderKeyValueName As String = "MaxOrderKey"


''''''''''''''''' Implementation ''''''''''''''''''''''

'
' Class initialisation
'

Private Sub Initialise()
    m_lBatchID = 0
    m_lCatalogueID = 0
    m_lEditsetID = 0
    Set m_objContext = GetObjectContext()
    Set m_oFWDBHelper = New FWDBHelper
    Set m_oBOCatalogueDBHelper = New BOCatalogueDBHelper
End Sub

Private Sub Terminate()
    Set m_objContext = Nothing
    Set m_oboParsedCatalogueEntry = Nothing
    Set m_oFWDBHelper = Nothing
    Set m_oBOCatalogueDBHelper = Nothing
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

Public Property Let CatalogueID(t_lCatalogueID As Long)
    m_lCatalogueID = t_lCatalogueID
End Property

Public Property Get CatalogueID() As Long
    CatalogueID = m_lCatalogueID
End Property

Public Property Let EditsetID(t_lEditSetID As Long)
    m_lEditsetID = t_lEditSetID
End Property

Public Property Get EditsetID() As Long
    EditsetID = m_lEditsetID
End Property

'Methods

Public Function Update() As Boolean
' Method to update a catalogue entry in an edit set with the
'
' Amendment history:
'   Version Name            Date            Comments
'   d1      D Johnston      19-Jul-2000     Created from VWImport
On Error GoTo ErrHandler

    Dim sSQL As String
    Dim eCurrentLevel As ECATALOGUELEVEL
    Dim colTemp As Collection
    Dim sSource As String
    Dim ofwResult As fwRecordList
    Dim rsmLink As Variant

    sSource = "BOEADUpdate.Update"
    
    ' Set default value
    Update = False
    
    ' Check that the ParsedCatalogueEntry has been set
    If IsNull(m_oboParsedCatalogueEntry) Then
        RaiseError sSource, PEGeneralInternalError, "ParsedCatalogueEntry not set before Update called"
    End If
    
    ' Validate the entries and clear variables
    CheckValues
    
    ' Get level number
    eCurrentLevel = GetCatalogueLevelEnum(m_oboParsedCatalogueEntry.Level)
    
    ' Log the update has started
    LogCurrentCatalogueUpdate
    
    ' Update the edit set tables with the
    sSQL = MainTableUpdateQuery(eCurrentLevel)
    If Not sSQL = "" Then m_oFWDBHelper.Execute (sSQL)
    
    LogMessage "Updated the Main table", EProfile.pComplete
    
    ' Update the Extension table
    sSQL = ExtensionTableUpdateQuery(eCurrentLevel)
    If Not sSQL = "" Then m_oFWDBHelper.Execute (sSQL)

    LogMessage "Updated the extenstion table", EProfile.pComplete

    'Insertion into accruals table
    sSQL = SingleFieldTableUpdateQuery(eCurrentLevel, ECATALOGUETABLETYPE.cttAccrualsTable)
    If Not sSQL = "" Then m_oFWDBHelper.Execute (sSQL)

    LogMessage "Updated the accruals table", EProfile.pComplete

    ' Update Admin history table
    sSQL = SingleFieldTableUpdateQuery(eCurrentLevel, ECATALOGUETABLETYPE.cttAdminHistoryTable)
    If Not sSQL = "" Then m_oFWDBHelper.Execute (sSQL)

    LogMessage "Updated the admin history table", EProfile.pComplete

    ' Update appraisal info table
    sSQL = SingleFieldTableUpdateQuery(eCurrentLevel, ECATALOGUETABLETYPE.cttAppDestructionTable)
    If Not sSQL = "" Then m_oFWDBHelper.Execute (sSQL)

    LogMessage "Updated the appraisal info table", EProfile.pComplete

    ' Update arrangement table
    sSQL = SingleFieldTableUpdateQuery(eCurrentLevel, ECATALOGUETABLETYPE.cttArrangementTable)
    If Not sSQL = "" Then m_oFWDBHelper.Execute (sSQL)

    LogMessage "Updated the arrangement table", EProfile.pComplete

    ' Update custodial history table
    sSQL = SingleFieldTableUpdateQuery(eCurrentLevel, ECATALOGUETABLETYPE.cttCustodialHistoryTable)
    If Not sSQL = "" Then m_oFWDBHelper.Execute (sSQL)

    LogMessage "Updated the custodial history table", EProfile.pComplete

    ' Update note table
    sSQL = SingleFieldTableUpdateQuery(eCurrentLevel, ECATALOGUETABLETYPE.cttNoteTable)
    If Not sSQL = "" Then m_oFWDBHelper.Execute (sSQL)

    LogMessage "Updated the note table", EProfile.pComplete

    ' Update scope and content table
    sSQL = SingleFieldTableUpdateQuery(eCurrentLevel, ECATALOGUETABLETYPE.cttScopeAndContentTable)
    If Not sSQL = "" Then m_oFWDBHelper.Execute (sSQL)

    LogMessage "Updated the scope and content table", EProfile.pComplete

    ' Update title table
    sSQL = SingleFieldTableUpdateQuery(eCurrentLevel, ECATALOGUETABLETYPE.cttTitleTable)
    If Not sSQL = "" Then m_oFWDBHelper.Execute (sSQL)

    'If the level is lettercode then insert data into the sorted title table
    If eCurrentLevel = ECATALOGUELEVEL.ctLetterCode Then
        sSQL = UpdateLettercodeSortedTitleQuery
        If Not sSQL = "" Then m_oFWDBHelper.Execute (sSQL)
    End If

    LogMessage "Updated the title table", EProfile.pComplete

    ' ---------------------------------
    ' --- START OF THE LINK SECTION ---
    ' ---------------------------------
    
    ' As the ParsedCatalogueEntry will have either the new, parsed values or the
    ' ones from the current version, it is OK to delete all the links now and reinsert
    ' them
    
    ' First thing to do is delete all the links that exist already
    sSQL = DeleteLinks(eCurrentLevel)
    If Not sSQL = "" Then m_oFWDBHelper.Execute (sSQL)

    LogMessage "Deleted all existing links from entry", EProfile.pComplete

    ' Then delete all the data elements that go with them
    sSQL = DeleteLinkDataElements(eCurrentLevel)
    If Not sSQL = "" Then m_oFWDBHelper.Execute (sSQL)

    LogMessage "Deleted all existing link data elements from entry", EProfile.pComplete

    ' --- Publication Note ---
    
    If PublicationNoteHasChanged(eCurrentLevel) Then
        ' Delete all the current SFA/Leaflet links that may exist
        sSQL = DeleteSFALinks(eCurrentLevel, ECATALOGUEPROPERTIES.cpPublicationNote)
        If Not sSQL = "" Then m_oFWDBHelper.Execute (sSQL)
        sSQL = DeleteLeafletLinks(eCurrentLevel, ECATALOGUEPROPERTIES.cpPublicationNote)
        If Not sSQL = "" Then m_oFWDBHelper.Execute (sSQL)
    End If
    
    ' Update of data for Publication note field
    sSQL = UpdateLinkDataElement(eCurrentLevel, ECATALOGUEPROPERTIES.cpPublicationNote)
    If Not sSQL = "" Then m_oFWDBHelper.Execute (sSQL)

    LogMessage "Updated Link data elements table for publication note", EProfile.pComplete
    
    ' --- Unpublished Finding Aids ---
    
    If UnpublishedFindingAidsHasChanged(eCurrentLevel) Then
        ' Delete all the current SFA/Leaflet links that may exist
        sSQL = DeleteSFALinks(eCurrentLevel, ECATALOGUEPROPERTIES.cpUnpublishedFindingAids)
        If Not sSQL = "" Then m_oFWDBHelper.Execute (sSQL)
        sSQL = DeleteLeafletLinks(eCurrentLevel, ECATALOGUEPROPERTIES.cpUnpublishedFindingAids)
        If Not sSQL = "" Then m_oFWDBHelper.Execute (sSQL)
    End If
        
    ' Update of data for unpublished finding aids field
    sSQL = UpdateLinkDataElement(eCurrentLevel, ECATALOGUEPROPERTIES.cpUnpublishedFindingAids)
    If Not sSQL = "" Then m_oFWDBHelper.Execute (sSQL)

    LogMessage "Updated Link data elements table for unpublished finding aids", EProfile.pComplete
    
    ' --- Related/Separated Material fields ---
    
    'Insertion of data for related materials field
    sSQL = UpdateLinkDataElement(eCurrentLevel, ECATALOGUEPROPERTIES.cpRelatedMaterials)
    If Not sSQL = "" Then m_oFWDBHelper.Execute (sSQL)

    LogMessage "Updated Link data elements table for related material", EProfile.pComplete

    ' Insertion of data for related materials field
    LogMessage "Inserting into the link table for related material links", EProfile.pComplete
    Dim colRelatedLinks As Collection
    Set colRelatedLinks = m_oboParsedCatalogueEntry.RelatedMaterialLinks
    For Each rsmLink In colRelatedLinks
        sSQL = UpdateLink(eCurrentLevel, ECATALOGUEPROPERTIES.cpRelatedMaterials, rsmLink)
        If Not sSQL = "" Then m_oFWDBHelper.Execute (sSQL)
    Next

    ' Insertion of data for seperated materials field
    LogMessage "Inserting into the link data elements table for separated material hanging text", EProfile.pComplete
    sSQL = UpdateLinkDataElement(eCurrentLevel, ECATALOGUEPROPERTIES.cpSeperatedMaterials)
    If Not sSQL = "" Then m_oFWDBHelper.Execute (sSQL)

    ' Add the actual links
    LogMessage "Inserting into the link table for separate material links", EProfile.pComplete
    Dim colSeparatedLinks As Collection
    Set colSeparatedLinks = m_oboParsedCatalogueEntry.SeparatedMaterialLinks
    For Each rsmLink In colSeparatedLinks
        sSQL = UpdateLink(eCurrentLevel, ECATALOGUEPROPERTIES.cpSeperatedMaterials, rsmLink)
        If Not sSQL = "" Then m_oFWDBHelper.Execute (sSQL)
    Next

    ' --- Copies Information ---

    ' Delete all the links that exist already
    sSQL = DeleteIndexLinks(eCurrentLevel, ECATALOGUETABLETYPE.cttCorporateBody, ECATALOGUEPROPERTIES.cpCopiesInformation)
    If Not sSQL = "" Then m_oFWDBHelper.Execute (sSQL)

    LogMessage "Deleted all existing copies info links from entry", EProfile.pComplete

    ' Update of data for Copies infomation TEXT ONLY field
    sSQL = UpdateLinkDataElement(eCurrentLevel, ECATALOGUEPROPERTIES.cpCopiesInformation)
    If Not sSQL = "" Then m_oFWDBHelper.Execute (sSQL)

    LogMessage "Updated link data elements table for copies info", EProfile.pComplete

    'Copies Info Corporate body ID description Insert
    SaveIDDescriptionInCollection m_oboParsedCatalogueEntry.CopiesCorpBodies, _
                            eCurrentLevel, _
                            ECATALOGUETABLETYPE.cttCorporateBody, _
                            ECATALOGUEPROPERTIES.cpCopiesInformation, _
                            True

    LogMessage "Updated pc corporate body table for copies info", EProfile.pComplete


    ' --- Immediate Source of Acquisition ---
    
    ' Delete all the links that exist already
    sSQL = DeleteIndexLinks(eCurrentLevel, ECATALOGUETABLETYPE.cttCorporateBody, ECATALOGUEPROPERTIES.cpImmediateSourceOfAcquisition)
    If Not sSQL = "" Then m_oFWDBHelper.Execute (sSQL)

    LogMessage "Deleted all existing immediate source of accq corp body links from entry", EProfile.pComplete

    sSQL = DeleteIndexLinks(eCurrentLevel, ECATALOGUETABLETYPE.cttPerson, ECATALOGUEPROPERTIES.cpImmediateSourceOfAcquisition)
    If Not sSQL = "" Then m_oFWDBHelper.Execute (sSQL)

    LogMessage "Deleted all existing immediate source of accq person links from entry", EProfile.pComplete

    ' Update of data for immediate source of accq TEXT ONLY field
    sSQL = UpdateLinkDataElement(eCurrentLevel, ECATALOGUEPROPERTIES.cpImmediateSourceOfAcquisition)
    If Not sSQL = "" Then m_oFWDBHelper.Execute (sSQL)

    LogMessage "Updated Link data elements table for immediate source of accq ", EProfile.pComplete

    'Immediate Source Of acquisition Corporate body ID description Insert
    SaveIDDescriptionInCollection m_oboParsedCatalogueEntry.SourceAcquisitionCorpBodies, _
                            eCurrentLevel, _
                            ECATALOGUETABLETYPE.cttCorporateBody, _
                            ECATALOGUEPROPERTIES.cpImmediateSourceOfAcquisition, _
                            True

    LogMessage "Updated pc corporate body table for Immediate source of accq", EProfile.pComplete

    'Immediate Source Of acquisition Person ID description Insert
    SaveIDDescriptionInCollection m_oboParsedCatalogueEntry.SourceAcquisitionPersons, _
                            eCurrentLevel, _
                            ECATALOGUETABLETYPE.cttPerson, _
                            ECATALOGUEPROPERTIES.cpImmediateSourceOfAcquisition, _
                            True
    LogMessage "Updated pc person table for Immediate source of accq", EProfile.pComplete


    ' --- Location of Originals ---

    ' Delete all the links that exist already
    sSQL = DeleteIndexLinks(eCurrentLevel, ECATALOGUETABLETYPE.cttCorporateBody, ECATALOGUEPROPERTIES.cpLocationOfOriginals)
    If Not sSQL = "" Then m_oFWDBHelper.Execute (sSQL)

    LogMessage "Deleted all existing location of originals links from entry", EProfile.pComplete

    'Location of Originals Corporate body ID description Insert
    SaveIDDescriptionInCollection m_oboParsedCatalogueEntry.LocationOriginalsCorpBodies, _
                            eCurrentLevel, _
                            ECATALOGUETABLETYPE.cttCorporateBody, _
                            ECATALOGUEPROPERTIES.cpLocationOfOriginals, _
                            True

    LogMessage "Updated pc corporate body table for location of originals", EProfile.pComplete


    ' --- Creator Names ---

    ' Delete all the links that exist already
    sSQL = DeleteIndexLinks(eCurrentLevel, ECATALOGUETABLETYPE.cttCorporateBody, ECATALOGUEPROPERTIES.cpCreatorNames)
    If Not sSQL = "" Then m_oFWDBHelper.Execute (sSQL)

    LogMessage "Deleted all existing creator names corp body links from entry", EProfile.pComplete

    sSQL = DeleteIndexLinks(eCurrentLevel, ECATALOGUETABLETYPE.cttPerson, ECATALOGUEPROPERTIES.cpCreatorNames)
    If Not sSQL = "" Then m_oFWDBHelper.Execute (sSQL)

    LogMessage "Deleted all existing creator names person links from entry", EProfile.pComplete

    'Creator names Corporate body ID ONLY Insert
    SaveIDInCollection m_oboParsedCatalogueEntry.CreatorNamesCorpBodies, _
                            eCurrentLevel, _
                            ECATALOGUETABLETYPE.cttCorporateBody, _
                            ECATALOGUEPROPERTIES.cpCreatorNames, _
                            False

    LogMessage "Updated pc corporate body table for creator names", EProfile.pComplete

    'Creator names Person ID ONLY Insert
    SaveIDInCollection m_oboParsedCatalogueEntry.CreatorNamesPersons, _
                            eCurrentLevel, _
                            ECATALOGUETABLETYPE.cttPerson, _
                            ECATALOGUEPROPERTIES.cpCreatorNames, _
                            False

    ' --- Place of Deposit ---
    
    ' Delete all the links that exist already
    sSQL = DeleteIndexLinks(eCurrentLevel, ECATALOGUETABLETYPE.cttCorporateBody, ECATALOGUEPROPERTIES.cpPlaceOfDeposit)
    If Not sSQL = "" Then m_oFWDBHelper.Execute (sSQL)

    LogMessage "Deleted all existing creator names corp body links from entry", EProfile.pComplete

    LogMessage "Inserting into the pc corporate body table for place of deposit", EProfile.pComplete

    'Creator names Person ID ONLY Insert
    SaveIDInCollection m_oboParsedCatalogueEntry.PlaceDepositCorpBodies, _
                            eCurrentLevel, _
                            ECATALOGUETABLETYPE.cttCorporateBody, _
                            ECATALOGUEPROPERTIES.cpPlaceOfDeposit, _
                            False

    LogMessage "Updated pc person table for creator names", EProfile.pComplete

    ' -------------------
    ' --- Index Links ---
    ' -------------------
    
    ' Delete all the links that exist already
    sSQL = DeleteIndexLinks(eCurrentLevel, ECATALOGUETABLETYPE.cttCorporateBody, ECATALOGUEPROPERTIES.cpIndexTermCorporateName)
    If Not sSQL = "" Then m_oFWDBHelper.Execute (sSQL)

    LogMessage "Deleted all existing corporate body links from entry", EProfile.pComplete

    'Corporate body ID ONLY Insert
    SaveIDInCollection m_oboParsedCatalogueEntry.CorpBodyIndex, _
                            eCurrentLevel, _
                            ECATALOGUETABLETYPE.cttCorporateBody, _
                            ECATALOGUEPROPERTIES.cpIndexTermCorporateName, _
                            False

    LogMessage "Updated pc corporate body table for Index term corporate body", EProfile.pComplete

    ' Delete all the links that exist already
    sSQL = DeleteIndexLinks(eCurrentLevel, ECATALOGUETABLETYPE.cttPerson, ECATALOGUEPROPERTIES.cpIndexTermPersonName)
    If Not sSQL = "" Then m_oFWDBHelper.Execute (sSQL)

    LogMessage "Deleted all existing person links from entry", EProfile.pComplete

    'Person ID ONLY Insert
    SaveIDInCollection m_oboParsedCatalogueEntry.PersonIndex, _
                            eCurrentLevel, _
                            ECATALOGUETABLETYPE.cttPerson, _
                            ECATALOGUEPROPERTIES.cpIndexTermPersonName, _
                            False

    LogMessage "Updated pc person table for Index term person", EProfile.pComplete

    ' Delete all the links that exist already
    sSQL = DeleteIndexLinks(eCurrentLevel, ECATALOGUETABLETYPE.cttPlace, ECATALOGUEPROPERTIES.cpIndexTermPlaceName)
    If Not sSQL = "" Then m_oFWDBHelper.Execute (sSQL)

    LogMessage "Deleted all existing place links from entry", EProfile.pComplete

    'Place ID ONLY Insert
    SaveIDInCollection m_oboParsedCatalogueEntry.PlaceIndex, _
                            eCurrentLevel, _
                            ECATALOGUETABLETYPE.cttPlace, _
                            ECATALOGUEPROPERTIES.cpIndexTermPlaceName, _
                            False

    LogMessage "Updated pc place table for Index term Place", EProfile.pComplete

    ' Delete all the links that exist already
    sSQL = DeleteIndexLinks(eCurrentLevel, ECATALOGUETABLETYPE.cttSubject, ECATALOGUEPROPERTIES.cpIndexTermSubject)
    If Not sSQL = "" Then m_oFWDBHelper.Execute (sSQL)

    LogMessage "Deleted all existing subject links from entry", EProfile.pComplete

    'Subject ID ONLY Insert
    SaveIDInCollection m_oboParsedCatalogueEntry.SubjectIndex, _
                            eCurrentLevel, _
                            ECATALOGUETABLETYPE.cttSubject, _
                            ECATALOGUEPROPERTIES.cpIndexTermSubject, _
                            False
    LogMessage "Updated pc place table for Index term Subject", EProfile.pComplete

    ' Set the entry type to modified
    m_oBOCatalogueDBHelper.SetCatalogueStatusAsModified _
        GetTableName(cttExtensionTable, SECeditorial, eCurrentLevel), _
        gk_sDBCOLEditSetID & "=" & m_lEditsetID & " AND " & _
        GetColumnName(cctIDColumn, eCurrentLevel) & "=" & m_lCatalogueID
    
    ' Return success
    Update = True
    
    Exit Function
    
ErrHandler:
    ConvertError sSource
End Function

Private Sub SaveIDDescriptionInCollection(t_colData As Collection, _
                            t_eCatalogueLevel As ECATALOGUELEVEL, _
                            t_eTableType As ECATALOGUETABLETYPE, _
                            t_eCatalogueProperty As ECATALOGUEPROPERTIES, _
                            t_bHasDescription As Boolean)
'Method to iterate throught a collection of ID, Description pairs
'  Inserting then into the appropriate table
'
' Amendment history:
'   Version Name            Date            Comments
'   d1      D Johnston      19-Jul-2000     Created from BOEADImport
On Error GoTo ErrHandle

    Dim typPair As GenericLinkRecord
    Dim sSQL As String
    Dim sSource As String
    Dim i As Integer

    sSource = "BOEADUpdate.SaveIDDescriptionInCollection"
                            
    For i = 1 To t_colData.Count
        typPair = t_colData.Item(i)
        
        sSQL = InsertForIndexTerms(t_eTableType, _
                        t_eCatalogueLevel, t_eCatalogueProperty, _
                        typPair.lID, _
                        StringToNullOrString(typPair.sOptionalDescription))
        
        m_oFWDBHelper.Execute (sSQL)
        
    Next
                            
    Exit Sub
    
ErrHandle:
    ConvertError sSource
End Sub

Private Sub SaveIDInCollection(t_colData As Collection, _
                            t_eCatalogueLevel As ECATALOGUELEVEL, _
                            t_eTableType As ECATALOGUETABLETYPE, _
                            t_eCatalogueProperty As ECATALOGUEPROPERTIES, _
                            t_bHasDescription As Boolean)
'Method to iterate throught a collection of ID values
'  Inserting then into the appropriate table
'
' Amendment history:
'   Version Name            Date            Comments
'   d1      D Johnston      19-Jul-2000     Created from BOEADImport
On Error GoTo ErrHandler

    Dim lIDInCollection As Long
    Dim i As Integer
    Dim sSQL As String
    Dim sSource As String

    sSource = "BOEADUpdate.SaveIDInCollection"
                            
    For i = 1 To t_colData.Count
        
        lIDInCollection = t_colData.Item(i)
        
        sSQL = InsertForIndexTerms(t_eTableType, _
                        t_eCatalogueLevel, t_eCatalogueProperty, _
                        lIDInCollection, _
                        Null)
        
        m_oFWDBHelper.Execute (sSQL)
        
    Next
    
    Exit Sub
ErrHandler:
    ConvertError sSource
End Sub


Private Function CheckValues()
' Checks initial setup of the BO object before update commences
'
' Amendment history:
'   Version Name            Date            Comments
'   d1      D Johnston      19-Jul-2000     Created from BOEADImport
On Error GoTo ErrHandler

    Dim sSource As String

    sSource = "BOEADUpdate.CheckValues"

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

Private Function MainTableUpdateQuery(t_eLevel As ECATALOGUELEVEL) As String
' SQL generator to update the main table.  Only updates the columns that are in
' the EAD file and require to be changed.
'
' Amendment history:
'   Version Name            Date            Comments
'   d1      D Johnston      19-Jul-2000     Created from BOEADImport
'   d4      A Manic         28-Feb-2001     Updates start / end dates

    Dim sSQL As String, sWhereSQL As String
    Dim sTablename As String
    Dim sIDColumnName As String
    
    ' Get the main table name for the edit set
    sTablename = GetTableName(ECATALOGUETABLETYPE.cttmainTable, _
                    ESectionType.SECeditorial, t_eLevel)
    
    ' Get the name of the ID column at this level
    sIDColumnName = GetColumnName(ECATALOGUECOLUMNTYPE.cctIDColumn, t_eLevel)
    
    ' Set the WHERE clause
    sWhereSQL = " WHERE " & sIDColumnName & " = " & m_lCatalogueID
    
    Select Case t_eLevel
    
    Case ECATALOGUELEVEL.ctLetterCode
                        
            ' Nothing to do for lettercodes as we can only update Closure Status - this
            ' is not applicable above piece and isn't even in the EAD file.
            sSQL = ""

    Case ECATALOGUELEVEL.ctDivision
    
            sSQL = "UPDATE " & sTablename & " SET " & _
                        gk_sDBCOLBatchID & " = " & m_lBatchID & _
                        sWhereSQL
            
    Case ECATALOGUELEVEL.ctClass
            
            ' d4 Added gk_sDBCOLFirstDate, gk_sDBCOLLastDate
            sSQL = "UPDATE " & sTablename & " SET " & _
                        gk_sDBCOLDateText & " = " & QuoteValueWithQuotes(m_oboParsedCatalogueEntry.DateText) & "," & _
                        gk_sDBCOLFirstDate & " = " & NullifBlank(m_oboParsedCatalogueEntry.StartDate) & "," & _
                        gk_sDBCOLLastDate & " = " & NullifBlank(m_oboParsedCatalogueEntry.EndDate) & "," & _
                        gk_sDBCOLBatchID & " = " & m_lBatchID & _
                        sWhereSQL
                        
    Case ECATALOGUELEVEL.ctHeaders, ECATALOGUELEVEL.ctSubheaders
            
            sSQL = "UPDATE " & sTablename & " SET " & _
                        gk_sDBCOLDateText & " = " & QuoteValueWithQuotes(m_oboParsedCatalogueEntry.DateText) & "," & _
                        gk_sDBCOLBatchID & " = " & m_lBatchID & _
                        sWhereSQL
                        

    Case ECATALOGUELEVEL.ctPiece, ECATALOGUELEVEL.ctItem
            
            ' d4 Added gk_sDBCOLFirstDate, gk_sDBCOLLastDate
            sSQL = "UPDATE " & sTablename & " SET " & _
                        gk_sDBCOLDateText & " = " & QuoteValueWithQuotes(m_oboParsedCatalogueEntry.DateText) & "," & _
                        gk_sDBCOLFirstDate & " = " & NullifBlank(m_oboParsedCatalogueEntry.StartDate) & "," & _
                        gk_sDBCOLLastDate & " = " & NullifBlank(m_oboParsedCatalogueEntry.EndDate) & "," & _
                        gk_sDBCOLLanguageID & " = " & NullifZero(CLng(m_oboParsedCatalogueEntry.LanguageID)) & ", " & _
                        gk_sDBCOLBatchID & " = " & m_lBatchID & _
                        sWhereSQL

    End Select

    MainTableUpdateQuery = sSQL
    
End Function


Private Function ExtensionTableUpdateQuery(t_eLevel As ECATALOGUELEVEL) As String
' SQL generator to update the extension tables.  Only updates the fields that are
' in the EAD file and can be changed.
'
' Amendment history:
'   Version Name            Date            Comments
'   d1      D Johnston      19-Jul-2000     Created from BOEADImport
'   d4      A Manic         28-Feb-2001     Updates start / end dates

    Dim sSQL As String, sWhereSQL As String
    Dim sTablename As String
    Dim sIDColumnName As String

    ' Get the name of the edit set extension table
    sTablename = GetTableName(ECATALOGUETABLETYPE.cttExtensionTable, _
                    ESectionType.SECeditorial, t_eLevel)

    ' Get the ID column name for this level
    sIDColumnName = GetColumnName(ECATALOGUECOLUMNTYPE.cctIDColumn, t_eLevel)

    ' Set the WHERE clause
    sWhereSQL = " WHERE " & sIDColumnName & " = " & m_lCatalogueID
    
    Select Case t_eLevel

    Case ECATALOGUELEVEL.ctLetterCode

            ' d4 Added gk_sDBCOLCoveringFromDate, gk_sDBCOLCoveringToDate
            sSQL = "UPDATE " & sTablename & " SET " & _
                        gk_sDBCOLAccessConditionID & " = " & NullifZero(CLng(m_oboParsedCatalogueEntry.AccessConditionID)) & ", " & _
                        gk_sDBCOLLanguageID & " = " & NullifZero(CLng(m_oboParsedCatalogueEntry.LanguageID)) & ", " & _
                        gk_sDBCOLAccumulationdates & " = " & QuoteValueWithQuotes(m_oboParsedCatalogueEntry.AccumulationDates) & ", " & _
                        gk_sDBCOLCoveringDates & " = " & QuoteValueWithQuotes(m_oboParsedCatalogueEntry.DateText) & ", " & _
                        gk_sDBCOLCoveringFromDate & " = " & NullifBlank(m_oboParsedCatalogueEntry.StartDate) & ", " & _
                        gk_sDBCOLCoveringToDate & " = " & NullifBlank(m_oboParsedCatalogueEntry.EndDate) & ", " & _
                        gk_sDBCOLLegalStatusID & " = " & NullifZero(CLng(m_oboParsedCatalogueEntry.LegalStatusID)) & ", " & _
                        gk_sDBCOLMapDesignation & " = " & QuoteValueWithQuotes(m_oboParsedCatalogueEntry.MapDesignation) & "," & _
                        gk_sDBCOLPhysicalCondition & " = " & QuoteValueWithQuotes(m_oboParsedCatalogueEntry.PhysdescCondition) & "," & _
                        gk_sDBCOLPhysicalDescriptionRecordQuantity & " = " & IIf(m_oboParsedCatalogueEntry.PhysdescExtentNull, "null", m_oboParsedCatalogueEntry.PhysdescExtent) & "," & _
                        gk_sDBCOLPhysicalDescriptionFormID & " = " & NullifZero(CLng(m_oboParsedCatalogueEntry.PhysdescFormID)) & "," & _
                        gk_sDBCOLPhysicalDescriptionDimension & " = " & QuoteValueWithQuotes(m_oboParsedCatalogueEntry.PhysdescDimensions) & "," & _
                        gk_sDBCOLRestrictionsOnUse & " = " & QuoteValueWithQuotes(m_oboParsedCatalogueEntry.RestrictionsOnUse) & "," & _
                        gk_sDBCOLMapScaleNumber & " = " & IIf(m_oboParsedCatalogueEntry.ScaleNumberNull, "null", m_oboParsedCatalogueEntry.ScaleNumber) & "," & _
                        gk_sDBCOLFormerDepartmentReference & " = " & QuoteValueWithQuotes(m_oboParsedCatalogueEntry.ReferenceDepartment) & "," & _
                        gk_sDBCOLFormerPROReference & " = " & QuoteValueWithQuotes(m_oboParsedCatalogueEntry.ReferencePRO) & "," & _
                        gk_sDBCOLBatchID & " = " & m_lBatchID & _
                        sWhereSQL
                    
     Case ECATALOGUELEVEL.ctDivision

            ' d4 Added gk_sDBCOLCoveringFromDate, gk_sDBCOLCoveringToDate
            sSQL = "UPDATE " & sTablename & " SET " & _
                        gk_sDBCOLAccessConditionID & " = " & NullifZero(CLng(m_oboParsedCatalogueEntry.AccessConditionID)) & ", " & _
                        gk_sDBCOLLanguageID & " = " & NullifZero(CLng(m_oboParsedCatalogueEntry.LanguageID)) & ", " & _
                        gk_sDBCOLAccumulationdates & " = " & QuoteValueWithQuotes(m_oboParsedCatalogueEntry.AccumulationDates) & ", " & _
                        gk_sDBCOLCoveringDates & " = " & QuoteValueWithQuotes(m_oboParsedCatalogueEntry.DateText) & ", " & _
                        gk_sDBCOLCoveringFromDate & " = " & NullifBlank(m_oboParsedCatalogueEntry.StartDate) & ", " & _
                        gk_sDBCOLCoveringToDate & " = " & NullifBlank(m_oboParsedCatalogueEntry.EndDate) & ", " & _
                        gk_sDBCOLLegalStatusID & " = " & NullifZero(CLng(m_oboParsedCatalogueEntry.LegalStatusID)) & ", " & _
                        gk_sDBCOLMapDesignation & " = " & QuoteValueWithQuotes(m_oboParsedCatalogueEntry.MapDesignation) & "," & _
                        gk_sDBCOLPhysicalCondition & " = " & QuoteValueWithQuotes(m_oboParsedCatalogueEntry.PhysdescCondition) & "," & _
                        gk_sDBCOLPhysicalDescriptionRecordQuantity & " = " & IIf(m_oboParsedCatalogueEntry.PhysdescExtentNull, "null", m_oboParsedCatalogueEntry.PhysdescExtent) & "," & _
                        gk_sDBCOLPhysicalDescriptionFormID & " = " & NullifZero(CLng(m_oboParsedCatalogueEntry.PhysdescFormID)) & "," & _
                        gk_sDBCOLPhysicalDescriptionDimension & " = " & QuoteValueWithQuotes(m_oboParsedCatalogueEntry.PhysdescDimensions) & "," & _
                        gk_sDBCOLRestrictionsOnUse & " = " & QuoteValueWithQuotes(m_oboParsedCatalogueEntry.RestrictionsOnUse) & "," & _
                        gk_sDBCOLMapScaleNumber & " = " & IIf(m_oboParsedCatalogueEntry.ScaleNumberNull, "null", m_oboParsedCatalogueEntry.ScaleNumber) & "," & _
                        gk_sDBCOLFormerDepartmentReference & " = " & QuoteValueWithQuotes(m_oboParsedCatalogueEntry.ReferenceDepartment) & "," & _
                        gk_sDBCOLFormerPROReference & " = " & QuoteValueWithQuotes(m_oboParsedCatalogueEntry.ReferencePRO) & _
                        sWhereSQL

    Case ECATALOGUELEVEL.ctClass

            sSQL = "UPDATE " & sTablename & " SET " & _
                        gk_sDBCOLAccessConditionID & " = " & NullifZero(CLng(m_oboParsedCatalogueEntry.AccessConditionID)) & ", " & _
                        gk_sDBCOLLegalStatusID & " = " & NullifZero(CLng(m_oboParsedCatalogueEntry.LegalStatusID)) & ", " & _
                        gk_sDBCOLLanguageID & " = " & NullifZero(CLng(m_oboParsedCatalogueEntry.LanguageID)) & ", " & _
                        gk_sDBCOLAccumulationdates & " = " & QuoteValueWithQuotes(m_oboParsedCatalogueEntry.AccumulationDates) & ", " & _
                        gk_sDBCOLPhysicalDescriptionFormID & " = " & NullifZero(CLng(m_oboParsedCatalogueEntry.PhysdescFormID)) & "," & _
                        gk_sDBCOLMapDesignation & " = " & QuoteValueWithQuotes(m_oboParsedCatalogueEntry.MapDesignation) & "," & _
                        gk_sDBCOLPhysicalDescriptionRecordQuantity & " = " & IIf(m_oboParsedCatalogueEntry.PhysdescExtentNull, "null", m_oboParsedCatalogueEntry.PhysdescExtent) & "," & _
                        gk_sDBCOLPhysicalDescriptionDimension & " = " & QuoteValueWithQuotes(m_oboParsedCatalogueEntry.PhysdescDimensions) & "," & _
                        gk_sDBCOLRestrictionsOnUse & " = " & QuoteValueWithQuotes(m_oboParsedCatalogueEntry.RestrictionsOnUse) & "," & _
                        gk_sDBCOLMapScaleNumber & " = " & IIf(m_oboParsedCatalogueEntry.ScaleNumberNull, "null", m_oboParsedCatalogueEntry.ScaleNumber) & "," & _
                        gk_sDBCOLFormerDepartmentReference & " = " & QuoteValueWithQuotes(m_oboParsedCatalogueEntry.ReferenceDepartment) & "," & _
                        gk_sDBCOLFormerPROReference & " = " & QuoteValueWithQuotes(m_oboParsedCatalogueEntry.ReferencePRO) & "," & _
                        gk_sDBCOLPhysicalCondition & " = " & QuoteValueWithQuotes(m_oboParsedCatalogueEntry.PhysdescCondition) & _
                        sWhereSQL
                        
    Case ECATALOGUELEVEL.ctHeaders
    
            ' d4 Added gk_sDBCOLCoveringFromDate, gk_sDBCOLCoveringToDate
            sSQL = "UPDATE " & sTablename & " SET " & _
                        gk_sDBCOLAccessConditionID & " = " & NullifZero(CLng(m_oboParsedCatalogueEntry.AccessConditionID)) & ", " & _
                        gk_sDBCOLLegalStatusID & " = " & NullifZero(CLng(m_oboParsedCatalogueEntry.LegalStatusID)) & ", " & _
                        gk_sDBCOLLanguageID & " = " & NullifZero(CLng(m_oboParsedCatalogueEntry.LanguageID)) & ", " & _
                        gk_sDBCOLAccumulationdates & " = " & QuoteValueWithQuotes(m_oboParsedCatalogueEntry.AccumulationDates) & ", " & _
                        gk_sDBCOLPhysicalDescriptionFormID & " = " & NullifZero(CLng(m_oboParsedCatalogueEntry.PhysdescFormID)) & "," & _
                        gk_sDBCOLMapDesignation & " = " & QuoteValueWithQuotes(m_oboParsedCatalogueEntry.MapDesignation) & "," & _
                        gk_sDBCOLPhysicalDescriptionRecordQuantity & " = " & IIf(m_oboParsedCatalogueEntry.PhysdescExtentNull, "null", m_oboParsedCatalogueEntry.PhysdescExtent) & "," & _
                        gk_sDBCOLPhysicalDescriptionDimension & " = " & QuoteValueWithQuotes(m_oboParsedCatalogueEntry.PhysdescDimensions) & "," & _
                        gk_sDBCOLRestrictionsOnUse & " = " & QuoteValueWithQuotes(m_oboParsedCatalogueEntry.RestrictionsOnUse) & "," & _
                        gk_sDBCOLMapScaleNumber & " = " & IIf(m_oboParsedCatalogueEntry.ScaleNumberNull, "null", m_oboParsedCatalogueEntry.ScaleNumber) & "," & _
                        gk_sDBCOLFormerDepartmentReference & " = " & QuoteValueWithQuotes(m_oboParsedCatalogueEntry.ReferenceDepartment) & "," & _
                        gk_sDBCOLFormerPROReference & " = " & QuoteValueWithQuotes(m_oboParsedCatalogueEntry.ReferencePRO) & "," & _
                        gk_sDBCOLCoveringFromDate & " = " & NullifBlank(m_oboParsedCatalogueEntry.StartDate) & ", " & _
                        gk_sDBCOLCoveringToDate & " = " & NullifBlank(m_oboParsedCatalogueEntry.EndDate) & ", " & _
                        gk_sDBCOLPhysicalCondition & " = " & QuoteValueWithQuotes(m_oboParsedCatalogueEntry.PhysdescCondition) & _
                        sWhereSQL
                        
    Case ECATALOGUELEVEL.ctSubheaders
    
            ' d4 Added gk_sDBCOLCoveringFromDate, gk_sDBCOLCoveringToDate
            sSQL = "UPDATE " & sTablename & " SET " & _
                        gk_sDBCOLAccessConditionID & " = " & NullifZero(CLng(m_oboParsedCatalogueEntry.AccessConditionID)) & ", " & _
                        gk_sDBCOLLegalStatusID & " = " & NullifZero(CLng(m_oboParsedCatalogueEntry.LegalStatusID)) & ", " & _
                        gk_sDBCOLLanguageID & " = " & NullifZero(CLng(m_oboParsedCatalogueEntry.LanguageID)) & ", " & _
                        gk_sDBCOLAccumulationdates & " = " & QuoteValueWithQuotes(m_oboParsedCatalogueEntry.AccumulationDates) & ", " & _
                        gk_sDBCOLPhysicalDescriptionFormID & " = " & NullifZero(CLng(m_oboParsedCatalogueEntry.PhysdescFormID)) & "," & _
                        gk_sDBCOLMapDesignation & " = " & QuoteValueWithQuotes(m_oboParsedCatalogueEntry.MapDesignation) & "," & _
                        gk_sDBCOLPhysicalDescriptionRecordQuantity & " = " & IIf(m_oboParsedCatalogueEntry.PhysdescExtentNull, "null", m_oboParsedCatalogueEntry.PhysdescExtent) & "," & _
                        gk_sDBCOLPhysicalDescriptionDimension & " = " & QuoteValueWithQuotes(m_oboParsedCatalogueEntry.PhysdescDimensions) & "," & _
                        gk_sDBCOLRestrictionsOnUse & " = " & QuoteValueWithQuotes(m_oboParsedCatalogueEntry.RestrictionsOnUse) & "," & _
                        gk_sDBCOLMapScaleNumber & " = " & IIf(m_oboParsedCatalogueEntry.ScaleNumberNull, "null", m_oboParsedCatalogueEntry.ScaleNumber) & "," & _
                        gk_sDBCOLFormerDepartmentReference & " = " & QuoteValueWithQuotes(m_oboParsedCatalogueEntry.ReferenceDepartment) & "," & _
                        gk_sDBCOLFormerPROReference & " = " & QuoteValueWithQuotes(m_oboParsedCatalogueEntry.ReferencePRO) & "," & _
                        gk_sDBCOLCoveringFromDate & " = " & NullifBlank(m_oboParsedCatalogueEntry.StartDate) & ", " & _
                        gk_sDBCOLCoveringToDate & " = " & NullifBlank(m_oboParsedCatalogueEntry.EndDate) & ", " & _
                        gk_sDBCOLPhysicalCondition & " = " & QuoteValueWithQuotes(m_oboParsedCatalogueEntry.PhysdescCondition) & _
                        sWhereSQL
    
    Case ECATALOGUELEVEL.ctPiece, ECATALOGUELEVEL.ctItem

            sSQL = "UPDATE " & sTablename & " SET " & _
                        gk_sDBCOLAccessConditionID & " = " & NullifZero(CLng(m_oboParsedCatalogueEntry.AccessConditionID)) & ", " & _
                        gk_sDBCOLLegalStatusID & " = " & NullifZero(CLng(m_oboParsedCatalogueEntry.LegalStatusID)) & ", " & _
                        gk_sDBCOLAccumulationdates & " = " & QuoteValueWithQuotes(m_oboParsedCatalogueEntry.AccumulationDates) & ", " & _
                        gk_sDBCOLPhysicalDescriptionFormID & " = " & NullifZero(CLng(m_oboParsedCatalogueEntry.PhysdescFormID)) & "," & _
                        gk_sDBCOLMapDesignation & " = " & QuoteValueWithQuotes(m_oboParsedCatalogueEntry.MapDesignation) & "," & _
                        gk_sDBCOLPhysicalDescriptionRecordQuantity & " = " & IIf(m_oboParsedCatalogueEntry.PhysdescExtentNull, "null", m_oboParsedCatalogueEntry.PhysdescExtent) & "," & _
                        gk_sDBCOLPhysicalDescriptionDimension & " = " & QuoteValueWithQuotes(m_oboParsedCatalogueEntry.PhysdescDimensions) & "," & _
                        gk_sDBCOLRestrictionsOnUse & " = " & QuoteValueWithQuotes(m_oboParsedCatalogueEntry.RestrictionsOnUse) & "," & _
                        gk_sDBCOLMapScaleNumber & " = " & IIf(m_oboParsedCatalogueEntry.ScaleNumberNull, "null", m_oboParsedCatalogueEntry.ScaleNumber) & "," & _
                        gk_sDBCOLFormerDepartmentReference & " = " & QuoteValueWithQuotes(m_oboParsedCatalogueEntry.ReferenceDepartment) & "," & _
                        gk_sDBCOLFormerPROReference & " = " & QuoteValueWithQuotes(m_oboParsedCatalogueEntry.ReferencePRO) & "," & _
                        gk_sDBCOLPhysicalCondition & " = " & QuoteValueWithQuotes(m_oboParsedCatalogueEntry.PhysdescCondition) & _
                        sWhereSQL

    End Select

    ExtensionTableUpdateQuery = sSQL

End Function


Private Function SingleFieldTableUpdateQuery(t_eLevel As ECATALOGUELEVEL, _
                                                t_eTableType As ECATALOGUETABLETYPE) As String
' Returns update string for one of the single, large field tables; eg Accruals or Scope/Content
'
' Amendment history:
'   Version Name            Date            Comments
'   d1      D Johnston      19-Jul-2000     Created from BOEADImport
    
    Dim sSQL As String
    Dim sTablename As String
    Dim sIDColumnName As String
    
    sTablename = GetTableName(t_eTableType, _
                    ESectionType.SECeditorial, t_eLevel)
                    
    sIDColumnName = GetColumnName(ECATALOGUECOLUMNTYPE.cctIDColumn, t_eLevel)

    sSQL = "UPDATE " & sTablename & " SET " & _
                GetSingleFieldTableFieldName(t_eTableType) & " = " & QuoteValueWithQuotes(GetSingleFieldTableFieldValue(t_eTableType)) & _
                " WHERE " & sIDColumnName & " = " & m_lCatalogueID

    SingleFieldTableUpdateQuery = sSQL

End Function


Private Function UpdateLettercodeSortedTitleQuery() As String
' Returns update string for sorted lettercode title table
'
' Amendment history:
'   Version Name            Date            Comments
'   d1      D Johnston      19-Jul-2000     Created from BOEADImport
    
    Dim sSQL As String
    Dim sData As String
    Dim sIDColumnName As String
    
    sData = GetSortedTitle(m_oboParsedCatalogueEntry.Title)
                    
    sIDColumnName = GetColumnName(ECATALOGUECOLUMNTYPE.cctIDColumn, ECATALOGUELEVEL.ctLetterCode)

    sSQL = "UPDATE " & gk_sDBTABLEEditSetLetterCodeSortedTitle & " SET " & _
                gk_sDBCOLSortTitle & " = " & QuoteValueWithQuotes(sData) & _
                " WHERE " & sIDColumnName & " = " & m_lCatalogueID

    UpdateLettercodeSortedTitleQuery = sSQL

End Function


Private Function UpdateLinkDataElement(t_eLevel As ECATALOGUELEVEL, t_eProperty As ECATALOGUEPROPERTIES) As String
' Creates the insert string for the PC_link_data_element table
'
' Amendment history:
'   Version Name            Date            Comments
'   d1      D Johnston      19-Jul-2000     Created from BOEADImport
   
    Dim sSQL As String
    
    sSQL = "INSERT INTO " & gk_sDBTABLEEditsetPCLinkDataElement & " ( " & _
                gk_sDBCOLPropertyID & "," & _
                gk_sDBCOLCatalogueLevel & "," & _
                gk_sDBCOLCatalogueID & "," & _
                gk_sDBCOLDescription & "," & _
                gk_sDBCOLEditSetID & _
                ")"

    sSQL = sSQL & " VALUES ( " & _
                t_eProperty & "," & _
                t_eLevel & "," & _
                m_lCatalogueID & "," & _
                QuoteValueWithQuotes(GetLinkTableDescriptionFieldValue(t_eProperty)) & "," & _
                m_lEditsetID & _
                ")"

    UpdateLinkDataElement = sSQL

End Function

Private Function DeleteLinkDataElements(t_eLevel As ECATALOGUELEVEL) As String
' Creates the delete string for all the entries in the PC_link_data_element table
'
' Amendment history:
'   Version Name            Date            Comments
'   d1      D Johnston      19-Jul-2000     Created from BOEADImport
    
    Dim sSQL As String
    
    sSQL = "DELETE " & gk_sDBTABLEEditsetPCLinkDataElement & _
                " WHERE " & gk_sDBCOLCatalogueLevel & " = " & t_eLevel & _
                " AND " & gk_sDBCOLCatalogueID & " = " & m_lCatalogueID & _
                " AND " & gk_sDBCOLEditSetID & " = " & m_lEditsetID

    DeleteLinkDataElements = sSQL

End Function

Private Function UpdateLink(t_eLevel As ECATALOGUELEVEL, _
                            t_eProperty As ECATALOGUEPROPERTIES, _
                            t_rsmLink As Variant) As String
' Creates the insert string for the PC_link table
'
' Amendment history:
'   Version Name            Date            Comments
'   d1      D Johnston      19-Jul-2000     Created from BOEADImport
    
    Dim sSQL As String
    
    ' Catalogue entry 1 is the source, 2 is the destination.
    
    sSQL = "INSERT INTO " & gk_sDBTABLEEditSetPCLink & " ( " & _
        gk_sDBCOLCatalogueLevel1 & "," & _
        gk_sDBCOLCatalogueID1 & "," & _
        gk_sDBCOLCatalogueLevel2 & "," & _
        gk_sDBCOLCatalogueID2 & "," & _
        gk_sDBCOLPropertyID & "," & _
        gk_sDBCOLDescription & "," & _
        gk_sDBCOLEditSetID & _
        ")"

    sSQL = sSQL & " VALUES ( " & _
        t_eLevel & "," & _
        m_lCatalogueID & "," & _
        t_rsmLink.iCatalogueLevelID & "," & _
        t_rsmLink.lCatalogueID & "," & _
        t_eProperty & "," & _
        QuoteValueWithQuotes(t_rsmLink.sOptionalDescription) & "," & _
        m_lEditsetID & _
        ")"
        
    UpdateLink = sSQL
    
End Function


Private Function DeleteLinks(t_eLevel As ECATALOGUELEVEL) As String
' Creates the delete string for all the entries in the PC_link table
'
' Amendment history:
'   Version Name            Date            Comments
'   d1      D Johnston      19-Jul-2000     Created from BOEADImport
'   d3      D Johnston      05-Oct-2000     Removed the code to delete reciprocal Related Material links, as
'                                           the release code no longer puts them in
    
    Dim sSQL As String
    
    ' Catalogue entry 1 is the source, 2 is the destination.  But ...
    ' For related materials there is a bi-directional link; we must
    ' delete the other link going back the way ...
    
    ' For both types of link, delete the 'outbound' links
    sSQL = "DELETE " & gk_sDBTABLEEditSetPCLink & _
        " WHERE (" & gk_sDBCOLCatalogueLevel1 & " = " & t_eLevel & _
        " AND " & gk_sDBCOLCatalogueID1 & " = " & m_lCatalogueID & _
        " AND " & gk_sDBCOLEditSetID & " = " & m_lEditsetID & ")"
    
    DeleteLinks = sSQL
    
End Function

Private Function DeleteIndexLinks(t_eLevel As ECATALOGUELEVEL, t_eTableType As ECATALOGUETABLETYPE, t_eProperty As ECATALOGUEPROPERTIES) As String
' Creates the delete string for all the entries in the PC_* tables for index terms
'
' Amendment history:
'   Version Name            Date            Comments
'   d1      D Johnston      19-Jul-2000     Created from BOEADImport
    
    Dim sSQL As String
    Dim sTablename As String
    
    sTablename = GetTableName(t_eTableType, SECeditorial)
    
    ' Catalogue entry 1 is the source, 2 is the destination.  But ...
    ' For related materials there is a bi-directional link; we must
    ' delete the other link going back the way ...
    
    ' For both types of link, delete the 'outbound' links
    sSQL = "DELETE " & sTablename & _
        " WHERE " & gk_sDBCOLCatalogueLevel & " = " & t_eLevel & _
        " AND " & gk_sDBCOLCatalogueID & " = " & m_lCatalogueID & _
        " AND " & gk_sDBCOLPropertyID & " = " & t_eProperty & _
        " AND " & gk_sDBCOLEditSetID & " = " & m_lEditsetID
    
    DeleteIndexLinks = sSQL
    
End Function

Private Function InsertForIndexTerms(t_eTableType As ECATALOGUETABLETYPE, _
                            t_eLevel As ECATALOGUELEVEL, _
                            t_eProperty As ECATALOGUEPROPERTIES, _
                            t_lID As Long, _
                            t_vDescription As Variant) As String
' Creates the insert string for entering the index links into the different tables
'
' Amendment history:
'   Version Name            Date            Comments
'   d1      D Johnston      19-Jul-2000     Created from BOEADImport

    Dim sTablename As String
    Dim sSQL  As String
    
    sTablename = GetTableName(t_eTableType, _
                    ESectionType.SECeditorial, t_eLevel)
                    
    sSQL = "INSERT INTO " & sTablename & " ( " & _
                gk_sDBCOLPropertyID & "," & _
                gk_sDBCOLCatalogueLevel & "," & _
                gk_sDBCOLCatalogueID & "," & _
                GetReferenceColumnNameInLinkTable(t_eTableType) & "," & _
                gk_sDBCOLDescription & "," & _
                gk_sDBCOLEditSetID & _
                ")"

    sSQL = sSQL & " VALUES ( " & _
                t_eProperty & "," & _
                t_eLevel & "," & _
                m_lCatalogueID & "," & _
                t_lID & "," & _
                QuoteValueWithQuotes(t_vDescription) & "," & _
                m_lEditsetID & _
                ")"

    InsertForIndexTerms = sSQL

End Function



Private Function GetSingleFieldTableFieldName(t_eTableType As ECATALOGUETABLETYPE) As String
'Returns the column name for the single column table names
'
' Amendment history:
'   Version Name            Date            Comments
'   d1      D Johnston      19-Jul-2000     Created from BOEADImport

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
' Returns the ID column name for the different link table types
'
' Amendment history:
'   Version Name            Date            Comments
'   d1      D Johnston      19-Jul-2000     Created from BOEADImport

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
'
' Amendment history:
'   Version Name            Date            Comments
'   d1      D Johnston      19-Jul-2000     Created from BOEADImport
'   d5      A Manic         26-Nov-2001     Modified to split adjacent paragraph marks
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
        
        GetSingleFieldTableFieldValue = SplitParagraphs(m_oboParsedCatalogueEntry.ScopeContent)
        
    Case ECATALOGUETABLETYPE.cttTitleTable
        
        GetSingleFieldTableFieldValue = SplitParagraphs(m_oboParsedCatalogueEntry.Title)
        
    End Select
    
End Function


Private Function GetLinkTableDescriptionFieldValue(t_eProperty As ECATALOGUEPROPERTIES) As Variant
'Return the description values for the different single field tables
'
' Amendment history:
'   Version Name            Date            Comments
'   d1      D Johnston      19-Jul-2000     Created from BOEADImport
    
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
' Function returns null (string) is the value is zero else returns the value
'
' Amendment history:
'   Version Name            Date            Comments
'   d1      D Johnston      19-Jul-2000     Created from BOEADImport

    If t_lValue = 0 Then
        NullifZero = "Null"
    Else
        NullifZero = t_lValue
    End If
End Function

Private Function NullifEmptyString(t_sValue As String) As Variant
'Function returns null (string) is the value is zero else returns the value
'
' Amendment history:
'   Version Name            Date            Comments
'   d1      D Johnston      19-Jul-2000     Created from BOEADImport

    If t_sValue = "" Then
        NullifEmptyString = Null
    Else
        NullifEmptyString = t_sValue
    End If
End Function

Private Function NullifBlank(t_sValue As String) As String
'Function returns 'null' (string) is the value is zero else returns
'       the value
    If t_sValue = "" Then
        NullifBlank = "Null"
    Else
        NullifBlank = t_sValue
    End If
End Function

Private Function LogCurrentCatalogueUpdate()
' Logs the fact that we've started the update of this entry
'
' Amendment history:
'   Version Name            Date            Comments
'   d1      D Johnston      19-Jul-2000     Created from BOEADImport

    Dim sMessage As String
    
    sMessage = "Updating catalogue reference " & m_oboParsedCatalogueEntry.Reference & " (" & m_oboParsedCatalogueEntry.ID & ")"
    sMessage = sMessage & " at level " & _
                    GetCatalogueLevelDescription(m_oboParsedCatalogueEntry.Level)
    
    LogLoadInfo m_lBatchID, sMessage
    LogMessage sMessage, EProfile.pProfileInformation

End Function

Private Function DeleteSFALinks(t_eLevel As ECATALOGUELEVEL, t_eProperty As ECATALOGUEPROPERTIES) As String
' Deletes all the SFA links for a catalogue entry
'
' Amendment history:
'   Version Name            Date            Comments
'   d1      D Johnston      19-Jul-2000     Created from BOEADImport
    
    Dim sSQL As String
    
    sSQL = "DELETE " & gk_sDBTABLEEditsetPCSFA & _
        " WHERE " & gk_sDBCOLCatalogueLevel & " = " & t_eLevel & _
        " AND " & gk_sDBCOLCatalogueID & " = " & m_lCatalogueID & _
        " AND " & gk_sDBCOLPropertyID & " = " & t_eProperty & _
        " AND " & gk_sDBCOLEditSetID & " = " & m_lEditsetID
    
    DeleteSFALinks = sSQL
    
End Function

Private Function DeleteLeafletLinks(t_eLevel As ECATALOGUELEVEL, t_eProperty As ECATALOGUEPROPERTIES) As String
' Deletes all the Leaflet links for a catalogue entry
'
' Amendment history:
'   Version Name            Date            Comments
'   d1      D Johnston      19-Jul-2000     Created from BOEADImport
    
    Dim sSQL As String
    
    sSQL = "DELETE " & gk_sDBTABLEEditsetPCLeaflet & _
        " WHERE " & gk_sDBCOLCatalogueLevel & " = " & t_eLevel & _
        " AND " & gk_sDBCOLCatalogueID & " = " & m_lCatalogueID & _
        " AND " & gk_sDBCOLPropertyID & " = " & t_eProperty & _
        " AND " & gk_sDBCOLEditSetID & " = " & m_lEditsetID
    
    DeleteLeafletLinks = sSQL
    
End Function

Private Function UnpublishedFindingAidsHasChanged(t_eLevel As ECATALOGUELEVEL) As Boolean
' Finds out if the Unpublished Finding Aid has changed from the original; if it has,
' we have loaded it from the file.  Otherwise, we haven't!
'
' Amendment history:
'   Version Name            Date            Comments
'   d1      D Johnston      19-Jul-2000     Created
'   d2      D Johnston      04-Sep-2000     Added level checking to *HasChanged functions, and tidied up

    Dim sSource As String
    sSource = "BOEADUpdate.UnpublishedFindingAidsHasChanged"
    
    Dim sSQL As String
    Dim fwRecordList As fwRecordList
    
    UnpublishedFindingAidsHasChanged = False
    
    sSQL = "SELECT " & gk_sDBCOLDescription & _
        " FROM " & gk_sDBTABLEPCLinkDataElement & _
        " WHERE " & gk_sDBCOLCatalogueID & " = " & m_lCatalogueID & _
        " AND " & gk_sDBCOLCatalogueLevel & " = " & t_eLevel & _
        " AND " & gk_sDBCOLPropertyID & " = " & ECATALOGUEPROPERTIES.cpUnpublishedFindingAids

    Set fwRecordList = m_oFWDBHelper.FindBySql(sSQL)
    
    If fwRecordList.RecordCount > 1 Then
        ' We should only have found one record, so there is a problem here.
        RaiseError sSource, PEGeneralInternalError, "Cannot find single Unpublished Finding Aid description"
        
    ElseIf fwRecordList.RecordCount = 0 Then
        ' We didn't find any records, so we should set it now
        UnpublishedFindingAidsHasChanged = True
        
    Else
        ' Return the result of the comparison
        UnpublishedFindingAidsHasChanged = Not (fwRecordList.GetFieldValue(gk_sDBCOLDescription) = GetLinkTableDescriptionFieldValue(ECATALOGUEPROPERTIES.cpUnpublishedFindingAids))
    End If
    
End Function

Private Function PublicationNoteHasChanged(t_eLevel As ECATALOGUELEVEL) As Boolean
' Finds out if the Publication Note has changed from the original; if it has,
' we have loaded it from the file.  Otherwise, we haven't!
'
' Amendment history:
'   Version Name            Date            Comments
'   d1      D Johnston      19-Jul-2000     Created
'   d2      D Johnston      04-Sep-2000     Added level checking to *HasChanged functions, and tidied up

    Dim sSource As String
    sSource = "BOEADUpdate.PublicationNoteHasChanged"
    
    Dim sSQL As String
    Dim fwRecordList As fwRecordList
    
    ' Set default value
    PublicationNoteHasChanged = False
    
    sSQL = "SELECT " & gk_sDBCOLDescription & _
        " FROM " & gk_sDBTABLEPCLinkDataElement & _
        " WHERE " & gk_sDBCOLCatalogueID & " = " & m_lCatalogueID & _
        " AND " & gk_sDBCOLCatalogueLevel & " = " & t_eLevel & _
        " AND " & gk_sDBCOLPropertyID & " = " & ECATALOGUEPROPERTIES.cpPublicationNote

    Set fwRecordList = m_oFWDBHelper.FindBySql(sSQL)
    
    If fwRecordList.RecordCount > 1 Then
        ' We should only have found one record, so there is a problem here.
        RaiseError sSource, PEGeneralInternalError, "Cannot find single Publication Note description"
        
    ElseIf fwRecordList.RecordCount = 0 Then
        ' We didn't find any records, so we should set it now
        PublicationNoteHasChanged = True
        
    Else
        ' Return the result of the comparison
        PublicationNoteHasChanged = Not (fwRecordList.GetFieldValue(gk_sDBCOLDescription) = GetLinkTableDescriptionFieldValue(ECATALOGUEPROPERTIES.cpPublicationNote))
    End If
    
    
End Function



