VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BOCatalogueIndexTerms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    BOCatalogueIndexTerms.cls
' System:       PROCAT, PRO
' Copyright:    (C) Quidnunc Limited
'
' Description:  Manages the functionality related to the index terms
'               of the catalogues at all levels
'
' Amendment history:
'   d1   ndk      25Aug1999       Created
'   d2   ndk      20Oct1999       Added error handling for duplicate link items
'   d3   ndk      18Nov1999       Update edit_set_entry_type_id after when catalogues is updated
'   d4   wdp      22Feb2000       If attempt is made to add duplicate AF term, ignore error and resume
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Option Explicit

'Implements
Implements PROCATTransientInterface


'Property variables
Private m_eSectionType As ESectionType
Private m_lEditsetID As Long
Private m_lCatalogueID As Long
Private m_eCatalogueLevelID As ECATALOGUELEVEL
Private m_colCorporateBodyNames As Collection
Private m_colPersonNames As Collection
Private m_colPlaces As Collection
Private m_colSubjects As Collection

'Other variables
Private m_oBOCatalogueDBHelper As BOCatalogueDBHelper
Private m_oBOCatalogueElementMatrix As BOCatalogueElementMatrix
Private m_oFWApplicationData As FWApplicationData

Private Sub Initialise()
    Set m_oBOCatalogueDBHelper = New BOCatalogueDBHelper
    Set m_oBOCatalogueElementMatrix = New BOCatalogueElementMatrix
    Set m_oFWApplicationData = New FWApplicationData
    
    Set m_colCorporateBodyNames = New Collection
    Set m_colPersonNames = New Collection
    Set m_colPlaces = New Collection
    Set m_colSubjects = New Collection
End Sub
Private Sub Terminate()
    Set m_oBOCatalogueDBHelper = Nothing
    Set m_oBOCatalogueElementMatrix = Nothing
    Set m_oFWApplicationData = Nothing
    
    Set m_colCorporateBodyNames = Nothing
    Set m_colPersonNames = Nothing
    Set m_colPlaces = Nothing
    Set m_colSubjects = Nothing
End Sub
Private Sub Class_Initialize()
    Initialise
End Sub
Private Sub Class_Terminate()
    Terminate
End Sub
Public Property Get Name()
    Name = "BOCatalogueIndexTerms"
End Property
Public Property Let SectionType(ByVal t_eSectionType As ESectionType)
    m_eSectionType = t_eSectionType
End Property
Public Property Let EditsetID(ByVal t_lEditSetID As Long)
    m_lEditsetID = t_lEditSetID
End Property
Public Property Let CatalogueID(ByVal t_lCatalogueID As Long)
    m_lCatalogueID = t_lCatalogueID
End Property
Public Property Let CatalogueLevelID(ByVal t_eCatalogueLevelID As ECATALOGUELEVEL)
    m_eCatalogueLevelID = t_eCatalogueLevelID
End Property
Public Property Get EditsetID() As Long
    EditsetID = m_lEditsetID
End Property
Public Property Get CatalogueID() As Long
    CatalogueID = m_lCatalogueID
End Property
Public Property Get CatalogueLevelID() As ECATALOGUELEVEL
    CatalogueLevelID = m_eCatalogueLevelID
End Property
Public Property Get CorporateBodyNames() As Collection
    Set CorporateBodyNames = m_colCorporateBodyNames
End Property
Public Property Get PersonalNames() As Collection
    Set PersonalNames = m_colPersonNames
End Property
Public Property Get Places() As Collection
    Set Places = m_colPlaces
End Property
Public Property Get Subjects() As Collection
    Set Subjects = m_colSubjects
End Property
Public Property Let CorporateBodyNames(ByVal t_colCorporateBodyNames As Collection)
    Set m_colCorporateBodyNames = t_colCorporateBodyNames
End Property
Public Property Let PersonalNames(ByVal t_colPersonNames As Collection)
    Set m_colPersonNames = t_colPersonNames
End Property
Public Property Let Places(ByVal t_colPlaces As Collection)
    Set m_colPlaces = t_colPlaces
End Property
Public Property Let Subjects(ByVal t_colSubjects As Collection)
    Set m_colSubjects = t_colSubjects
End Property


Public Sub Load( _
    ByVal t_eLoadType As ECatalogueEditFormLoadType, _
    Optional ByVal t_lLinkTermID As Long, _
    Optional ByVal t_eLinkTermType As ELinkTermType)
'
'Description: Loads the details of the bo,depending on the value of t_eLoadType either the details are
'loaded from the loader class or from the stored stage
'
'Amendment history:
'   d1   ndk      25Aug1999      Created
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String
    
    sSource = Me.Name & ".Load"
    
    'The user clicked on Edit button and control has come here
    If t_eLoadType = ECatalogueEditFormLoadType.etLoadFromDatabaseEditCatalogueEntry Then
        
        'Delete the state information (if any)
        PROCATTransientInterface_DeleteState
        
        'Load from the loader class
        LoadFromDatabase
    Else
        
        Select Case t_eLoadType
        
            Case ECatalogueEditFormLoadType.etLoadFromSavedStateAndAddTermEditCatalogueEntry
                'The user has selected a pick term to add and the control has come here
                
                'Read the stored state info
                PROCATTransientInterface_GetState
                
                'Add the new link term to the collection
                AddNewLinkTerm t_lLinkTermID, t_eLinkTermType
            
            Case ECatalogueEditFormLoadType.etLoadFromSavedStateAndRemoveTermEditCatalogueEntry
                'The user has selected a pick term to remove and the control has come here
                
                'Read the stored state info
                PROCATTransientInterface_GetState

                'Remove the link term from the collection
                RemoveLinkTerm t_lLinkTermID, t_eLinkTermType
                
            Case ECatalogueEditFormLoadType.etLoadFromSavedStateNoChangeEditCatalogueEntry
                'The user has selected cancel operation while adding/removing a link term and the control has come here
                
                'Just Read the stored stage info
                PROCATTransientInterface_GetState
                
        End Select
        
        'Recreate the collections with all details
        PopulateRecordSetsUsingIDsInCollection
    End If
    
    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub

Private Sub AddNewLinkTerm( _
    t_lLinkTermID As Long, _
    t_eLinkTermType As ELinkTermType)
'
'Description: Adds a link item to the collection
'
'Amendment history:
'   d1   ndk      25Aug1999      Created
'   d2   ndk      20Oct1999      Added error handling for duplicate link items
'   d4   wdp      22Feb2000       If attempt is made to add duplicate AF term, ignore error and resume

    On Error GoTo ErrorHandler
    
    Dim sSource As String
    
    sSource = Me.Name & ".AddNewLinkTerm"
    
    Dim recLinkItemRecord  As GenericLinkRecord
    
    recLinkItemRecord.lID = t_lLinkTermID
    recLinkItemRecord.sOptionalDescription = ""
    
    Select Case t_eLinkTermType
        Case ELinkTermType.elIndexTermCorporateBody
            m_colCorporateBodyNames.Add recLinkItemRecord, gk_sPrefixForLinkTerm & recLinkItemRecord.lID
        Case ELinkTermType.elIndexTermPerson
            m_colPersonNames.Add recLinkItemRecord, gk_sPrefixForLinkTerm & recLinkItemRecord.lID
        Case ELinkTermType.elIndexTermPlace
            m_colPlaces.Add recLinkItemRecord, gk_sPrefixForLinkTerm & recLinkItemRecord.lID
        Case ELinkTermType.elIndexTermSubject
            m_colSubjects.Add recLinkItemRecord, gk_sPrefixForLinkTerm & recLinkItemRecord.lID
    End Select
    
    Exit Sub
ErrorHandler:
    If Err.Number = gk_lErrorNumberDuplicateKeyValueForCollection Then
        'Dim sErrorMessage As String
        
        'sErrorMessage = "You have choosen a duplicate link term for "
        
        'If t_eLinkTermType = elIndexTermCorporateBody Then
        '    sErrorMessage = sErrorMessage & _
        '        m_oFWApplicationData.GetDisplayInformation(ESharedResourceFile.srfTabDisplayIndexTermsLeftHeaderCorporateNames)
        'ElseIf t_eLinkTermType = elIndexTermPerson Then
        '    sErrorMessage = sErrorMessage & _
        '        m_oFWApplicationData.GetDisplayInformation(ESharedResourceFile.srfTabDisplayIndexTermsLeftHeaderPersonalNames)
        'ElseIf t_eLinkTermType = elPlaceOfDeposit Then
        '    sErrorMessage = sErrorMessage & _
        '        m_oFWApplicationData.GetDisplayInformation(ESharedResourceFile.srfTabDisplayIndexTermsLeftHeaderPlaces)
        'ElseIf t_eLinkTermType = elIndexTermSubject Then
        '    sErrorMessage = sErrorMessage & _
        '        m_oFWApplicationData.GetDisplayInformation(ESharedResourceFile.srfTabDisplayIndexTermsLeftHeaderSubject)
        'End If
        
        'RaiseError sSource, peDuplicateLinkItemChoosen, sErrorMessage
        Resume Next
    Else
        ConvertError sSource
    End If

End Sub


Private Sub RemoveLinkTerm( _
    t_lLinkTermID As Long, _
    t_eLinkTermType As ELinkTermType)
'
'Description: Removes the link item from the collection
'
'Amendment history:
'   d1   ndk      25Aug1999      Created
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String
    
    sSource = Me.Name & ".RemoveLinkTerm"
    
    Select Case t_eLinkTermType
        Case ELinkTermType.elIndexTermCorporateBody
            m_colCorporateBodyNames.Remove gk_sPrefixForLinkTerm & t_lLinkTermID
        Case ELinkTermType.elIndexTermPerson
            m_colPersonNames.Remove gk_sPrefixForLinkTerm & t_lLinkTermID
        Case ELinkTermType.elIndexTermPlace
            m_colPlaces.Remove gk_sPrefixForLinkTerm & t_lLinkTermID
        Case ELinkTermType.elIndexTermSubject
            m_colSubjects.Remove gk_sPrefixForLinkTerm & t_lLinkTermID
    End Select
    
    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub


Private Sub PopulateRecordSetsUsingIDsInCollection()
'
'Description: Recreates the link term details using the ids stored in the collection
'These were stored in the collection by retreieving them from the stored state
'
'Amendment history:
'   d1   ndk      25Aug1999      Created
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String
    
    sSource = Me.Name & ".PopulateRecordSetsUsingIDsInCollection"
    
    'Corporate Body
    Set m_colCorporateBodyNames = _
        m_oBOCatalogueDBHelper.GenerateDetailsUsingSavedIDs( _
            m_lEditsetID, _
            m_lCatalogueID, _
            m_eCatalogueLevelID, _
            ELinkTermType.elIndexTermCorporateBody, _
            m_colCorporateBodyNames)
    
    'Person
    Set m_colPersonNames = _
        m_oBOCatalogueDBHelper.GenerateDetailsUsingSavedIDs( _
            m_lEditsetID, _
            m_lCatalogueID, _
            m_eCatalogueLevelID, _
            ELinkTermType.elIndexTermPerson, _
            m_colPersonNames)
    
    'Place
    Set m_colPlaces = _
        m_oBOCatalogueDBHelper.GenerateDetailsUsingSavedIDs( _
            m_lEditsetID, _
            m_lCatalogueID, _
            m_eCatalogueLevelID, _
            ELinkTermType.elIndexTermPlace, _
            m_colPlaces)
    
    'Subject
    Set m_colSubjects = _
        m_oBOCatalogueDBHelper.GenerateDetailsUsingSavedIDs( _
            m_lEditsetID, _
            m_lCatalogueID, _
            m_eCatalogueLevelID, _
            ELinkTermType.elIndexTermSubject, _
            m_colSubjects)
    
    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub
Private Sub LoadFromDatabase()
'
'Description: Loads the object details from the loader class
'
' Amendment history:
'   d1   ndk      25Aug1999      Created
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim oBOCatalogueIndexTermsLd As BOCatalogueIndexTermsLd
    
    sSource = Me.Name & ".LoadFromDatabase"
    
    Set oBOCatalogueIndexTermsLd = New BOCatalogueIndexTermsLd
    
    'Set the properties
    oBOCatalogueIndexTermsLd.EditsetID = m_lEditsetID
    oBOCatalogueIndexTermsLd.CatalogueID = m_lCatalogueID
    oBOCatalogueIndexTermsLd.CatalogueLevelID = m_eCatalogueLevelID
    oBOCatalogueIndexTermsLd.SectionType = m_eSectionType
    
    'Load the details
    oBOCatalogueIndexTermsLd.Load
    
    'Move details from the loader class to current object
    Set m_colCorporateBodyNames = oBOCatalogueIndexTermsLd.CorporateBodyNames
    Set m_colPersonNames = oBOCatalogueIndexTermsLd.PersonalNames
    Set m_colPlaces = oBOCatalogueIndexTermsLd.Places
    Set m_colSubjects = oBOCatalogueIndexTermsLd.Subjects
    
    Set oBOCatalogueIndexTermsLd = Nothing
    
    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub


Private Sub PROCATTransientInterface_DeleteState(Optional t_sInstanceReference As String)
'
'Description: Deletes the state of the index term details
'
' Amendment history:
'   d1   ndk      25Aug1999      Created
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String
    
    Dim sInstanceReference As String
    Dim ofwTransientManager As FWTransientManager
    
    sSource = Me.Name & ".PROCATTransientInterface_DeleteState"
    
    'Initialise object
    Set ofwTransientManager = New FWTransientManager
    
    'Create the instance reference
    sInstanceReference = CreateDelimitedString(Me.Name, t_sInstanceReference)
    
    'Delete from the transient store
    ofwTransientManager.DeleteData GetObjectContext.Item(gk_sSessionObject).SessionID, sInstanceReference
        
    LogMessage "Deleted " & Me.Name & " state" & sInstanceReference, EProfile.pComplete

    Set ofwTransientManager = Nothing
    
    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub

Public Sub PROCATTransientInterface_GetState(Optional t_sInstanceReference As String)
'
'Description: gets the the current state of the index term details of the catalogue whose details are
'set in the properties
'
' Amendment history:
'   d1   ndk      25Aug1999      Created
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String
    
    Dim sInstanceReference As String
    Dim sSerialisedString As String
    Dim ofwTransientManager As FWTransientManager
    
    sSource = Me.Name & ".PROCATTransientInterface_GetState"
    
    'Initialise object
    Set ofwTransientManager = New FWTransientManager
    
    'Create the instance reference
    sInstanceReference = CreateDelimitedString(Me.Name, t_sInstanceReference)
    
    'Read from the transient store
    sSerialisedString = _
        ofwTransientManager.GetData(GetObjectContext.Item(gk_sSessionObject).SessionID, sInstanceReference)

    If Trim(sSerialisedString) = "" Then
        LogMessage Me.Name & "state not found " & sInstanceReference, EProfile.pProfileInformation
    Else
        'Simple Strings
        m_eSectionType = CInt(GetDataFromSerialisedString(sSerialisedString))
        m_lEditsetID = CLng(GetDataFromSerialisedString(sSerialisedString))
        m_lCatalogueID = CLng(GetDataFromSerialisedString(sSerialisedString))
        m_eCatalogueLevelID = GetCatalogueLevelEnum(CInt(GetDataFromSerialisedString(sSerialisedString)))
        
        'Collections
        Set m_colPersonNames = m_oBOCatalogueDBHelper.ExtractLinkItemsFromSerialIsedString(sSerialisedString)
        Set m_colCorporateBodyNames = m_oBOCatalogueDBHelper.ExtractLinkItemsFromSerialIsedString(sSerialisedString)
        Set m_colPlaces = m_oBOCatalogueDBHelper.ExtractLinkItemsFromSerialIsedString(sSerialisedString)
        Set m_colSubjects = m_oBOCatalogueDBHelper.ExtractLinkItemsFromSerialIsedString(sSerialisedString)
            
        LogMessage "Read " & Me.Name & " state" & sInstanceReference, EProfile.pComplete
    End If
    
    Set ofwTransientManager = Nothing
    
    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub

Public Sub PROCATTransientInterface_SetState(Optional t_sInstanceReference As String)
'
'Description: Saves the the current state of the index term details of the catalogue whose details are
'set in the properties
'
' Amendment history:
'   d1   ndk      25Aug1999      Created
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String
    
    Dim sInstanceReference As String
    Dim sSerialisedString As String
    Dim ofwTransientManager As FWTransientManager
    
    sSource = Me.Name & ".PROCATTransientInterface_SetState"
    
    'Initialise object
    Set ofwTransientManager = New FWTransientManager
    
    'Create the instance reference
    sInstanceReference = CreateDelimitedString(Me.Name, t_sInstanceReference)

    'Simple Strings
    sSerialisedString = ""
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, CStr(m_eSectionType))
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, CStr(m_lEditsetID))
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, CStr(m_lCatalogueID))
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, CStr(GetCatalogueLevelCode(m_eCatalogueLevelID)))
    
    'Collections
    sSerialisedString = m_oBOCatalogueDBHelper.SerialiseLinkItem(sSerialisedString, m_colPersonNames)
    sSerialisedString = m_oBOCatalogueDBHelper.SerialiseLinkItem(sSerialisedString, m_colCorporateBodyNames)
    sSerialisedString = m_oBOCatalogueDBHelper.SerialiseLinkItem(sSerialisedString, m_colPlaces)
    sSerialisedString = m_oBOCatalogueDBHelper.SerialiseLinkItem(sSerialisedString, m_colSubjects)
    
    'Store in the transient manager
    ofwTransientManager.AddData _
        GetObjectContext.Item(gk_sSessionObject).SessionID, sInstanceReference, sSerialisedString

    LogMessage "Stored " & Me.Name & sInstanceReference, EProfile.pComplete

    Set ofwTransientManager = Nothing
    
    Exit Sub
    
ErrorHandler:
    ConvertError sSource
End Sub

Public Sub Save()
'
'Description: Saves the index term details of the catalogue whose details are set in the properties
'
'Properties required to be set before calling this method
'
'CatalogueID
'CatalogueLevelID
'EditSetID
'
' Amendment history:
'   d1   ndk      25Aug1999      Created
'   d3   ndk      18Nov1999      Update edit_set_entry_type_id after when catalogues is updated
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim sConditionForLinkTables As String
    
    sSource = Me.Name & ".Save"
    
    'Generic condition
    sConditionForLinkTables = _
        gk_sDBCOLEditSetID & "=" & m_lEditsetID & " AND " & _
        gk_sDBCOLCatalogueID & "=" & m_lCatalogueID & " AND " & _
        gk_sDBCOLCatalogueLevel & "=" & GetCatalogueLevelCode(m_eCatalogueLevelID)
    
    'Corporate Body
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(m_eCatalogueLevelID, cpIndexTermCorporateName) Then
        m_oBOCatalogueDBHelper.UpdateLinkTable _
            m_lEditsetID, m_lCatalogueID, m_eCatalogueLevelID, elIndexTermCorporateBody, m_colCorporateBodyNames, False

    End If
    
    'Person
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(m_eCatalogueLevelID, cpIndexTermPersonName) Then
        m_oBOCatalogueDBHelper.UpdateLinkTable _
            m_lEditsetID, m_lCatalogueID, m_eCatalogueLevelID, elIndexTermPerson, m_colPersonNames, False

    End If
    
    'Place
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(m_eCatalogueLevelID, cpIndexTermPlaceName) Then
        m_oBOCatalogueDBHelper.UpdateLinkTable _
            m_lEditsetID, m_lCatalogueID, m_eCatalogueLevelID, elIndexTermPlace, m_colPlaces, False

    End If
    
    'Subject
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(m_eCatalogueLevelID, cpIndexTermSubject) Then
        m_oBOCatalogueDBHelper.UpdateLinkTable _
            m_lEditsetID, m_lCatalogueID, m_eCatalogueLevelID, elIndexTermSubject, m_colSubjects, False

    End If
    
    'update the status that catalogue was modified
    m_oBOCatalogueDBHelper.SetCatalogueStatusAsModified _
        GetTableName(cttExtensionTable, SECEditorial, m_eCatalogueLevelID), _
        gk_sDBCOLEditSetID & "=" & m_lEditsetID & " AND " & _
        GetColumnName(cctIDColumn, m_eCatalogueLevelID) & "=" & m_lCatalogueID
    
    Exit Sub
    
ErrorHandler:
    ConvertError sSource
End Sub


