VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BOCatalogueAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    BOCatalogueAccess.cls
' System:       PROCAT, PRO
' Copyright:    (C) Quidnunc Limited
'
' Description:  Manages the functionality related to the access details
'               of the catalogues at all levels
'
' Amendment history:
'   d1   ndk        25Aug1999      Created
'   d2   ndk        20Oct1999      Added error handling for duplicate link items
'                                  Added code to handle fields closure type and closure code
'   d3   ndk        03Nov1999      Changed the way closure type and code are handled
'                                  Put validations for lengths of text fields
'   d4   ndk        18Nov1999      Update edit_set_entry_type_id after when catalogues is updated
'   d5   ndk        19Nov1999      Used quotes before updating closure status
'   d6   wdp        03Dec1999      Added record status flags
'   d7   ndk        17Dec1999      Record opening date handled properly
'   d8   ndk        29Dec1999      Using tbl_piece/item.open_date column instead of piece/item_extension.Record_Opening_date
'   d9   gdb        13Feb2001      Changed Save function to use oBOCatalogueSummaryLd to pass data between Summary and Acess
'   d10   gdb       13Feb2001      Changed Save function to invoke EditorialFunctions.IsValidOpenDate and .IsValidStatus
'   d11  fjd        05Feb2003       Renamed LCI flag to AR
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'Implements
Implements PROCATTransientInterface

'Property variables
Private m_eSectionType As ESectionType
Private m_lEditsetID As Long
Private m_lCatalogueID As Long
Private m_eCatalogueLevelID As ECATALOGUELEVEL
Private m_vAccessConditionsID As Variant
Private m_vClosureType As Variant
Private m_vClosureCode As Variant
Private m_vRecordOpeningDate As Variant
Private m_vClosureStatus As Variant
Private m_vRestrictionsOnUse As Variant
Private m_vCustodialHistory As Variant
Private m_vAccumulationDates As Variant
Private m_vAppraisalDestructionInfo As Variant
Private m_vAccruals As Variant
Private m_colImmedieteSourceOfAcquisitionPerson As Collection
Private m_colImmedieteSourceOfAcquisitionCorporateBody As Collection
Private m_colLocationOfOriginals As Collection
Private m_colCopiesInformation As Collection
Private m_vUnknownImmedieteSourceOfAcquisition As Variant
Private m_vUnknownCopiesInformation As Variant

'Flags that represent parts of the record_status field
Private m_bAccessionedAsCopy As Boolean
Private m_bExtractAtTransfer As Boolean
Private m_bOrderable As Boolean
Private m_bAR As Boolean

'Other variables
Private m_oBOCatalogueDBHelper As BOCatalogueDBHelper
Private m_oBOCatalogueElementMatrix As BOCatalogueElementMatrix
Private m_oFWReferenceData As FWReferenceData
Private m_oFWApplicationData As FWApplicationData

Private Sub Initialise()
    Set m_oBOCatalogueDBHelper = New BOCatalogueDBHelper
    Set m_oBOCatalogueElementMatrix = New BOCatalogueElementMatrix
    Set m_oFWReferenceData = New FWReferenceData
    Set m_oFWApplicationData = New FWApplicationData
    
    Set m_colImmedieteSourceOfAcquisitionPerson = New Collection
    Set m_colImmedieteSourceOfAcquisitionCorporateBody = New Collection
    Set m_colCopiesInformation = New Collection
    Set m_colLocationOfOriginals = New Collection
End Sub
Private Sub Terminate()
    Set m_oBOCatalogueDBHelper = Nothing
    Set m_oBOCatalogueElementMatrix = Nothing
    Set m_oFWReferenceData = Nothing
    Set m_oFWApplicationData = Nothing
    
    Set m_colImmedieteSourceOfAcquisitionPerson = Nothing
    Set m_colImmedieteSourceOfAcquisitionCorporateBody = Nothing
    Set m_colCopiesInformation = Nothing
    Set m_colLocationOfOriginals = Nothing
End Sub
Private Sub Class_Initialize()
    Initialise
End Sub
Private Sub Class_Terminate()
    Terminate
End Sub
Public Property Get Name()
    Name = "BOCatalogueAccess"
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
Public Property Let AccessConditionID(ByVal t_vAccessConditionID As Variant)
    m_vAccessConditionsID = t_vAccessConditionID
End Property
Public Property Let ClosureType(ByVal t_vClosureType As Variant)
    m_vClosureType = t_vClosureType
End Property
Public Property Let ClosureCode(ByVal t_vClosureCode As Variant)
    m_vClosureCode = t_vClosureCode
End Property
Public Property Let RecordOpeningDate(ByVal t_vRecordOpeningDate As Variant)
    m_vRecordOpeningDate = t_vRecordOpeningDate
End Property
Public Property Let ClosureStatus(ByVal t_vClosureStatus As Variant)
    m_vClosureStatus = t_vClosureStatus
End Property
Public Property Let RestrictionsOnUse(ByVal t_vRestrictionsOnUse)
    m_vRestrictionsOnUse = t_vRestrictionsOnUse
End Property
Public Property Let CustodialHistory(ByVal t_vCustodialHistory)
    m_vCustodialHistory = t_vCustodialHistory
End Property
Public Property Let AccumulationDate(ByVal t_vAccumulationDate)
    m_vAccumulationDates = t_vAccumulationDate
End Property
Public Property Let AppraisalDestructionInfo(ByVal t_vAppraisalDestructionInfo)
    m_vAppraisalDestructionInfo = t_vAppraisalDestructionInfo
End Property
Public Property Let Accruals(ByVal t_vAccruals)
    m_vAccruals = t_vAccruals
End Property
Public Property Let ImmedieteSourceOfAcquisitionPerson(ByVal t_colImmedieteSourceOfAcquisitionPerson As Collection)
    Set m_colImmedieteSourceOfAcquisitionPerson = t_colImmedieteSourceOfAcquisitionPerson
End Property
Public Property Let ImmedieteSourceOfAcquisitionCorporateBody(ByVal t_colImmedieteSourceOfAcquisitionCorporateBody As Collection)
    Set m_colImmedieteSourceOfAcquisitionCorporateBody = t_colImmedieteSourceOfAcquisitionCorporateBody
End Property
Public Property Let UnknownImmedieteSourceOfAcquisition(ByVal t_vUnknownImmedieteSourceOfAcquisition As Variant)
    m_vUnknownImmedieteSourceOfAcquisition = t_vUnknownImmedieteSourceOfAcquisition
End Property
Public Property Let LocationOfOriginals(ByVal t_colLocationOfOriginals As Collection)
    Set m_colLocationOfOriginals = t_colLocationOfOriginals
End Property
Public Property Let CopiesInformation(ByVal t_colCopiesInformation As Collection)
    Set m_colCopiesInformation = t_colCopiesInformation
End Property
Public Property Let UnknownCopiesInformation(ByVal t_vUnknownCopiesInformation As Variant)
    m_vUnknownCopiesInformation = t_vUnknownCopiesInformation
End Property

Public Property Let AccessionedAsCopy(t_bAccessionedAsCopy As Boolean)
    m_bAccessionedAsCopy = t_bAccessionedAsCopy
End Property

Public Property Let ExtractAtTransfer(t_bExtractAtTransfer As Boolean)
    m_bExtractAtTransfer = t_bExtractAtTransfer
End Property

Public Property Let AR(t_bAR As Boolean)
    m_bAR = t_bAR
End Property

Public Property Let Orderable(t_bOrderable As Boolean)
    m_bOrderable = t_bOrderable
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
Public Property Get AccessConditionID() As Variant
    AccessConditionID = m_vAccessConditionsID
End Property
Public Property Get RecordOpeningDate() As Variant
    RecordOpeningDate = m_vRecordOpeningDate
End Property
Public Property Get ClosureStatus() As Variant
    ClosureStatus = m_vClosureStatus
End Property
Public Property Get RestrictionsOnUse() As Variant
    RestrictionsOnUse = m_vRestrictionsOnUse
End Property
Public Property Get CustodialHistory() As Variant
    CustodialHistory = m_vCustodialHistory
End Property
Public Property Get AccumulationDate() As Variant
    AccumulationDate = m_vAccumulationDates
End Property
Public Property Get AppraisalDestructionInfo() As Variant
    AppraisalDestructionInfo = m_vAppraisalDestructionInfo
End Property
Public Property Get Accruals() As Variant
    Accruals = m_vAccruals
End Property
Public Property Get ImmedieteSourceOfAcquisitionPerson() As Collection
    Set ImmedieteSourceOfAcquisitionPerson = m_colImmedieteSourceOfAcquisitionPerson
End Property
Public Property Get ImmedieteSourceOfAcquisitionCorporateBody() As Collection
    Set ImmedieteSourceOfAcquisitionCorporateBody = m_colImmedieteSourceOfAcquisitionCorporateBody
End Property
Public Property Get UnknownImmedieteSourceOfAcquisition() As Variant
    UnknownImmedieteSourceOfAcquisition = m_vUnknownImmedieteSourceOfAcquisition
End Property
Public Property Get LocationOfOriginals() As Collection
    Set LocationOfOriginals = m_colLocationOfOriginals
End Property
Public Property Get CopiesInformation() As Collection
    Set CopiesInformation = m_colCopiesInformation
End Property
Public Property Get UnknownCopiesInformation() As Variant
    UnknownCopiesInformation = m_vUnknownCopiesInformation
End Property
Public Property Get ClosureType() As Variant
    ClosureType = m_vClosureType
End Property
Public Property Get ClosureCode() As Variant
    ClosureCode = m_vClosureCode
End Property

Public Property Get AccessionedAsCopy() As Boolean
    AccessionedAsCopy = m_bAccessionedAsCopy
End Property

Public Property Get ExtractAtTransfer() As Boolean
    ExtractAtTransfer = m_bExtractAtTransfer
End Property

Public Property Get AR() As Boolean
    AR = m_bAR
End Property

Public Property Get Orderable() As Boolean
    Orderable = m_bOrderable
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
            
            Case ECatalogueEditFormLoadType.etLoadFromSavedStateAndRemoveTermEditCatalogueEntry 'The user has selected a pick term to remove and the control has come here
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

    On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim recLinkItemRecord  As GenericLinkRecord
    
    sSource = Me.Name & ".AddNewLinkTerm"
    
    recLinkItemRecord.lID = t_lLinkTermID
    recLinkItemRecord.sOptionalDescription = ""
    
    Select Case t_eLinkTermType
        Case ELinkTermType.elImmedieteSourceOfAcquisitionPerson
            m_colImmedieteSourceOfAcquisitionPerson.Add recLinkItemRecord, gk_sPrefixForLinkTerm & recLinkItemRecord.lID
        Case ELinkTermType.elImmedieteSourceOfAcquisitionCorporateBody
            m_colImmedieteSourceOfAcquisitionCorporateBody.Add recLinkItemRecord, gk_sPrefixForLinkTerm & recLinkItemRecord.lID
        Case ELinkTermType.elCopiesInformation
            m_colCopiesInformation.Add recLinkItemRecord, gk_sPrefixForLinkTerm & recLinkItemRecord.lID
        Case ELinkTermType.elLocationOfOriginals
            m_colLocationOfOriginals.Add recLinkItemRecord, gk_sPrefixForLinkTerm & recLinkItemRecord.lID
        
    End Select
    
    Exit Sub
ErrorHandler:
    If Err.Number = gk_lErrorNumberDuplicateKeyValueForCollection Then
        Dim sErrorMessage As String
        
        sErrorMessage = "You have choosen a duplicate link term for "
        
        If t_eLinkTermType = elImmedieteSourceOfAcquisitionPerson Then
            sErrorMessage = sErrorMessage & _
                m_oFWApplicationData.GetDisplayInformation(ESharedResourceFile.srfTabDisplayAccessDetailsLeftHeaderImmediateSourceOfAcquisition)
        ElseIf t_eLinkTermType = elImmedieteSourceOfAcquisitionCorporateBody Then
            sErrorMessage = sErrorMessage & _
                m_oFWApplicationData.GetDisplayInformation(ESharedResourceFile.srfTabDisplayAccessDetailsLeftHeaderImmediateSourceOfAcquisition)

        ElseIf t_eLinkTermType = elCopiesInformation Then
            sErrorMessage = sErrorMessage & _
                m_oFWApplicationData.GetDisplayInformation(ESharedResourceFile.srfTabDisplayAccessDetailsLeftHeaderCopiesInformation)

        ElseIf t_eLinkTermType = elLocationOfOriginals Then
            sErrorMessage = sErrorMessage & _
                m_oFWApplicationData.GetDisplayInformation(ESharedResourceFile.srfTabDisplayAccessDetailsLeftHeaderLocationOfOriginals)
        End If
        
        RaiseError sSource, peDuplicateLinkItemChoosen, sErrorMessage
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
        Case ELinkTermType.elImmedieteSourceOfAcquisitionPerson
            m_colImmedieteSourceOfAcquisitionPerson.Remove gk_sPrefixForLinkTerm & t_lLinkTermID
        Case ELinkTermType.elImmedieteSourceOfAcquisitionCorporateBody
            m_colImmedieteSourceOfAcquisitionCorporateBody.Remove gk_sPrefixForLinkTerm & t_lLinkTermID
        Case ELinkTermType.elCopiesInformation
            m_colCopiesInformation.Remove gk_sPrefixForLinkTerm & t_lLinkTermID
        Case ELinkTermType.elLocationOfOriginals
            m_colLocationOfOriginals.Remove gk_sPrefixForLinkTerm & t_lLinkTermID
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
    
    'Immediate source of acquisition person
    Set m_colImmedieteSourceOfAcquisitionPerson = _
        m_oBOCatalogueDBHelper.GenerateDetailsUsingSavedIDs( _
            m_lEditsetID, _
            m_lCatalogueID, _
            m_eCatalogueLevelID, _
            ELinkTermType.elImmedieteSourceOfAcquisitionPerson, _
            m_colImmedieteSourceOfAcquisitionPerson)
    
    'Immediate source of acquisition corporate body
    Set m_colImmedieteSourceOfAcquisitionCorporateBody = _
        m_oBOCatalogueDBHelper.GenerateDetailsUsingSavedIDs( _
            m_lEditsetID, _
            m_lCatalogueID, _
            m_eCatalogueLevelID, _
            ELinkTermType.elImmedieteSourceOfAcquisitionCorporateBody, _
            m_colImmedieteSourceOfAcquisitionCorporateBody)
    
    'Copies Information
    Set m_colCopiesInformation = _
        m_oBOCatalogueDBHelper.GenerateDetailsUsingSavedIDs( _
            m_lEditsetID, _
            m_lCatalogueID, _
            m_eCatalogueLevelID, _
            ELinkTermType.elCopiesInformation, _
            m_colCopiesInformation)
            
    'Location of originals
    Set m_colLocationOfOriginals = _
        m_oBOCatalogueDBHelper.GenerateDetailsUsingSavedIDs( _
            m_lEditsetID, _
            m_lCatalogueID, _
            m_eCatalogueLevelID, _
            ELinkTermType.elLocationOfOriginals, _
            m_colLocationOfOriginals)
   
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
'   d6   wdp      03Dec1999      Added copying of record status flags
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim oBOCatalogueAccessld As BOCatalogueAccessLd
    
    sSource = Me.Name & ".LoadFromDatabase"
    
    Set oBOCatalogueAccessld = New BOCatalogueAccessLd
        
    'Set the properties
    oBOCatalogueAccessld.EditsetID = m_lEditsetID
    oBOCatalogueAccessld.CatalogueID = m_lCatalogueID
    oBOCatalogueAccessld.CatalogueLevelID = m_eCatalogueLevelID
    oBOCatalogueAccessld.SectionType = m_eSectionType
    
    'Load the details
    oBOCatalogueAccessld.Load
    
    'Move details from the loader class to current object
    
    'Simple strings
    m_vAccessConditionsID = oBOCatalogueAccessld.AccessConditionID
    m_vClosureType = oBOCatalogueAccessld.ClosureType
    m_vClosureCode = oBOCatalogueAccessld.ClosureCode
    m_vRecordOpeningDate = oBOCatalogueAccessld.RecordOpeningDate
    m_vClosureStatus = oBOCatalogueAccessld.ClosureStatus
    m_vRestrictionsOnUse = oBOCatalogueAccessld.RestrictionsOnUse
    m_vCustodialHistory = oBOCatalogueAccessld.CustodialHistory
    m_vAccumulationDates = oBOCatalogueAccessld.AccumulationDate
    m_vAppraisalDestructionInfo = oBOCatalogueAccessld.AppraisalDestructionInfo
    m_vAccruals = oBOCatalogueAccessld.Accruals
                
    With oBOCatalogueAccessld
        m_bAccessionedAsCopy = .AccessionedAsCopy
        m_bExtractAtTransfer = .ExtractAtTransfer
        m_bAR = .AR
        m_bOrderable = .Orderable
    End With
    
    'Unknown Descriptions
    m_vUnknownImmedieteSourceOfAcquisition = oBOCatalogueAccessld.UnknownImmedieteSourceOfAcquisition
    m_vUnknownCopiesInformation = oBOCatalogueAccessld.UnknownCopiesInformation
    
    'Collections
    Set m_colImmedieteSourceOfAcquisitionPerson = oBOCatalogueAccessld.ImmedieteSourceOfAcquisitionPerson
    Set m_colImmedieteSourceOfAcquisitionCorporateBody = oBOCatalogueAccessld.ImmedieteSourceOfAcquisitionCorporateBody
    Set m_colCopiesInformation = oBOCatalogueAccessld.CopiesInformation
    Set m_colLocationOfOriginals = oBOCatalogueAccessld.LocationOfOriginals
    
    Set oBOCatalogueAccessld = Nothing
    
    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub

Private Sub PROCATTransientInterface_DeleteState(Optional t_sInstanceReference As String)
'
'Description: Deletes the state of the access details
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
'Description: gets the the current state of the access details of the catalogue whose details are
'set in the properties
'
' Amendment history:
'   d1   ndk      25Aug1999      Created
'   d6   wdp      03Dec1999      Added retrieval of record status flags
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
        'Common properties
        m_eSectionType = CInt(GetDataFromSerialisedString(sSerialisedString))
        m_lEditsetID = CLng(GetDataFromSerialisedString(sSerialisedString))
        m_lCatalogueID = CLng(GetDataFromSerialisedString(sSerialisedString))
        m_eCatalogueLevelID = GetCatalogueLevelEnum(CInt(GetDataFromSerialisedString(sSerialisedString)))
        
        'Simple strings
        m_vAccessConditionsID = GetDataFromSerialisedString(sSerialisedString)
        m_vClosureType = GetDataFromSerialisedString(sSerialisedString)
        m_vClosureCode = GetDataFromSerialisedString(sSerialisedString)
        m_vRecordOpeningDate = GetDataFromSerialisedString(sSerialisedString)
        m_vClosureStatus = GetDataFromSerialisedString(sSerialisedString)
        m_vRestrictionsOnUse = GetDataFromSerialisedString(sSerialisedString)
        m_vCustodialHistory = GetDataFromSerialisedString(sSerialisedString)
        m_vAccumulationDates = GetDataFromSerialisedString(sSerialisedString)
        m_vAppraisalDestructionInfo = GetDataFromSerialisedString(sSerialisedString)
        m_vAccruals = GetDataFromSerialisedString(sSerialisedString)
            
        'record status flags
        m_bAccessionedAsCopy = CBool(GetDataFromSerialisedString(sSerialisedString))
        m_bExtractAtTransfer = CBool(GetDataFromSerialisedString(sSerialisedString))
        m_bAR = CBool(GetDataFromSerialisedString(sSerialisedString))
        m_bOrderable = CBool(GetDataFromSerialisedString(sSerialisedString))
        
        'Unknown Descriptions
        m_vUnknownImmedieteSourceOfAcquisition = GetDataFromSerialisedString(sSerialisedString)
        m_vUnknownCopiesInformation = GetDataFromSerialisedString(sSerialisedString)
    
        'Collections
        Set m_colImmedieteSourceOfAcquisitionPerson = m_oBOCatalogueDBHelper.ExtractLinkItemsFromSerialIsedString(sSerialisedString)
        Set m_colImmedieteSourceOfAcquisitionCorporateBody = m_oBOCatalogueDBHelper.ExtractLinkItemsFromSerialIsedString(sSerialisedString)
        Set m_colCopiesInformation = m_oBOCatalogueDBHelper.ExtractLinkItemsFromSerialIsedString(sSerialisedString)
        Set m_colLocationOfOriginals = m_oBOCatalogueDBHelper.ExtractLinkItemsFromSerialIsedString(sSerialisedString)
            
        LogMessage "Read " & Me.Name & " state" & sInstanceReference, EProfile.pComplete
    End If
    
    Set ofwTransientManager = Nothing
    
    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub

Public Sub PROCATTransientInterface_SetState(Optional t_sInstanceReference As String)
'
'Description: Saves the the current state of the content details of the catalogue whose details are
'set in the properties
'
' Amendment history:
'   d1   ndk      25Aug1999      Created
'   d6   wdp      03Dec1999      Added storage of record status flags
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

    'Common Properties
    sSerialisedString = ""
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, CStr(m_eSectionType))
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, CStr(m_lEditsetID))
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, CStr(m_lCatalogueID))
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, CStr(GetCatalogueLevelCode(m_eCatalogueLevelID)))
    
    'Simple strings
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, VariantStringAsString(m_vAccessConditionsID))
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, VariantStringAsString(m_vClosureType))
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, VariantStringAsString(m_vClosureCode))
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, VariantStringAsString(m_vRecordOpeningDate))
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, VariantStringAsString(m_vClosureStatus))
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, VariantStringAsString(m_vRestrictionsOnUse))
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, VariantStringAsString(m_vCustodialHistory))
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, VariantStringAsString(m_vAccumulationDates))
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, VariantStringAsString(m_vAppraisalDestructionInfo))
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, VariantStringAsString(m_vAccruals))
    
    'Record status flags
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, CStr(m_bAccessionedAsCopy))
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, CStr(m_bExtractAtTransfer))
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, CStr(m_bAR))
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, CStr(m_bOrderable))
    
    'Unknown Descriptions
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, VariantStringAsString(m_vUnknownImmedieteSourceOfAcquisition))
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, VariantStringAsString(m_vUnknownCopiesInformation))
    
    'Collections
    sSerialisedString = m_oBOCatalogueDBHelper.SerialiseLinkItem(sSerialisedString, m_colImmedieteSourceOfAcquisitionPerson)
    sSerialisedString = m_oBOCatalogueDBHelper.SerialiseLinkItem(sSerialisedString, m_colImmedieteSourceOfAcquisitionCorporateBody)
    sSerialisedString = m_oBOCatalogueDBHelper.SerialiseLinkItem(sSerialisedString, m_colCopiesInformation)
    sSerialisedString = m_oBOCatalogueDBHelper.SerialiseLinkItem(sSerialisedString, m_colLocationOfOriginals)
    
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
'Description: Saves the access details of the catalogue whose details are set in the properties
'
'Properties required to be set before calling this method
'
'CatalogueID
'CatalogueLevelID
'EditSetID
'
' Amendment history:
'   d1   ndk      25Aug1999      Created
'   d3   ndk      03Nov1999      Changed the way closure type and code are handled
'                                Put validations for lengths of text fields
'   d4   ndk      18Nov1999      Update edit_set_entry_type_id after when catalogues is updated
'   d5   ndk      19Nov1999      Used quotes before updating closure status
'   d6   wdp      03Dec1999      Added saving of record status flags
'   d7   ndk      17Dec1999      Record opening date handled properly
'   d8   ndk      29Dec1999      Using tbl_piece/item.open_date column instead of piece/item_extension.Record_Opening_date
'   d9   gdb      13Feb2001      Changed Save function to use oBOCatalogueSummaryLd to pass data between Summary and Acess
'   d10  gdb      13Feb2001      Changed Save function to invoke EditorialFunctions.IsValidOpenDate and isValidStatus

    On Error GoTo ErrorHandler
    
    Dim sSource As String, sErrorMessageForClosureTypeAndCode As String, sErrorMessageForClosureStatus As String
    Dim sConditionForMainAndExtnTables As String, sConditionForLinkTables As String
    Dim sMainTableName As String, sExtensionTableName As String
    
    Dim sRecordOpeningDate As String
    
    sSource = Me.Name & ".Save"
        
    'd9 - need to create summary object so we can extract field covering dates
    ' values not available from Access object
    
    'Business objects
    Dim oBOCatalogueSummaryLd As BOCatalogueSummaryLd
    
    'Instantiate the object
    Set oBOCatalogueSummaryLd = New BOCatalogueSummaryLd
    
    'Set the properties of the catalogue to be displayed
    oBOCatalogueSummaryLd.CatalogueID = m_lCatalogueID
    oBOCatalogueSummaryLd.CatalogueLevelID = m_eCatalogueLevelID
    oBOCatalogueSummaryLd.EditsetID = m_lEditsetID
    oBOCatalogueSummaryLd.SectionType = m_eSectionType
    
    'Get the details
    oBOCatalogueSummaryLd.Load

    ' the following are the start & final date of field covering dates
    Dim vStart As Variant
    vStart = oBOCatalogueSummaryLd.StartDate
    Dim vEnd As Variant
    vEnd = oBOCatalogueSummaryLd.EndDate

    
    'Generic condition
    sConditionForMainAndExtnTables = _
        gk_sDBCOLEditSetID & "=" & m_lEditsetID & " AND " & _
        GetColumnName(ECATALOGUECOLUMNTYPE.cctIDColumn, m_eCatalogueLevelID) & "=" & m_lCatalogueID
        
    sExtensionTableName = GetTableName(cttExtensionTable, SECeditorial, m_eCatalogueLevelID)
    sMainTableName = GetTableName(cttMainTable, SECeditorial, m_eCatalogueLevelID)
    
    'Generic condition
    sConditionForLinkTables = _
        gk_sDBCOLEditSetID & "=" & m_lEditsetID & " AND " & _
        gk_sDBCOLCatalogueID & "=" & m_lCatalogueID & " AND " & _
        gk_sDBCOLCatalogueLevel & "=" & GetCatalogueLevelCode(m_eCatalogueLevelID)
    
    'Access condition id
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(m_eCatalogueLevelID, cpAccessCondition) Then
        m_oBOCatalogueDBHelper.UpdateMainExtnTable sConditionForMainAndExtnTables, sExtensionTableName, _
            gk_sDBCOLAccessConditionID, m_vAccessConditionsID
    End If

    'Closure type and Closure code
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(m_eCatalogueLevelID, cpClosureType) And _
        m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(m_eCatalogueLevelID, cpClosureCode) Then
        
        'Check the validity of closure type and closure code
        If Not IsValidClosureTypeAndCode(m_vClosureType, m_vClosureCode, _
                sErrorMessageForClosureTypeAndCode, vStart, vEnd) Then
            RaiseError sSource, peInvalidCombinationOfClosureTypeAndCode, sErrorMessageForClosureTypeAndCode
        End If
    
        'closure type
        m_oBOCatalogueDBHelper.UpdateMainExtnTable sConditionForMainAndExtnTables, sMainTableName, _
            gk_sDBCOLClosureType, QuoteValueWithQuotes(m_vClosureType)
        
        'closure code
        m_oBOCatalogueDBHelper.UpdateMainExtnTable sConditionForMainAndExtnTables, sMainTableName, _
             gk_sDBCOLClosureCode, QuoteValueWithQuotes(m_vClosureCode)
        
    End If
    
    'Record Opening Date
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(m_eCatalogueLevelID, cpRecordOpeningDate) Then
        
        'd10 call IsValidOpenDate function that calculate the correct open date depending
        ' on the values of field closure type & code
        sRecordOpeningDate = IsValidOpenDate(m_vClosureType, m_vClosureCode, _
                 m_vRecordOpeningDate, vStart, vEnd)
        
        m_oBOCatalogueDBHelper.UpdateMainExtnTable sConditionForMainAndExtnTables, sMainTableName, _
             gk_sDBCOLOpenDate, sRecordOpeningDate
             
    End If
    
    'Closure Status (Main table)
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(m_eCatalogueLevelID, cpClosureStatus) Then
    
        'Check the validity of closure status
        If Not IsValidClosureStatus(m_vClosureType, m_vClosureCode, _
                sRecordOpeningDate, m_vClosureStatus, sErrorMessageForClosureTypeAndCode, vStart, vEnd) Then
            RaiseError sSource, peInvalidCombinationOfClosureTypeAndCode, sErrorMessageForClosureTypeAndCode
        End If
    
    
        m_oBOCatalogueDBHelper.UpdateMainExtnTable sConditionForMainAndExtnTables, sMainTableName, _
             gk_sDBCOLClosureStatus, QuoteValueWithQuotes(m_vClosureStatus)
    End If
    
    'Restrictions on use
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(m_eCatalogueLevelID, cpRestrictionsOnUse) Then
        m_oBOCatalogueDBHelper.UpdateMainExtnTable sConditionForMainAndExtnTables, sExtensionTableName, _
             gk_sDBCOLRestrictionsOnUse, QuoteValueWithQuotes(m_vRestrictionsOnUse)
    End If
    
    'Accumulation Date text
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(m_eCatalogueLevelID, cpAccumulationDates) Then
        m_oBOCatalogueDBHelper.UpdateMainExtnTable sConditionForMainAndExtnTables, sExtensionTableName, _
             gk_sDBCOLAccumulationdates, QuoteValueWithQuotes(m_vAccumulationDates)
    End If
    
    'Custodial history
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(m_eCatalogueLevelID, cpCustodialHistory) Then
        
        If Len(m_vCustodialHistory) > m_oBOCatalogueElementMatrix.CataloguePropertyLength(cpCustodialHistory) Then
            RaiseError sSource, peIncorrectLength, _
                "Too many characters for item " & _
                m_oFWApplicationData.GetDisplayInformation( _
                    ESharedResourceFile.srfTabDisplayAccessDetailsLeftHeaderCustodialHistory)
        End If
        
        m_oBOCatalogueDBHelper.UpdateMainExtnTable sConditionForMainAndExtnTables, _
            GetTableName(cttCustodialHistoryTable, SECeditorial, m_eCatalogueLevelID), _
            gk_sDBCOLCustodialHistory, QuoteValueWithQuotes(m_vCustodialHistory)
    End If
    
    'App Dest Information
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(m_eCatalogueLevelID, cpAppraisal_DestructionInformation) Then
        
        If Len(m_vAppraisalDestructionInfo) > _
            m_oBOCatalogueElementMatrix.CataloguePropertyLength(cpAppraisal_DestructionInformation) Then
            RaiseError sSource, peIncorrectLength, _
                "Too many characters for item " & _
                m_oFWApplicationData.GetDisplayInformation( _
                    ESharedResourceFile.srfTabDisplayAccessDetailsLeftHeaderAppDestructionInformation)
        End If
    
        m_oBOCatalogueDBHelper.UpdateMainExtnTable sConditionForMainAndExtnTables, _
            GetTableName(cttAppDestructionTable, SECeditorial, m_eCatalogueLevelID), _
            gk_sDBCOLAppraisalDestructionInformation, QuoteValueWithQuotes(m_vAppraisalDestructionInfo)
    End If

    'Accruals text
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(m_eCatalogueLevelID, cpAccruals) Then
        If Len(m_vAccruals) > m_oBOCatalogueElementMatrix.CataloguePropertyLength(cpAccruals) Then
            RaiseError sSource, peIncorrectLength, _
                "Too many characters for item " & _
                m_oFWApplicationData.GetDisplayInformation( _
                    ESharedResourceFile.srfTabDisplayAccessDetailsLeftHeaderAccruals)
        End If
            
        m_oBOCatalogueDBHelper.UpdateMainExtnTable sConditionForMainAndExtnTables, _
            GetTableName(cttAccrualsTable, SECeditorial, m_eCatalogueLevelID), _
            gk_sDBCOLAccruals, QuoteValueWithQuotes(m_vAccruals)
    End If

    'Immediete source of acquisition
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(m_eCatalogueLevelID, cpImmediateSourceOfAcquisition) Then
        
        'Person
        m_oBOCatalogueDBHelper.UpdateLinkTable _
            m_lEditsetID, m_lCatalogueID, m_eCatalogueLevelID, elImmedieteSourceOfAcquisitionPerson, _
            m_colImmedieteSourceOfAcquisitionPerson, True
        
        'Corporate body
        m_oBOCatalogueDBHelper.UpdateLinkTable _
            m_lEditsetID, m_lCatalogueID, m_eCatalogueLevelID, elImmedieteSourceOfAcquisitionCorporateBody, _
            m_colImmedieteSourceOfAcquisitionCorporateBody, True
        
        'Unknown description
        m_oBOCatalogueDBHelper.UpdateUnknownDescription _
            m_lEditsetID, m_lCatalogueID, m_eCatalogueLevelID, cpImmediateSourceOfAcquisition, _
            m_vUnknownImmedieteSourceOfAcquisition
        
    End If
    
    'Location of originals
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(m_eCatalogueLevelID, cpLocationOfOriginals) Then
        m_oBOCatalogueDBHelper.UpdateLinkTable _
            m_lEditsetID, m_lCatalogueID, m_eCatalogueLevelID, elLocationOfOriginals, m_colLocationOfOriginals, True

    End If
        
    'Copies Information
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(m_eCatalogueLevelID, cpCopiesInformation) Then
        m_oBOCatalogueDBHelper.UpdateLinkTable _
            m_lEditsetID, m_lCatalogueID, m_eCatalogueLevelID, elCopiesInformation, m_colCopiesInformation, True

        'Unknown description
        m_oBOCatalogueDBHelper.UpdateUnknownDescription _
            m_lEditsetID, m_lCatalogueID, m_eCatalogueLevelID, cpCopiesInformation, m_vUnknownCopiesInformation
    End If
    
    

    'Record status flags for pieces and items
    If m_eCatalogueLevelID >= ctPiece Then
        Dim lFlagsValue As Long
        lFlagsValue = 0
        If m_bAccessionedAsCopy = True Then
            lFlagsValue = lFlagsValue Or gk_lFLAGAccessionAsCopy
        End If
        If m_bExtractAtTransfer = True Then
            lFlagsValue = lFlagsValue Or gk_lFLAGExtractAtTransfer
        End If
        If m_bAR = True Then
            lFlagsValue = lFlagsValue Or gk_lFLAGAR
        End If
        If m_bOrderable = True Then
            lFlagsValue = lFlagsValue Or gk_lFLAGOrderable
        Else
            lFlagsValue = lFlagsValue Or gk_lFLAGNotOrderable
        End If

        m_oBOCatalogueDBHelper.UpdateRecordStatus m_lEditsetID, m_lCatalogueID, m_eCatalogueLevelID, lFlagsValue
    End If
        
    'update the status that catalogue was modified
    m_oBOCatalogueDBHelper.SetCatalogueStatusAsModified _
        GetTableName(cttExtensionTable, SECeditorial, m_eCatalogueLevelID), _
        gk_sDBCOLEditSetID & "=" & m_lEditsetID & " AND " & _
        GetColumnName(cctIDColumn, m_eCatalogueLevelID) & "=" & m_lCatalogueID
    
    Exit Sub
    
ErrorHandler:
    ConvertError sSource
End Sub




