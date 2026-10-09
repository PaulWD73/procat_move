VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BOCatalogueContent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    BOCatalogueContent.cls
' System:       PROCAT, PRO
' Copyright:    (C) Quidnunc Limited
'
' Description:  Manages the functionality related to the content details
'               of the catalogues at all levels
'
' Amendment history:
'   d1   ndk      25Aug1999      Created
'   d2   ndk      20Oct1999      Added error handling for duplicate link items
'   d3   ndk      03Nov1999      Put validations for lengths of text fields
'   d4   ndk      03Nov1999      Added validation to avoid the catalogue linking to itself when
'                                choosing related/separated material
'   d5   ndk      18Nov1999      Update edit_set_entry_type_id after when catalogues is updated
'   d6   ndk      24Dec1999      Validating the length of scope and content and arrangement
'   d7   ndk      24Dec1999      Correction in error for validation of arrangement
'   d8   crm      02Feb2000       Updated code to provide links for Leaflets to catalogues
'   d9   jcc      24Mar2000       Changed error message
'   d10  tmj      29Mar2000      Fixed mistake in Save
'   d11  alm      25Sep2001      Moved property 'Notes' to BOCatalogueSummary
'   d12  alm      26Nov2001      Save modified so that adjacent paragraphs in scopecontent split
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'Implements
Implements PROCATTransientInterface

'Property variables
Private m_eSectionType As ESectionType
Private m_lEditsetID As Long
Private m_lCatalogueID As Long
Private m_eCatalogueLevelID As ECATALOGUELEVEL
Private m_vScopeAndContent As Variant
Private m_vArrangement As Variant
Private m_colPublicationNote As Collection
Private m_vUnknownPublicationNote As Variant
Private m_colUnpublishedFindingAids As Collection
Private m_vUnknownUnpublishedFindingAids As Variant
Private m_colRelatedMaterial As Collection
Private m_vUnknownRelatedMaterial As Variant
Private m_colSeparateMaterial As Collection
Private m_vUnknownSeparateMaterial As Variant

'Other variables
Private m_oBOCatalogueDBHelper As BOCatalogueDBHelper
Private m_oBOCatalogueElementMatrix As BOCatalogueElementMatrix
Private m_oFWApplicationData As FWApplicationData
'   d8   crm      02Feb2000
Private m_colPublicationNoteLeaflet As Collection
Private m_colUnpublishedFindingAidsLeaflet As Collection

Private Sub Initialise()
    Set m_oBOCatalogueDBHelper = New BOCatalogueDBHelper
    Set m_oBOCatalogueElementMatrix = New BOCatalogueElementMatrix
    Set m_oFWApplicationData = New FWApplicationData
    
    Set m_colPublicationNote = New Collection
    Set m_colUnpublishedFindingAids = New Collection
    Set m_colRelatedMaterial = New Collection
    Set m_colSeparateMaterial = New Collection
    '   d8   crm      02Feb2000
    Set m_colPublicationNoteLeaflet = New Collection
    Set m_colUnpublishedFindingAidsLeaflet = New Collection
    
End Sub
Private Sub Terminate()
    Set m_oBOCatalogueDBHelper = Nothing
    Set m_oBOCatalogueElementMatrix = Nothing
    Set m_oFWApplicationData = Nothing
    
    Set m_colPublicationNote = Nothing
    Set m_colUnpublishedFindingAids = Nothing
    Set m_colRelatedMaterial = Nothing
    Set m_colSeparateMaterial = Nothing
    '   d8   crm      02Feb2000
    Set m_colPublicationNoteLeaflet = Nothing
    Set m_colUnpublishedFindingAidsLeaflet = Nothing
End Sub
Private Sub Class_Initialize()
    Initialise
End Sub
Private Sub Class_Terminate()
    Terminate
End Sub
Public Property Get Name()
    Name = "BOCatalogueContent"
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
Public Property Let CatalogueLevelID(ByVal t_m_eCatalogueLevelID As ECATALOGUELEVEL)
    m_eCatalogueLevelID = t_m_eCatalogueLevelID
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
Public Property Get ScopeAndContent() As Variant
    ScopeAndContent = m_vScopeAndContent
End Property
Public Property Get Arrangement() As Variant
    Arrangement = m_vArrangement
End Property
Public Property Get PublicationNote() As Collection
    Set PublicationNote = m_colPublicationNote
End Property
Public Property Get UnknownPublicationNote() As Variant
    UnknownPublicationNote = m_vUnknownPublicationNote
End Property
Public Property Get UnpublishedFindingAids() As Collection
    Set UnpublishedFindingAids = m_colUnpublishedFindingAids
End Property
Public Property Get UnknownUnpublishedFindingAids() As Variant
    UnknownUnpublishedFindingAids = m_vUnknownUnpublishedFindingAids
End Property
Public Property Get RelatedMaterial() As Collection
    Set RelatedMaterial = m_colRelatedMaterial
End Property
Public Property Get UnknownRelatedMaterial() As Variant
    UnknownRelatedMaterial = m_vUnknownRelatedMaterial
End Property
Public Property Get SeparateMaterial() As Collection
    Set SeparateMaterial = m_colSeparateMaterial
End Property
Public Property Get UnknownSeparateMaterial() As Variant
    UnknownSeparateMaterial = m_vUnknownSeparateMaterial
End Property
'   d8   crm      02Feb2000
Public Property Get PublicationNoteLeaflet() As Collection
Set PublicationNoteLeaflet = m_colPublicationNoteLeaflet
End Property
Public Property Get UnpublishedFindingAidsLeaflet() As Collection
Set UnpublishedFindingAidsLeaflet = m_colUnpublishedFindingAidsLeaflet
End Property
Public Property Let ScopeAndContent(ByVal t_vScopeAndContent As Variant)
    m_vScopeAndContent = t_vScopeAndContent
End Property
Public Property Let Arrangement(ByVal t_vArrangement As Variant)
    m_vArrangement = t_vArrangement
End Property
Public Property Let PublicationNote(ByVal t_colPublicationNote As Collection)
    Set m_colPublicationNote = t_colPublicationNote
End Property
Public Property Let UnknownPublicationNote(ByVal t_vUnknownPublicationNote As Variant)
    m_vUnknownPublicationNote = t_vUnknownPublicationNote
End Property
Public Property Let UnpublishedFindingAids(ByVal t_colUnpublishedFindingAids As Collection)
    Set m_colUnpublishedFindingAids = t_colUnpublishedFindingAids
End Property
Public Property Let UnknownUnpublishedFindingAids(ByVal t_vUnknownUnpublishedFindingAids As Variant)
    m_vUnknownUnpublishedFindingAids = t_vUnknownUnpublishedFindingAids
End Property
Public Property Let RelatedMaterial(ByVal t_colRelatedMaterial As Collection)
    Set m_colRelatedMaterial = t_colRelatedMaterial
End Property
Public Property Let UnknownRelatedMaterial(ByVal t_vUnknownRelatedMaterial As Variant)
    m_vUnknownRelatedMaterial = t_vUnknownRelatedMaterial
End Property
Public Property Let SeparateMaterial(ByVal t_colSeparateMaterial As Collection)
    Set m_colSeparateMaterial = t_colSeparateMaterial
End Property
Public Property Let UnknownSeparateMaterial(ByVal t_vUnknownSeparateMaterial As Variant)
    m_vUnknownSeparateMaterial = t_vUnknownSeparateMaterial
End Property
'   d8   crm      02Feb2000
Public Property Let PublicationNoteLeaflet(ByVal t_colPublicationNoteLeaflet As Collection)
Set m_colPublicationNoteLeaflet = t_colPublicationNoteLeaflet
End Property
Public Property Let UnpublishedFindingAidsLeaflet(ByVal t_colUnpublishedFindingAidsLeaflet As Collection)
Set m_colUnpublishedFindingAidsLeaflet = t_colUnpublishedFindingAidsLeaflet
End Property


Public Sub Load( _
    ByVal t_eLoadType As ECatalogueEditFormLoadType, _
    Optional ByVal t_lLinkTermID As Long, _
    Optional ByVal t_eLinkTermType As ELinkTermType, _
    Optional ByVal t_iLinkTermLevelID As Integer, _
    Optional ByVal t_iMainCollectionKey As Integer)
    
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
                AddNewLinkTerm t_lLinkTermID, t_eLinkTermType, t_iLinkTermLevelID
            
            Case ECatalogueEditFormLoadType.etLoadFromSavedStateAndRemoveTermEditCatalogueEntry
                'The user has selected a pick term to remove and the control has come here
            
                'Read the stored state info
                PROCATTransientInterface_GetState
                
                'Remove the link term from the collection
                RemoveLinkTerm t_lLinkTermID, t_eLinkTermType, t_iMainCollectionKey
                
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
    t_eLinkTermType As ELinkTermType, _
    Optional ByVal t_iLinkTermLevelID As Integer)
'
'Description: Adds a link item to the collection
'
'Amendment history:
'   d1   ndk      25Aug1999      Created
'   d2   ndk      20Oct1999      Added error handling for duplicate link items
'   d4   ndk      03Nov1999      Added validation to avoid the catalogue linking to itself when
'                                choosing related/separated material
'   d8   crm      02Feb2000       Updated code to provide links for Leaflets to catalogues
'   d9   jcc      24Mar2000       Changed error message

    On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim recLinkItemRecord  As GenericLinkRecord, recRelatedSeparateMaterial As RelatedSeparateMaterial, varForDuplicateCheck As Variant
        
    sSource = Me.Name & ".AddNewLinkTerm"

    'check whether the catalogue is linking to itself
    If t_eLinkTermType = ELinkTermType.elRelatedMaterial Or t_eLinkTermType = ELinkTermType.elSeparateMaterial Then
        If m_lCatalogueID = t_lLinkTermID And _
            m_eCatalogueLevelID = GetCatalogueLevelEnum(t_iLinkTermLevelID) Then
                
            RaiseError sSource, peTryingToAddLinkToItSelf, _
                "You cannot create a link from a catalogue entry to itself "
                
        End If
    End If

    Select Case t_eLinkTermType
        Case ELinkTermType.elPublicationNote
            recLinkItemRecord.lID = t_lLinkTermID
            recLinkItemRecord.sOptionalDescription = ""
        
            'Add to the collection
            m_colPublicationNote.Add recLinkItemRecord, gk_sPrefixForLinkTerm & recLinkItemRecord.lID
                
        Case ELinkTermType.elUnPublishedFindingAids
            recLinkItemRecord.lID = t_lLinkTermID
            recLinkItemRecord.sOptionalDescription = ""

            'Add to the collection
            m_colUnpublishedFindingAids.Add recLinkItemRecord, gk_sPrefixForLinkTerm & recLinkItemRecord.lID
                
        Case ELinkTermType.elRelatedMaterial
            
            'check for duplicate
            For Each varForDuplicateCheck In m_colRelatedMaterial
                If varForDuplicateCheck.lCatalogueID = t_lLinkTermID And _
                    varForDuplicateCheck.iCatalogueLevelID = t_iLinkTermLevelID Then
                    
                    Err.Raise gk_lErrorNumberDuplicateKeyValueForCollection
                    
                End If
            Next
            
            recRelatedSeparateMaterial.lCatalogueID = t_lLinkTermID
            recRelatedSeparateMaterial.iCatalogueLevelID = t_iLinkTermLevelID
            
            'Add to the collection
            m_colRelatedMaterial.Add recRelatedSeparateMaterial
        
        Case ELinkTermType.elSeparateMaterial
            'check for duplicate
            For Each varForDuplicateCheck In m_colSeparateMaterial
                If varForDuplicateCheck.lCatalogueID = t_lLinkTermID And _
                    varForDuplicateCheck.iCatalogueLevelID = t_iLinkTermLevelID Then
                    
                    Err.Raise gk_lErrorNumberDuplicateKeyValueForCollection
                    
                End If
            Next
            
            recRelatedSeparateMaterial.lCatalogueID = t_lLinkTermID
            recRelatedSeparateMaterial.iCatalogueLevelID = t_iLinkTermLevelID
            
            'Add to the collection
            m_colSeparateMaterial.Add recRelatedSeparateMaterial
        '   d8   crm      02Feb2000
        'Links for leaflet against Publication Note
        Case ELinkTermType.elPublicationNoteLeaflet
            recLinkItemRecord.lID = t_lLinkTermID
            recLinkItemRecord.sOptionalDescription = ""
            
            m_colPublicationNoteLeaflet.Add recLinkItemRecord, gk_sPrefixForLinkTerm & recLinkItemRecord.lID
            
        'Links for leaflet against Unpublished finding aids
        Case ELinkTermType.elUnPublishedFindingAidLeaflet
            recLinkItemRecord.lID = t_lLinkTermID
            recLinkItemRecord.sOptionalDescription = ""
            
            m_colUnpublishedFindingAidsLeaflet.Add recLinkItemRecord, gk_sPrefixForLinkTerm & recLinkItemRecord.lID
            
    End Select
    
    Exit Sub
ErrorHandler:
    If Err.Number = gk_lErrorNumberDuplicateKeyValueForCollection Then
        Dim sErrorMessage As String
        
        sErrorMessage = "You have choosen a duplicate link term for "
        '   d8   crm      02Feb2000
        If t_eLinkTermType = elPublicationNote Or t_eLinkTermType = elPublicationNoteLeaflet Then
            sErrorMessage = sErrorMessage & _
                m_oFWApplicationData.GetDisplayInformation(ESharedResourceFile.srfTabDisplayContentDetailsLeftHeaderPublicationNote)
                
        ElseIf t_eLinkTermType = elUnPublishedFindingAids Or t_eLinkTermType = elUnPublishedFindingAidLeaflet Then
            sErrorMessage = sErrorMessage & _
                m_oFWApplicationData.GetDisplayInformation(ESharedResourceFile.srfTabDisplayContentDetailsLeftHeaderUnpublishedFindingAids)
        
        ElseIf t_eLinkTermType = elRelatedMaterial Then
            sErrorMessage = sErrorMessage & _
                m_oFWApplicationData.GetDisplayInformation(ESharedResourceFile.srfTabDisplayContentDetailsLeftHeaderRelatedMaterial)

        ElseIf t_eLinkTermType = elSeparateMaterial Then
            sErrorMessage = sErrorMessage & _
                m_oFWApplicationData.GetDisplayInformation(ESharedResourceFile.srfTabDisplayContentDetailsLeftHeaderSeparateMaterial)
        End If
        
        RaiseError sSource, peDuplicateLinkItemChoosen, sErrorMessage
    Else
        ConvertError sSource
    End If
End Sub

Private Sub RemoveLinkTerm( _
    ByVal t_lLinkTermID As Long, _
    ByVal t_eLinkTermType As ELinkTermType, _
    Optional ByVal t_iMainCollectionKey As Integer)

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
        Case ELinkTermType.elPublicationNote
            m_colPublicationNote.Remove gk_sPrefixForLinkTerm & t_lLinkTermID
        Case ELinkTermType.elUnPublishedFindingAids
            m_colUnpublishedFindingAids.Remove gk_sPrefixForLinkTerm & t_lLinkTermID
        Case ELinkTermType.elRelatedMaterial
            'No prefix is attached to the collection key for related and separate material
            m_colRelatedMaterial.Remove t_iMainCollectionKey
        Case ELinkTermType.elSeparateMaterial
            m_colSeparateMaterial.Remove t_iMainCollectionKey
        Case ELinkTermType.elPublicationNoteLeaflet
            m_colPublicationNoteLeaflet.Remove gk_sPrefixForLinkTerm & t_lLinkTermID
        Case ELinkTermType.elUnPublishedFindingAidLeaflet
            m_colUnpublishedFindingAidsLeaflet.Remove gk_sPrefixForLinkTerm & t_lLinkTermID
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
'   d8   crm      02Feb2000       Updated code to provide links for Leaflets to catalogues
    On Error GoTo ErrorHandler
    
    Dim sSource As String
    
    sSource = Me.Name & ".PopulateRecordSetsUsingIDsInCollection"
    
    'Publication Note
    Set m_colPublicationNote = _
        m_oBOCatalogueDBHelper.GenerateDetailsUsingSavedIDs( _
            m_lEditsetID, _
            m_lCatalogueID, _
            m_eCatalogueLevelID, _
            ELinkTermType.elPublicationNote, _
            m_colPublicationNote)
    
    'UnPublished Finding Aids
    Set m_colUnpublishedFindingAids = _
        m_oBOCatalogueDBHelper.GenerateDetailsUsingSavedIDs( _
            m_lEditsetID, _
            m_lCatalogueID, _
            m_eCatalogueLevelID, _
            ELinkTermType.elUnPublishedFindingAids, _
            m_colUnpublishedFindingAids)
    '   d8   crm      02Feb2000
    'Publication Note Leaflet
    Set m_colPublicationNoteLeaflet = _
        m_oBOCatalogueDBHelper.GenerateDetailsUsingSavedIDs( _
            m_lEditsetID, _
            m_lCatalogueID, _
            m_eCatalogueLevelID, _
            ELinkTermType.elPublicationNoteLeaflet, _
            m_colPublicationNoteLeaflet)
    'UnPublished Finding Aids Leaflet
    Set m_colUnpublishedFindingAidsLeaflet = _
        m_oBOCatalogueDBHelper.GenerateDetailsUsingSavedIDs( _
            m_lEditsetID, _
            m_lCatalogueID, _
            m_eCatalogueLevelID, _
            ELinkTermType.elUnPublishedFindingAidLeaflet, _
            m_colUnpublishedFindingAidsLeaflet)
    'Note : Related and Separate materials are not set here as they would already be in a shape to be
    '       displayed and no additional data is to be retrieved from the database.
    
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
'   d8   crm      02Feb2000       Updated code to provide links for Leaflets to catalogues
'   d11  alm      25Sep2001      Moved property 'Notes' to BOCatalogueSummary

    On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim oBOCatalogueContentLd As BOCatalogueContentLd
    
    sSource = Me.Name & ".LoadFromDatabase"
    
    Set oBOCatalogueContentLd = New BOCatalogueContentLd
    
    'Set the properties
    oBOCatalogueContentLd.EditsetID = m_lEditsetID
    oBOCatalogueContentLd.CatalogueID = m_lCatalogueID
    oBOCatalogueContentLd.CatalogueLevelID = m_eCatalogueLevelID
    oBOCatalogueContentLd.SectionType = m_eSectionType
    
    'Load the details
    oBOCatalogueContentLd.Load
    
    'Move details from the loader class to current object
    m_vScopeAndContent = oBOCatalogueContentLd.ScopeAndContent
    m_vArrangement = oBOCatalogueContentLd.Arrangement
    
    'Unknown Descriptions
    m_vUnknownPublicationNote = oBOCatalogueContentLd.UnknownPublicationNote
    m_vUnknownUnpublishedFindingAids = oBOCatalogueContentLd.UnknownUnpublishedFindingAids
    m_vUnknownRelatedMaterial = oBOCatalogueContentLd.UnknownRelatedMaterial
    m_vUnknownSeparateMaterial = oBOCatalogueContentLd.UnknownSeparateMaterial

    'Collections
    Set m_colPublicationNote = oBOCatalogueContentLd.PublicationNote
    Set m_colUnpublishedFindingAids = oBOCatalogueContentLd.UnpublishedFindingAids
    Set m_colRelatedMaterial = oBOCatalogueContentLd.RelatedMaterial
    Set m_colSeparateMaterial = oBOCatalogueContentLd.SeparateMaterial
    '   d8   crm      02Feb2000
    Set m_colPublicationNoteLeaflet = oBOCatalogueContentLd.PublicationNoteLeaflet
    Set m_colUnpublishedFindingAidsLeaflet = oBOCatalogueContentLd.UnpubFindingAidLeaflet
    
    Set oBOCatalogueContentLd = Nothing
    
    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub


    

Private Sub PROCATTransientInterface_DeleteState(Optional t_sInstanceReference As String)
'
'Description: Deletes the state of the content details
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
'Description: gets the the current state of the content details of the catalogue whose details are
'set in the properties
'
' Amendment history:
'   d1   ndk      25Aug1999      Created
'   d8   crm      02Feb2000       Updated code to provide links for Leaflets to catalogues
'   d11  alm      25Sep2001      Moved property 'Notes' to BOCatalogueSummary
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
        m_eSectionType = GetCatalogueLevelEnum(CInt(GetDataFromSerialisedString(sSerialisedString)))
        m_lEditsetID = CLng(GetDataFromSerialisedString(sSerialisedString))
        m_lCatalogueID = CLng(GetDataFromSerialisedString(sSerialisedString))
        m_eCatalogueLevelID = GetCatalogueLevelEnum(CInt(GetDataFromSerialisedString(sSerialisedString)))

        'Simple Strings
        m_vScopeAndContent = GetDataFromSerialisedString(sSerialisedString)
        m_vArrangement = GetDataFromSerialisedString(sSerialisedString)

        'Unknown Descriptions
        m_vUnknownPublicationNote = GetDataFromSerialisedString(sSerialisedString)
        m_vUnknownUnpublishedFindingAids = GetDataFromSerialisedString(sSerialisedString)
        m_vUnknownRelatedMaterial = GetDataFromSerialisedString(sSerialisedString)
        m_vUnknownSeparateMaterial = GetDataFromSerialisedString(sSerialisedString)
        
        'Collections
        Set m_colPublicationNote = m_oBOCatalogueDBHelper.ExtractLinkItemsFromSerialIsedString(sSerialisedString)
        Set m_colUnpublishedFindingAids = m_oBOCatalogueDBHelper.ExtractLinkItemsFromSerialIsedString(sSerialisedString)
        Set m_colRelatedMaterial = m_oBOCatalogueDBHelper.ExtractRelatedSeparateMaterialFromSerialIsedString(sSerialisedString)
        Set m_colSeparateMaterial = m_oBOCatalogueDBHelper.ExtractRelatedSeparateMaterialFromSerialIsedString(sSerialisedString)
        '   d8   crm      02Feb2000
        Set m_colPublicationNoteLeaflet = m_oBOCatalogueDBHelper.ExtractLinkItemsFromSerialIsedString(sSerialisedString)
        Set m_colUnpublishedFindingAidsLeaflet = m_oBOCatalogueDBHelper.ExtractLinkItemsFromSerialIsedString(sSerialisedString)
        
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
'   d8   crm      02Feb2000       Updated code to provide links for Leaflets to catalogues
'   d11  alm      25Sep2001      Moved property 'Notes' to BOCatalogueSummary
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

    'Common properties
    sSerialisedString = ""
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, CStr(GetCatalogueLevelCode(m_eSectionType)))
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, CStr(m_lEditsetID))
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, CStr(m_lCatalogueID))
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, CStr(GetCatalogueLevelCode(m_eCatalogueLevelID)))

    'Simple Strings
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, VariantStringAsString(m_vScopeAndContent))
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, VariantStringAsString(m_vArrangement))
    
    'Unknown Descriptions
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, VariantStringAsString(m_vUnknownPublicationNote))
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, VariantStringAsString(m_vUnknownUnpublishedFindingAids))
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, VariantStringAsString(m_vUnknownRelatedMaterial))
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, VariantStringAsString(m_vUnknownSeparateMaterial))
    
    'Collections
    sSerialisedString = m_oBOCatalogueDBHelper.SerialiseLinkItem(sSerialisedString, m_colPublicationNote)
    sSerialisedString = m_oBOCatalogueDBHelper.SerialiseLinkItem(sSerialisedString, m_colUnpublishedFindingAids)
    sSerialisedString = m_oBOCatalogueDBHelper.SerialiseRelatedSeparatedMaterial(sSerialisedString, m_colRelatedMaterial)
    sSerialisedString = m_oBOCatalogueDBHelper.SerialiseRelatedSeparatedMaterial(sSerialisedString, m_colSeparateMaterial)
    '   d8   crm      02Feb2000
    sSerialisedString = m_oBOCatalogueDBHelper.SerialiseLinkItem(sSerialisedString, m_colPublicationNoteLeaflet)
    sSerialisedString = m_oBOCatalogueDBHelper.SerialiseLinkItem(sSerialisedString, m_colUnpublishedFindingAidsLeaflet)
    
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
'Description: Saves the context details of the catalogue whose details are set in the properties
'
'Properties required to be set before calling this method
'
'CatalogueID
'CatalogueLevelID
'EditSetID
'
' Amendment history:
'   d1   ndk      25Aug1999      Created
'   d3   ndk      03Nov1999      Put validations for lengths of text fields
'   d5   ndk      18Nov1999      Update edit_set_entry_type_id after when catalogues is updated
'   d6   ndk      24Dec1999      Validating the length of scope and content and arrangement
'   d8   crm      02Feb2000       Updated code to provide links for Leaflets to catalogues
'   d10  tmj      29Mar2000      Fixed mistake, where m_vScopeAndContent rather than
'                                m_vArrangement
'   d11  alm      25Sep2001      Moved property 'Notes' to BOCatalogueSummary
'   d12  alm      26Nov2001      Adjacent paragraphs in scopecontent split
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim sConditionForMainAndExtnTables As String, sConditionForLinkTables As String

    sSource = Me.Name & ".Save"
    
    'Generic condition
    sConditionForMainAndExtnTables = _
        gk_sDBCOLEditSetID & "=" & m_lEditsetID & " AND " & _
        GetColumnName(ECATALOGUECOLUMNTYPE.cctIDColumn, m_eCatalogueLevelID) & "=" & m_lCatalogueID
            
    'Generic condition
    sConditionForLinkTables = _
        gk_sDBCOLEditSetID & "=" & m_lEditsetID & " AND " & _
        gk_sDBCOLCatalogueID & "=" & m_lCatalogueID & " AND " & _
        gk_sDBCOLCatalogueLevel & "=" & GetCatalogueLevelCode(m_eCatalogueLevelID)
    
    'Properties always present in extension table
    'Scope and content
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(m_eCatalogueLevelID, cpScopeAndContent) Then
        
        'scope and content is must for piece and item
        If m_eCatalogueLevelID = ctPiece Or m_eCatalogueLevelID = ctItem Then
            If VariantStringAsString(m_vScopeAndContent) = "" Then
                RaiseError sSource, peScopeAndContentCanNotBeEmpty, _
                    "Can not be null : " & _
                    m_oFWApplicationData.GetDisplayInformation( _
                        ESharedResourceFile.srfTabDisplayContentDetailsLeftHeaderScopeAndContent)
            End If
        End If
    
        If Len(VariantStringAsString(m_vScopeAndContent)) > m_oBOCatalogueElementMatrix.CataloguePropertyLength(cpScopeAndContent) Then
            RaiseError sSource, peIncorrectLength, _
                "Too many characters for : " & _
                m_oFWApplicationData.GetDisplayInformation( _
                    ESharedResourceFile.srfTabDisplayContentDetailsLeftHeaderScopeAndContent)
        End If
        
        m_oBOCatalogueDBHelper.UpdateMainExtnTable sConditionForMainAndExtnTables, GetTableName(cttScopeAndContentTable, SECeditorial, m_eCatalogueLevelID), _
            gk_sDBCOLScopeAndContent, QuoteValueWithQuotes(SplitParagraphs(m_vScopeAndContent))
    End If
    
    'Arrangement
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(m_eCatalogueLevelID, cpArrangement) Then
        
        If Len(VariantStringAsString(m_vArrangement)) > m_oBOCatalogueElementMatrix.CataloguePropertyLength(cpArrangement) Then
            RaiseError sSource, peIncorrectLength, _
                "Too many characters for : " & _
                m_oFWApplicationData.GetDisplayInformation( _
                    ESharedResourceFile.srfTabDisplayContentDetailsLeftHeaderArrangement)
        End If
        
        m_oBOCatalogueDBHelper.UpdateMainExtnTable sConditionForMainAndExtnTables, GetTableName(cttArrangementTable, SECeditorial, m_eCatalogueLevelID), _
            gk_sDBCOLArrangement, QuoteValueWithQuotes(m_vArrangement)
    End If

    'Publication Note
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(m_eCatalogueLevelID, cpPublicationNote) Then
        m_oBOCatalogueDBHelper.UpdateLinkTable _
            m_lEditsetID, m_lCatalogueID, m_eCatalogueLevelID, elPublicationNote, m_colPublicationNote, True
        '   d8   crm      02Feb2000
        m_oBOCatalogueDBHelper.UpdateLinkTable _
            m_lEditsetID, m_lCatalogueID, m_eCatalogueLevelID, elPublicationNoteLeaflet, m_colPublicationNoteLeaflet, True
            
        'Unknown description
        m_oBOCatalogueDBHelper.UpdateUnknownDescription _
            m_lEditsetID, m_lCatalogueID, m_eCatalogueLevelID, cpPublicationNote, m_vUnknownPublicationNote
        
    End If
    
    'Unpublished Finding Aids
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(m_eCatalogueLevelID, cpUnpublishedFindingAids) Then
        m_oBOCatalogueDBHelper.UpdateLinkTable _
            m_lEditsetID, m_lCatalogueID, m_eCatalogueLevelID, elUnPublishedFindingAids, m_colUnpublishedFindingAids, True
        
        '   d8   crm      02Feb2000
        m_oBOCatalogueDBHelper.UpdateLinkTable _
            m_lEditsetID, m_lCatalogueID, m_eCatalogueLevelID, elUnPublishedFindingAidLeaflet, m_colUnpublishedFindingAidsLeaflet, True
        
        'Unknown description
        m_oBOCatalogueDBHelper.UpdateUnknownDescription _
            m_lEditsetID, m_lCatalogueID, m_eCatalogueLevelID, cpUnpublishedFindingAids, m_vUnknownUnpublishedFindingAids
        
    End If
    
    'Related Material
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(m_eCatalogueLevelID, cpRelatedMaterials) Then
        m_oBOCatalogueDBHelper.UpdateRelatedSeparateMaterialLinks _
            m_lEditsetID, m_lCatalogueID, m_eCatalogueLevelID, elRelatedMaterial, m_colRelatedMaterial
        
        'Unknown description
        m_oBOCatalogueDBHelper.UpdateUnknownDescription _
            m_lEditsetID, m_lCatalogueID, m_eCatalogueLevelID, cpRelatedMaterials, m_vUnknownRelatedMaterial
        
    End If
    
    'Separate Material
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(m_eCatalogueLevelID, cpSeperatedMaterials) Then
        m_oBOCatalogueDBHelper.UpdateRelatedSeparateMaterialLinks _
            m_lEditsetID, m_lCatalogueID, m_eCatalogueLevelID, elSeparateMaterial, m_colSeparateMaterial
        
        'Unknown description
        m_oBOCatalogueDBHelper.UpdateUnknownDescription _
            m_lEditsetID, m_lCatalogueID, m_eCatalogueLevelID, cpSeperatedMaterials, m_vUnknownSeparateMaterial
        
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








