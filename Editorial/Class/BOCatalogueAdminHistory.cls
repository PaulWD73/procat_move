VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BOCatalogueAdminHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    BOCatalogueAdminHistory.cls
' System:       PROCAT, PRO
' Copyright:    (C) Quidnunc Limited
'
' Description:  Manages the functionality related to the admin history
'               of the catalogues at all levels
'
' Amendment history:
'   d1   ndk      25Aug1999      Created
'   d2   ndk      03Nov1999      Put validations for lengths of text fields
'   d3   ndk      18Nov1999      Update edit_set_entry_type_id after when catalogues is updated
'   d4   ndk      07Nov1999      Added persistence
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'Implements
Implements PROCATTransientInterface

'Property variables
Private m_eSectionType As ESectionType
Private m_lEditsetID As Long
Private m_lCatalogueID As Long
Private m_eCatalogueLevelID As ECATALOGUELEVEL
Private m_vAdminHistory As Variant

'Business objects and framework objects
Private m_oFWApplicationData As FWApplicationData
Private m_oBOCatalogueDBHelper As BOCatalogueDBHelper
Private m_oBOCatalogueElementMatrix As BOCatalogueElementMatrix

Private Sub Initialise()
    Set m_oFWApplicationData = New FWApplicationData
    Set m_oBOCatalogueDBHelper = New BOCatalogueDBHelper
    Set m_oBOCatalogueElementMatrix = New BOCatalogueElementMatrix
End Sub
Private Sub Terminate()
    Set m_oFWApplicationData = Nothing
    Set m_oBOCatalogueDBHelper = Nothing
    Set m_oBOCatalogueElementMatrix = Nothing
End Sub
Private Sub Class_Initialize()
    Initialise
End Sub
Private Sub Class_Terminate()
    Terminate
End Sub
Public Property Get Name()
    Name = "BOCatalogueAdminHistory"
End Property
Public Property Let SectionType(ByVal t_eSectionType As ESectionType)
    m_eSectionType = t_eSectionType
End Property
Public Property Let EditSetId(ByVal t_lEditSetID As Long)
    m_lEditsetID = t_lEditSetID
End Property
Public Property Let CatalogueID(ByVal t_lCatalogueID As Long)
    m_lCatalogueID = t_lCatalogueID
End Property
Public Property Let CatalogueLevelID(ByVal t_eCatalogueLevelID As ECATALOGUELEVEL)
    m_eCatalogueLevelID = t_eCatalogueLevelID
End Property
Public Property Let AdminHistory(ByVal t_vAdminHistory As Variant)
    m_vAdminHistory = t_vAdminHistory
End Property
Public Property Get AdminHistory() As Variant
    AdminHistory = m_vAdminHistory
End Property

Public Sub Load( _
    ByVal t_eLoadType As ECatalogueEditFormLoadType)
   
'
'Description: Loads the details of the bo,depending on the value of t_eLoadType either the details are
'loaded from the loader class or from the stored stage
'
'Amendment history:
'   d1   ndk      25Aug1999      Created
'   d4   ndk      07Nov1999      Persistence related changes
'
    On Error GoTo errorhandler
    
    Dim sSource As String
    
    sSource = Me.Name & ".Load"
    
    'The user clicked on Edit button and control has come here
    If t_eLoadType = ECatalogueEditFormLoadType.etLoadFromDatabaseEditCatalogueEntry Then
        
        'Delete the state information (if any)
        PROCATTransientInterface_DeleteState
        
        'Load from the loader class
        LoadFromDatabase
    Else
        'it will be
        'ECatalogueEditFormLoadType.etLoadFromSavedStateNoChangeEditCatalogueEntry
                
        'Just Read the stored state info
        PROCATTransientInterface_GetState
                
    End If

    Exit Sub

errorhandler:
    ConvertError sSource
End Sub

Private Sub LoadFromDatabase()
'
'Description: Loads the object details from the loader class
'
' Amendment history:
'   d1   ndk      25Aug1999      Created
'
    On Error GoTo errorhandler
    
    Dim sSource As String
    Dim oBOCatalogueAdminHistoryLd As BOCatalogueAdminHistoryLd
    
    sSource = Me.Name & ".LoadFromDatabase"
    
    Set oBOCatalogueAdminHistoryLd = New BOCatalogueAdminHistoryLd
    
    'Set the properties
    oBOCatalogueAdminHistoryLd.EditSetId = m_lEditsetID
    oBOCatalogueAdminHistoryLd.CatalogueID = m_lCatalogueID
    oBOCatalogueAdminHistoryLd.CatalogueLevelID = m_eCatalogueLevelID
    oBOCatalogueAdminHistoryLd.SectionType = m_eSectionType
    
    'Load the details
    oBOCatalogueAdminHistoryLd.Load
    
    'Move details from the loader class to current object
    m_vAdminHistory = oBOCatalogueAdminHistoryLd.AdminHistory
    
    Set oBOCatalogueAdminHistoryLd = Nothing
    
    Exit Sub
errorhandler:
    ConvertError sSource
End Sub

Public Sub Save()
'
'Description: Saves the admin history details of the catalogue whose details are set in the properties
'
'Properties required to be set before calling this method
'
'CatalogueID
'CatalogueLevelID
'EditSetID
'
' Amendment history:
'   d1   ndk      25Aug1999      Created
'   d2   ndk      03Nov1999      Put validations for lengths of text fields
'   d3   ndk      18Nov1999      Update edit_set_entry_type_id after when catalogues is updated
'
    On Error GoTo errorhandler
    
    Dim sSource As String
    Dim sConditionForMainAndExtnTables As String
    
    sSource = Me.Name & ".Save"

    'Generic condition
    sConditionForMainAndExtnTables = _
        gk_sDBCOLEditSetID & "=" & m_lEditsetID & " AND " & _
        GetColumnName(ECATALOGUECOLUMNTYPE.cctIDColumn, m_eCatalogueLevelID) & "=" & m_lCatalogueID
        
    'Admin History
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(m_eCatalogueLevelID, _
        cpAdministrative_BiographicalHistory) Then
        
        If Len(m_vAdminHistory) > m_oBOCatalogueElementMatrix.CataloguePropertyLength( _
                cpAdministrative_BiographicalHistory) Then
                
            RaiseError sSource, peIncorrectLength, _
                "Too many characters for item " & _
                m_oFWApplicationData.GetDisplayInformation( _
                    ESharedResourceFile.srfTabDisplayAdminHistoryLeftHeaderAdminHistory)
        End If
        
        m_oBOCatalogueDBHelper.UpdateMainExtnTable sConditionForMainAndExtnTables, GetTableName(cttAdminHistoryTable, SECEditorial, m_eCatalogueLevelID), _
             gk_sDBCOLAdministrativeHistory, QuoteValueWithQuotes(m_vAdminHistory)
    End If
    
    'update the status that catalogue was modified
    m_oBOCatalogueDBHelper.SetCatalogueStatusAsModified _
        GetTableName(cttExtensionTable, SECEditorial, m_eCatalogueLevelID), _
        gk_sDBCOLEditSetID & "=" & m_lEditsetID & " AND " & _
        GetColumnName(cctIDColumn, m_eCatalogueLevelID) & "=" & m_lCatalogueID
    
    Exit Sub
errorhandler:
    ConvertError sSource
End Sub
    

Private Sub PROCATTransientInterface_DeleteState(Optional t_sInstanceReference As String)
'
'Description: Deletes the state of the adminhistory details
'
' Amendment history:
'   d4   ndk      07Nov1999      Created
'
    On Error GoTo errorhandler
    
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
errorhandler:
    ConvertError sSource
End Sub


Public Sub PROCATTransientInterface_GetState(Optional t_sInstanceReference As String)
'
'Description: gets the the current state of the adminhistory details of the catalogue whose details are
'set in the properties
'
' Amendment history:
'   d4   ndk      07Nov1999      Created
'
    On Error GoTo errorhandler
    
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
        m_vAdminHistory = GetDataFromSerialisedString(sSerialisedString)

        LogMessage "Read " & Me.Name & " state" & sInstanceReference, EProfile.pComplete
    End If
    
    Set ofwTransientManager = Nothing
    
    Exit Sub
errorhandler:
    ConvertError sSource

End Sub

Public Sub PROCATTransientInterface_SetState(Optional t_sInstanceReference As String)
'
'Description: Saves the the current state of the adminhistory details of the catalogue whose details are
'set in the properties
'
' Amendment history:
'   d4   ndk      07Nov1999      Created
'
    On Error GoTo errorhandler
    
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
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, VariantStringAsString(m_vAdminHistory))

    'Store in the transient manager
    ofwTransientManager.AddData _
        GetObjectContext.Item(gk_sSessionObject).SessionID, sInstanceReference, sSerialisedString

    LogMessage "Stored " & Me.Name & sInstanceReference, EProfile.pComplete

    Set ofwTransientManager = Nothing
    
    Exit Sub
    
errorhandler:
    ConvertError sSource
End Sub
