VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BOFilteredEditsetList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    BOFilteredEditsetList.cls
' System:       PROCAT, PRO
' Copyright:    (C) Quidnunc Limited
'
' Description:
'
'       Business object Displays a filtered editset list. This business object
'   internally implements and manipulates the transient persistence of state
'
' Amendment history:
'
' d1    ksk     04Oct1999       Created.
' d2    ndk     21Oct1999       Moved enum ECATALOGUEENTRYSTATUS to FWTypes
' d5    ksk     22Oct1999       Changed Catalogue_level_no to Level_No
' d6    ksk     12Nov1999       Removed check for number if get does not find
'                               a number then set to defaults
' d7    ksk     07Dec1999       Changed the name of public method GetEditsetDescription
'                               to GetEditsetStatusDescription.
'                               Changed the description text
' d8    spb     08Feb2000       In GetEditSetEntries, retrieve a max of
'                               gk_lMaxEditsetSizeToRetrieve.
' d9    tmj     03Apr2000       Added parameter in GetEditsetEntries so full title
'                               length is returned
' d10   daj     12Jan2001       Modified object to position the edit set viewer at a particular entry,
'                               instead of a page
' d11   pnp     17Jan2001       Added CatalogueViewFiltered functionality
' d12   pnp     17Jan2001       Added GetEditsetViewList() functionality
' d13   pnp     17Jan2001       Added View functionality to PROCATTransientInterface
' d14   pnp     19Jan2001       Added GetEditsetColumnList functionality
' d15   pnp     22Jan2001       Added AddColumnHeading functionality
' d16   pnp     23Jan2001       Added AddContentCell functionality
' d17   pnp     24Jan2001       Added Get EditSetView functionality
' d18   pnp     08Feb2001       Added EditSetOrderList functionality
' d19   pnp     16Feb2001       Updated GetEditSetEntries function
' d20   alm     19Sep2001       Changed the default ordering to KeyOrder
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Private m_iMaxPerPage As Integer
Private m_eStatus  As ECATALOGUEENTRYSTATUS
Private m_eCatalogueLevel As ECATALOGUELEVEL
Private m_lEditsetID As Long
Private m_lSessionID As Long

' d10 - Added
Private m_sPositionCatalogueReference As String
Private m_lEditsetPosition As Long
Private m_bPositionValid As Boolean

' d11 - Added
Private m_eCatalogueView As ECATALOGUEVIEW
'd18 - Added
Private m_eCatalogueOrder As ECATALOGUEORDER




Implements PROCATTransientInterface
Implements ObjectControl





''''''''''''''''''''''''''''''''''''''''''''''Properties'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Name() As String
    Name = "BOFilteredEditsetList"
End Property

Public Property Let MaxCataloguePerPage(t_iMaxCataloguePerPage As Integer)
    m_iMaxPerPage = t_iMaxCataloguePerPage
End Property
Public Property Get MaxCataloguePerPage() As Integer
    MaxCataloguePerPage = m_iMaxPerPage
End Property


Public Property Let CatalogueEntryStatus(t_eCatalogueEntryStatus As ECATALOGUEENTRYSTATUS)
    m_eStatus = t_eCatalogueEntryStatus
End Property
Public Property Get CatalogueEntryStatus() As ECATALOGUEENTRYSTATUS
    CatalogueEntryStatus = m_eStatus
End Property


Public Property Let CatalogueLevelFiltered(t_eCatalogueLevel As ECATALOGUELEVEL)
    m_eCatalogueLevel = t_eCatalogueLevel
End Property
Public Property Get CatalogueLevelFiltered() As ECATALOGUELEVEL
    CatalogueLevelFiltered = m_eCatalogueLevel
End Property
' d11 - Added
Public Property Let CatalogueViewFiltered(t_eCatalogueView As ECATALOGUEVIEW)
    m_eCatalogueView = t_eCatalogueView
End Property
Public Property Get CatalogueViewFiltered() As ECATALOGUEVIEW
    CatalogueViewFiltered = m_eCatalogueView
End Property
'd18 - Added
Public Property Let CatalogueOrderFiltered(t_eCatalogueOrder As ECATALOGUEORDER)
    m_eCatalogueOrder = t_eCatalogueOrder
End Property
Public Property Get CatalogueOrderFiltered() As ECATALOGUEORDER
    CatalogueOrderFiltered = m_eCatalogueOrder
End Property


Public Property Let EditsetID(t_lEditSetID As Long)
    m_lEditsetID = t_lEditSetID
End Property
Public Property Get EditsetID() As Long
    EditsetID = m_lEditsetID
End Property

' d10 - Added
Public Property Let PositionCatalogueReference(t_sPositionCatalogueReference As String)
    m_sPositionCatalogueReference = ReplaceString(LCase(t_sPositionCatalogueReference), " ", "")
End Property
Public Property Get PositionCatalogueReference() As String
    PositionCatalogueReference = m_sPositionCatalogueReference
End Property

Public Property Let EditsetPosition(t_lEditsetPosition As Long)
    m_lEditsetPosition = t_lEditsetPosition
End Property
Public Property Get EditsetPosition() As Long
    EditsetPosition = m_lEditsetPosition
End Property

Public Property Let EditsetPositionValid(t_bPositionValid As Boolean)
    m_bPositionValid = t_bPositionValid
End Property
Public Property Get EditsetPositionValid() As Boolean
    EditsetPositionValid = m_bPositionValid
End Property



''''''''''''''''''''''''''''''''''''''''''''''Methods'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function Initialise()
' d10   daj     12Jan2001       Added new property initialisers
' d20   alm     19Sep2001       Changed the default ordering to KeyOrder

    m_iMaxPerPage = gk_iMaxFilteredEditsetEntriesPerPage
    m_eStatus = ECATALOGUEENTRYSTATUS.cesNONE
    m_eCatalogueLevel = ECATALOGUELEVEL.ctNone
    
    m_lSessionID = GetObjectContext.Item(gk_sSessionObject).SessionID
    
    ' d10 - Added
    m_sPositionCatalogueReference = ""
    m_lEditsetPosition = 0
    m_bPositionValid = True
    
    ' d13 - Added
     m_eCatalogueView = ECATALOGUEVIEW.cevDefault
     
    ' d18 - Added
    m_eCatalogueOrder = ECATALOGUEORDER.ceoKeyOrder
    
    
End Function

Private Function Terminate()

End Function


Public Sub PROCATTransientInterface_DeleteState(Optional t_sInstanceReference As String)
'Method the delete the instant state
    Dim sInstanceReference As String
    Dim ofwTransientManager As FWTransientManager

    sInstanceReference = CreateDelimitedString(Me.Name, t_sInstanceReference)

    Set ofwTransientManager = New FWTransientManager
    ofwTransientManager.DeleteData m_lSessionID, _
                                    sInstanceReference

    LogMessage "Deleted " & Me.Name & " state" & sInstanceReference, EProfile.pComplete

End Sub

Public Sub PROCATTransientInterface_GetState(Optional t_sInstanceReference As String)
On Error GoTo ErrorHandler
'Method to retrive the instant state
' d10   daj     12Jan2001       Added new properties

    Dim sInstanceReference As String
    Dim sSerialisedString As String
    Dim ofwTransientManager As FWTransientManager
    Dim sSource As String
    
    sSource = Me.Name & "PROCATTransientInterface_GetState"
    
    sInstanceReference = CreateDelimitedString(Me.Name, t_sInstanceReference)

    Set ofwTransientManager = New FWTransientManager
    sSerialisedString = ofwTransientManager.GetData(m_lSessionID, _
                                            sInstanceReference)
                                            
    'If data found then reassign the values else let it as it is
    If Trim(sSerialisedString) = "" Then
    
        LogMessage "BOUser state not found for " & sInstanceReference & ". Creating new state info." _
                                                , EProfile.pProfileInformation
        PROCATTransientInterface_SetState
    Else
        m_iMaxPerPage = CInt(GetDataFromSerialisedString(sSerialisedString))
        m_eStatus = CInt(GetDataFromSerialisedString(sSerialisedString))
        m_eCatalogueLevel = CInt(GetDataFromSerialisedString(sSerialisedString))
        m_lEditsetID = CLng(GetDataFromSerialisedString(sSerialisedString))
        
        ' d10 - Added
        m_sPositionCatalogueReference = GetDataFromSerialisedString(sSerialisedString)
        m_lEditsetPosition = CLng(GetDataFromSerialisedString(sSerialisedString))
        m_bPositionValid = CBool(GetDataFromSerialisedString(sSerialisedString))
        
        ' d13 - Added
        m_eCatalogueView = CInt(GetDataFromSerialisedString(sSerialisedString))
        
        ' d18 - Added
        m_eCatalogueOrder = CInt(GetDataFromSerialisedString(sSerialisedString))
        
        
        LogMessage "Read " & Me.Name & " state" & sInstanceReference, EProfile.pComplete
    End If
    Exit Sub
ErrorHandler:
    ' d6    ksk     12Nov1999
    'Removed check for number if get does not find a number then set to defaults
    If InStr(Err.Description, sInstanceReference & " does not exist") Then
        
        Exit Sub
    Else
        ConvertError sSource
    End If
End Sub

Public Sub PROCATTransientInterface_SetState(Optional t_sInstanceReference As String)
'Method to retrive the instant state
' d10   daj     12Jan2001       Added new properties

    Dim sInstanceReference As String
    Dim sSerialisedString As String
    Dim ofwTransientManager As FWTransientManager
    
    sInstanceReference = CreateDelimitedString(Me.Name, t_sInstanceReference)

        
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, CStr(m_iMaxPerPage))
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, CStr(m_eStatus))
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, CStr(m_eCatalogueLevel))
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, CStr(m_lEditsetID))

    ' d10 - Added
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, m_sPositionCatalogueReference)
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, CStr(m_lEditsetPosition))
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, CStr(m_bPositionValid))

    'd13 - Added
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, CStr(m_eCatalogueView))

    'd18 - Added
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, CStr(m_eCatalogueOrder))

    Set ofwTransientManager = New FWTransientManager
    ofwTransientManager.AddData m_lSessionID, _
                                sInstanceReference, sSerialisedString

    LogMessage "Stored " & Me.Name & " state" & sInstanceReference, EProfile.pComplete

End Sub


Private Sub Class_Initialize()
    Initialise
End Sub

Private Sub Class_Terminate()
    Terminate
End Sub

Private Sub ObjectControl_Activate()
    Initialise
End Sub

Private Sub ObjectControl_Deactivate()
    Terminate
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = True
End Function


Public Function GetEditsetEntries() As FWRecordList
On Error GoTo ErrHandler
'Lists the editset entries
' d8    spb     08Feb2000       In GetEditSetEntries, retrieve a max of
'                               gk_lMaxEditsetSizeToRetrieve.
' d9    tmj     03Apr2000       Ensured full title length is returned
' d10   daj     12Jan2001       Modified object to position the edit set viewer at a particular entry,
'                               instead of a page
' d11   pnp     08Feb2001       Added edit set order funtionality

    Dim ofwQuery As FWDBHelper
    Dim sSQL As String
    Dim ofwResult As FWRecordList
    Dim sSource As String
    
    sSource = Me.Name & ".GetEditsetEntries"
    
    'd11
    If m_eCatalogueOrder = ceoHierarchical Then
        sSQL = "EXECUTE sp_Edit_GetEditSetListForDisplay"
    ElseIf m_eCatalogueOrder = ceoKeyOrder Then
        sSQL = "EXECUTE sp_Edit_GetEditSetListPKOrderForDisplay"
    End If
    
    sSQL = sSQL & " @t_iEditsetID = " & CStr(m_lEditsetID)
    sSQL = sSQL & ", @t_iFilteredStatus = " & CStr(m_eStatus)
    sSQL = sSQL & ", @t_iFilteredLevel = " & CStr(m_eCatalogueLevel)
    sSQL = sSQL & ", @t_iTitleLengthToReturn = " & CStr(gk_iDBCOLLenShortTitle) ' d19
    
    Set ofwQuery = New FWDBHelper

    Set ofwResult = ofwQuery.FindBySql(sSQL, gk_lMaxEditsetSizeToRetrieve)
        
    ' d10 - Find the position of the selected entry if specified
    If Not m_sPositionCatalogueReference = "" Then
        m_lEditsetPosition = FindEditsetPosition(ofwResult)
        If m_lEditsetPosition > -1 Then
            m_bPositionValid = True
        Else
            m_bPositionValid = False
        End If
    End If
    
    Set GetEditsetEntries = ofwResult
    Exit Function
    
ErrHandler:
    ConvertError sSource
End Function

Private Function FindEditsetPosition(t_ofwResults As FWRecordList) As Long
' Goes through the edit set and finds the catalogue entry specified in the member variables.  If
' found, returns the position of that entry in the editset
'
' d10   daj     12Jan2001       Created
On Error GoTo ErrHandler

    Dim sSource As String
    Dim sReference As String
    Dim lCount
    
    sSource = Me.Name & ".GetEditsetStatusList"
    
    ' Initialise counter and set default
    lCount = -1
    FindEditsetPosition = -1
    
    t_ofwResults.MoveFirst
    While Not t_ofwResults.IsAtEnd
        lCount = lCount + 1
        
        ' Construct the reference of the current entry
        sReference = CreateReference(VariantStringAsString(t_ofwResults.GetFieldValue(gk_sDBCOLAliasLettercodeRef)), _
            VariantStringAsString(t_ofwResults.GetFieldValue(gk_sDBCOLAliasDivisionRef)), _
            VariantStringAsString(t_ofwResults.GetFieldValue(gk_sDBCOLAliasClassRef)), _
            VariantStringAsString(t_ofwResults.GetFieldValue(gk_sDBCOLAliasHeaderRef)), _
            VariantStringAsString(t_ofwResults.GetFieldValue(gk_sDBCOLAliasSubHeaderRef)), _
            VariantStringAsString(t_ofwResults.GetFieldValue(gk_sDBCOLAliasPieceRef)), _
            VariantStringAsString(t_ofwResults.GetFieldValue(gk_sDBCOLAliasItemRef)), _
            t_ofwResults.GetFieldValue(gk_sDBCOLCatalogueLevel), _
            VariantStringAsString(t_ofwResults.GetFieldValue(gk_sDBCOLAliasSubClassRef)))
        
        sReference = ReplaceString(LCase(sReference), " ", "")
        
        ' Check for a match
        If sReference = m_sPositionCatalogueReference Then
            FindEditsetPosition = lCount
            Exit Function
        End If
        
        t_ofwResults.MoveNext
    Wend
    
    Exit Function
    
ErrHandler:
    ConvertError sSource
End Function
Public Function GetEditsetView(t_iView) As String
' function to take View id and return View_name
' d17   pnp     24Jan2001       Added

On Error GoTo ErrHandler
    
    Dim sSource As String
    Dim ofwQuery As FWDBHelper
    Dim sSQL As String
    Dim ofwResult As FWRecordList
    
    sSource = Me.Name & ".GetEditsetView"

    sSQL = "Select " & gk_sDBCOLViewName & " as " & gk_sDBCOLValue
    sSQL = sSQL & " from " & gk_sDBTABLEViewType & " (nolock) Where " & gk_sDBCOLViewID & " = " & t_iView
    Set ofwQuery = New FWDBHelper
    
    Set ofwResult = ofwQuery.FindBySql(sSQL)
    
    ofwResult.MoveFirst
    GetEditsetView = ofwResult.GetFieldValue(gk_sDBCOLValue)
    

Exit Function
    
ErrHandler:
    ConvertError sSource
End Function
' d12 - Added

Public Function GetEditsetViewList() As FWRecordList
On Error GoTo ErrHandler
'Lists the editset entries
    Dim ofwQuery As FWDBHelper
    Dim sSQL As String
    Dim ofwResult As FWRecordList
    Dim sSource As String
    
    sSource = Me.Name & ".GetEditsetViewList"
    
    sSQL = "Select " & gk_sDBCOLViewID & " as ID , " & gk_sDBCOLViewName & " as " & gk_sDBCOLValue
    sSQL = sSQL & " from " & gk_sDBTABLEViewType & " (nolock) Order by ID"
    Set ofwQuery = New FWDBHelper

    Set ofwResult = ofwQuery.FindBySql(sSQL)
    
    Set GetEditsetViewList = ofwResult
    Exit Function
    
ErrHandler:
    ConvertError sSource
End Function

Public Function GetEditsetOrder(t_iOrder) As String
' function to take Order id and return Order_name
' d18   pnp     08Feb2001       Added

On Error GoTo ErrHandler
    
    Dim sSource As String
    Dim ofwQuery As FWDBHelper
    Dim sSQL As String
    Dim ofwResult As FWRecordList
    
    sSource = Me.Name & ".GetEditsetOrder"

    sSQL = "Select " & gk_sDBCOLOrderName & " as " & gk_sDBCOLValue
    sSQL = sSQL & " from " & gk_sDBTABLEOrderType & " (nolock) Where " & gk_sDBCOLOrderID & " = " & t_iOrder
    Set ofwQuery = New FWDBHelper
    
    Set ofwResult = ofwQuery.FindBySql(sSQL)
    
    ofwResult.MoveFirst
    GetEditsetOrder = ofwResult.GetFieldValue(gk_sDBCOLValue)
    

Exit Function
    
ErrHandler:
    ConvertError sSource
End Function

' d18 - Added
Public Function GetEditsetOrderList() As FWRecordList
On Error GoTo ErrHandler
'Lists the editset entries
    Dim ofwQuery As FWDBHelper
    Dim sSQL As String
    Dim ofwResult As FWRecordList
    Dim sSource As String
    
    sSource = Me.Name & ".GetEditsetOrderList"
    
    sSQL = "Select " & gk_sDBCOLOrderID & " as ID , " & gk_sDBCOLOrderName & " as " & gk_sDBCOLValue
    sSQL = sSQL & " from " & gk_sDBTABLEOrderType & " (nolock) Order by ID"
    Set ofwQuery = New FWDBHelper

    Set ofwResult = ofwQuery.FindBySql(sSQL)
    
    Set GetEditsetOrderList = ofwResult
    Exit Function
    
ErrHandler:
    ConvertError sSource
End Function

Public Function GetEditsetColumnList(t_iView As Integer) As FWRecordList
'Method to return a record list of the columns required for the custom view
' d14  Added

On Error GoTo ErrHandler

    Dim ofwQuery As FWDBHelper
    Dim sSQL As String
    Dim ofwResult As FWRecordList
    Dim sSource As String
    
    sSource = Me.Name & ".GetEditsetColumnList"
    
    sSQL = "Select column_name as " & gk_sDBCOLValue & ", column_width as " & gk_sDBCOLWidth
    sSQL = sSQL & " from edit_set_view_columns esvc (nolock)"
    sSQL = sSQL & " WHERE esvc.view_id = '" & t_iView & "' Order by order_id"
    Set ofwQuery = New FWDBHelper

    Set ofwResult = ofwQuery.FindBySql(sSQL)
    
    Set GetEditsetColumnList = ofwResult
    Exit Function
    
ErrHandler:
    ConvertError sSource
End Function
' d16
Public Function AddContentCell(ByVal t_ofwRecord, ByVal t_sColName As String, ByRef t_sColText As String, _
                                ByRef t_sColWidth As String, ByRef t_iColSpan As Integer, ByRef t_bAsString As Boolean)

' Method to return cell properties for custom columns
'd16    pnp     22Jan2001       Added

On Error GoTo ErrHandler

Dim sSource As String
      Select Case t_sColName
        
            Case gk_sDBCOLClosureCode
            
                t_sColText = VariantStringAsString(t_ofwRecord.GetValueForField(gk_sDBCOLClosureCode))
                t_sColWidth = ""
                t_iColSpan = 1
                t_bAsString = True
            
     
            Case gk_sDBCOLClosureStatus
            
                t_sColText = VariantStringAsString(t_ofwRecord.GetValueForField(gk_sDBCOLClosureStatus))
                t_sColWidth = ""
                t_iColSpan = 1
                t_bAsString = True
                       
         
            Case gk_sDBCOLClosureType
            
                t_sColText = VariantStringAsString(t_ofwRecord.GetValueForField(gk_sDBCOLClosureType))
                t_sColWidth = ""
                t_iColSpan = 1
                t_bAsString = True
             
             Case gk_sDBColClosureStatusTypeCode
            
                t_sColText = VariantStringAsString(t_ofwRecord.GetValueForField(gk_sDBColClosureStatusTypeCode))
                t_sColWidth = ""
                t_iColSpan = 1
                t_bAsString = True
             
            Case gk_sDBCOLCoveringDate
            
                t_sColText = VariantStringAsString(t_ofwRecord.GetValueForField(gk_sDBCOLCoveringDate))
                t_sColWidth = ""
                t_iColSpan = 1
                t_bAsString = True
                    
            
            Case gk_sDBCOLformerDepRef
            
                t_sColText = VariantStringAsString(t_ofwRecord.GetValueForField(gk_sDBCOLformerDepRef))
                t_sColWidth = ""
                t_iColSpan = 1
                t_bAsString = True
  
            Case gk_sDBCOLLegalStatus
            
                t_sColText = VariantStringAsString(t_ofwRecord.GetValueForField(gk_sDBCOLLegalStatus))
                t_sColWidth = ""
                t_iColSpan = 1
                t_bAsString = True
               
               
            Case gk_sDBCOLRecordOpeningDate
            
                If IsNull(t_ofwRecord.GetValueForField(gk_sDBCOLRecordOpeningDate)) Then
                    t_sColText = VariantStringAsString(t_ofwRecord.GetValueForField(gk_sDBCOLRecordOpeningDate))
                Else
                    t_sColText = VariantStringAsString(VariantDateAsDDMMMYYYY(t_ofwRecord.GetValueForField(gk_sDBCOLRecordOpeningDate)))
                End If
                t_sColWidth = ""
                t_iColSpan = 1
                t_bAsString = True
            
               
  
            Case gk_sDBCOLRecordStatus
            
                t_sColText = VariantStringAsString(t_ofwRecord.GetValueForField(gk_sDBCOLRecordStatus))
                t_sColWidth = ""
                t_iColSpan = 1
                t_bAsString = True
            
           
            Case Else
            
                t_sColText = "*ERR*"
                t_sColWidth = ""
                t_iColSpan = 1
                t_bAsString = True
      
        End Select
        
 Exit Function

ErrHandler:
    ConvertError sSource
End Function
Public Function AddColumnHeading(ByVal t_sColName As String, ByRef t_sColText As String, ByRef t_iColSpan As Integer)
                                                        
' Returns column header values for custom view columns
' d15       pnp         22Jan2001           Added

On Error GoTo ErrHandler

    Dim sSource As String
    
    Dim m_oFWAppData As FWApplicationData
    Set m_oFWAppData = New FWApplicationData
    
    Select Case t_sColName
        
                
        Case gk_sDBCOLClosureStatus
        
            t_sColText = m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfCatalogueClosureStatus)
            t_iColSpan = 1
            
        
        Case gk_sDBColClosureStatusTypeCode
            
            t_sColText = m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfCatalogueClosureStatusTypecode)
            t_iColSpan = 1
            
        
        Case gk_sDBColClosureTypeCode
        
            t_sColText = m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfCatalogueClosureTypeAndCode)
            t_iColSpan = 1
        
        
        Case gk_sDBCOLCoveringDate
        
            t_sColText = m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfCatalogueCoveringDate)
            t_iColSpan = 1
            
        
        Case gk_sDBCOLformerDepRef
        
            t_sColText = m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfCatalogueformerDeptRef)
            t_iColSpan = 1
          
        Case gk_sDBCOLLegalStatus
        
            t_sColText = m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfCatalogueLegalStatus)
            t_iColSpan = 1
            
        Case gk_sDBColPiecekeyOrder
        
            t_sColText = m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfCatalogueKeyOrder)
            t_iColSpan = 1
            
        Case gk_sDBCOLRecordOpeningDate
        
            t_sColText = m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfCatalogueRecordOpeningDate)
            t_iColSpan = 1
            
        Case gk_sDBCOLRecordStatus
        
            t_sColText = m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfCatalogueRecordStatus)
            t_iColSpan = 1
            
        Case gk_sDBCOLTitle
        
            t_sColText = m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfCatalogueTitle)
            t_iColSpan = 1
            
        Case Else
        
            t_sColText = t_sColName
            t_iColSpan = 1
           
        End Select
        
        
    Set m_oFWAppData = Nothing
    Exit Function

ErrHandler:
    ConvertError sSource
End Function

Public Function GetEditsetStatusList() As FWRecordList
On Error GoTo ErrHandler
'Lists the editset entries
    Dim ofwQuery As FWDBHelper
    Dim sSQL As String
    Dim ofwResult As FWRecordList
    Dim sSource As String
    
    sSource = Me.Name & ".GetEditsetStatusList"
    
'    sSql = "Select 1 as ID, 'Accessioned' as Value " & " Union "
'    sSql = sSql & "Select 2 as ID, 'Added' as Value " & " Union "
'    sSql = sSql & "Select 3 as ID, 'Created' as Value " & " Union "
'    sSql = sSql & "Select 4 as ID, 'Read Only' as Value " & " Union "
'    sSql = sSql & "Select 5 as ID, 'Modified' as Value "
    
    sSQL = "Select " & gk_sDBCOLEntryTypeID & " as ID , " & gk_sDBCOLEntryTypeName & " as Value "
    sSQL = sSQL & " from " & gk_sDBTABLEEntryType & " (nolock) Order by ID"
    Set ofwQuery = New FWDBHelper

    Set ofwResult = ofwQuery.FindBySql(sSQL)
    
    Set GetEditsetStatusList = ofwResult
    Exit Function
    
ErrHandler:
    ConvertError sSource
End Function

Public Function GetCatalogueLevelList() As FWRecordList
On Error GoTo ErrHandler
'Lists the editset entries
    Dim ofwQuery As FWDBHelper
    Dim sSQL As String
    Dim ofwResult As FWRecordList
    Dim sSource As String
    
    sSource = Me.Name & ".GetCatalogueLevelList"
    
    sSQL = "Select " & gk_sDBCOLCatalogueLevel & " as ID, " & gk_sDBCOLLevelDesc & " as Value " & _
            " from " & gk_sDBTABLETblLevels & " (nolock) order by " & gk_sDBCOLCatalogueLevel
    
    Set ofwQuery = New FWDBHelper

    Set ofwResult = ofwQuery.FindBySql(sSQL)
    
    Set GetCatalogueLevelList = ofwResult
    Exit Function
    
ErrHandler:
    ConvertError sSource
End Function


Public Function GetEditsetStatusDescription(t_eEditsetStatus As ECATALOGUEENTRYSTATUS) As String
' d7    ksk     07Dec1999       Changed the name from GetEditsetDescription to GetEditsetStatusDescription
'                               Changed the description text
'
    Select Case t_eEditsetStatus
        Case ECATALOGUEENTRYSTATUS.cesNONE
            GetEditsetStatusDescription = "None"
        Case ECATALOGUEENTRYSTATUS.cesAddedFromLive
            GetEditsetStatusDescription = "entries added from the live"
        Case ECATALOGUEENTRYSTATUS.cesAddedFromLiveModified
            GetEditsetStatusDescription = "entries added from live and modified"
        Case ECATALOGUEENTRYSTATUS.cesImported
            GetEditsetStatusDescription = "accessioned entries"
        Case ECATALOGUEENTRYSTATUS.cesImportedModified
            GetEditsetStatusDescription = "accessioned and modified entries"
        Case ECATALOGUEENTRYSTATUS.cesPROCATCreated
            GetEditsetStatusDescription = "created entries"
        Case ECATALOGUEENTRYSTATUS.cesPROCATCreatedModified
            GetEditsetStatusDescription = "created and modified entries"
        Case ECATALOGUEENTRYSTATUS.cesReadeOnly
            GetEditsetStatusDescription = "readonly catalogue entries"
    End Select

End Function
