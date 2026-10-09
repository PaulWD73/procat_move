VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BOLeaflet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    BOLeaflet.cls
' System:       PROCAT, PRO
' Copyright:    (C) Quidnunc Limited
'
' Description:  The Leaflet Business object
'
' Amendment history:
'
'   d1      crm     05Aug1999      Created.
'   d2      crm     03Nov1999      Added Save(),Load()
'   d3      crm     09Nov1999      (integrated by ndk) Added LoadIndexTermsForLeaflet,ClearIndexTerms
'                                  SaveIndexTermsForLeaflet,AddIndexTerm
'                                  RemoveIndexTerm,SerialiseLinkItem,
'                                  ExtractLinkItemsFromSerialIsedString
'   d4      ndk     19Nov1999      Removed all references to note text
'   d5      crm     22Nov1999      CreatedOn returns current date for a new Leaflet
'   d6      ksk     22Nov1999      Changed to use the application object for the
'                                  Source and Destination file paths for leaflets
'   d7      wdp     02Dec1999      Added verity indexing: IndexInVerity
'   d8      crm     06Dec1999      Modified Save() to use VBA file handling functions
'                                  instead of MicrosoftScriptingRuntime library
'                                  Added  functions to transform Leaflets
'   d9      crm     13Dec1999      Modified Alllefaletfiles to return null when no files are present
'   d10     WDP     27Jan2000      Added DCOM call to do Verity indexing on a remote server.
'   d11     WDP     31Jan2000      Altered link URL in transformerd leaflet to be
'                                  relative to Leaflet URL
'   d12     CRM     11Feb2000      Made leaflet transformation non case sensitive
'                                  Use freefile instead of #1 and using PRINT
'                                  instead of write.
'   d13     WDP     15Feb2000      Dont delete leaflet source file. This is done at a later stage
'                                  Using DeleteSourceFile
'   d14     TMJ     21Feb2000      Modified AllLeafletFiles so the directory value can be changed
'   d15     TMJ     15Mar2000      Changed so Leaflets are forced to be normal files
'   d16     DAJ     10Jul2000      Added Delete, DeleteFromVerity and DeleteLiveFile
'   d17     DAJ     04Oct2000      Fixed Date$ function that returns unlocalised date strings
'   d18     DAJ     06Nov2000      Added Trim functions to prevent extra spaces in leaflet directories causing problems
'   d19     MAP     08Mar2001      Amended ApplyTransformation to raise an informative error message
'   d20     DAJ     14Mar2001      Changed the Reader URL constant to reflect the new web site structure
'   d21     ALM     02Oct2001      Fixed bug where leaflet end was never found in ApplyTransformation
'   d22     TomH    25Feb2002      Fixed defect (0215)
'                                  problem in loop not being exited properly
'   d23     TomH    22Mar2002      Changed the Reader URL constant steve asked for this as currently
'                                  site works on internet but not internally
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Implemented interfaces
Implements PROCATTransientInterface
Implements ObjectControl

'object specific properties
Private m_sLeafletFile As String
Private m_lSFAReferenceID As Long
Private m_sCreatedon As String
Private m_sInfoLeafletURL As String
Private m_lLeafletOrder As Long
Private m_sLeafletTitleText As String

Private m_sSourceOfLeafletFiles As String
Private m_sDestinationOfLeafletFiles As String

Private m_typData As LeafletDetails

'private collection to hold the IndexTermIDs
Private m_colPerson As Collection
Private m_colSubject As Collection
Private m_colCorporateBody As Collection
Private m_colPlace As Collection
Private m_ofwApplication As FWApplicationData

'private array to hold the registered leaflet files
Private m_vArrRegisteredFiles As Variant

Private Const gk_sDisplayCatDetailsGoto As String = "DisplayCatDetailsGoto.asp"
Private Const gk_sPSRunSearch As String = "PSRunSearch.asp"

' d20 - Changed
' d23 - Changed from "/" to include reader see comment above
Private Const mk_sLinkBackToReaderURL As String = "/Reader/"

Public Property Get Data() As LeafletDetails
    Data = m_typData
End Property

Public Property Let Data(t_typData As LeafletDetails)
    m_typData = t_typData
End Property

Private Sub Initialise()
    'put stuff to be initialised here
    Set m_colCorporateBody = New Collection
    Set m_colSubject = New Collection
    Set m_colPerson = New Collection
    Set m_colPlace = New Collection
    Set m_ofwApplication = New FWApplicationData

    m_sSourceOfLeafletFiles = m_ofwApplication.GetApplicationConfigurationInfomation( _
                                ECONFIGURATIONDATA.pLeafletSourceFileDirectory)
    m_sSourceOfLeafletFiles = IIf(Right(m_sSourceOfLeafletFiles, 1) = "\", m_sSourceOfLeafletFiles, _
                                m_sSourceOfLeafletFiles & "\")
                                
    m_sDestinationOfLeafletFiles = m_ofwApplication.GetApplicationConfigurationInfomation( _
                                ECONFIGURATIONDATA.pLeafletDestinationFileDirectory)
    m_sDestinationOfLeafletFiles = IIf(Right(m_sDestinationOfLeafletFiles, 1) = "\", m_sDestinationOfLeafletFiles, _
                                m_sDestinationOfLeafletFiles & "\")
                                
    'd13 added to allow comparison
    m_vArrRegisteredFiles = AllLeafletFiles(False)
                                
End Sub

Private Sub Terminate()
    'put stuff to be terminated here
    Set m_colCorporateBody = Nothing
    Set m_colSubject = Nothing
    Set m_colPerson = Nothing
    Set m_colPlace = Nothing
    Set m_ofwApplication = Nothing
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

Private Function ObjectControl_CanBePooled() As Boolean
ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
Terminate
End Sub

Public Sub PROCATTransientInterface_DeleteState(Optional t_sInstanceReference As String)

    Dim sInstanceReference As String
    Dim ofwTransienceManager As FWTransientManager
    
    sInstanceReference = CreateDelimitedString(Me.Name, t_sInstanceReference)
    Set ofwTransienceManager = New FWTransientManager
    ofwTransienceManager.DeleteData GetObjectContext.Item(gk_sSessionObject).SessionID, sInstanceReference
    
    LogMessage "Deleted BOLeaflet state " & sInstanceReference, EProfile.pComplete
    
End Sub

Public Sub PROCATTransientInterface_GetState(Optional t_sInstanceReference As String)
    Dim sInstanceReference As String
    Dim sSerialisedString As String
    Dim ofwTransienceManager As FWTransientManager
    
    sInstanceReference = CreateDelimitedString(Me.Name, t_sInstanceReference)
    
    Set ofwTransienceManager = New FWTransientManager
    
    sSerialisedString = ofwTransienceManager.GetData(GetObjectContext.Item(gk_sSessionObject).SessionID, sInstanceReference)
    
    If Trim(sSerialisedString) = "" Then
        LogMessage "BOUser state not found " & sInstanceReference, EProfile.pProfileInformation
    Else
        m_sLeafletFile = GetDataFromSerialisedString(sSerialisedString)
        m_lSFAReferenceID = CLng(GetDataFromSerialisedString(sSerialisedString))
        m_sCreatedon = GetDataFromSerialisedString(sSerialisedString)
        m_sInfoLeafletURL = GetDataFromSerialisedString(sSerialisedString)
        m_lLeafletOrder = CLng(GetDataFromSerialisedString(sSerialisedString))
        m_sLeafletTitleText = GetDataFromSerialisedString(sSerialisedString)
        
        'extract the collection from serialised string
        Set m_colCorporateBody = ExtractLinkItemsFromSerialIsedString(sSerialisedString)
        Set m_colSubject = ExtractLinkItemsFromSerialIsedString(sSerialisedString)
        Set m_colPerson = ExtractLinkItemsFromSerialIsedString(sSerialisedString)
        Set m_colPlace = ExtractLinkItemsFromSerialIsedString(sSerialisedString)
        
        LogMessage "Read BOLeaflet state" & sInstanceReference, EProfile.pComplete
    End If
    
    
    
End Sub

Public Sub PROCATTransientInterface_SetState(Optional t_sInstanceReference As String)
    Dim sInstanceReference As String
    Dim sSerialisedString As String
    Dim ofwTransienceManager As FWTransientManager
    
    sInstanceReference = CreateDelimitedString(Me.Name, t_sInstanceReference)
    
    sSerialisedString = ""
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, m_sLeafletFile)
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, CStr(m_lSFAReferenceID))
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, m_sCreatedon)
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, m_sInfoLeafletURL)
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, CStr(m_lLeafletOrder))
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, CStr(m_sLeafletTitleText))
    
    'the collections to be serialised
    sSerialisedString = SerialiseLinkItem(sSerialisedString, m_colCorporateBody)
    sSerialisedString = SerialiseLinkItem(sSerialisedString, m_colSubject)
    sSerialisedString = SerialiseLinkItem(sSerialisedString, m_colPerson)
    sSerialisedString = SerialiseLinkItem(sSerialisedString, m_colPlace)
    
    Set ofwTransienceManager = New FWTransientManager
    
    ofwTransienceManager.AddData GetObjectContext.Item(gk_sSessionObject).SessionID, _
                                 sInstanceReference, sSerialisedString
                                 
    LogMessage "BOLeaflet saved state " & t_sInstanceReference, EProfile.pComplete
    
End Sub


Public Function AllLeaflets() As FWRecordList
'Description    :   Function returns a list of available leaflets
'
'Amendment History  :
'   d1      crm     04AUG1999      Created
'   d2      crm     03Nov1999      Modified to pick leaflets from 'Leaflet' table
On Error GoTo ErrorHandler

    Dim sSource As String
    Dim ofwDBHelper As FWDBHelper
    Dim ofwLeaflet As FWRecordList
    Dim sSQL As String
        
    sSource = "BOLeaflet.AllLeaflets"
    
    'Create the DB helper object
    Set ofwDBHelper = New FWDBHelper
    
    sSQL = "SELECT * FROM " & gk_sDBTABLELeaflet
           
    'Fetch the results
    Set ofwLeaflet = ofwDBHelper.FindBySql(sSQL)
    
    'Return the object reference to the caller
    Set AllLeaflets = ofwLeaflet
    
    'Destroy object references
    Set ofwLeaflet = Nothing
    Set ofwDBHelper = Nothing

    Exit Function
ErrorHandler:
    ConvertError sSource
End Function

Public Function AllLeafletFiles(t_bDir As Boolean) As Variant
'Description    :       This fuction returns the list of leaflet files that can be
'                       registered against a Leaflet,it scans a marked directory on the web-server
'                       for files.Assumes that the directory has been marked for this purpose.
'Variables      :       t_bDir if true gets from the source, if false gets from the destination
'Amedment History
'   d1      crm     05Aug1999      Created.
'   d8      crm     06Dec1999      Using VBA file handler functions instead of
'                                  FileSystemObject library,removed reference to
'                                  MicrosoftScriptingRuntime Library
'   d9      crm     13Dec1999      Modified Alllefaletfiles to return null when no files are present
'   d14     TMJ     21Feb2000      Added boolean variable to determine where selected from
'   d18     DAJ     06Nov2000      Added Trim functions to prevent extra spaces in leaflet directories causing problems
On Error GoTo ErrorHandler
    Dim sSource As String
    Dim sArrFileNames() As String
    Dim lIndex As Long
    Dim sTempFileName As String
    Dim sDirectory As String
       
    sSource = "BOLeaflet.AllLeafletFiles()"
    
    ' d18 - Added Trim calls
    If t_bDir Then
        sDirectory = Trim(m_sSourceOfLeafletFiles)
    Else
        sDirectory = Trim(m_sDestinationOfLeafletFiles)
    End If
    
    'get the first file
    sTempFileName = Dir(sDirectory)
    
    Do While sTempFileName <> ""   ' Start the loop.
      'check the attribute to ensure it is a file
      If (GetAttr(sDirectory & _
            sTempFileName) And vbNormal) = vbNormal Then
         lIndex = lIndex + 1
         'if a file store in array
         ReDim Preserve sArrFileNames(lIndex)
         sArrFileNames(lIndex - 1) = sTempFileName
      End If
    sTempFileName = Dir   ' Get next entry.
    Loop
    
    'if no elements return null
    If lIndex = 0 Then
    AllLeafletFiles = Null
    Else
    'return the array with the file names
    AllLeafletFiles = sArrFileNames
    End If
    
    Exit Function
ErrorHandler:
    ConvertError sSource
End Function

Public Property Get LeafletFile() As Variant
LeafletFile = m_sLeafletFile
End Property

Public Property Let LeafletFile(ByVal t_sLeafletFile As Variant)
m_sLeafletFile = t_sLeafletFile
End Property

Public Property Get Name() As Variant
Name = "BOLeaflet"
End Property

Public Property Get SFAReferenceID() As Long
SFAReferenceID = m_lSFAReferenceID
End Property

Public Property Let SFAReferenceID(ByVal t_lSFAReferenceID As Long)
m_lSFAReferenceID = t_lSFAReferenceID
End Property

Public Property Get CreatedOn() As String
'   d5      crm     22Nov1999
'   d17     DAJ     04Oct2000      Fixed Date$ function that returns unlocalised date strings
If m_lSFAReferenceID > 0 Then
    'if the ID > o then it is an old Leaflet
    CreatedOn = m_sCreatedon
Else
    'else it is a new Leaflet so return current date
    CreatedOn = Date
End If
End Property

Public Property Let CreatedOn(ByVal t_sCreatedon As String)
m_sCreatedon = t_sCreatedon
End Property

Public Sub Save(Optional t_bFileChange As Boolean = False)
'Description    :       This method saves a lefalet whose details have been provided
'                       by way of assigment to the properties of this object
'
'Amedment History
'd2      crm     03Nov1999      Created
'd8      crm     06Dec1999      Modified the file handler functions
'                               now using VBA file functions
On Error GoTo ErrorHandler
    Dim sSource As String
    Dim sSQL As String
    Dim ofwDBHelper As New FWDBHelper
    
    sSource = "BOLeaflet.Save()"
        
    If m_lSFAReferenceID > 0 Then
        'this is a old SFA,so do an update
        sSQL = "UPDATE " & gk_sDBTABLELeaflet & " SET "
        
        If Len(Trim(m_sInfoLeafletURL)) > 0 Then
        
        End If
        
        If m_lLeafletOrder > 0 Then
        sSQL = sSQL & gk_sDBCOLLeafletOrder & " = " & m_lLeafletOrder & ","
        End If
        
        If Len(Trim(m_sLeafletTitleText)) > 0 Then
        sSQL = sSQL & gk_sDBCOLLeafletTitleText & " = " & QuoteValueWithQuotes(m_sLeafletTitleText) & ","
        End If
        'do not update the date field
'        If Len(Trim(m_sCreatedon)) > 0 Then
'        sSQL = sSQL & gk_sDBCOLLeafletCreatedOn & " = " & QuoteValueWithQuotes(m_sCreatedon) & ","
'        End If
            
        If Len(Trim(m_sLeafletFile)) > 0 Then
        'first set the leaflet URL
        sSQL = sSQL & gk_sDBCOLInfoLeafletURL & " = " & _
                QuoteValueWithQuotes(m_sDestinationOfLeafletFiles) & ","
        'now the leaflet name
        sSQL = sSQL & gk_sDBCOLLeafletFilename & " = " & QuoteValueWithQuotes(m_sLeafletFile)
        End If
        
        sSQL = sSQL & " WHERE " & gk_sDBCOLLeafletReferenceID & " = " & m_lSFAReferenceID
        'execute the SQL
        ofwDBHelper.Execute sSQL
    Else
        'this is a new SFA, so create one
        sSQL = "INSERT " & gk_sDBTABLELeaflet & " ( "
        sSQL = sSQL & gk_sDBCOLLeafletOrder & ","
        sSQL = sSQL & gk_sDBCOLLeafletTitleText & ","
        sSQL = sSQL & gk_sDBCOLLeafletCreatedOn & ","
        sSQL = sSQL & gk_sDBCOLLeafletFilename & ","
        sSQL = sSQL & gk_sDBCOLInfoLeafletURL & ")"
        sSQL = sSQL & " VALUES ( " & m_lLeafletOrder & ","
        sSQL = sSQL & QuoteValueWithQuotes(m_sLeafletTitleText) & ","
        sSQL = sSQL & QuoteValueWithQuotes(m_sCreatedon) & ","
        sSQL = sSQL & QuoteValueWithQuotes(m_sLeafletFile) & ","
        sSQL = sSQL & QuoteValueWithQuotes(m_sDestinationOfLeafletFiles) & ")"
        m_lSFAReferenceID = ofwDBHelper.InsertAndReturnIdentity(sSQL)
        t_bFileChange = True
    End If
    
    ClearIndexTerms m_lSFAReferenceID
    
    SaveIndexTermsForLeaflet m_lSFAReferenceID
    
    'if all actions and updates are successful, then move leaflet file to live area
    'move the leaflet-file from the source to the destination folder
    'if something goes wrong here ? Ha Ha...View class should rollback...
    If t_bFileChange Then
    
        TransformLeaflet m_sSourceOfLeafletFiles & m_sLeafletFile
        '   d8      crm     06Dec1999
        'copy the file from the source to the destination
        FileCopy m_sSourceOfLeafletFiles & m_sLeafletFile, _
                 m_sDestinationOfLeafletFiles & m_sLeafletFile
        
        'delete file from source directory
        'This is now done later to avoid problems if
        'transaction fails
        'Kill m_sSourceOfLeafletFiles & m_sLeafletFile
    End If
    
    'release object references
    Set ofwDBHelper = Nothing
'    Set oFileSys = Nothing
    
Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub


Public Sub IndexInVerity(t_bIsNew As Boolean)
'   This method triggers the indexing of the leaflet in Verity
'
'   d7  WDP     2-Dec-1999      Created
'   d10 WDP     27-Jan-2000     Added call to do Verity indexing on remote server

On Error GoTo ErrorHandler
    Dim sSource As String
    Dim oboVerIndexer
    sSource = "BOLeaflet.IndexInVerity()"
        
    Dim eIndexAction As EIndexingAction
    
    If t_bIsNew = False Then
        eIndexAction = IDXUpdate
    Else
        eIndexAction = IDXInsert
    End If
    
    Dim oboIndexer As New BOLeafletIndexer
    
    Dim sActionLog As String
    'This will generate the SGML and BLK data files for indexing
    'and place them on the Verity server, the indexing actions
    'are held in the log file returned.
    sActionLog = oboIndexer.IndexLeaflet(m_lSFAReferenceID, eIndexAction)
    
    'The Verity server could be remote and so we make a (D)COM
    'call to the indexer to actually index the leaflet into
    'the collection
    Set oboVerIndexer = CreateObject("PROCATEdVerIdx.BOIndexer")
    oboVerIndexer.IndexFromIndexLog sActionLog, False, True
    Set oboVerIndexer = Nothing
Exit Sub
ErrorHandler:
    Set oboVerIndexer = Nothing
    ConvertError sSource
End Sub

Public Sub DeleteFromVerity()
' Removes the Verity indexes for this leaflet
'
' Amendment history:
'   d16     DAJ     10Jul2000      Added Delete, DeleteFromVerity and DeleteLiveFile

On Error GoTo ErrorHandler

    Dim sSource As String
    Dim oboVerIndexer
    Dim oboIndexer As New BOLeafletIndexer
    Dim sActionLog As String
    
    sSource = "BOLeaflet.DeleteFromVerity()"
    
    ' Choose the Delete indexing function
    Dim eIndexAction As EIndexingAction
    eIndexAction = IDXDelete
    
    'This will generate the SGML and BLK data files for indexing
    'and place them on the Verity server, the indexing actions
    'are held in the log file returned.
    sActionLog = oboIndexer.IndexLeaflet(m_lSFAReferenceID, eIndexAction)
    
    'The Verity server could be remote and so we make a (D)COM
    'call to the indexer to actually index the leaflet into
    'the collection
    Set oboVerIndexer = CreateObject("PROCATEdVerIdx.BOIndexer")
    oboVerIndexer.IndexFromIndexLog sActionLog, False, True
    Set oboVerIndexer = Nothing
    
Exit Sub
ErrorHandler:
    Set oboIndexer = Nothing
    Set oboVerIndexer = Nothing
    ConvertError sSource
End Sub


Public Sub DeleteSourceFile()
'   Delete leaflet source file - this is done from VWleafet.Save
'   after all other actions are complete
'
'   d13 WDP     15-Feb-2000     Created
On Error GoTo ErrorHandler
    Dim sSource As String
    sSource = "BOLeaflet.DeleteSourceFile()"
        
    Kill m_sSourceOfLeafletFiles & m_sLeafletFile
    
Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub

Public Sub DeleteLiveFile()
'   Delete leaflet HTML file in the live area
'
'   d16     DAJ     10Jul2000      Added Delete, DeleteFromVerity and DeleteLiveFile
On Error GoTo ErrorHandler

    Dim sSource As String
    sSource = "BOLeaflet.DeleteLiveFile()"
        
    Kill m_sDestinationOfLeafletFiles & m_sLeafletFile
    
Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub


Public Sub Load(ByVal t_eLoadType As EAFTermLoadType, _
                Optional ByVal t_lID As Long, Optional ByVal t_lLinkTermID As Long, _
                Optional ByVal t_sLinkTermValue As String, Optional t_sLinkTermType As String)
                
'Description    :       This method loads a lefalet from the database
'
'
'Amedment History
'd2      crm     03Nov1999      Created
On Error GoTo ErrorHandler
    Dim sSource As String
    Dim sSQL As String
    Dim ofwDBHelper As New FWDBHelper
    Dim ofwRecList As FWRecordList
    Dim eTermType As EGroup
    
    sSource = "BOLeaflet.Load()"
       
   'The user clicked on Edit button and control has come here
    If t_eLoadType = EAFTermLoadType.EAFLoadFromDatabase Then
        
        'Delete the state information (if any)
        PROCATTransientInterface_DeleteState Me.Name
        
        'Load from the loader class
        LoadFromDatabase t_lID
    Else
        
        Select Case t_eLoadType
        
            Case EAFTermLoadType.EAFLoadFromStateAddTerm
            'The user has selected a pick term to add and the control has come here
                
                'Read the stored state info
                PROCATTransientInterface_GetState Me.Name
                'Add the new link term to the collection
                If Not IsMissing(t_sLinkTermValue) Then
                    AddIndexTerm t_lLinkTermID, eTermType
                Else
                    AddIndexTerm t_lLinkTermID, eTermType
                End If
            
            Case EAFTermLoadType.EAFLoadFromStateRemoveTerm
                'The user has selected a pick term to remove and the control has come here
                
                'Read the stored state info
                PROCATTransientInterface_GetState Me.Name
                            
                'Remove the link term from the collection
                If Not IsMissing(t_sLinkTermValue) Then
                    RemoveIndexTerm t_lLinkTermID, eTermType
                Else
                    RemoveIndexTerm t_lLinkTermID, eTermType
                End If
                
            Case EAFTermLoadType.EAFLoadFromStateTermNoChange
                'The user has selected cancel operation while adding/removing a link term and the control has come here
                
                'Just Read the stored stage info
                PROCATTransientInterface_GetState Me.Name
            Case Else
                PROCATTransientInterface_GetState Me.Name
                
        End Select
        
    End If
    
    'release object references
    Set ofwDBHelper = Nothing
    Set ofwRecList = Nothing
Exit Sub
ErrorHandler:

    ConvertError sSource
End Sub

Public Sub Delete()
' Deletes a leaflet from the database
'
' Amedment History
'   d16     DAJ     10Jul2000      Added Delete, DeleteFromVerity and DeleteLiveFile
On Error GoTo ErrorHandler

    Dim sSource As String
    Dim sSQL As String
    Dim ofwDBHelper As New FWDBHelper
    Dim ofwRecList As FWRecordList
    
    sSource = "BOLeaflet.Delete()"
    
    ' Delete the terms first
    DeleteIndexTermsForLeaflet m_lSFAReferenceID
    
    ' Construct SQL
    sSQL = "DELETE " & gk_sDBTABLELeaflet
    sSQL = sSQL & " WHERE " & gk_sDBCOLLeafletReferenceID & _
           " = " & m_lSFAReferenceID
           
    ' Delete the details from DB
    ofwDBHelper.Execute (sSQL)
    
    ' Release object references
    Set ofwDBHelper = Nothing
    Set ofwRecList = Nothing
    
Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub

Public Function AbleToDelete() As Boolean
' Checks whether this leaflet can be deleted from the database without breaking constraints
'
' Amedment History
'   d16     DAJ     10Jul2000      Added Delete, DeleteFromVerity and DeleteLiveFile
On Error GoTo ErrorHandler

    Dim sSource As String
    sSource = "BOLeaflet.AbleToDelete()"

    Dim sSQL As String
    Dim ofwDBHelper As New FWDBHelper
    Dim ofwRecList As New FWRecordList
    
    ' Set default of failure
    AbleToDelete = False
    
    ' Are there any PC_Leaflet links?
    sSQL = "SELECT * FROM " & gk_sDBTABLEPCLeaflet & _
        " WHERE " & gk_sDBCOLLeafletReferenceID & " = " & m_lSFAReferenceID
    Set ofwRecList = ofwDBHelper.FindBySql(sSQL)
    If ofwRecList.RecordCount > 0 Then
        Exit Function
    End If
    
    ' Are there any edit set links?
    sSQL = "SELECT * FROM " & gk_sDBTABLEEditsetPCLeaflet & _
        " WHERE " & gk_sDBCOLLeafletReferenceID & " = " & m_lSFAReferenceID
    Set ofwRecList = ofwDBHelper.FindBySql(sSQL)
    If ofwRecList.RecordCount > 0 Then
        Exit Function
    End If
    
    ' Set success
    AbleToDelete = True
    
Exit Function
ErrorHandler:
    ConvertError sSource
End Function

Public Property Get Title() As String
Title = m_sLeafletTitleText
End Property

Public Property Let Title(ByVal t_sTitle As String)
m_sLeafletTitleText = t_sTitle
End Property



Public Property Get LeafletURL() As String
LeafletURL = m_sInfoLeafletURL
End Property

Public Property Let LeafletURL(ByVal t_sLeafletURL As String)
m_sInfoLeafletURL = t_sLeafletURL
End Property

Public Property Get PersonTerms() As Collection
Set PersonTerms = m_colPerson
End Property

Public Property Get PlaceTerms() As Collection
Set PlaceTerms = m_colPlace
End Property
Public Property Get SubjectTerms() As Variant
Set SubjectTerms = m_colSubject
End Property
Public Property Get CorporateBodyTerms() As Variant
Set CorporateBodyTerms = m_colCorporateBody
End Property

Private Sub LoadFromDatabase(ByVal t_lLeafletID As Long)
'Description    :       This method saves a lefalet whose details have been provided
'                       by way of assigment to the properties of this object
'
'Amedment History
'd2      crm     03Nov1999      Created
On Error GoTo ErrorHandler
    Dim sSource As String
    Dim sSQL As String
    Dim ofwDBHelper As New FWDBHelper
    Dim ofwRecList As FWRecordList
    
    sSource = "BOLeaflet.Load()"
       
    If t_lLeafletID > 0 Then
        'construct SQL
        sSQL = "SELECT * FROM " & gk_sDBTABLELeaflet
        sSQL = sSQL & " WHERE " & gk_sDBCOLLeafletReferenceID & _
               " = " & t_lLeafletID
               
        'Load the details from DB
        Set ofwRecList = ofwDBHelper.FindBySql(sSQL)
        
        'populate the private variables
        m_sCreatedon = IIf(IsNull(ofwRecList.GetFieldValue(gk_sDBCOLLeafletCreatedOn)), _
                       "", ofwRecList.GetFieldValue(gk_sDBCOLLeafletCreatedOn))
        m_sLeafletFile = ofwRecList.GetFieldValue(gk_sDBCOLLeafletFilename)
        m_sLeafletTitleText = ofwRecList.GetFieldValue(gk_sDBCOLLeafletTitleText)
        m_lLeafletOrder = IIf(IsNull(ofwRecList.GetFieldValue(gk_sDBCOLLeafletOrder)), _
                          0, ofwRecList.GetFieldValue(gk_sDBCOLLeafletOrder))
        m_sInfoLeafletURL = ofwRecList.GetFieldValue(gk_sDBCOLInfoLeafletURL)
    
        LoadIndexTermsForLeaflet t_lLeafletID
    End If
    
    
    'release object references
    Set ofwDBHelper = Nothing
    Set ofwRecList = Nothing
Exit Sub
ErrorHandler:

    ConvertError sSource
End Sub

Private Sub LoadIndexTermsForLeaflet(ByVal t_lLeafletID As Long)
'
'Description: Loads Index terms against a Leaflet from the database
'
'
' Amendment history:
'd3     crm     09Nov1999     Created
'
On Error GoTo ErrorHandler
    Dim sSource As String
    Dim sSQL As String
    Dim ofwDBHelper As New FWDBHelper
    Dim ofwRecList As FWRecordList
    Dim lTempID As Long
    
    
    sSource = "BOLeaflet.Load()"
       
    If t_lLeafletID > 0 Then
        
        'construct SQL for Corporatebody
        sSQL = "SELECT " & gk_sDBTABLELeafletCorpBody & "." & gk_sDBCOLCorpBodyReferenceID & "," & gk_sDBCOLCorpBodyNameText
        sSQL = sSQL & " FROM " & gk_sDBTABLELeafletCorpBody & "," & gk_sDBTABLECorporateBody
        sSQL = sSQL & " WHERE " & gk_sDBTABLELeafletCorpBody & "." & gk_sDBCOLCorpBodyReferenceID
        sSQL = sSQL & "=" & gk_sDBTABLECorporateBody & "." & gk_sDBCOLCorpBodyReferenceID
        sSQL = sSQL & " AND " & gk_sDBCOLLeafletReferenceID & "=" & t_lLeafletID
        
        'Load the details from DB
        Set ofwRecList = ofwDBHelper.FindBySql(sSQL)
        If ofwRecList.RecordCount > 0 Then
            Do Until ofwRecList.IsAtEnd
                lTempID = ofwRecList.GetFieldValue(gk_sDBCOLCorpBodyReferenceID)
                m_colCorporateBody.Add lTempID, CStr(lTempID)
                ofwRecList.MoveNext
            Loop
        End If
        
        sSQL = ""
        
        'construct SQL for Subject
        sSQL = "SELECT " & gk_sDBTABLELeafletSubject & "." & gk_sDBCOLSubjectReferenceID & "," & gk_sDBCOLSubjectTermText
        sSQL = sSQL & " FROM " & gk_sDBTABLELeafletSubject & "," & gk_sDBTABLESubject
        sSQL = sSQL & " WHERE " & gk_sDBTABLELeafletSubject & "." & gk_sDBCOLSubjectReferenceID
        sSQL = sSQL & "=" & gk_sDBTABLESubject & "." & gk_sDBCOLSubjectReferenceID
        sSQL = sSQL & " AND " & gk_sDBCOLLeafletReferenceID & "=" & t_lLeafletID
        
        'Load the details from DB
        Set ofwRecList = ofwDBHelper.FindBySql(sSQL)
        If ofwRecList.RecordCount > 0 Then
            Do Until ofwRecList.IsAtEnd
                lTempID = ofwRecList.GetFieldValue(gk_sDBCOLSubjectReferenceID)
                m_colSubject.Add lTempID, CStr(lTempID)
                ofwRecList.MoveNext
            Loop
        End If
        
        'construct SQL for Person
        sSQL = "SELECT " & gk_sDBTABLELeafletPerson & "." & gk_sDBCOLPersonReferenceID
        sSQL = sSQL & " FROM " & gk_sDBTABLELeafletPerson & "," & gk_sDBTABLEPerson
        sSQL = sSQL & " WHERE " & gk_sDBTABLELeafletPerson & "." & gk_sDBCOLPersonReferenceID
        sSQL = sSQL & "=" & gk_sDBTABLEPerson & "." & gk_sDBCOLPersonReferenceID
        sSQL = sSQL & " AND " & gk_sDBCOLLeafletReferenceID & "=" & t_lLeafletID
        
        'Load the details from DB
        Set ofwRecList = ofwDBHelper.FindBySql(sSQL)
        If ofwRecList.RecordCount > 0 Then
            Do Until ofwRecList.IsAtEnd
                lTempID = ofwRecList.GetFieldValue(gk_sDBCOLPersonReferenceID)
                m_colPerson.Add lTempID, CStr(lTempID)
                ofwRecList.MoveNext
            Loop
        End If
        
        'construct SQL for Place
        sSQL = "SELECT " & gk_sDBTABLELeafletPlace & "." & gk_sDBCOLPlaceReferenceID & "," & gk_sDBCOLPlaceNameText
        sSQL = sSQL & " FROM " & gk_sDBTABLELeafletPlace & "," & gk_sDBTABLEPlace
        sSQL = sSQL & " WHERE " & gk_sDBTABLELeafletPlace & "." & gk_sDBCOLPlaceReferenceID
        sSQL = sSQL & "=" & gk_sDBTABLEPlace & "." & gk_sDBCOLPlaceReferenceID
        sSQL = sSQL & " AND " & gk_sDBCOLLeafletReferenceID & "=" & t_lLeafletID
        
        'Load the details from DB
        Set ofwRecList = ofwDBHelper.FindBySql(sSQL)
        If ofwRecList.RecordCount > 0 Then
            Do Until ofwRecList.IsAtEnd
                lTempID = ofwRecList.GetFieldValue(gk_sDBCOLPlaceReferenceID)
                m_colPlace.Add lTempID, CStr(lTempID)
                ofwRecList.MoveNext
            Loop
        End If
        
    End If
        
    'release object references
    Set ofwDBHelper = Nothing
    Set ofwRecList = Nothing
Exit Sub
ErrorHandler:
    If Err.Number = 457 Then Resume Next
    ConvertError sSource
End Sub

Private Sub DeleteIndexTermsForLeaflet(ByVal t_lLeafletID As Long)
' Delete all the link terms between a particular leaflet and index terms
'
' Amendment history:
'   d16     DAJ     10Jul2000      Added Delete, DeleteFromVerity and DeleteLiveFile

On Error GoTo ErrorHandler

    Dim sSource As String
    Dim sSQL As String
    Dim ofwDBHelper As New FWDBHelper
    sSource = "BOLeaflet.DeleteIndexTermsForLeaflet()"
       
    'construct SQL for Corporatebody
    sSQL = "DELETE " & gk_sDBTABLELeafletCorpBody
    sSQL = sSQL & " WHERE " & gk_sDBCOLLeafletReferenceID & "=" & t_lLeafletID
    
    ' Execute the SQL
    ofwDBHelper.Execute (sSQL)
    
    'construct SQL for Subject
    sSQL = "DELETE " & gk_sDBTABLELeafletSubject
    sSQL = sSQL & " WHERE " & gk_sDBCOLLeafletReferenceID & "=" & t_lLeafletID
    
    ' Execute the SQL
    ofwDBHelper.Execute (sSQL)
    
    'construct SQL for Person
    sSQL = "DELETE " & gk_sDBTABLELeafletPerson
    sSQL = sSQL & " WHERE " & gk_sDBCOLLeafletReferenceID & "=" & t_lLeafletID
    
    ' Execute the SQL
    ofwDBHelper.Execute (sSQL)
    
    'construct SQL for Place
    sSQL = "DELETE " & gk_sDBTABLELeafletPlace
    sSQL = sSQL & " WHERE " & gk_sDBCOLLeafletReferenceID & "=" & t_lLeafletID
    
    ' Execute the SQL
    ofwDBHelper.Execute (sSQL)
        
    'release object references
    Set ofwDBHelper = Nothing
    
Exit Sub
ErrorHandler:
    If Err.Number = 457 Then Resume Next
    ConvertError sSource
End Sub

Private Sub ClearIndexTerms(ByVal t_lLeafletID As Long)
'
'Description: Deletes all Index terms against a leaflet
'             This avoids having to keep count of which Item to delete and which to add
'
'
' Amendment history:
'd3     crm     09Nov1999     Created
'
On Error GoTo ErrorHandler
    Dim sSource As String
    Dim sSQL As String
    Dim ofwDBHelper As New FWDBHelper
    
    
    sSource = "BOLeaflet.ClearIndexTerms()"
    
    sSQL = "DELETE " & gk_sDBTABLELeafletPerson & " WHERE " & _
           gk_sDBCOLLeafletReferenceID & " = " & t_lLeafletID
    ofwDBHelper.Execute sSQL
               
    sSQL = "DELETE " & gk_sDBTABLELeafletPlace & " WHERE " & _
           gk_sDBCOLLeafletReferenceID & " = " & t_lLeafletID
    ofwDBHelper.Execute sSQL
    
    sSQL = "DELETE " & gk_sDBTABLELeafletSubject & " WHERE " & _
           gk_sDBCOLLeafletReferenceID & " = " & t_lLeafletID
    ofwDBHelper.Execute sSQL

    sSQL = "DELETE " & gk_sDBTABLELeafletCorpBody & " WHERE " & _
           gk_sDBCOLLeafletReferenceID & " = " & t_lLeafletID
    ofwDBHelper.Execute sSQL
    
    'release object references
    Set ofwDBHelper = Nothing
Exit Sub
ErrorHandler:
   ConvertError sSource
End Sub
Private Sub SaveIndexTermsForLeaflet(ByVal t_lLeafletID As Long)
'
'Description: Saves the Index terms against a Leaflet in the database
'
'
' Amendment history:
'd3     crm     09Nov1999     Created
'
On Error GoTo ErrorHandler
    Dim sSource As String
    Dim vTemp As Variant
    Dim sSQL As String

    Dim ofwDBHelper As New FWDBHelper
    
    
    sSource = "BOLeaflet.SaveIndexTermsForLeaflet"
    
    'save Corporatebody terms for leaflet
    For Each vTemp In m_colCorporateBody
    sSQL = "INSERT " & gk_sDBTABLELeafletCorpBody & " VALUES ( "
    sSQL = sSQL & t_lLeafletID & "," & CLng(vTemp) & ")"
    ofwDBHelper.Execute sSQL
    Next
    
    'save subject terms for leaflet
    For Each vTemp In m_colSubject
    sSQL = "INSERT " & gk_sDBTABLELeafletSubject & " VALUES ( "
    sSQL = sSQL & t_lLeafletID & "," & CLng(vTemp) & ")"
    ofwDBHelper.Execute sSQL
    Next
    
    'save person terms for leaflet
    For Each vTemp In m_colPerson
    sSQL = "INSERT " & gk_sDBTABLELeafletPerson & " VALUES ( "
    sSQL = sSQL & CLng(vTemp) & "," & t_lLeafletID & ")"
    ofwDBHelper.Execute sSQL
    Next
    
    'save place terms for leaflet
    For Each vTemp In m_colPlace
    sSQL = "INSERT " & gk_sDBTABLELeafletPlace & " VALUES ( "
    sSQL = sSQL & t_lLeafletID & "," & CLng(vTemp) & ")"
    ofwDBHelper.Execute sSQL
    Next
    
    Set ofwDBHelper = Nothing
    
Exit Sub

ErrorHandler:
    ConvertError sSource
End Sub
Public Function AddIndexTerm(ByVal t_lTermID As Long, ByVal t_eGroup As EGroup) As Boolean
'
'Description: Adds Index term based on the ID into the private collection
'
'
' Amendment history:
'd3     crm     09Nov1999     Created
'
On Error GoTo ErrorHandler
   
    Select Case t_eGroup
            Case EGroup.gCorporateBody
                m_colCorporateBody.Add t_lTermID, CStr(t_lTermID)
            Case EGroup.gPerson
                m_colPerson.Add t_lTermID, CStr(t_lTermID)
            Case EGroup.gPlace
                m_colPlace.Add t_lTermID, CStr(t_lTermID)
            Case EGroup.gSubject
                m_colSubject.Add t_lTermID, CStr(t_lTermID)
    End Select
    
    AddIndexTerm = True
    
Exit Function
ErrorHandler:
    AddIndexTerm = False
    
End Function

Public Function RemoveIndexTerm(ByVal t_lLinkTermID As Long, ByVal t_eGroup As EGroup) As Boolean
'
'Description: Removes Index term based on the ID from the private collection
'
'
' Amendment history:
'd3     crm     09Nov1999     Created
'

    On Error GoTo ErrorHandler
   
    Select Case t_eGroup
            Case EGroup.gCorporateBody
                m_colCorporateBody.Remove CStr(t_lLinkTermID)
            Case EGroup.gPerson
                m_colPerson.Remove CStr(t_lLinkTermID)
            Case EGroup.gPlace
                m_colPlace.Remove CStr(t_lLinkTermID)
            Case EGroup.gSubject
                m_colSubject.Remove CStr(t_lLinkTermID)
    End Select
    
    RemoveIndexTerm = True
    
Exit Function
ErrorHandler:
    RemoveIndexTerm = False
End Function

Public Function SerialiseLinkItem( _
    ByVal t_sSerialIsedLinkItem As String, _
    ByVal t_vColLinkItem As Collection) As String
'
'Description: Convert the link items (id ) stored in the collection into a string
'             to be stored in persistence manager
'
' Amendment history:
'd3     crm     04Nov1999     Created
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim vTemp  As Variant
                
    sSource = "BOAFPerson.SerialiseLinkItem"
    
    SerialiseLinkItem = t_sSerialIsedLinkItem
    
    'If there are no items in the collection store 0 in the serialised string
    If t_vColLinkItem.Count = 0 Then
        SerialiseLinkItem = AddDataToSerialisedString(SerialiseLinkItem, CStr(0))
    Else
        'Store the number of items in the collection in the string
        SerialiseLinkItem = AddDataToSerialisedString(SerialiseLinkItem, CStr(t_vColLinkItem.Count))
        For Each vTemp In t_vColLinkItem
            'Serialise the id
            SerialiseLinkItem = _
                AddDataToSerialisedString(SerialiseLinkItem, CStr(vTemp))
        Next
    End If
    
    Exit Function
    
ErrorHandler:
    ConvertError sSource
End Function
Public Function ExtractLinkItemsFromSerialIsedString( _
        ByRef t_sSerialisedString As String) As Collection
'
'Description: Convert the link items (id) stored in the string into
'             a collection
'
' Amendment history:
'd3     crm     03Nov,1999     Created
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim lTemp  As Variant, iTotalRecords As Integer, iCounter As Integer
    
    sSource = "BOAFPerson.ExtractLinkItemsFromSerialIsedString"
    
    Set ExtractLinkItemsFromSerialIsedString = New Collection
    
    'Get the total items to be put in the collection,  if there are 0 items then the collection will remain empty
    iTotalRecords = CInt(GetDataFromSerialisedString(t_sSerialisedString))
    
    For iCounter = 1 To iTotalRecords
        'Get the id
        lTemp = CLng(GetDataFromSerialisedString(t_sSerialisedString))
        'Store it on the collection
        ExtractLinkItemsFromSerialIsedString.Add lTemp, CStr(lTemp)
    Next
    
    Exit Function
    
ErrorHandler:
    ConvertError sSource
End Function

Private Sub TransformLeaflet(ByVal t_sLeafletFilePath As String)
'
'Description: Transform the leaflet
'
'
' Amendment history:
'd8     crm     07Dec1999     Created
'd12    CRM     11Feb2000     Using freefile instead of hard-coded number and
'                             PRINT instead of WRITE(which puts in extra quotes in the file)
'd15    TMJ     20Mar2000     Forced file to be normal before writing to it

On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim LinesFromFile As String
    Dim NextLine As String
    Dim sTransformedFile As String
    Dim lFileNumber As Long
        
    sSource = "BOLeaflet.TransformLeaflet()"
    
    'open the file, read it line by line and collect
    'in the string variable
    lFileNumber = FreeFile
    Open t_sLeafletFilePath For Input As #lFileNumber
    Do Until EOF(lFileNumber)
        Line Input #lFileNumber, NextLine
        LinesFromFile = LinesFromFile + NextLine
    Loop
    Close #lFileNumber
    
    'call function to transform each tag
    sTransformedFile = ApplyTransformation(LinesFromFile, "ref")
    sTransformedFile = ApplyTransformation(sTransformedFile, "search")
    
    sTransformedFile = Replace(sTransformedFile, """", "'")
    
    VBA.FileSystem.SetAttr t_sLeafletFilePath, vbNormal
    
    Open t_sLeafletFilePath For Output As #1
        Print #1, sTransformedFile
    Close #1
    
Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub

Private Function ApplyTransformation(ByVal t_sLeafletFile As String, t_sRefOrSearch As String) As String
'
'Description: Transform the leaflet
'
'   Arguments   :   t_sLeafletFile - The file collected in a string variable
'                   t_sRefOrSearch - The tag to search for....
'   Note        :   The tags and their transformations are given below
'
'   1) <ref target="FCO 10">See FCO 10</ref> transformed to:
'      <A TARGET="bottom" HREF="DisplayCatDetailsGoto.asp?CATREF=FCO+10">See FCO10</A>
'   2) <search id="search 2">FCO 10/1-10</search> transformed to:
'      <A TARGET="bottom" HREF="PSRunSearch.asp?sName=search+2">FCO 10/1-10</A>
'
'
' Amendment history:
'd8     crm     07Dec1999     Created
'd11    WDP     31Jan2000     Made hyperlinks relative to Leaflet pages
'd12    CRM     11Feb2000     Made leaflet transformation non case sensitive
'd19    MAP     08Mar2001     Amended to raise an informative error message in case of incorrectly structured tags
'd21    ALM     02Oct2001     Fixed bug where leaflet end was never found
'd22    TomH    25Feb2002     Fixed defect (0215)
'                             problem in loop not being exited properly not sure why
'                             adrian changed in first place.

On Error GoTo ErrorHandler

    Dim sSource As String
    
    Dim lStartPos As Long
    Dim lEndPos As Long
    Dim sReplaceText As String
    Dim sDisplayText As String
    Dim sCATRef As String
    Dim ofwHTML As New FWHTML
    Dim sHTML As String
    
    sSource = "BOLeaflet.ApplyTransformation()"
    
    'Please do not expect me to write comments for this....
    'see the note above and figure out how I have done the transformation
    
    'DO NOT REMOVE the Debug.Print statements...can simplify life if things go wrong
    
    lStartPos = InStr(1, t_sLeafletFile, "<" & t_sRefOrSearch, vbTextCompare)
    lEndPos = InStr(1, t_sLeafletFile, "</" & t_sRefOrSearch & ">", vbTextCompare)
    
    ' (Tomh d22 25-02-02) have changed this back to lStartPos
    ' unsure why adrian changed in first place. Remain to be proved wrong.
    Do Until lStartPos = 0 'd21
        sReplaceText = Mid$(t_sLeafletFile, lStartPos, lEndPos - lStartPos + _
                       Len("</" & t_sRefOrSearch & ">"))
        Debug.Print "The contained string" & " = " & sReplaceText
        sDisplayText = Mid$(Left(sReplaceText, Len(sReplaceText) - Len("</" & _
                        t_sRefOrSearch & ">")), InStr(1, sReplaceText, ">") + 1)
        Debug.Print "The display text " & " = " & sDisplayText
        sCATRef = Mid$(sReplaceText, InStr(1, sReplaceText, "=") + 2, InStr(InStr(1, sReplaceText, "="), sReplaceText, ">") - InStr(1, sReplaceText, "=") - 3)
        Debug.Print "The CAT Ref " & " = " & sCATRef
        
        'if the tag is a 'ref'
        If UCase(t_sRefOrSearch) = "REF" Then
            sHTML = ofwHTML.Hyperlink(mk_sLinkBackToReaderURL & gk_sDisplayCatDetailsGoto & "?" & _
                    "CATREF" & "=" & URLEncode(sCATRef) & "&" & gk_sParamAccessMethod & "=" & _
                    EAccessMethod.EAMLeafletLink, sDisplayText, gk_sFRAMEBottomFrame)
            t_sLeafletFile = Replace(t_sLeafletFile, sReplaceText, sHTML)
        End If
        'if the tag is a search
        If UCase(t_sRefOrSearch) = "SEARCH" Then
            sHTML = ofwHTML.Hyperlink(mk_sLinkBackToReaderURL & gk_sPSRunSearch & "?" & "sName" & "=" & _
                    URLEncode(sCATRef), sDisplayText, gk_sFRAMEBottomFrame)
            t_sLeafletFile = Replace(t_sLeafletFile, sReplaceText, sHTML)
        End If
                
        'd19 New lStartPos now found by looking from end of current lStartPos, rather than after lEndPos
        '    If this new lStartPos is smaller/before the current lEndPos then raise an error
        lStartPos = InStr(lStartPos + Len("<" & t_sRefOrSearch & ">") + 1, _
                    t_sLeafletFile, _
                    "<" & t_sRefOrSearch, vbTextCompare)
        
        
        'lStartPos = InStr(lEndPos + Len("</" & t_sRefOrSearch & ">") + 1, _
                    t_sLeafletFile, "<" & t_sRefOrSearch, vbTextCompare)
                    
        If lStartPos > 0 Then
            If lStartPos < lEndPos Then
                RaiseError sSource, peInvalidLeafletStructure, _
                    "Leaflet has invalid nesting of reference or search tags, and cannot be released"
            Else
                lEndPos = InStr(lStartPos, t_sLeafletFile, "</" & t_sRefOrSearch & ">", vbTextCompare)
            End If
        End If
        
    Loop
    
    'return the transformed string
    ApplyTransformation = t_sLeafletFile
    
Exit Function
ErrorHandler:
    ConvertError sSource
End Function
Function Replace(ByVal t_sValuein As String, ByVal t_sWhatToReplace As _
                       String, ByVal t_sReplacevalue As String) As String
'
'Description: Transform the leaflet
'
'
' Amendment history:
'd8     crm     07Dec1999     Created
'd12    CRM     11Feb2000     Made leaflet transformation non case sensitive
On Error GoTo ErrorHandler
         Dim sSource As String
         Dim sTemp As String
         Dim lPos As Long
         
         sTemp = t_sValuein
         lPos = InStr(1, sTemp, t_sWhatToReplace, vbTextCompare)
         Do While lPos > 0
            sTemp = Left(sTemp, lPos - 1) & t_sReplacevalue & _
                Mid(sTemp, lPos + Len(t_sWhatToReplace))
            lPos = InStr(lPos + Len(t_sWhatToReplace), sTemp, t_sWhatToReplace, vbTextCompare)
         Loop
         Replace = sTemp
Exit Function
ErrorHandler:
        ConvertError sSource
End Function

Function AlreadyExists(t_sFilename As String) As Boolean
'
'Description: Tests a file name to determine whether a leaflet file
'             of the same name already exists in the system
'
'Parameters:  t_sFileName the name of the file to be tested
'
'Return:      A boolean true if a file with the name exists, false
'             otherwise
'
' Amendment history:
'd14    TMJ     22Feb2000       Created

On Error GoTo ErrorHandler
    Dim sSource As String
    Dim lIndex As Long
    
    sSource = "BOLeaflet.AlreadyExists(" & t_sFilename & ")"
    
    If IsNull(m_vArrRegisteredFiles) Then
        AlreadyExists = False
        Exit Function
    Else
        For lIndex = LBound(m_vArrRegisteredFiles) To UBound(m_vArrRegisteredFiles) - 1
            If CStr(m_vArrRegisteredFiles(lIndex)) = t_sFilename Then
                AlreadyExists = True
                Exit Function
            End If
        Next
    End If
    
    AlreadyExists = False
    
Exit Function
ErrorHandler:
    ConvertError sSource
End Function
