VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "VWCatalogueIndexer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    VWCatalogueIndexer.cls
' System:       PRO, PROCAT
' Copyright:    (C) Quidnunc Limited
'
' View class to control the indexing of the entire catalogue
' and view logs
'
' Amendment history:
'   d1  WDP 7-Mar-2000   Created
'   d2  WDP 12-Apr-2000  Added AssignLettercodeCollections
'   d3  DAJ 18-Apr-2000  Removed temporary creation code for TempIndexingProgress
'   d4  DAJ 19-Jul-2000  Changed NOLOCKs for READPASTs in AssignLettercodeCollections
'   d5  Tom 14-Aug-2002  Addded IndexAllLeaflets and IndexAllPopSearches
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Implements ObjectControl
Private m_objContext As Object
Private m_aspResponse As Response
Private m_aspRequest As Request
Private m_aspServer As Server
Private m_aspSession As Session
Private m_ofwProgress As FWRecordList
Private m_ofwTable As FWTable
    
' Constants


Public Sub IndexCatalogue()
'   trigger generation of data files for indexing
'
'   d1  WDP     07-Mar-2000     Created
'   d3  DAJ     18-Apr-2000     Removed temporary creation code for TempIndexingProgress
On Error GoTo ErrorHandler
    
    Dim sSource As String
    sSource = "VWCatalogueIndexer.IndexCatalogue"
                       
    Dim oboGetDat As New BOCatDataExtractor
    
    ' Check to see if indexing steps
    If oboGetDat.IndexStepExists(IXSTEPStart) Then
        oboGetDat.LogIndexStep IXSTEPStart, "Restarting indexing process"
    Else
        oboGetDat.LogIndexStep IXSTEPStart, "Starting indexing process (first pass)"
    End If
        
    If Not oboGetDat.IndexStepExists(IXSTEPIndexingComplete) Then
        
        If Not oboGetDat.IndexStepExists(IXSTEPDataFileCreationCompleted) Then
            oboGetDat.CreateAllCatalogueDataFromLive
        Else
            oboGetDat.LogIndexStep IXSTEPDataFileCreationCompleted, "Data file creation already complete."
        End If
        
        Set oboGetDat = Nothing
        Dim oboIndexer
        Dim sMasterLog As String
        Dim ofwAppData As New FWApplicationData
        
        sMasterLog = ofwAppData.GetApplicationConfigurationInfomation(pRootForIndexingFiles) & gk_sMasterIDXLogFile
        Set oboIndexer = CreateObject("PROCATEdVerIdx.BOIndexer")
        
        oboIndexer.IndexFromIndexLog sMasterLog, False
        Set oboGetDat = New BOCatDataExtractor
        
        oboGetDat.LogIndexStep IXSTEPIndexingComplete, "Verity has finished indexing from the master log"
    Else
        oboGetDat.LogIndexStep IXSTEPIndexingComplete, "Creation of data files and indexing already completed successfully"
        m_aspResponse.Write "Indexing process has already been run completely."
    End If
    
    m_objContext.SetComplete
    
Exit Sub
ErrorHandler:
    ' Only use SetAbort if this method did an update,
    ' (and then there should have been no transient storage).
    ' Otherwise use SetComplete
    Set oboGetDat = New BOCatDataExtractor
    oboGetDat.LogIndexStep IXSTEPError, "Error occurred in indexing process. Check log files and re-run."
    m_objContext.SetComplete
    
    ShowError sSource
End Sub

Public Sub ShowProgress()
'   Show entries from progress log
'
'   d1  WDP     7-Mar-2000   Created

On Error GoTo ErrorHandler
    
    Dim sSource As String
    sSource = "VWCatalogueIndexer.ShowProgess"
                       
    Dim ofwHelper As New FWDBHelper
    Set m_ofwTable = New FWTable
    Dim typTabStyle As TableStyle
    typTabStyle.eALIGN = HALLeft
    typTabStyle.lCELLPADDING = 2
    typTabStyle.lCELLSPACING = 2
    typTabStyle.sWidth = "100%"
    
    m_ofwTable.TableStyle = typTabStyle
    
    Dim sSQL As String
    
    sSQL = " SELECT * FROM " & gk_sDBTABLEIndexingProgress & " (nolock) ORDER BY entry_time desc"
    Set m_ofwProgress = ofwHelper.FindBySql(sSQL)
    
    AddProgressTableHeader
    While Not m_ofwProgress.IsAtEnd
        AddProgressTableRow
        m_ofwProgress.MoveNext
    Wend
    
    m_aspResponse.Write m_ofwTable.GetHTML
    
    m_objContext.SetComplete
    
Exit Sub
ErrorHandler:
    ' Only use SetAbort if this method did an update,
    ' (and then there should have been no transient storage).
    ' Otherwise use SetComplete
    m_objContext.SetComplete
    
    ShowError sSource
End Sub


Public Sub AssignLettercodeCollections()
'   Assign collections to a new lettercode by copying those
'   for an existing one.
'
'   d2  WDP     12-Apr-2000   Created
'   d4  DAJ 19-Jul-2000  Changed NOLOCKs for READPASTs in AssignLettercodeCollections

On Error GoTo ErrorHandler
    
    Dim sSource As String
    sSource = "VWCatalogueIndexer.AssignLettercodeCollections"
                       
    Dim sNewLcode As String, sExistingLcode As String
    sNewLcode = m_aspRequest.Item("newlcode")
    sExistingLcode = m_aspRequest.Item("existinglcode")
    
    Dim lNewLcodeID As Long, lExistingLcodeID As Long
    
    Dim ofwHelper As New FWDBHelper
    Dim ofwResults As FWRecordList
    Dim sSQL As String
    
    
    If sNewLcode = "" Then
        m_aspResponse.Write "You must enter a new lettercode."
        Exit Sub
    End If
    If sExistingLcode = "" Then
        m_aspResponse.Write "You must enter an existing lettercode."
        Exit Sub
    End If
    
    ' d4 - Changed the NOLOCKs to READPASTs; this prevents us reading a lettercode that is being
    '      loaded in another transaction which has a lock on it - this will cause the INSERT
    '      below to block until the load is finished.
    sSQL = " SELECT " & gk_sDBCOLLetterCodeID & " FROM " & gk_sDBTABLELCode & " (READPAST) WHERE " & gk_sDBCOLLetterCodeCode & _
            "=" & QuoteValueWithQuotes(sNewLcode)
            
    Set ofwResults = ofwHelper.FindBySql(sSQL)
    
    If ofwResults.RecordCount < 1 Then
        m_aspResponse.Write "Error: the new lettercode does not exist."
        Exit Sub
    Else
        lNewLcodeID = ofwResults.GetFieldValue(gk_sDBCOLLetterCodeID)
    End If
    
    ' d4 - Changed the NOLOCKs to READPASTs
    sSQL = " SELECT " & gk_sDBCOLLetterCodeID & " FROM " & gk_sDBTABLELCodeVCollection & " (READPAST) WHERE " & gk_sDBCOLLetterCodeID & _
            "=" & lNewLcodeID
            
    Set ofwResults = ofwHelper.FindBySql(sSQL)
    
    If Not ofwResults.RecordCount < 1 Then
        m_aspResponse.Write "Error: the new lettercode already has collections assigned to it."
        Exit Sub
    End If
    
    ' d4 - Changed the NOLOCKs to READPASTs
    sSQL = " SELECT l." & gk_sDBCOLLetterCodeID & " FROM " & gk_sDBTABLELCode & " l (READPAST), " & _
            gk_sDBTABLELCodeVCollection & " v (READPAST) WHERE l." & gk_sDBCOLLetterCodeCode & _
            "=" & QuoteValueWithQuotes(sExistingLcode) & " AND l." & gk_sDBCOLLetterCodeID & "=v." & gk_sDBCOLLetterCodeID

            
    Set ofwResults = ofwHelper.FindBySql(sSQL)
    
    If ofwResults.RecordCount < 1 Then
        m_aspResponse.Write "Error: the existing lettercode does not exist or has no collections assigned to it."
        Exit Sub
    Else
        lExistingLcodeID = ofwResults.GetFieldValue(gk_sDBCOLLetterCodeID)
    End If
    
    sSQL = "INSERT INTO " & gk_sDBTABLELCodeVCollection & "(" & _
            gk_sDBCOLLetterCodeID & ", " & gk_sDBCOLVerityCollectionID & ", " & _
            gk_sDBCOLCollectionContextID & ", " & gk_sDBCOLLcodeIsDepthRestricted & ")" & _
            " SELECT " & lNewLcodeID & ", " & gk_sDBCOLVerityCollectionID & ", " & _
            gk_sDBCOLCollectionContextID & ", " & gk_sDBCOLLcodeIsDepthRestricted & _
            " FROM " & gk_sDBTABLELCodeVCollection & " lvc WHERE " & _
            " lvc." & gk_sDBCOLLetterCodeID & "=" & lExistingLcodeID
    
    ofwHelper.Execute sSQL
    
    m_aspResponse.Write "Done."
    
    m_objContext.SetComplete
    
Exit Sub
ErrorHandler:
    ' Only use SetAbort if this method did an update,
    ' (and then there should have been no transient storage).
    ' Otherwise use SetComplete
    m_objContext.SetComplete
    
    ShowError sSource
End Sub

Public Sub IndexAllLeaflets()
'   Recurses through the leaflets db table indexing all the leaflets
'
'   d5  TH      13-Aug-2002      Created
'
On Error GoTo ErrorHandler

    Dim sSource As String
    sSource = "VWCatalogueIndexer.IndexAllLeaflets"
    
    Dim oboLeafletIndexer As BOLeafletIndexer
    Set oboLeafletIndexer = New BOLeafletIndexer
    
    oboLeafletIndexer.IndexBulkLeaflet
    
    m_objContext.SetComplete
    
Exit Sub
ErrorHandler:
    m_objContext.SetComplete
    
    ShowError sSource
End Sub

Public Sub IndexAllPopSearches()
'   Recurses through the pop searches db table indexing all the popsearches
'
'   d5  TH      13-Aug-2002      Created
'
On Error GoTo ErrorHandler

    Dim sSource As String
    sSource = "VWCatalogueIndexer.IndexAllPopSearches"
    
    Dim oboPopSearchIndexer As BOPopularSearchIndexer
    Set oboPopSearchIndexer = New BOPopularSearchIndexer
    
    oboPopSearchIndexer.IndexAllPopSearches
    
    m_objContext.SetComplete
    
Exit Sub
ErrorHandler:
    m_objContext.SetComplete
    
    ShowError sSource
End Sub


Private Sub AddProgressTableHeader()
'   Add row to member table
'
'   d1  WDP     7-Mar-2000   Created

On Error GoTo ErrorHandler
    
    Dim sSource As String
    sSource = "VWCatalogueIndexer.AddProgessTAbleRow"
                
    m_ofwTable.StartRow
    m_ofwTable.AddHCellAsString "Time"
    m_ofwTable.AddHCellAsString "Message"
    m_ofwTable.AddHCellAsString "(Lettercode ID)"
    m_ofwTable.AddHCellAsString "(Class ID)"
    m_ofwTable.EndRow
    
Exit Sub
ErrorHandler:
    ' Only use SetAbort if this method did an update,
    ' (and then there should have been no transient storage).
    ' Otherwise use SetComplete
    ConvertError sSource
End Sub

Private Sub AddProgressTableRow()
'   Add row to member table
'
'   d1  WDP     7-Mar-2000   Created

On Error GoTo ErrorHandler
    
    Dim sSource As String
    sSource = "VWCatalogueIndexer.AddProgessTAbleRow"
                
    Dim sDate As String, sMessage As String, sLcodeID As String, sClassID As String
    sDate = VariantStringAsString(m_ofwProgress.GetFieldValue(gk_sDBCOLIndexingProgressTime))
    sMessage = VariantStringAsString(m_ofwProgress.GetFieldValue(gk_sDBCOLIndexingProgressMessage))
    sLcodeID = VariantStringAsString(m_ofwProgress.GetFieldValue(gk_sDBCOLIndexingProgressLettercodeID))
    sClassID = VariantStringAsString(m_ofwProgress.GetFieldValue(gk_sDBCOLIndexingProgressClassID))
    
    m_ofwTable.StartRow
    m_ofwTable.AddCellAsString sDate
    m_ofwTable.AddCellAsString sMessage
    m_ofwTable.AddCellAsString sLcodeID
    m_ofwTable.AddCellAsString sClassID
    m_ofwTable.EndRow
    
Exit Sub
ErrorHandler:
    ' Only use SetAbort if this method did an update,
    ' (and then there should have been no transient storage).
    ' Otherwise use SetComplete
    ConvertError sSource
End Sub
''''''''''''''''' Implementation ''''''''''''''''''''''

'
' Private functions
'


'
' Class initialisation and MTS event handlers
'

Private Sub Initialise()
    Set m_objContext = GetObjectContext()
    
    ' Set asp objects
    Set m_aspResponse = m_objContext.Item("Response")
    Set m_aspRequest = m_objContext.Item("Request")
    Set m_aspServer = m_objContext.Item("Server")
    Set m_aspSession = m_objContext.Item("Session")
    
    ' Specific initialisation
    ' <Specific initialisation here>
    
End Sub

Private Sub ObjectControl_Activate()
    Initialise
End Sub

Private Sub ObjectControl_Deactivate()
    Set m_objContext = Nothing
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = True
End Function










