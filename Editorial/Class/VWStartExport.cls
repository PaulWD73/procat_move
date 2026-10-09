VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "VWStartExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    VWStartExport.cls
' System:       PRO, PROCAT
' Copyright:    (C) Quidnunc Limited
'
' Description:  This class is used to start the export of information from the
'               database
'
'       This is done by calling a shell command to run an executable
'
'
'
' Amendment history from VWStartExport:
'   Vers Auth    Date        Comment
'   d1   ADF     30Jun1999   Created
'   d2 MAP     22Oct2001    Added code to handle EAD scheduling
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Implements ObjectControl
Private m_objContext As ObjectContext
Private m_aspResponse As Response
Private m_aspRequest As Request
Private m_aspServer As Server
Private m_aspSession As Session

Private m_lBatchID As Long
Private m_sFileName As String
Private m_sBatchDescription As String
Private m_oFWAppData As FWApplicationData

Private m_lLoggedUserID  As Long
Private m_oBOEditorialPermissions As BOEditorialPermissions

Private Const m_ksCallMethod As String = "CM:"
Private Const m_ksFileName As String = "FN:"
Private Const m_ksBatchID As String = "BID:"
Private Const m_ksEditSetID As String = "EID:"






Public Sub StartExportToEAD()
'
'   Description :- This method starts the export to an EAD file by calling a shell command
'   and returning control back to the user
'
'   Vers  Auth  Date        Comments
'   d1    ADF   15Nov1999   Created
'   d2  MAP     09Oct2001   Added scheduling functionality
'
On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim sShellCommand As String
    Dim sCommandArgs As String
    Dim ofwAppConfig As New FWApplicationData
    ' d2
    Dim oboScheduler As New BOScheduler
    Dim bScheduledJob As Boolean
    
    sSource = "VWStartExport.StartExportToEAD"

    m_sFileName = m_aspRequest(gk_sParamFileName)

    m_sBatchDescription = m_sFileName & "_" & GetTimeStamp
    
    ' d2 - Check if job is to be scheduled
    If m_aspRequest(gk_sParamJobStartTime) <> "" Then
        bScheduledJob = True
    Else
        bScheduledJob = False
    End If
    
    ' d2 - Added scheduled job flag to end
    m_lBatchID = GenerateBatchRecord(m_sFileName, m_sBatchDescription, bScheduledJob)
    
    ' We now put all the export form info into a table to be retrived once the
    ' asyncloader has triggered the Export componenent
    GenerateExportTableRecord
    
    ' d2 If Scheduling the export, load the Scheduler object
    If bScheduledJob = True Then
        
        oboScheduler.StartTime = m_aspRequest(gk_sParamJobStartTime)
        oboScheduler.StartDate = m_aspRequest(gk_sParamJobStartDate)
        oboScheduler.BatchID = m_lBatchID
        oboScheduler.JobType = jtEADExport
        
        oboScheduler.CreateScheduledJob
        
    Else
    
        LogLoadInfo m_lBatchID, "Started Export to file '" & m_sFileName & "'", , , pMandatoryInfo
        LogLoadInfo m_lBatchID, "Batch description - " & m_sBatchDescription
        
        
        sCommandArgs = m_ksCallMethod & ECallMethod.cmEADExporting & " " & _
                       m_ksBatchID & m_lBatchID
        
        sShellCommand = ofwAppConfig.GetApplicationConfigurationInfomation(pAsyncLoaderPath) & " " & sCommandArgs
        Shell sShellCommand
            
    End If
            
    Set ofwAppConfig = Nothing
    
    'Return disabling the commit so that the Insertion starts
    '   in the same transaction
    m_objContext.SetComplete
    
    ' d2 - display scheduled job confirmation page if appropriate
    If bScheduledJob = True Then
        m_aspResponse.Redirect gk_sLinkConfirmJobCreation & "?" & _
            gk_sParamFileName & "=" & m_sFileName & "&" & gk_sParamJobType & "=" & CStr(jtEADExport)
    Else
        m_aspResponse.Redirect gk_sLinkDisplayBatchList
    End If
    
Exit Sub
ErrorHandler:
    m_objContext.SetComplete
    ShowError sSource
End Sub








''''''''''''''''' Implementation ''''''''''''''''''''''

'
' Private functions
'

Private Function GenerateBatchRecord(t_sFilename As String, _
                                t_sDescription As String, _
                                Optional t_bScheduledJob As Boolean = False) As Long
                                
                                '
'   Description :- This method generates a batch id
'
'   Vers  Auth  Date            Comments
'   d1    ADF   24-Jul-2000     Created
'   d2    MAP   17-Oct-2001     Added scheduled job flag to header
'
On Error GoTo ErrorHandler
'Thismethod needs to be called before ExportToEADFile
' this sets the member varaiables and creates the batch ID required for
'   the import
'
Dim sSource As String

    sSource = "VWStartImport.GenerateBatchRecord"
    
    m_sFileName = t_sFilename
    m_sBatchDescription = t_sDescription
    
    Dim oboBatchExport As New BOAccessionBatch

    'Add a row into the Accession Batch tables
    oboBatchExport.Description = m_sBatchDescription
    ' d2 - Added
    If t_bScheduledJob = False Then
        oboBatchExport.ImportStatus = EImportStatus.isEADExportStarted
    Else
        oboBatchExport.ImportStatus = EImportStatus.isEADExportScheduled
    End If
    oboBatchExport.ImportType = EImportLoadType.iltEAD
    oboBatchExport.ImportFileName = t_sFilename
    oboBatchExport.UPDATE
    
    m_lBatchID = oboBatchExport.BatchID
    
    GenerateBatchRecord = m_lBatchID
    
    LogLoadInfo m_lBatchID, "Added an entry into the batch accession table in DRUID; ID = " & m_lBatchID
    
    m_objContext.SetComplete
    
    Exit Function
ErrorHandler:
    
    m_objContext.SetAbort
    
    ConvertError sSource
End Function


'
' Class initialisation and MTS event handlers
'

Private Sub Initialise()

    Dim sSource As String
    
    '  Just testing that we get all the params correctly from the request object
    Dim sTestparam As String
    Dim sTestparam1 As String
    Dim sTestparam2 As String
    Dim sTestparam3 As String
    Dim sTestparam4 As String
    Dim sTestparam5 As String
    Dim sTestparam6 As String

    sSource = "VWExport.Initialise"
    
    Set m_objContext = GetObjectContext()
    
    ' Set asp objects only if called from a asp
    If m_objContext.Count > 0 Then
        Set m_aspResponse = m_objContext.Item("Response")
        Set m_aspRequest = m_objContext.Item("Request")
        Set m_aspServer = m_objContext.Item("Server")
        Set m_aspSession = m_objContext.Item("Session")
        
        m_lLoggedUserID = GetDataFromSession(dsUserID, m_objContext.Item(gk_sSessionObject))
        If m_lLoggedUserID = 0 Then     'Not a valid user
            RaiseError sSource, EPROCATINTERNALERROR.peDBInvalidUser, "Invalid User ID : " & m_lLoggedUserID
        End If
        
    End If
    
    Set m_oFWAppData = New FWApplicationData
    Set m_oBOEditorialPermissions = New BOEditorialPermissions
    
End Sub

Private Sub Terminate()
    
    Set m_aspResponse = Nothing
    Set m_aspRequest = Nothing
    Set m_aspServer = Nothing
    Set m_aspSession = Nothing
    
    Set m_oFWAppData = Nothing
    Set m_objContext = Nothing
    Set m_oBOEditorialPermissions = Nothing
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



Public Sub GenerateExportTableRecord()
'
'   Description :- This sub puts all the details of the export form into a table to be referenced
'                   later.
'
'   Vers  Auth          Date            Comments
'   d1    A.Farrow      24-Jul-2000     Created
'

    Dim oboExportTable As New BOExportTable
    
    '  Set the boolean values
    If m_aspRequest("IsDeptWanted") = "" Then
        oboExportTable.IsDepartmentWanted = False
    Else
        oboExportTable.IsDepartmentWanted = True
    End If
    
    If m_aspRequest("IsDivWanted") = "" Then
        oboExportTable.IsDivWanted = False
    Else
        oboExportTable.IsDivWanted = True
    End If
    
    If m_aspRequest("IsSeriesWanted") = "" Then
        oboExportTable.IsSeriesWanted = False
    Else
        oboExportTable.IsSeriesWanted = True
    End If
    
    If m_aspRequest("IsSubseriesWanted") = "" Then
        oboExportTable.IsSubSeriesWanted = False
    Else
        oboExportTable.IsSubSeriesWanted = True
    End If
    
     If m_aspRequest("IsSubsubseriesWanted") = "" Then
        oboExportTable.IsSubSubSeriesWanted = False
    Else
        oboExportTable.IsSubSubSeriesWanted = True
    End If
    
    If m_aspRequest("IsPieceWanted") = "" Then
        oboExportTable.IsPieceWanted = False
    Else
        oboExportTable.IsPieceWanted = True
    End If
    
    If m_aspRequest("IsItemWanted") = "" Then
        oboExportTable.IsItemWanted = False
    Else
        oboExportTable.IsItemWanted = True
    End If
    
    ' Get the IDS and then send them to the update table sub
    Dim oboExportIds As New BOExportIDs
    
    oboExportIds.LetterCode = m_aspRequest("LetterCodeDBID")
    oboExportIds.MaximumPieceKeyOrder = m_aspRequest("MaxPOrder")
    oboExportIds.MinimumPieceKeyOrder = m_aspRequest("MinPOrder")
    
    Dim dummy As String
    dummy = oboExportTable.UpdateExportTable(m_sFileName, oboExportIds, m_aspRequest("SeriesIDs"), m_lBatchID)
    
End Sub
