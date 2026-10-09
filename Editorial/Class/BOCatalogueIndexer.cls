VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "BOIndexer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    BOCatalogueIndexer.cls
' System:       PRO, PROCAT
' Copyright:    (C) Quidnunc Limited
'
' Description:
'
' Amendment history:
'   d1  WDP     5-Oct-1999      Created
'   d2  WDP     25-Oct-1999     Added Link data retrival and indexing
'   d3  WDP     29-Oct-1999     Added large edit set capabilities
'   d4  WDP     11-Nov-1999     Fixed defects, removed TruncateTitle
'                               to GlobalFunctions
'   d5  WDP     16-Nov-1999     removed GetDateString to GlobalFunctions
'                               Separated writing to file from
'                               performing the indexing
'   d6  WDP     19-Nov-1999     Added piece moving facilities and
'                               indexing from live
'   d7  WDP     26-Npv-1999     Added CreateIndexDataForEntry
'   d8  WDP     13-Dec-1999     Enabled selection of entire level
'                               for subheader upwards
'   d9  WDP     27-Jan-2000     Split functionality into this and BOCatDataExtractor
'                               This class only performs indexing from logfiles.
'   d10 WDP     06-Apr-2000     Added optional flag params to IndexFromIndexLog
'   d11 TomH    23-Aug-2002     Altered mkvdk so now uses fastsearch mode
'
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

' action log file text
Private m_sActionLogText As String

'
' Public Methods
'

Public Sub IndexFromIndexLog(t_sLogFile As String, Optional t_bForRelease As Boolean = False, Optional t_bRaiseError As Boolean = False)
'   This method runs MKVDK from instructions
'   in the given logfile. The t_bForRelease parameter is used
'   to distinguish actions that are for editset release (as messages are
'   logged in a different way. The t_bRaiseError is used to enable errors
'   to be raised when we are indexing pop searches and leaflets
'   (we dont raise errors when indexing from live)
'
'   d1  WDP     18-Nov-1999     Created
'   d10 WDP    06-Apr-2000     Added optional params
On Error GoTo ErrorHandler

    Dim sSource As String
    sSource = "BOCatalogueIndexer.IndexFromIndexLog"
    
    If t_bForRelease = True Then
        LogReleaseMessage "Indexing from log: " & t_sLogFile, RELInfo
    End If
    
    Dim lFileNum As Long
    
    lFileNum = FreeFile
    Open t_sLogFile For Input Access Read As lFileNum
    If lFileNum > 0 Then
        m_sActionLogText = StrConv(InputB(LOF(lFileNum), lFileNum), vbUnicode)
        Close #lFileNum
           
        Dim sCollection As String, sBulkfile As String, sLetterCode As String
        Dim eAction As EIndexingAction, lEditsetID As Long
        
        While Not (FindTaggedData(m_sActionLogText, gk_sTAGEntry)) = ""
            ReadAction sBulkfile, sLetterCode, sCollection, eAction, lEditsetID
            RunMKVDK eAction, sCollection, sBulkfile
            CheckLogFileForErrors sBulkfile & ".log", t_bForRelease, t_bRaiseError
        Wend
    Else
        If t_bForRelease = True Then
            LogReleaseMessage "Unable to open logfile: " & t_sLogFile & ". This could be because no indexing was required", RELError
        Else
            RaiseError sSource, PEGeneralInternalError, "Couldnt open log file"
        End If
    End If
    Exit Sub
ErrorHandler:
    If t_bForRelease = True Then
        LogReleaseMessage "Error indexing from logfile: " & t_sLogFile & " " & Err.Description, RELError
    Else
        ConvertError sSource
    End If
End Sub

Private Sub RunMKVDK(t_eAction As EIndexingAction, t_sCollection As String, t_sBLKFilepath As String)
'   Run the MKVDK bulk command on given collection
'   with the appropriate command (insert, update etc.)
'   using the given bulk file
'
'   d1  WDP     7-Oct-1999      Created
'   d11 TomH    23-Aug-2002     Altered so now uses fast search mode
On Error GoTo ErrorHandler

    Dim sSource As String
    sSource = "BOCatalogueIndexer.RunMKVDK"
    
    Dim sBulkCommand As String
    Dim sCommandToRun As String
    Dim sCommand As String
    'Dim iRes As Integer
    
    'set the collection command (update or insert)
    sCommand = "-" & GetIndexActionString(t_eAction) & " "
    'create the command to run
    sCommandToRun = "mkvdk -mode fastsearch -quiet -collection " & t_sCollection & _
                    " -bulk " & sCommand & t_sBLKFilepath & _
                    " -logfile " & t_sBLKFilepath & ".log"
    
    LogMessage "Command to run: " & sCommandToRun, pComplete
    If Not ExecCmd(sCommandToRun) Then
        '+ log this somewhere else
        RaiseError sSource, PEGeneralInternalError, "Error running mkvdk."
    End If
    Exit Sub
    
ErrorHandler:
    ConvertError sSource
    
End Sub

Private Sub ReadAction( _
    ByRef t_sBulkFilePath As String, _
    ByRef t_sLettercode As String, _
    ByRef t_sCollection As String, _
    ByRef t_eAction As EIndexingAction, _
    ByRef t_lEditsetID As Long)
'   Reads indexing info from the log file (stored as a member variable)
'   string into the parameters
'   passed by reference and removes the read section from the string
'
'   d5  WDP     18-Nov-1999     Created
On Error GoTo ErrorHandler

    Dim sSource As String
    sSource = "BOCatalogueIndexer.ReadAction"
    
    Dim sData As String
    Dim sEndOfEntry As String
    Dim lPos As Long
    
    sEndOfEntry = "</" & gk_sTAGEntry & ">"
    lPos = InStr(m_sActionLogText, sEndOfEntry) + Len(sEndOfEntry)
    sData = Left(m_sActionLogText, lPos)
    m_sActionLogText = Mid(m_sActionLogText, lPos + 1)
    
    t_sBulkFilePath = Trim(FindTaggedData(sData, gk_sTAGBulkFile))
    t_sLettercode = Trim(FindTaggedData(sData, gk_sTAGLettercode))
    t_sCollection = Trim(FindTaggedData(sData, gk_sTAGCollection))
    t_eAction = GetIndexActionEnum(FindTaggedData(sData, gk_sTAGAction))
    Dim sEditsetID As String
    sEditsetID = Trim(FindTaggedData(sData, gk_sTAGEditsetID))
    If sEditsetID = "" Then
        t_lEditsetID = 0
    Else
        t_lEditsetID = CLng(Trim(FindTaggedData(sData, gk_sTAGEditsetID)))
    End If
    
    LogMessage "Action read from logfile: " & "bulkfile: " & t_sBulkFilePath & ". Collection: " & t_sCollection & ". Lettercode: " & t_sLettercode, pComplete
    
    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub

Private Sub Initialise()
    
    ' Specific initialisation
    ' <Specific initialisation here>
    
End Sub

Private Sub Class_Initialize()
    Initialise
End Sub

Private Sub Class_Terminate()
    ' specific clear up
End Sub


