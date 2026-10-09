VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "VWReleaser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    VWTReleaseEditSet.cls
' System:       PRO, PROCAT
' Copyright:    (C) Quidnunc Limited
'
' Description:  Class to trigger the process of releasing
'               edit sets from
'
' Amendment history:
'   d1  WDP     12-Nov-1999   Created
'   d2  WDP     11-Dec-2000   Modified Release function to create a verity index
'                             file if it does not already exist
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Implements ObjectControl

Private m_objContext As Object
Private m_aspResponse As Response
Private m_aspRequest As Request
Private m_aspServer As Server
Private m_aspSession As Session

Private m_colEditSets As Collection
Private m_colMovedCatalogueEntries As Collection

' Constants


Public Sub Release()
'   Master control for releasing edit sets
'
'   d1  WDP     16-Nov-1999   Created

On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim oboReleaser
    Dim oboIndexer
    sSource = "VWReleaseEditSet.Release"
    
    LogReleaseMessage "Beginning release process", RELInfo
    
    'Get which edit sets are releasable
    Set oboReleaser = CreateObject("PROCATEdReleas.BOReleaser")
    LogReleaseMessage "Finding releasable edit sets...", RELInfo
    Set m_colEditSets = oboReleaser.FindReleasableEditSets
    
    If m_colEditSets.Count > 0 Then
        
        'Log which editsets we are going to attempt to release
        Dim i As Long
        Dim sMsg As String
        For i = 1 To m_colEditSets.Count
            If Not sMsg = "" Then
                sMsg = sMsg & ", "
            End If
            sMsg = sMsg & m_colEditSets.Item(i)
        Next
        sMsg = "Editset IDs ready for release: " & sMsg
        LogReleaseMessage sMsg, RELInfo
        
        Dim l As Long
        Dim sDate As String
        'We use a date string to identify files used in the
        'release process
        sDate = GetDateString()
        'We go through the releasable edit sets one by one
        'executing all the processes needed to release them
        For l = 1 To m_colEditSets.Count
            Dim lID As Long
            lID = m_colEditSets.Item(l)
            
            LogReleaseMessage "Started release of editset " & lID, RELInfo
            'The release procedure writes data to files
            'and creates a log for indexing which has a
            'time filename prefix. This string is returned by the
            'release procedure and needs to be passed to
            'the indexer so it can read the appropriate log
            
            Dim sTime As String, sMoveTime As String
            'Collection to hold the IDs of any pieces or classes
            'that have been moved within the catalogue (this
            'operation can only be applied to pieces and classes)
            Dim colMovedEntries As New Collection
            Dim bReleaseSucceeded As Boolean
            Dim bMoveRes As Boolean
                    
            'Run the database release and generation of indexing files
            'This also finds any moved pieces or classes
            bReleaseSucceeded = oboReleaser.ReleaseEditsetInDB(lID, sDate, sTime, colMovedEntries)
            If bReleaseSucceeded Then
                'Run the indexing from the indexing log
                Dim sLogFile As String
                'Create the indexer component (this runs
                'on the Verity server via DCOM)
                Set oboIndexer = CreateObject("PROCATEdVerIdx.BOIndexer")
                
                Dim oboCatData As New BOCatDataExtractor
                sLogFile = oboCatData.LogFileName(sDate & sTime, lID)
                Set oboCatData = Nothing
                LogReleaseMessage "Editset " & lID & " indexing data files created, actions to perform in logfile: " & sLogFile, RELInfo
                
                oboIndexer.IndexFromIndexLog sLogFile, True
                Set oboIndexer = Nothing
                'If there are any moved pieces or classes then we
                'must reindex all of the catalogue below them as
                'the reference parts of their children will have
                'changed so need reindexing
                If colMovedEntries.Count > 0 Then
                    
                    Dim sTxt As String, sText As String
                    Dim j As Long
                    For j = 1 To colMovedEntries.Count
                        If Not sTxt = "" Then
                            sTxt = sTxt & ", "
                        End If
                        sTxt = sTxt & colMovedEntries.Item(j)
                    Next
                    sText = "Editset " & lID & ", entries that have been moved: " & sTxt
                    LogReleaseMessage sText, RELInfo
        
                    'Create the data files for the moved entries
                    bMoveRes = oboReleaser.CreateMovedDataFiles(colMovedEntries, sDate, sMoveTime, lID)
                    If bMoveRes = False Then
                        LogReleaseMessage "Creation of indexing data files for moved entries failed." & _
                                    "Run this procedure manually on the following catalogue entries: " & sTxt, RELError
                                    
                    Else
                        Set oboIndexer = CreateObject("PROCATEdVerIdx.BOIndexer")
                        Dim oboCatDExtr As New BOCatDataExtractor
                        sLogFile = oboCatDExtr.LogFileName(sDate & sMoveTime, lID, True)
                        'd2  WDP  11-Dec-2000
                        'Log a dummy action just in case there are no entries to reindex
                        'so we can be sure that the action log exists
                        LogAction sLogFile, t_bDummyAction:=True
                        Set oboCatDExtr = Nothing
                        LogReleaseMessage "Editset " & lID & _
                            ": indexing data files created, action to perform in log: " & sLogFile, RELInfo
                        'Run the indexing from the indexing log
                        oboIndexer.IndexFromIndexLog sLogFile
                    End If
                End If
                'Finally we delete the edit set
                Dim bDeleteSucceeded As Boolean
                bDeleteSucceeded = oboReleaser.DeleteEditset(lID)
                If bDeleteSucceeded = False Then
                    'log fact that delete failed
                    LogReleaseMessage "Deletion of editset " & lID & _
                        " failed, run manual deletion of this edit set before next scheduled release.", RELError
                Else
                    LogReleaseMessage "Editset " & lID & " deleted from database.", RELInfo
                    If bMoveRes = True And bReleaseSucceeded = True Then
                        LogReleaseMessage "Release of editset " & lID & " completed successfully. ", RELInfo
                    End If
                End If
                
            Else
                'log message saying release of that editset completely unsuccessful
                LogReleaseMessage "Release of Editset " & lID & " unsuccessful, all actions rolled back. Run manual release for that edit set or leave until next scheduled release", RELError
            End If
            
        Next
            
    Else
        LogReleaseMessage "No edit sets are releasable", RELInfo
    End If
    LogReleaseMessage "Release process complete.", RELInfo
    
    Set oboReleaser = Nothing
    Set oboIndexer = Nothing
    'The object context is not obtained
    'm_objContext.SetComplete
Exit Sub
ErrorHandler:
    Set oboReleaser = Nothing
    Set oboIndexer = Nothing
    ' Only use SetAbort if this method did an update,
    ' (and then there should have been no transient storage).
    ' Otherwise use SetComplete
    '+ track what needs to be done still...
    LogReleaseMessage "Error in release process, release not complete. Error: " & Err.Description, RELError
    'm_objContext.SetComplete
    
    ShowError sSource
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
    'Set m_aspResponse = m_objContext.Item("Response")
    'Set m_aspRequest = m_objContext.Item("Request")
    'Set m_aspServer = m_objContext.Item("Server")
    
    
End Sub

Public Sub TestTransaction(t_bRaiseError As Boolean)
'   test transaction rollback/commit
'
'   d2  WDP     24-Nov-1999   Created

On Error GoTo ErrorHandler
    
    Dim sSource As String
    sSource = "VWReleaseEditSet.TestTransaction"
    
    'Get which edit sets are releasable
    Dim oboReleaser
    Set oboReleaser = CreateObject("PROCATEdReleas.BOReleaser")
    oboReleaser.TestTransaction t_bRaiseError
    oboReleaser.TestTransaction False
    
    m_objContext.SetComplete
Exit Sub
ErrorHandler:
   
    ShowError sSource
    'm_objContext.SetComplete
End Sub

Public Sub TestMethod()
'Test method to test scheduled tasks are working
On Error GoTo ErrorHandler
    
    Dim sSource As String
    sSource = "VWReleaser.TestMethod"
    
    Dim obj As New FWDBHelper
    obj.Execute ("INSERT INTO zTempTable(id) VALUES (1)")
    
    'Dim obj As New BOReleaser
    'obj.TestMethod False
    
    'm_objContext.SetComplete
Exit Sub
ErrorHandler:
    ShowError sSource
    
End Sub

Public Function TestSum(x As Long, y As Long) As Long
'Test method to test scheduled tasks are working
On Error GoTo ErrorHandler
    
    Dim sSource As String
    sSource = "VWReleaser.TestSum"
    
    TestSum = x + y
    'Dim obj As New BOReleaser
    'obj.TestMethod False
    
    'm_objContext.SetComplete
Exit Function
ErrorHandler:
    ShowError sSource
    
End Function

Private Sub ObjectControl_Activate()
    Initialise
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = True
End Function

Private Sub ObjectControl_Deactivate()

End Sub
