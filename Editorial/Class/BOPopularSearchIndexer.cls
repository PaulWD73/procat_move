VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BOPopularSearchIndexer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    BOPopularSearchIndexer.cls
' System:       PRO, PROCAT
' Copyright:    (C) Quidnunc Limited
'
' Description:  Business object used to insert, update or delete
'               Popular searches from the popular search verity
'               collection
'
' Amendment history:
'   d1  WDP     12-Oct-1999      Created
'   d2  WDP     27-Jan-2000      Removed mkvdk call, this is now done
'                                by BOIndexer running on the Verity server
'                                called using DCOM
'   d3  WDP     15-Feb-2000      Altered bulk file nameing to avoid name clash
'   d4  WDP     29-Mar-2000      Added different behaviour for delete
'   d5  DAJ     09-Jun-2000      Added some special character stripping from filenames
'   d6  Tomh    14-Aug-2002      Added IndexAllPopSearches
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Class constants

Private Const mk_sFilepath As String = "VerityWrapper\PROCAT\PopularSearchData\"
Private Const mk_sBulkActionLogFileName As String = "MasterPopSearchIndexingLogFile"


Private m_oboSearch As BOPopSearchLoader

Public Function IndexPopSearch(t_sSearchName As String, t_eAction As EIndexingAction, Optional t_sActionLogFileName As String = "") As String
'   Index Popular search that has been set as search member
'
'   d1  WDP     14-Oct-1999      Created
'   d2  WDP     27-Jan-2000      Removed mkvdk call, put actions
'                                in logfile for indexing later.
'   d3  WDP     15-Feb-2000      Altered bulk file nameing to avoid name clash
'   d4  WDP     29-Mar-2000      Added different behaviour for delete
'   d5  DAJ     09-Jun-2000      Added some special character stripping from filenames
'   d6  TomH    13-Aug-2002      Added optional argument for if is index all pop searches
On Error GoTo ErrorHandler

    Dim sSource As String
    sSource = "BOPopularSearchIndexer.IndexPopSearch"
    
    If Not t_eAction = IDXDelete Then
        Set m_oboSearch = New BOPopSearchLoader
        m_oboSearch.LoadData t_sSearchName
    End If
    
    Dim sSearchName As String, iIsReaderSearch As Integer
    Dim sFilename As String, sSGMLFilename As String, sBLKFilename As String
    Dim sSGMLFilenameForVerity As String
    Dim lSGMLFileNum As Long, lBLKFileNum As Long
    Dim sCollectionDrive As String, sIndexFilesDrive As String
    
    Dim ofwAppData As New FWApplicationData
    sCollectionDrive = ofwAppData.GetApplicationConfigurationInfomation(pRootForCollections)
    sIndexFilesDrive = ofwAppData.GetApplicationConfigurationInfomation(pRootForIndexingFiles)
    Dim sFileIdentifier As String
    sFileIdentifier = CStr(Hour(Now())) & CStr(Minute(Now())) & CStr(Second(Now()))
    
    ' d5 - Remove special filename characters
    sSearchName = EncodeFileName(t_sSearchName)
    
    sFilename = sSearchName & "_" & GetIndexActionString(t_eAction)
    sSGMLFilename = sIndexFilesDrive & mk_sFilepath & sFileIdentifier & sFilename & ".sgml"
    sSGMLFilenameForVerity = Replace(sSGMLFilename, "\", "/")
    sBLKFilename = sIndexFilesDrive & mk_sFilepath & sFileIdentifier & sFilename & ".blk"
    
    'Open the files, if they already exist they will be overwritten
    lBLKFileNum = FreeFile
    Open (sBLKFilename) For Output As lBLKFileNum
    If lBLKFileNum = 0 Then
        '+raise error
    End If
    
    'For deletes we just need to write the doc key to the bulk file
    'and neednt bother with the SGML file
    If Not t_eAction = IDXDelete Then
        
        lSGMLFileNum = FreeFile
        Open (sSGMLFilename) For Output As lSGMLFileNum
        If lSGMLFileNum = 0 Then
            '+raise error
        End If
        
        Dim sSGML As String, sBLK As String
        sSGML = "<" & gk_sVZONEPopSearchDescription & ">" & _
                m_oboSearch.Description & "</" & gk_sVZONEPopSearchDescription & ">"
        
        Dim lDocSize As Long
        lDocSize = 0
        
        Print #lSGMLFileNum, sSGML
            lDocSize = Len(sSGML)
            
        If m_oboSearch.IsReaderSearch = True Then
            iIsReaderSearch = 1
        Else
            iIsReaderSearch = 0
        End If
        
        sBLK = gk_sVFIELDVerityKey & ": " & DoubleQuotes(sSearchName) & vbCrLf & _
                "DOC_FN: " & DoubleQuotes(sSGMLFilenameForVerity) & vbCrLf & _
                "DOC_OF: 0" & vbCrLf & _
                "DOC_SZ: " & lDocSize & vbCrLf & _
                gk_sVFIELDPopSearchDescription & ": " & DoubleQuotes(m_oboSearch.Description) & vbCrLf & _
                gk_sVFIELDPopSearchIsRdSearch & ": " & iIsReaderSearch & vbCrLf & _
                "<<EOD>>" & vbCrLf & vbCrLf
        Close #lSGMLFileNum
        
    Else
        sBLK = gk_sVFIELDVerityKey & ": " & DoubleQuotes(sSearchName) & vbCrLf
    End If
    Print #lBLKFileNum, sBLK
    Close #lBLKFileNum
    
    Dim sIndexingLogFile As String
    ' d5 - tomh
    ' if the optional argument has a value then make that the action log file name
    If t_sActionLogFileName = "" Then
        sIndexingLogFile = sIndexFilesDrive & mk_sFilepath & sFileIdentifier & sSearchName & ".log"
    Else
        sIndexingLogFile = sIndexFilesDrive & mk_sFilepath & t_sActionLogFileName & ".log"
    End If

    
    Dim sCollection As String
    sCollection = sCollectionDrive & gk_sPopSearchCollectionPath
    
    'Log the indexing actions to perform in a log file
    LogAction sIndexingLogFile, sBLKFilename, sCollection, t_eAction
            
    'return the indexing action logfile name so we can trigger the
    'indexing process on the Verity server
    IndexPopSearch = sIndexingLogFile
        
    Exit Function
ErrorHandler:
    ConvertError sSource
        
End Function


Public Sub IndexAllPopSearches()
'   Recurses through the popular searches in db table indexing all the pop searches
'
'   d6  TomH      14-Aug-2002      Created
'
On Error GoTo ErrorHandler

    Dim sSource As String
    sSource = "BOPopSearcheIndexer.IndexAllPopSearches"
    
    Dim sIndexingLogFile As String
    Dim sPopSearchName As String
    Dim oboVerIndexer
            
    'Load leaflet data from DB
    Dim sSQL As String
    Dim ofwDBHelper As FWDBHelper
    Set ofwDBHelper = New FWDBHelper
    Dim orsResults As FWRecordList
        
    sSQL = "SELECT " & gk_sDBCOLPSSearchName & _
            " FROM " & gk_sDBTABLEPopularSearch
            
    Set orsResults = ofwDBHelper.FindBySql(sSQL)
    
    LogMessage "Getting pop search file names.", pProfileInformation
        
    orsResults.MoveFirst
    
            
    ' recurse through results and index files
    While Not orsResults.IsAtEnd
    
        ' get pop search info from results
        sPopSearchName = orsResults.GetFieldValue(gk_sDBCOLPSSearchName)
        
        sIndexingLogFile = IndexPopSearch(sPopSearchName, IDXInsert, mk_sBulkActionLogFileName)
        LogMessage "Indexed file " & sPopSearchName, pProfileInformation
                             
       orsResults.MoveNext
        
    Wend
    
    'The Verity server could be remote and so we make a (D)COM
    'call to the indexer to actually index the leaflet into
    'the collection
    Set oboVerIndexer = CreateObject("PROCATEdVerIdx.BOIndexer")
    oboVerIndexer.IndexFromIndexLog sIndexingLogFile, False, True
    Set oboVerIndexer = Nothing
        
    Exit Sub

ErrorHandler:
    ConvertError sSource
End Sub


