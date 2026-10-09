VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BOLeafletIndexer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    BOLeafletIndexer.cls
' System:       PRO, PROCAT
' Copyright:    (C) Quidnunc Limited
'
' Description:  Business object used to insert, update or delete
'               Leaflets from the leaflet Verity collection.
'
' Amendment history:
'   d1  WDP     22-Oct-1999      Created
'   d2  WDP     26-Oct-1999      Added index term indexing
'   d3  WDP     3-Nov-1999       Removed calls to add linked AF terms
'   d4  WDP     27-Jan-2000      Removed mkvdk call, this is done in another component
'                                on the verity server called by DCOM
'   d5  WDP     15-Feb-2000      Give BLK file a identifier based on the time
'                                to avoid name clashes
'   d6  TomH    13-Aug-2002      Added bulk log file indexer
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Class constants

Private Const mk_sFilepath As String = "VerityWrapper\PROCAT\LeafletData\"
Private Const mk_sFilepathForVerity As String = "VerityWrapper/PROCAT/LeafletData/"
Private Const mk_sBulkActionLogFileName As String = "MasterLeafletIndexingLogFile"

Private m_oboLeaflet As BOLeafletLoader

Public Function IndexLeaflet(t_lLeafletID As Long, t_eAction As EIndexingAction, Optional t_sActionLogFileName As String = "") As String
'   Index Leaflet given the ID and indexing action
'   (insert, update or delete)
'
'   d1  WDP     22-Oct-1999      Created
'   d2  WDP     26-Oct-1999      Added index term fields to bulk file
'   d4  WDP     27-Jan-2000      put indexing actions in a logfile and
'                                Leave actual indexing to BOIndexer
'   d5  WDP     15-Feb-2000      Give BLK file a identitfier based on the time
'                                to avoid name clashes
'   d6  TomH    13-Aug-2002      Added optional argument

On Error GoTo Errorhandler

    Dim sSource As String
    sSource = "BOLeafletIndexer.IndexLeaflet"
    
    Set m_oboLeaflet = New BOLeafletLoader
    'Load leaflet data from DB
    m_oboLeaflet.LoadData t_lLeafletID
    
    Dim sFilename As String, sBLKFilename As String, sLeafletFile As String
    Dim lBLKFileNum As Long, lLeafFileNum As Long
    Dim sBLK As String
    Dim sCollectionDrive As String
    Dim sIndexFilesDrive As String
    Dim sLeafletDir As String
    Dim sLeafletDirForVerity As String
    Dim ofwAppData As New FWApplicationData
    
    sCollectionDrive = ofwAppData.GetApplicationConfigurationInfomation(pRootForCollections)
    sIndexFilesDrive = ofwAppData.GetApplicationConfigurationInfomation(pRootForIndexingFiles)
    sLeafletDir = ofwAppData.GetApplicationConfigurationInfomation(pLeafletDestinationFileDirectory)
    sLeafletDirForVerity = Replace(sLeafletDir, "\", "/")
    
    sLeafletFile = sLeafletDir & m_oboLeaflet.Data.sFilename
    sFilename = Replace(Replace(m_oboLeaflet.Data.sFilename, " ", ""), vbTab, "") & "_" & GetIndexActionString(t_eAction)
    Dim sFileIdentifier As String
    sFileIdentifier = CStr(Hour(Now())) & CStr(Minute(Now())) & CStr(Second(Now()))
    sBLKFilename = sIndexFilesDrive & mk_sFilepath & sFileIdentifier & sFilename & ".blk"
    
    'Open the files, if they already exist they will be overwritten
    lBLKFileNum = FreeFile
    Open (sBLKFilename) For Output As lBLKFileNum
    
    If lBLKFileNum > 0 Then
        lLeafFileNum = FreeFile
        'We need to read in from the leaflet HTML file as we need
        'to write the file length to the bulk file.
        Open sLeafletFile For Input Access Read As lLeafFileNum
    
        If lLeafFileNum > 0 Then
            Dim sLeafHTML As String
            'Read into a string.
            sLeafHTML = StrConv(InputB(LOF(lLeafFileNum), lLeafFileNum), vbUnicode)
            Close #lLeafFileNum
                
            Dim lDocSize As Long
            lDocSize = Len(sLeafHTML)
                
            'Field data for the bulk file
            sBLK = gk_sVFIELDVerityKey & ": " & m_oboLeaflet.Data.lID & vbCrLf & _
                    "DOC_FN: " & DoubleQuotes(sLeafletDirForVerity & m_oboLeaflet.Data.sFilename) & vbCrLf & _
                    "DOC_OF: 0" & vbCrLf & _
                    "DOC_SZ: " & lDocSize & vbCrLf & _
                    gk_sVFIELDLeafletTitle & ": " & DoubleQuotes( _
                        VariantStringAsString(m_oboLeaflet.Data.sTitle)) & vbCrLf & _
                    gk_sVFIELDLeafletOrder & ": " & m_oboLeaflet.Data.lOrderNumber & vbCrLf & _
                    gk_sVFIELDLeafletID & ": " & m_oboLeaflet.Data.lID & vbCrLf & _
                    "<<EOD>>" & vbCrLf & vbCrLf
            Print #lBLKFileNum, sBLK
            
            Close #lBLKFileNum
            
            Dim sIndexingLogFile As String
            
            'if the optional argument has a value then make that the action log file name
            If t_sActionLogFileName = "" Then
                sIndexingLogFile = sIndexFilesDrive & mk_sFilepath & sFilename & ".log"
            Else
                sIndexingLogFile = sIndexFilesDrive & mk_sFilepath & t_sActionLogFileName & ".log"
            End If
            
            Dim sCollection As String
            sCollection = sCollectionDrive & gk_sLeafletCollectionPath
            
            'Log the indexing actions to perform in a log file
            LogAction sIndexingLogFile, sBLKFilename, sCollection, t_eAction
                    
            'return the indexing action logfile name so we can trigger the
            'indexing process on the Verity server
            IndexLeaflet = sIndexingLogFile
             
        Else
            '+Give this a new error type
            RaiseError sSource, PEGeneralInternalError, "Could not open leaflet source file"
        End If
    Else
        '+Give this a new error type
        RaiseError sSource, PEGeneralInternalError, "Could not open bulk indexing file"
        
    End If
    
    Exit Function
Errorhandler:
    ConvertError sSource
        
End Function

Public Sub IndexBulkLeaflet()
'   Recurses through the leaflets db table indexing all the leaflets
'
'   d6  TomH      13-Aug-2002      Created
'
On Error GoTo Errorhandler

    Dim sSource As String
    sSource = "BOLeafletIndexer.IndexBulkLeaflet"
    
    Dim lLeafFileNum As Long
    Dim sIndexingLogFile As String
    Dim lCount As Long
    Dim lLeafletID As Long
    Dim sLeafletFile As String
    Dim sLeafletDir As String
    Dim lFileCount As Long
    Dim ofwAppData As FWApplicationData
    Set ofwAppData = New FWApplicationData
    Dim objfs As FileSystemObject
    Set objfs = New FileSystemObject
    Dim oboVerIndexer
            
    sLeafletDir = ofwAppData.GetApplicationConfigurationInfomation(pLeafletDestinationFileDirectory)
    
    'Load leaflet data from DB
    Dim sSql As String
    Dim ofwDBHelper As FWDBHelper
    Set ofwDBHelper = New FWDBHelper
    Dim orsResults As FWRecordList
        
    sSql = "SELECT " & gk_sDBCOLLeafletReferenceID & ", " & _
            gk_sDBCOLLeafletFilename & _
            " FROM " & gk_sDBTABLELeaflet
            
    Set orsResults = ofwDBHelper.FindBySql(sSql)
    
    LogMessage "Getting leaflet file names and IDs from database.", pProfileInformation
        
    orsResults.MoveFirst
    
    lFileCount = orsResults.RecordCount
            
    ' recurse through results and index files
    For lCount = 1 To lFileCount
    
        ' get leaflet info from results
        lLeafletID = orsResults.GetFieldValue(gk_sDBCOLLeafletReferenceID)
        sLeafletFile = sLeafletDir & orsResults.GetFieldValue(gk_sDBCOLLeafletFilename)
        
        If objfs.FileExists(sLeafletFile) Then
                                
            lLeafFileNum = FreeFile
            Open sLeafletFile For Input Access Read As lLeafFileNum
            LogMessage "Opened file " & sLeafletFile, pProfileInformation
                        
            ' check file exists in directory, if so index
            If lLeafFileNum > 0 Then
                sIndexingLogFile = IndexLeaflet(lLeafletID, IDXInsert, mk_sBulkActionLogFileName)
                LogMessage "Indexed file " & sLeafletFile, pProfileInformation
            
            Else
                LogMessage sLeafletFile & "was empty", pProfileInformation
            End If
                
        Else
            LogMessage sLeafletFile & " was not present in directory", pProfileInformation
                            
        End If
        orsResults.MoveNext
        
    Next
    
    'The Verity server could be remote and so we make a (D)COM
    'call to the indexer to actually index the leaflet into
    'the collection
    Set oboVerIndexer = CreateObject("PROCATEdVerIdx.BOIndexer")
    oboVerIndexer.IndexFromIndexLog sIndexingLogFile, False, True
    Set oboVerIndexer = Nothing
        
    Exit Sub

Errorhandler:
    ConvertError sSource
End Sub

Private Function IndexTermBulkData(t_eTermType As EAFTermType)
'   Add the leaflet's index terms to the bulk file
'   NOTE: this is no longer done. Searches over
'   leaflet's index terms are done on the database
'
'   d2  WDP     26-Oct-1999      Created
On Error GoTo Errorhandler

    Dim sSource As String
    sSource = "BOLeafletIndexer.IndexTermBulkData"
    
    Dim sRes As String
    Dim sSql As String
    Dim sVField As String
    Dim sLinkTable As String
    Dim sTermIDColumn As String
    Dim ofwResults As New FWRecordList
    Dim ofwDBHelper As New FWDBHelper
    
    Select Case t_eTermType
    Case AFCorporateBody
        sVField = gk_sVFIELDCorporateName
        sLinkTable = gk_sDBTABLELeafletCorpBody
        sTermIDColumn = gk_sDBCOLCorporateBodyReferenceID
    Case AFPerson
        sVField = gk_sVFIELDPerson
        sLinkTable = gk_sDBTABLELeafletPerson
        sTermIDColumn = gk_sDBCOLPersonReferenceID
    Case AFPlace
        sVField = gk_sVFIELDPlace
        sLinkTable = gk_sDBTABLELeafletPlace
        sTermIDColumn = gk_sDBCOLPlaceReferenceID
    Case AFSubject
        sVField = gk_sVFIELDSubject
        sLinkTable = gk_sDBTABLELeafletSubject
        sTermIDColumn = gk_sDBCOLSubjectReferenceID
    Case Else
        RaiseError sSource, PEGeneralInternalError, "Invalid AF term type enum."
        
    End Select
    sSql = " SELECT " & sTermIDColumn & " AS " & gk_sDILinkDataID & _
            " FROM " & sLinkTable & " WHERE " & _
            gk_sDBCOLLeafletReferenceID & "=" & _
            m_oboLeaflet.Data.lID
    
    Set ofwResults = ofwDBHelper.FindBySql(sSql)
    
    sRes = sVField & ": """
    
    While Not ofwResults.IsAtEnd
        sRes = sRes & gk_sLinkDataDelimiter & _
            ofwResults.GetFieldValue(gk_sDILinkDataID) & _
            gk_sLinkDataDelimiter
        ofwResults.MoveNext
    Wend
    
    IndexTermBulkData = sRes & """" & vbCrLf
    
    Exit Function
Errorhandler:
    ConvertError sSource
        
End Function

