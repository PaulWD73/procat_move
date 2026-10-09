VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BOCatDataExtractor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    BOCatDataExtractor.cls
' System:       PRO, PROCAT
' Copyright:    (C) Quidnunc Limited
'
' Description:      This along with BOCatalogueIndexer performs the
'                   Verity indexing functionality. This class is responsible
'                   for extracting data from the DB into data files
'                   BOCatalogueIndexer actually indexes these files into
'                   the Verity collections
'
' Amendment history:
'   d1 to 8 WDP     27-Jan-2000     Created by splitting functionality in BOCatalogueIndexer
'   d9      WDP     14-Feb-2000     Select non-modified entries that have been accessioned or created
'                               when generating new editset data
'   d10     WDP     2-Mar-2000      Convert DB NULLs into defined values for numeric fields
'                                   in WriteToFile (Verity does not have NULLs)
'
'   d11     WDP     6-Mar-2000      Added progress mapping functions so indexing
'                                   can recover and restart
'   d12     WDP     14-Mar-2000     Altered SQL generation to only select live records
'                                   when selecting from LIVE (live_flag = 0 or 1)
'                                   Added logging of each entry written to file for checking
'                                   NOTE! this should be removed after testing
'   d13     WDP     22-Mar-2000     Cope with "" as a valid closure type
'   d14     TMJ     30-Mar-2000     Added calls to SpecialCharacterDecode
'   d15     DAJ     18-Apr-2000     Removed temporary table create/drop code
'   d16     DAJ     11-May-2000     Added MaxRecords overide to GetAdvancedData (limit of 100
'                                   causes an EOF error during release as there will be a mismatch
'                                   between the number of 'basic' records and 'advanced' ones)
'   d17     DAJ     06-Feb-2001     Added extra stripping of zones to prevent multi-line values that cause parsing errors
'   d18     ALM     29-Oct-2001     Added a new editset entry type - 'delete' and supports the
'                                   creation of index files to remove entries from the verity
'                                   collections
'   d19     TomH    23-Aug-2002     Have added ability to index subset of catalogue
'                                   in CreateSomeFromLetterCodes
'   d20     TomH    23-Aug-2002     Altered so basic files not created for indexing whole catalogue
'                                   Also it now added <basic> in SGML files created.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Public Enum EIndexingStep
    IXSTEPStart
    IXSTEPDataFileCreationCompleted
    IXSTEPMessage
    IXSTEPSmallLettercodes
    IXSTEPMediumLettercode
    IXSTEPLargeLettercode
    IXSTEPClassUpData
    IXSTEPWholeLevel
    IXSTEPError
    IXSTEPIndexingComplete
End Enum



' Class constants

'Where to close sentence from for truncated title/scope
Private Const mk_lTitleTruncationMark As Long = 255

'The maximum limit to truncate to
Private Const mk_lTitleTruncationLimit As Long = 1000

' max records to select at one time
Private Const mk_lMaxRecords As Long = 1000000

' Threshhold for splitting indexing of an edit set into
' sub processes
Private Const mk_lEditSetMedSize As Long = 10000

'Filepath to blk and sgml files
Private Const mk_sFilepath As String = "VerityWrapper\PROCAT\ReleaseFiles\"
Private Const mk_sFilepathForVerity As String = "VerityWrapper/PROCAT/ReleaseFiles/"

'Link term property DB Ids
Private Const mk_lCreatorPropID As Long = ECATALOGUEPROPERTIES.cpCreatorNames
Private Const mk_lPlaceOfDepPropID As Long = ECATALOGUEPROPERTIES.cpPlaceOfDeposit
Private Const mk_lLocOfOrigPropID As Long = ECATALOGUEPROPERTIES.cpLocationOfOriginals
Private Const mk_lPlacePropID As Long = ECATALOGUEPROPERTIES.cpIndexTermPlaceName
Private Const mk_lPersonPropID As Long = ECATALOGUEPROPERTIES.cpIndexTermPersonName
Private Const mk_lSubjectPropID As Long = ECATALOGUEPROPERTIES.cpIndexTermSubject
Private Const mk_lCorpBodyPropID As Long = ECATALOGUEPROPERTIES.cpIndexTermCorporateName
Private Const mk_lSourceOfAcqPropID As Long = ECATALOGUEPROPERTIES.cpImmediateSourceOfAcquisition
Private Const mk_lCopiesPropID As Long = ECATALOGUEPROPERTIES.cpCopiesInformation

'Const used to select pieces with null header refs when
'selecting from live
Private Const mk_lNullRef As Long = -1

'Threshold above which we must break lettercodes up
Private Const mk_lMaxLettercodeSize As Long = 100000
Private Const mk_lMedLettercodeSize As Long = 25000
'Threshold above which we must break classes up
Private Const mk_lMaxClassSize As Long = 40000
'Threshold above which we must break classes up
Private Const mk_lSmallClassSize As Long = 500


'Constant used to represent NULLs for closure type ("" is a valid value in ILDB)
Private Const mk_sNeverUsedClosureType As String = "-1"


' Instance variables

' Record set to hold edit set entry main data to index
Private m_ofwMainData As FWRecordList
' Record sets to hold the long text fields for zone indexing
Private m_ofwTitleData As FWRecordList
Private m_ofwScopeData As FWRecordList
Private m_ofwAdminData As FWRecordList
Private m_ofwArrangeData As FWRecordList
Private m_ofwCustHistData As FWRecordList
Private m_ofwAppInfoData As FWRecordList
Private m_ofwAccrualData As FWRecordList
Private m_ofwNoteData As FWRecordList
Private m_ofwLinkData As FWRecordList

'RecordSet to hold all the lettercodes for collection mapping
Private m_ofwEditSetLettercodes As FWRecordList

'holds sizes of lettercodes for live indexing
Private m_ofwLcodeSize As FWRecordList
'index at which we are looking at the above record set
'we need to track this so we can jump around
Private m_lLcodeSizeRecordsetIndex

' ID of the edit set
Private m_lEditsetID As Long

Public Enum EEditSetEntryType
    EDITNew
    EDITUpdated
    EDITDelete 'd18 Added
End Enum

'The fields which have larger text fields thus separate record sets
Public Enum ELargeField
    LARGETitle
    LARGEScopeContent
    LARGEAdminHistory
    LARGEArrangement
    LARGECustodialHistory
    LARGEAppraisalInfo
    LARGEAccruals
    LARGENote
End Enum

'
' Public Methods
'

Public Function CreateEditSetIndexData(t_lEditSetID As Long, t_sDate As String, _
                                       Optional t_bDelete As Boolean = False) As String
'   This method is the master indexing function that triggers
'   selecting and writing the indexing data to files.
'   It returns a timestamp that is used to identify the
'   index action log file when the indexing is performed later.
'
'   d1  WDP     5-Oct-1999      Created
'   d2  WDP     18-Nov-1999     Changed name from IndexEditset and added date parameter
'   d18 ALM     25-Oct-2001     Handles entries that are to be deleted from the collections
'   d20 TomH    23-Aug-2002     Removed all basic file generation as no longer required
On Error GoTo ErrorHandler

    Dim sSource As String
    Dim sTime As String
    
    sSource = "BOCatalogueIndexer.CreateEditSetIndexData"
    
    m_lEditsetID = t_lEditSetID
    sTime = GetTimeString
        
    If t_bDelete = False Then
    
        'Index the new and updated record data over all collections
        LogReleaseMessage "Getting data and creating indexing files for new entries...", RELInfo
        GetBasicData EDITNew
        ' d20 tomh - removed
        'CreateDataForAllEntries EDITNew, COLBasic, t_sDate, sTime
        GetAdvancedData EDITNew
        CreateDataForAllEntries EDITNew, COLAdvanced, t_sDate, sTime
        CreateDataForAllEntries EDITNew, COLClassUp, t_sDate, sTime
        
        LogReleaseMessage "Getting data and creating indexing files for updated entries...", RELInfo
        GetBasicData EDITUpdated
        ' d20 tomh - removed
        'CreateDataForAllEntries EDITUpdated, COLBasic, t_sDate, sTime
        GetAdvancedData EDITUpdated
        CreateDataForAllEntries EDITUpdated, COLAdvanced, t_sDate, sTime
        CreateDataForAllEntries EDITUpdated, COLClassUp, t_sDate, sTime
        
    Else
    
        'Index the deletable record data over all collections
        LogReleaseMessage "Getting data and creating indexing files for deleted entries...", RELInfo
        GetBasicData EDITDelete
        ' d20 tomh - removed
        ' CreateDataForAllEntries EDITDelete, COLBasic, t_sDate, sTime
        GetAdvancedData EDITDelete
        CreateDataForAllEntries EDITDelete, COLAdvanced, t_sDate, sTime
        CreateDataForAllEntries EDITDelete, COLClassUp, t_sDate, sTime
        
    End If
        
    CreateEditSetIndexData = sTime
    LogReleaseMessage "Editset " & m_lEditsetID & ": All indexing datafiles complete...", RELInfo
        
    Exit Function
ErrorHandler:
    ConvertError sSource
    
End Function


Public Function CreateAllClassUpData(t_sDate As String, t_sTime As String, Optional ByVal t_sLetterCodesIn As String = "")
'   This method creates data for classes, divisions and lettercodes
'   to index in the class up collection.
'
'   d8  WDP     13-Dec-1999     Created
'   d19 TomH    16-Jul-2002     Added optional parameter

On Error GoTo ErrorHandler

    Dim sSource As String
    
    sSource = "BOCatalogueIndexer.CreateAllClassUpData"
    
    Dim ofwAppData As New FWApplicationData
    Dim sFileDrive As String
    Dim sCollectionDrive As String
    
    Dim eLevel As ECATALOGUELEVEL
    eLevel = ctLetterCode
    Dim eAction As EIndexingAction
    eAction = IDXInsert
    Dim oboVerConfig As New BOVerityConfig
    Dim sCollection As String
    sCollection = sCollectionDrive & oboVerConfig.GetLettercodeCollection("E", COLClassUp)
    Dim sLogFile As String
    sLogFile = LogFileName(t_sDate & t_sTime)
    
    sFileDrive = ofwAppData.GetApplicationConfigurationInfomation(pRootForIndexingFiles)
    sCollectionDrive = ofwAppData.GetApplicationConfigurationInfomation(pRootForCollections)
    
    'we get data for lettercode then division then class
    While eLevel < ctHeaders
        If Not IndexStepExists(IXSTEPWholeLevel, COLClassUp, eLevel) Then
            Dim sFilenamePrefix As String
            Dim sBLKFilepath As String
            
            ' Get data from DB
            GetBasicDataFromLive eLevel, t_sLettercodeIn:=t_sLetterCodesIn
            GetAdvancedDataFromLive eLevel, t_sLettercodeIn:=t_sLetterCodesIn
                
            'Get the collection path from verityconfig
            '+ use E as dummy lettercode to get collection for now
            sFilenamePrefix = GetFileName(t_sDate, True, COLClassUp, t_eAction:=eAction, t_eLevel:=eLevel)
            sBLKFilepath = sFileDrive & mk_sFilepath & sFilenamePrefix & ".blk"
            
            If WriteToFile(sFilenamePrefix, t_eCollectionType:=COLClassUp) Then
                'The indexing actions required are logged to a file
                'which can later be read to perform the indexing itself
                'from the bulk and sgml files generated
                LogAction sLogFile, sBLKFilepath, sCollection, eAction
                LogIndexStep IXSTEPWholeLevel, "Created files for cat level: " & eLevel & " for class up collection.", COLClassUp, eLevel
                DumpLogToMasterLog LogFileName(t_sDate & t_sTime)
            End If
        End If
        eLevel = eLevel + 1
        
    Wend
    Exit Function
ErrorHandler:
    ConvertError sSource
    
End Function

Public Function CreateSubheaderUpData(t_sDate As String, t_sTime As String, t_eCollectionType As ECollectionType, Optional ByVal t_sLetterCodes As String)
'   This method creates data for collections of the given tpe
'   for subheaders upwards.
'
'   d8  WDP     13-Dec-1999     Created
'   d19 Tomh    16-Jul-2002     Added optional var to see if want just some lettercodes

On Error GoTo ErrorHandler

    Dim sSource As String
    
    sSource = "BOCatalogueIndexer.CreateSubheaderUpData"
    
    Dim eAction As EIndexingAction
    eAction = IDXInsert
    Dim oboVerConfig As New BOVerityConfig
    Dim ofwAppData As New FWApplicationData
    Dim sFileDrive As String
    Dim sCollectionDrive As String
    sFileDrive = ofwAppData.GetApplicationConfigurationInfomation(pRootForIndexingFiles)
    sCollectionDrive = ofwAppData.GetApplicationConfigurationInfomation(pRootForCollections)
    
    Dim colCollectionClusters As New Collection
    Set colCollectionClusters = oboVerConfig.CollectionLettercodes(t_eCollectionType)
    
    Dim sLogFile As String
    sLogFile = LogFileName(t_sDate & t_sTime)
    
    'we get data for lettercode then division then class
    'then header then subheader
    Dim lCount As Long
    For lCount = 1 To colCollectionClusters.Count
        Dim eLevel As ECATALOGUELEVEL
        eLevel = ctLetterCode
    
        While eLevel < ctPiece
            
            If Not IndexStepExists(IXSTEPWholeLevel, t_eCollectionType, eLevel, lCount) Then
                
                Dim sFilenamePrefix As String
                Dim sBLKFilepath As String
                Dim sLettercodeIn As String
                Dim lLettercodeID As Long
                Dim sCollection As String
                Dim sData As String
                
                sData = colCollectionClusters.Item(lCount)
                sCollection = Left(sData, InStr(sData, gk_sClusterPathDelimiter) - 1)
                
                ' d19 added test to see if have been passed lettercodes
                If Not t_sLetterCodes = "" Then
                
                    ' set lettercodes to what we have been passed in
                    sLettercodeIn = t_sLetterCodes

                ElseIf colCollectionClusters.Count = 1 Then
                    'If there is only 1 collection we need not
                    'build the IN clause
                    sLettercodeIn = ""
                    lLettercodeID = 0
                Else
                    sLettercodeIn = Mid(sData, InStr(sData, gk_sClusterPathDelimiter) + Len(gk_sClusterPathDelimiter))
                    lLettercodeID = 0
                End If
                
                ' Get data from DB
                GetBasicDataFromLive eLevel, t_lLettercodeID:=lLettercodeID, t_sLettercodeIn:=sLettercodeIn
                If Not t_eCollectionType = COLBasic Then
                    GetAdvancedDataFromLive eLevel, t_lLettercodeID:=lLettercodeID, t_sLettercodeIn:=sLettercodeIn
                End If
                
                'Get the collection path from verityconfig
                '+ use E as dummy lettercode to get collection for now
                sFilenamePrefix = GetFileName(t_sDate & "part" & lCount, True, t_eCollectionType, t_eAction:=eAction, t_eLevel:=eLevel)
                sBLKFilepath = sFileDrive & mk_sFilepath & sFilenamePrefix & ".blk"
                
                If WriteToFile(sFilenamePrefix, t_eCollectionType:=t_eCollectionType) Then
                    'The indexing actions required are logged to a file
                    'which can later be read to perform the indexing itself
                    'from the bulk and sgml files generated
                    LogAction sLogFile, sBLKFilepath, sCollectionDrive & sCollection, eAction
                    Dim sColl As String
                    
                    If t_eCollectionType = COLBasic Then
                        sColl = " basic collections"
                    Else
                        sColl = " advanced collections"
                    End If
                    
                    
                    LogIndexStep IXSTEPWholeLevel, "Created files for cat level: " & eLevel & " for " & sColl & " part " & lCount, t_eCollectionType, eLevel, lCount
                    DumpLogToMasterLog LogFileName(t_sDate & t_sTime)
                End If
                
            End If
            eLevel = eLevel + 1
            
        Wend
    Next
    
    Exit Function
ErrorHandler:
    ConvertError sSource
    
End Function


Private Sub CreateDataForAllEntries( _
    t_eEntryType As EEditSetEntryType, _
    t_eCollectionType As ECollectionType, _
    t_sDate As String, t_sTime As String)
'   Process data in record sets to create indexing files
'
'   d1  WDP     5-Oct-1999      Created
'   d5  WDP     18-Nov-1999     Changed name to CreateDataForAllEntries
'                               from IndexAllEntries and removed the
'                                call to RunMKVDK
'   d18 ALM     25-Oct-2001     Supports a new type of entry - delete
'
On Error GoTo ErrorHandler

    Dim sSource As String
    sSource = "BOCatalogueIndexer.CreateDataForAllEntries"
    Dim lSize As Long
    Dim ofwAppData As New FWApplicationData
    Dim sIndexFileDrive As String
    Dim bHaveWrittenToActionLog As Boolean
    bHaveWrittenToActionLog = False
    
    sIndexFileDrive = ofwAppData.GetApplicationConfigurationInfomation(pRootForIndexingFiles)
    Set ofwAppData = Nothing
    
    If Not m_ofwEditSetLettercodes.RecordCount = 0 Then
        m_ofwEditSetLettercodes.MoveFirst
    End If
    
    'We create a separate SGML and bulk file for each lettercode
    'in the edit set, this enables us to index appropriate
    'data over the correct collections
    
    Dim eAction As EIndexingAction
    
    ' d18 - Rewrote using case
    Select Case t_eEntryType
    
        Case EDITNew:
            eAction = IDXInsert
            
        Case EDITUpdated:
            eAction = IDXUpdate
            
        Case EDITDelete:
            eAction = IDXDelete
        
    End Select
            
    Dim oboVerConfig As New BOVerityConfig
    
    Dim sLogFile As String
    sLogFile = LogFileName(t_sDate & t_sTime, m_lEditsetID)
                                 
    While Not m_ofwEditSetLettercodes.IsAtEnd
        
        Dim sLcode As String
        sLcode = VariantStringAsString( _
            m_ofwEditSetLettercodes.GetFieldValue(gk_sDILetterCodeRef))
        
        Dim sFilenamePrefix As String
        sFilenamePrefix = GetFileName(t_sDate, False, t_eCollectionType, sLcode, eAction)
        
        'If there was data to write then we get the collection path
        'that is appropriate for the lettercode and log the action
        'required
        If WriteToFile(sFilenamePrefix, sLcode, t_eCollectionType) Then
            Dim sCollection As String, sBLKFilepath As String
        
            'Get the collection path from verityconfig
            sCollection = oboVerConfig.GetLettercodeCollection(sLcode, t_eCollectionType)
            sBLKFilepath = sIndexFileDrive & mk_sFilepath & _
                GetFileName(t_sDate, False, t_eCollectionType, sLcode, eAction) & ".blk"
                
            'The indexing actions required are logged to a file
            'which can later be read to perform the indexing itself
            'from the bulk and sgml files generated
            LogAction sLogFile, sBLKFilepath, sCollection, eAction, sLcode, m_lEditsetID
            bHaveWrittenToActionLog = True
        End If
        
        m_ofwEditSetLettercodes.MoveNext
    Wend
    
    If bHaveWrittenToActionLog = False Then
        'Write a blank action (to create the file) if we
        'have no actions to perform
        LogAction sLogFile, t_bDummyAction:=True
    End If
    
    
    Exit Sub
    
ErrorHandler:
    ConvertError sSource
    
End Sub

Private Sub GetBasicData(t_eEntryType As EEditSetEntryType)
'   Reads basic indexing data into the record sets
'   The entry type parameter determines whether we select new
'   or updated records
'
'   d5  WDP     18-Nov-1999     Moved from CreateDataForAllEntries
On Error GoTo ErrorHandler

    Dim sSource As String
    sSource = "BOCatalogueIndexer.GetBasicData"
    
    Dim ofwDBHelper As New FWDBHelper
    
    Dim sMainSql As String
    
    'We create queries to extract data from all catalogue levels
    'in the edit set and union the results together so we
    'can index them all at once in bulk
    
    sMainSql = GetLettercodeSql(t_eEntryType) & " UNION " & _
                GetDivisionSql(t_eEntryType) & " UNION " & _
                GetClassSql(CATPlain, t_eEntryType) & " UNION " & _
                GetClassSql(CATWithDiv, t_eEntryType) & " UNION " & _
                GetHeaderSql(CATPlain, t_eEntryType) & " UNION " & _
                GetHeaderSql(CATWithDiv, t_eEntryType) & " UNION " & _
                GetSubheaderSql(CATPlain, t_eEntryType) & " UNION " & _
                GetSubheaderSql(CATWithDiv, t_eEntryType) & " UNION " & _
                GetPieceSql(CATPlain, t_eEntryType) & " UNION " & _
                GetPieceSql(CATWithDiv, t_eEntryType) & " UNION " & _
                GetPieceSql(CATWithHeader, t_eEntryType) & " UNION " & _
                GetPieceSql(CATWithSubheader, t_eEntryType) & " UNION " & _
                GetPieceSql(CATWithHeaderAndDiv, t_eEntryType) & " UNION " & _
                GetPieceSql(CATWithSubheaderAndDiv, t_eEntryType) & " UNION " & _
                GetItemSql(CATPlain, t_eEntryType) & " UNION " & _
                GetItemSql(CATWithDiv, t_eEntryType) & " UNION " & _
                GetItemSql(CATWithHeader, t_eEntryType) & " UNION " & _
                GetItemSql(CATWithSubheader, t_eEntryType) & " UNION " & _
                GetItemSql(CATWithHeaderAndDiv, t_eEntryType) & " UNION " & _
                GetItemSql(CATWithSubheaderAndDiv, t_eEntryType) & " ORDER BY " & gk_sDICatalogueLevel & ", " & gk_sDIID
            
    'We get the large text fields containing data for
    'verity zones in separate record sets as SQLServer cannot
    'cope with processing wide result sets
    
    'Main entry details
    Set m_ofwMainData = ofwDBHelper.FindBySql(sMainSql, mk_lMaxRecords)
    'Scope-content
    Set m_ofwScopeData = ofwDBHelper.FindBySql(GetLargeFieldSql(LARGEScopeContent, t_eEntryType), mk_lMaxRecords)
    'Title
    Set m_ofwTitleData = ofwDBHelper.FindBySql(GetLargeFieldSql(LARGETitle, t_eEntryType), mk_lMaxRecords)
    'Admin history
    Set m_ofwAdminData = ofwDBHelper.FindBySql(GetLargeFieldSql(LARGEAdminHistory, t_eEntryType), mk_lMaxRecords)
    
    'The set of all lettercodes present in the edit set
    Set m_ofwEditSetLettercodes = ofwDBHelper.FindBySql(GetAllLettercodesSQL(), mk_lMaxRecords)
    
    Exit Sub
    
ErrorHandler:
    ConvertError sSource
    
End Sub

Private Sub GetBasicDataFromLive( _
    t_eLevel As ECATALOGUELEVEL, _
    Optional t_lLettercodeID As Long = 0, _
    Optional t_lClassID As Long, _
    Optional t_lHeaderID As Long, _
    Optional t_sLettercodeIn As String = "", _
    Optional t_sClassIn As String = "", _
    Optional t_bGetExtraExtnData As Boolean = True _
    )
'   Reads basic indexing data into the record sets
'   from live tables depending on the parameters supplied
'
'   d5  WDP     19-Nov-1999     Moved from IndexLIVEEntries
'   d7  WDP     30-Nov-1999     Altered to use ID cols
'   d8  WDP     13-Dec-1999     Made lettercode id optional
'                               If omitted all entries are selected
'                               Added LettercodeIn parameter to allow
'                               selection from multiple lettercodes
'
On Error GoTo ErrorHandler

    Dim sSource As String
    sSource = "BOCatalogueIndexer.GetBasicDataFromLive"
    
    Dim ofwDBHelper As New FWDBHelper
    
    Dim sMainSql As String
    
    'We create queries to extract data from all catalogue levels
    'in the edit set and union the results together so we
    'can index them all at once in bulk
    
    Select Case t_eLevel
        Case ctLetterCode:
            sMainSql = GetLettercodeSql(t_bFromLive:=True, t_lLettercodeID:=t_lLettercodeID, t_sLettercodeIn:=t_sLettercodeIn, t_bGetExtraExtnData:=t_bGetExtraExtnData)
        Case ctDivision:
            sMainSql = GetDivisionSql(t_bFromLive:=True, t_lLettercodeID:=t_lLettercodeID, t_sLettercodeIn:=t_sLettercodeIn, t_bGetExtraExtnData:=t_bGetExtraExtnData)
        Case ctClass:
            sMainSql = GetClassSql(CATPlain, t_bFromLive:=True, t_lLettercodeID:=t_lLettercodeID, t_sLettercodeIn:=t_sLettercodeIn, t_bGetExtraExtnData:=t_bGetExtraExtnData) & _
            " UNION " & GetClassSql(CATWithDiv, t_bFromLive:=True, t_lLettercodeID:=t_lLettercodeID, t_sLettercodeIn:=t_sLettercodeIn, t_bGetExtraExtnData:=t_bGetExtraExtnData)
        Case ctHeaders:
            sMainSql = GetHeaderSql(CATPlain, t_bFromLive:=True, t_lLettercodeID:=t_lLettercodeID, t_lClassID:=t_lClassID, t_sLettercodeIn:=t_sLettercodeIn, t_bGetExtraExtnData:=t_bGetExtraExtnData) & _
            " UNION " & GetHeaderSql(CATWithDiv, t_bFromLive:=True, t_lLettercodeID:=t_lLettercodeID, t_lClassID:=t_lClassID, t_sLettercodeIn:=t_sLettercodeIn, t_bGetExtraExtnData:=t_bGetExtraExtnData)
        Case ctSubHeaders:
            sMainSql = GetSubheaderSql(CATPlain, t_bFromLive:=True, t_lLettercodeID:=t_lLettercodeID, t_lClassID:=t_lClassID, t_sLettercodeIn:=t_sLettercodeIn, t_bGetExtraExtnData:=t_bGetExtraExtnData) & _
            " UNION " & GetSubheaderSql(CATWithDiv, t_bFromLive:=True, t_lLettercodeID:=t_lLettercodeID, t_lClassID:=t_lClassID, t_sLettercodeIn:=t_sLettercodeIn, t_bGetExtraExtnData:=t_bGetExtraExtnData)
        Case ctPiece:
            sMainSql = GetPieceSql(CATPlain, t_bFromLive:=True, _
                t_lLettercodeID:=t_lLettercodeID, t_lClassID:=t_lClassID, t_lHeaderID:=t_lHeaderID, t_sClassIn:=t_sClassIn, t_sLettercodeIn:=t_sLettercodeIn, t_bGetExtraExtnData:=t_bGetExtraExtnData) & _
            " UNION " & GetPieceSql(CATWithDiv, t_bFromLive:=True, _
                t_lLettercodeID:=t_lLettercodeID, t_lClassID:=t_lClassID, t_lHeaderID:=t_lHeaderID, t_sClassIn:=t_sClassIn, t_sLettercodeIn:=t_sLettercodeIn, t_bGetExtraExtnData:=t_bGetExtraExtnData) & _
            " UNION " & GetPieceSql(CATWithHeader, t_bFromLive:=True, _
                t_lLettercodeID:=t_lLettercodeID, t_lClassID:=t_lClassID, t_lHeaderID:=t_lHeaderID, t_sClassIn:=t_sClassIn, t_sLettercodeIn:=t_sLettercodeIn, t_bGetExtraExtnData:=t_bGetExtraExtnData) & _
            " UNION " & GetPieceSql(CATWithSubheader, t_bFromLive:=True, _
                t_lLettercodeID:=t_lLettercodeID, t_lClassID:=t_lClassID, t_lHeaderID:=t_lHeaderID, t_sClassIn:=t_sClassIn, t_sLettercodeIn:=t_sLettercodeIn, t_bGetExtraExtnData:=t_bGetExtraExtnData) & _
            " UNION " & GetPieceSql(CATWithHeaderAndDiv, t_bFromLive:=True, _
                t_lLettercodeID:=t_lLettercodeID, t_lClassID:=t_lClassID, t_lHeaderID:=t_lHeaderID, t_sClassIn:=t_sClassIn, t_sLettercodeIn:=t_sLettercodeIn, t_bGetExtraExtnData:=t_bGetExtraExtnData) & _
            " UNION " & GetPieceSql(CATWithSubheaderAndDiv, t_bFromLive:=True, _
                t_lLettercodeID:=t_lLettercodeID, t_lClassID:=t_lClassID, t_lHeaderID:=t_lHeaderID, t_sClassIn:=t_sClassIn, t_sLettercodeIn:=t_sLettercodeIn, t_bGetExtraExtnData:=t_bGetExtraExtnData)
            
        Case ctItem:
            sMainSql = GetItemSql(CATPlain, t_bFromLive:=True, _
                t_lLettercodeID:=t_lLettercodeID, t_lClassID:=t_lClassID, t_lHeaderID:=t_lHeaderID, t_sClassIn:=t_sClassIn, t_sLettercodeIn:=t_sLettercodeIn, t_bGetExtraExtnData:=t_bGetExtraExtnData) & _
            " UNION " & GetItemSql(CATWithSubheader, t_bFromLive:=True, _
                t_lLettercodeID:=t_lLettercodeID, t_lClassID:=t_lClassID, t_lHeaderID:=t_lHeaderID, t_sClassIn:=t_sClassIn, t_sLettercodeIn:=t_sLettercodeIn, t_bGetExtraExtnData:=t_bGetExtraExtnData) & _
            " UNION " & GetItemSql(CATWithDiv, t_bFromLive:=True, _
                t_lLettercodeID:=t_lLettercodeID, t_lClassID:=t_lClassID, t_lHeaderID:=t_lHeaderID, t_sClassIn:=t_sClassIn, t_sLettercodeIn:=t_sLettercodeIn, t_bGetExtraExtnData:=t_bGetExtraExtnData) & _
            " UNION " & GetItemSql(CATWithHeader, t_bFromLive:=True, _
                t_lLettercodeID:=t_lLettercodeID, t_lClassID:=t_lClassID, t_lHeaderID:=t_lHeaderID, t_sClassIn:=t_sClassIn, t_sLettercodeIn:=t_sLettercodeIn, t_bGetExtraExtnData:=t_bGetExtraExtnData) & _
            " UNION " & GetItemSql(CATWithHeaderAndDiv, t_bFromLive:=True, _
                t_lLettercodeID:=t_lLettercodeID, t_lClassID:=t_lClassID, t_lHeaderID:=t_lHeaderID, t_sClassIn:=t_sClassIn, t_sLettercodeIn:=t_sLettercodeIn, t_bGetExtraExtnData:=t_bGetExtraExtnData) & _
            " UNION " & GetItemSql(CATWithSubheaderAndDiv, t_bFromLive:=True, _
                t_lLettercodeID:=t_lLettercodeID, t_lClassID:=t_lClassID, t_lHeaderID:=t_lHeaderID, t_sClassIn:=t_sClassIn, t_sLettercodeIn:=t_sLettercodeIn, t_bGetExtraExtnData:=t_bGetExtraExtnData)
            
    End Select
    sMainSql = sMainSql & " ORDER BY " & gk_sDICatalogueLevel & ", " & gk_sDIID
    
    'We get the large text fields containing data for
    'verity zones in separate record sets as SQLServer cannot
    'cope with processing wide result sets
    
    'Title
    Set m_ofwTitleData = ofwDBHelper.FindBySql( _
        GetLargeFieldSql(LARGETitle, t_bFromLive:=True, _
            t_eLevel:=t_eLevel, t_lLettercodeID:=t_lLettercodeID, _
            t_lClassID:=t_lClassID, t_lHeaderID:=t_lHeaderID, t_sLettercodeIn:=t_sLettercodeIn, t_sClassIn:=t_sClassIn), _
        mk_lMaxRecords)
    
    'We only bother attempting to get data if there are > 0 entries
    'The title query is chosen as the first to run as it takes less
    'time than the main data query
    If m_ofwTitleData.RecordCount > 0 Then
        'Main entry details
        Set m_ofwMainData = ofwDBHelper.FindBySql(sMainSql, mk_lMaxRecords)
        
        'Scope-content
        Set m_ofwScopeData = ofwDBHelper.FindBySql( _
            GetLargeFieldSql(LARGEScopeContent, t_bFromLive:=True, _
                t_eLevel:=t_eLevel, t_lLettercodeID:=t_lLettercodeID, _
                t_lClassID:=t_lClassID, t_lHeaderID:=t_lHeaderID, t_sLettercodeIn:=t_sLettercodeIn, t_sClassIn:=t_sClassIn), _
            mk_lMaxRecords)
        'Admin history
        Set m_ofwAdminData = ofwDBHelper.FindBySql( _
            GetLargeFieldSql(LARGEAdminHistory, t_bFromLive:=True, _
                t_eLevel:=t_eLevel, t_lLettercodeID:=t_lLettercodeID, _
                t_lClassID:=t_lClassID, t_lHeaderID:=t_lHeaderID, t_sLettercodeIn:=t_sLettercodeIn, t_sClassIn:=t_sClassIn), _
            mk_lMaxRecords)
    Else
        'Just set the record sets to be empty (to prevent errors
        'if an attempt is made to access them)
        Set m_ofwMainData = New FWRecordList
        Set m_ofwScopeData = New FWRecordList
        Set m_ofwAdminData = New FWRecordList
        
    End If
    
    Exit Sub
    
ErrorHandler:
    ConvertError sSource
    
End Sub


Private Sub GetAdvancedData(t_eEntryType As EEditSetEntryType)
'   Get extra data needed for advanced collections
'   The entry type parameter determines whether we get
'   New or Updated records
'
'   d5  WDP     18-Nov-1999     Moved from CreateDataForAllEntries
'   d16     DAJ     11-May-2000     Added MaxRecords overide to GetAdvancedData (limit of 100
'                                   causes an EOF error during release as there will be a mismatch
'                                   between the number of 'basic' records and 'advanced' ones)

On Error GoTo ErrorHandler

    Dim sSource As String
    sSource = "BOCatalogueIndexer.GetAdvancedData"
    
    Dim ofwDBHelper As New FWDBHelper
    
    ' d16 - Added mk_lMaxRecords option
    Set m_ofwAccrualData = ofwDBHelper.FindBySql(GetLargeFieldSql(LARGEAccruals, t_eEntryType), mk_lMaxRecords)
    Set m_ofwAppInfoData = ofwDBHelper.FindBySql(GetLargeFieldSql(LARGEAppraisalInfo, t_eEntryType), mk_lMaxRecords)
    Set m_ofwArrangeData = ofwDBHelper.FindBySql(GetLargeFieldSql(LARGEArrangement, t_eEntryType), mk_lMaxRecords)
    Set m_ofwCustHistData = ofwDBHelper.FindBySql(GetLargeFieldSql(LARGECustodialHistory, t_eEntryType), mk_lMaxRecords)
    Set m_ofwNoteData = ofwDBHelper.FindBySql(GetLargeFieldSql(LARGENote, t_eEntryType), mk_lMaxRecords)
    Set m_ofwLinkData = ofwDBHelper.FindBySql(LinksSQL(), mk_lMaxRecords)

    Exit Sub
    
ErrorHandler:
    ConvertError sSource
    
End Sub


Private Sub GetAdvancedDataFromLive( _
    t_eLevel As ECATALOGUELEVEL, _
    Optional t_lLettercodeID As Long = 0, _
    Optional t_lClassID As Long, _
    Optional t_lHeaderID As Long, _
    Optional t_sLettercodeIn As String = "", _
    Optional t_sClassIn As String = "")
'   Reads advanced indexing data into the record sets
'   from live tables depending on the parameters supplied
'
'   d6  WDP     19-Nov-1999     Moved from IndexLIVEEntries
'   d8  WDP     13-Dec-1999     Made lettercode id optional
'                               If omitted all entries are selected
'                               Added LettercodeIn parameter to allow
'                               selection from multiple lettercodes
On Error GoTo ErrorHandler

    Dim sSource As String
    sSource = "BOCatalogueIndexer.GetAdvancedDataFromLive"
    
    Dim ofwDBHelper As New FWDBHelper
    
    'Only try to get data if there is some to get
    If m_ofwMainData.RecordCount > 0 Then
        Set m_ofwAccrualData = ofwDBHelper.FindBySql( _
            GetLargeFieldSql(LARGEAccruals, t_bFromLive:=True, _
                t_eLevel:=t_eLevel, t_lLettercodeID:=t_lLettercodeID, _
                t_lClassID:=t_lClassID, t_lHeaderID:=t_lHeaderID, t_sLettercodeIn:=t_sLettercodeIn, t_sClassIn:=t_sClassIn), _
            mk_lMaxRecords)
        Set m_ofwAppInfoData = ofwDBHelper.FindBySql( _
            GetLargeFieldSql(LARGEAppraisalInfo, t_bFromLive:=True, _
                t_eLevel:=t_eLevel, t_lLettercodeID:=t_lLettercodeID, _
                t_lClassID:=t_lClassID, t_lHeaderID:=t_lHeaderID, t_sLettercodeIn:=t_sLettercodeIn, t_sClassIn:=t_sClassIn), _
            mk_lMaxRecords)
        Set m_ofwArrangeData = ofwDBHelper.FindBySql( _
            GetLargeFieldSql(LARGEArrangement, t_bFromLive:=True, _
                t_eLevel:=t_eLevel, t_lLettercodeID:=t_lLettercodeID, _
                t_lClassID:=t_lClassID, t_lHeaderID:=t_lHeaderID, t_sLettercodeIn:=t_sLettercodeIn, t_sClassIn:=t_sClassIn), _
            mk_lMaxRecords)
        Set m_ofwCustHistData = ofwDBHelper.FindBySql( _
            GetLargeFieldSql(LARGECustodialHistory, t_bFromLive:=True, _
                t_eLevel:=t_eLevel, t_lLettercodeID:=t_lLettercodeID, _
                t_lClassID:=t_lClassID, t_lHeaderID:=t_lHeaderID, t_sLettercodeIn:=t_sLettercodeIn, t_sClassIn:=t_sClassIn), _
            mk_lMaxRecords)
        Set m_ofwNoteData = ofwDBHelper.FindBySql( _
            GetLargeFieldSql(LARGENote, t_bFromLive:=True, _
                t_eLevel:=t_eLevel, t_lLettercodeID:=t_lLettercodeID, _
                t_lClassID:=t_lClassID, t_lHeaderID:=t_lHeaderID, t_sLettercodeIn:=t_sLettercodeIn, t_sClassIn:=t_sClassIn), _
            mk_lMaxRecords)
        Set m_ofwLinkData = ofwDBHelper.FindBySql( _
            LinksSQL(True, t_eLevel, t_lLettercodeID:=t_lLettercodeID, _
                t_lClassID:=t_lClassID, t_lHeaderID:=t_lHeaderID, t_sLettercodeIn:=t_sLettercodeIn, t_sClassIn:=t_sClassIn), _
            mk_lMaxRecords)
    Else
        'Set the recordsets to be empty in case an attempt is
        'made to access them
        Set m_ofwAccrualData = New FWRecordList
        Set m_ofwAppInfoData = New FWRecordList
        Set m_ofwArrangeData = New FWRecordList
        Set m_ofwCustHistData = New FWRecordList
        Set m_ofwNoteData = New FWRecordList
        Set m_ofwLinkData = New FWRecordList
        
    End If
    Exit Sub
    
ErrorHandler:
    ConvertError sSource
    
End Sub



Public Sub CreateIndexDataForEntry(t_sEntry As String, t_sDate As String, t_sTime As String, t_lEditSetID As Long)
'   Create indexing data files for all
'   children of a catalogue entry (must be a piece or class)
'   For indexing moved catalogue entries
'
'   d7  WDP     26-Nov-1999     Created
On Error GoTo ErrorHandler

    Dim sSource As String
    sSource = "BOCatalogueIndexer.CreateIndexDataForEntry"
    
    m_lEditsetID = t_lEditSetID
    
    Dim eLevel As ECATALOGUELEVEL
    Dim lDBID As Long
    
    lDBID = DBIDFromDocKey(t_sEntry, eLevel)
    
    Select Case eLevel
        Case ctPiece
            CreateIndexDataForPiece lDBID, t_sDate, t_sTime
        Case ctClass
            CreateIndexDataForClass lDBID, t_sDate, t_sTime
        Case Else
            RaiseError sSource, PEGeneralInternalError, "Entry is not a piece or class"
    End Select
    Exit Sub
ErrorHandler:
    ConvertError sSource
    
End Sub

Public Sub CreateIndexDataForPiece(t_lPieceID As Long, t_sDate As String, t_sTime As String)
'   Create indexing data files for all
'   children of a piece
'
'   d7  WDP     26-Nov-1999     Created
On Error GoTo ErrorHandler

    Dim sSource As String
    sSource = "BOCatalogueIndexer.CreateIndexDataForPiece"
    
    Dim ofwDBHelper As New FWDBHelper
    Dim ofwRecords As FWRecordList
    Dim lCount As Long
    Dim ePieceStructure As ECatStructure
    Dim sLcode As String
    Dim sBasFilePrefix As String
    Dim sAdvFilePrefix As String
    Dim sIndexFileDrive As String
    
    Dim ofwAppData As New FWApplicationData
    sIndexFileDrive = ofwAppData.GetApplicationConfigurationInfomation(pRootForIndexingFiles)
    Set ofwAppData = Nothing
    
    'Determine catalogue structure above the piece
    Dim sSQL As String
    sSQL = "SELECT CASE WHEN c." & gk_sDBCOLDivisionID & " = 0 THEN " & _
            " CASE WHEN p." & gk_sDBCOLHeaderID & " IS NULL THEN " & CATPlain & _
            " ELSE CASE WHEN p." & gk_sDBCOLSubHeaderID & " IS NULL THEN " & CATWithHeader & _
            " ELSE " & CATWithSubheader & " END END ELSE " & _
            " CASE WHEN p." & gk_sDBCOLHeaderID & " IS NULL THEN " & CATWithDiv & _
            " ELSE CASE WHEN p." & gk_sDBCOLSubHeaderID & " IS NULL THEN " & CATWithHeaderAndDiv & _
            " ELSE " & CATWithSubheaderAndDiv & " END END END " & _
            " AS " & gk_sDICatStructure & ", " & _
            " lc." & gk_sDBCOLLetterCodeCode & " AS " & gk_sDILetterCodeRef & _
            " FROM " & gk_sDBTABLEPiece & " p, " & _
            gk_sDBTABLEClass & " c, " & _
            gk_sDBTABLELCode & " lc " & _
            " WHERE p." & gk_sDBCOLPieceID & "=" & t_lPieceID & _
            " AND p." & gk_sDBCOLClassID & "=c." & gk_sDBCOLClassID & _
            " AND c." & gk_sDBCOLLetterCodeID & "=lc." & gk_sDBCOLLetterCodeID
    
    Set ofwRecords = ofwDBHelper.FindBySql(sSQL)
    lCount = ofwRecords.RecordCount
    If lCount <> 1 Then
        RaiseError sSource, PEGeneralInternalError, "Query returned " & lCount & "records"
    Else
        ePieceStructure = ofwRecords.GetFieldValue(gk_sDICatStructure)
        sLcode = ofwRecords.GetFieldValue(gk_sDILetterCodeRef)
    End If
    
    Set m_ofwMainData = ofwDBHelper.FindBySql(GetItemSql(ePieceStructure, t_bFromLive:=True, t_lPieceID:=t_lPieceID))
    Set m_ofwAccrualData = ofwDBHelper.FindBySql(GetLargeFieldItemsInPieceSQL(t_lPieceID, LARGEAccruals))
    Set m_ofwAdminData = ofwDBHelper.FindBySql(GetLargeFieldItemsInPieceSQL(t_lPieceID, LARGEAdminHistory))
    Set m_ofwAppInfoData = ofwDBHelper.FindBySql(GetLargeFieldItemsInPieceSQL(t_lPieceID, LARGEAppraisalInfo))
    Set m_ofwArrangeData = ofwDBHelper.FindBySql(GetLargeFieldItemsInPieceSQL(t_lPieceID, LARGEArrangement))
    Set m_ofwCustHistData = ofwDBHelper.FindBySql(GetLargeFieldItemsInPieceSQL(t_lPieceID, LARGECustodialHistory))
    Set m_ofwLinkData = ofwDBHelper.FindBySql(LinksSQL(True, t_lPieceID:=t_lPieceID))
    Set m_ofwNoteData = ofwDBHelper.FindBySql(GetLargeFieldItemsInPieceSQL(t_lPieceID, LARGENote))
    Set m_ofwScopeData = ofwDBHelper.FindBySql(GetLargeFieldItemsInPieceSQL(t_lPieceID, LARGEScopeContent))
    Set m_ofwTitleData = ofwDBHelper.FindBySql(GetLargeFieldItemsInPieceSQL(t_lPieceID, LARGETitle))
    
    sBasFilePrefix = GetMovedEntryFileName(t_sDate, ctPiece, t_lPieceID, COLBasic)
    sAdvFilePrefix = GetMovedEntryFileName(t_sDate, ctPiece, t_lPieceID, COLAdvanced)
    
    Dim oboVerConfig As New BOVerityConfig
    Dim sLogFile As String
    sLogFile = LogFileName(t_sDate & t_sTime, m_lEditsetID, True)
    Dim sCollection As String, sBLKFilepath As String
        
    If WriteToFile(sBasFilePrefix, sLcode, COLBasic) Then
        'Only index if we actually wrote some data to file
        
        'find which collection the lettercode
        'maps to and run the mkvdk utility to index the bulk and sgml files
        
        'Get the collection path from verityconfig
        
        sCollection = oboVerConfig.GetLettercodeCollection(sLcode, COLBasic)
        sBLKFilepath = sIndexFileDrive & mk_sFilepath & sBasFilePrefix & ".blk"
        
        LogAction sLogFile, sBLKFilepath, sCollection, IDXUpdate, sLcode, m_lEditsetID
    End If
    
    If WriteToFile(sAdvFilePrefix, sLcode, COLAdvanced) Then
        'Only index if we actually wrote some data to file
        'find which collection the lettercode
        'maps to and run the mkvdk utility to index the bulk and sgml files
        
        'Get the collection path from verityconfig
        
        sCollection = oboVerConfig.GetLettercodeCollection(sLcode, COLAdvanced)
        sBLKFilepath = sIndexFileDrive & mk_sFilepath & sAdvFilePrefix & ".blk"
        
        LogAction sLogFile, sBLKFilepath, sCollection, IDXUpdate, sLcode, m_lEditsetID
    End If
    
    Exit Sub
ErrorHandler:
    ConvertError sSource
    
End Sub


Public Sub CreateIndexDataForClass(t_lClassID As Long, t_sDate As String, t_sTime As String)
'   Create indexing data files for all
'   children of a class
'
'   d7  WDP     26-Nov-1999     Created stub
On Error GoTo ErrorHandler

    Dim sSource As String
    sSource = "BOCatalogueIndexer.CreateIndexDataForClass"
    
    '+possibly break class up if its too big to index in one go.
    'Dim lPieceCount As Long, lItemCount As Long
    'Dim sPieceSql As String, sItemSql As String
    'Dim ofwHelper As New FWDBHelper
    'Dim ofwResults As FWRecordList
    
    'sPieceSql = " SELECT COUNT(*) AS " & gk_sDICount & " FROM " & gk_sDBTABLEPiece & _
                " WHERE " & gk_sDBCOLClassID & "=" & t_lClassID
    
    'sItemSql = " SELECT COUNT(*) AS " & gk_sDICount & _
                " FROM " & gk_sDBTABLEItem & " i, " & _
                gk_sDBTABLEPiece & " p " & _
                " WHERE p." & gk_sDBCOLClassID & "=" & t_lClassID & _
                " AND p." & gk_sDBCOLPieceID & "=i." & gk_sDBCOLPieceID
    
    'Set ofwResults = ofwHelper.FindBySql(sPieceSql)
    'lPieceCount = ofwResults.GetFieldValue(gk_sDICount)
    
    'Set ofwResults = ofwHelper.FindBySql(sItemSql)
    'lItemCount = ofwResults.GetFieldValue(gk_sDICount)
    
    ' dont bother using above data to break up should be OK for all classes
    CreateLiveIndexingData t_lClassID, ctClass, ctHeaders, t_sDate, t_sTime, True, IDXUpdate
    CreateLiveIndexingData t_lClassID, ctClass, ctSubHeaders, t_sDate, t_sTime, True, IDXUpdate
    CreateLiveIndexingData t_lClassID, ctClass, ctPiece, t_sDate, t_sTime, True, IDXUpdate
    CreateLiveIndexingData t_lClassID, ctClass, ctItem, t_sDate, t_sTime, True, IDXUpdate
    
    Exit Sub
ErrorHandler:
    ConvertError sSource
    
End Sub





Private Function GetEntryTypeString(t_eEntryType As EEditSetEntryType) As String
'   Convert enumerated entry type into string
'
'   d1  WDP     6-Oct-1999      Created
On Error GoTo ErrorHandler

    Dim sSource As String
    sSource = "BOCatalogueIndexer.GetEntryTypeString"
    
    Select Case t_eEntryType
    Case EDITNew
        GetEntryTypeString = "New"
    Case EDITUpdated
        GetEntryTypeString = "Updated"
    Case Else
        RaiseError sSource, PEGeneralInternalError, "Invalid editet entry type enum."
    End Select
    
    Exit Function
    
ErrorHandler:
    ConvertError sSource
    
End Function

Private Function GetCollectionTypeString(t_eCollectionType As ECollectionType) As String
'   Convert enumerated collection type into string
'
'   d1  WDP     7-Oct-1999      Created
On Error GoTo ErrorHandler

    Dim sSource As String
    sSource = "BOCatalogueIndexer.GetCollectionTypeString"
    
    Select Case t_eCollectionType
    Case COLAdvanced
        GetCollectionTypeString = "Adv"
    Case COLBasic
        GetCollectionTypeString = "Bas"
    Case COLClassUp
        GetCollectionTypeString = "CUp"
    Case Else
        RaiseError sSource, PEGeneralInternalError, "Invalid editet entry type enum."
    End Select
    
    Exit Function
    
ErrorHandler:
    ConvertError sSource
    
End Function

Private Function GetAllLettercodesSQL() As String
'   Generate sql to read all distinct lettercodes within
'   the edit set
'
'   d1  WDP     6-Oct-1999      Created
On Error GoTo ErrorHandler

    Dim sSource As String
    sSource = "BOCatalogueIndexer.GetAllLettercodesSQL"
    
    Dim sSQL As String
    sSQL = "SELECT DISTINCT " & _
            gk_sDBCOLLetterCodeCode & " AS " & gk_sDILetterCodeRef & _
            " FROM " & gk_sDBTABLEEditsetLettercode & " WHERE " & _
            gk_sDBCOLEditSetID & "=" & m_lEditsetID
            
    GetAllLettercodesSQL = sSQL
    
    Exit Function
    
ErrorHandler:
    ConvertError sSource
    
End Function


Private Function ProcessZone(t_sSource As String, t_sTag As String) As String
'   Strip all EAD tags from the source string and add enclosing
'   tags of the type given. Used to process text fields in
'   preparation for indexing as zones.
'
'   d1  WDP     11-Oct-1999      Created
'   d14 TMJ     30-Mar-2000      Added call to SpecialCharacterDecode
'   d17     DAJ     06-Feb-2001     Added extra stripping of zones to prevent multi-line values that cause parsing errors
On Error GoTo ErrorHandler

    Dim sSource As String
    sSource = "BOCatalogueIndexer.ProcessZone"
    
    ' d17 - Remove the tags, special characters and any newline characters
    Dim sText As String
    sText = SpecialCharacterDecode(RemoveTagsFromString(t_sSource))
    sText = ReplaceString(sText, vbNewLine, " ")
    sText = ReplaceString(sText, vbCrLf, " ")
    
    ProcessZone = "<" & t_sTag & ">" & sText & "</" & t_sTag & ">"
    
    Exit Function
ErrorHandler:
    ConvertError sSource
    
End Function



Private Function BuildReference(t_sLettercode As String, _
                                t_sClass As String, _
                                t_sSubclass As String, _
                                t_sPiece As String, _
                                t_sItem As String, _
                                t_eCatLevel As ECATALOGUELEVEL) As String
'   Build a reference from given parts and level
'
'   d1  WDP     6-Oct-1999      Created
On Error GoTo ErrorHandler

    Dim sSource As String
    sSource = "BOCatalogueIndexer.BuildReference"
    
    Dim sRef As String
    
    If t_sLettercode = "" Then
        RaiseError sSource, PEGeneralInternalError, "Empty lettercode"
    End If
    sRef = t_sLettercode
    If t_eCatLevel = ctLetterCode Or t_eCatLevel = ctDivision Then
        BuildReference = sRef
        Exit Function
    End If
    
    If t_sClass = "" Then
        RaiseError sSource, PEGeneralInternalError, "Empty class"
    End If
    sRef = sRef & t_sClass
        
    If Not t_sSubclass = "" Then
        sRef = sRef & "/" & t_sSubclass
    End If
    
    If t_eCatLevel = ctClass Or t_eCatLevel = ctHeaders _
            Or t_eCatLevel = ctSubHeaders Then
        BuildReference = sRef
        Exit Function
    End If
    
    If t_sPiece = "" Then
        RaiseError sSource, PEGeneralInternalError, "Empty piece"
    End If
    sRef = sRef & "/" & t_sPiece
    If t_eCatLevel = ctPiece Then
        BuildReference = sRef
        Exit Function
    End If
    
    If Not t_sItem = "" Then
        sRef = sRef & "/" & t_sItem
    End If
    BuildReference = sRef
    
    Exit Function
    
ErrorHandler:
    ConvertError sSource
End Function


Private Sub CreateLiveData( _
        t_eCollectionType As ECollectionType, _
        t_eLevel As ECATALOGUELEVEL, _
        t_sLettercode As String, _
        t_sDate As String, _
        t_sTime As String, _
        Optional t_eAction As EIndexingAction = IDXUpdate, _
        Optional t_lClassID As Long = 0, _
        Optional t_lHeaderID As Long = 0, _
        Optional t_lPieceID As Long, _
        Optional t_bForMove As Boolean = False, _
        Optional t_sPart As String = "")
'   Create test data from live tables
'
'   d7  WDP     30-Nov-1999     Created from IndexLIVEEntries

On Error GoTo ErrorHandler

    Dim sSource As String
    sSource = "BOCatalogueIndexer.CreateLiveData"
    
    Dim sIndexFileDrive As String
    
    Dim ofwAppData As New FWApplicationData
    sIndexFileDrive = ofwAppData.GetApplicationConfigurationInfomation(pRootForIndexingFiles)
    Set ofwAppData = Nothing
    
    Dim sFilePrefix As String
    sFilePrefix = GetFileName(t_sDate, True, _
            t_eCollectionType, t_sLettercode & t_sPart, t_eAction, t_eLevel, t_lClassID, t_lHeaderID, t_lPieceID)
    
    Dim sLogFile As String
    sLogFile = LogFileName(t_sDate & t_sTime, m_lEditsetID, t_bForMove)
    
    If WriteToFile(sFilePrefix, t_sLettercode, t_eCollectionType) Then
        'Only index if we actually wrote some data to file
        Dim oboVerConfig As New BOVerityConfig
        'find which collection the lettercode
        'maps to and run the mkvdk utility to index the bulk and sgml files
        
        'Get the collection path from verityconfig
        Dim sCollection As String, sBLKFilepath As String
        
        sCollection = oboVerConfig.GetLettercodeCollection(t_sLettercode, t_eCollectionType)
        sBLKFilepath = sIndexFileDrive & mk_sFilepath & sFilePrefix & ".blk"
        
        LogAction sLogFile, sBLKFilepath, sCollection, t_eAction, t_sLettercode, m_lEditsetID
        
    End If
    
    Exit Sub
    
ErrorHandler:
    ConvertError sSource
    
End Sub

Public Sub CreateLiveIndexingData(ByVal t_lParentCatalogueID As Long, _
        t_eParentLevel As ECATALOGUELEVEL, _
        t_eIndexingLevel As ECATALOGUELEVEL, _
        t_sDate As String, _
        t_sTime As String, _
        Optional t_bForMove As Boolean = False, _
        Optional t_eIndexAction As EIndexingAction = IDXUpdate, _
        Optional t_sClassIn As String = "", _
        Optional t_sPart As String = "")
'   This method is used to index sections of the catalogue from live
'   The parameters are as follows:
'   t_lParentCatalogueID - the DB id of the catalogue entry we are
'                          indexing beneath
'   t_eParentLevel       - the catalogue level of the parent
'   t_eIndexingLevel     - the level of the cat hierarchy below the parent
'                          that we wish to index
'   t_bForMove           - if we are using this method to index for
'                          moving cat entries on release this is True
'   t_eIndexingAction    - The type of indexing we are doing (insert, update, delete)

'   d7  WDP     30-Nov-1999      Created
'   d20 TomH    23-Aug-2002      Altered so doesn't create basic collections

On Error GoTo ErrorHandler

    Dim sSource As String
    sSource = "BOCatalogueIndexer.CreateLiveIndexingData"
    
    Dim lLettercodeID As Long, sLetterCode As String
    lLettercodeID = GetParentLettercodeID(t_lParentCatalogueID, t_eParentLevel)
    sLetterCode = GetParentLettercode(t_lParentCatalogueID, t_eParentLevel)
    
    'If we have a class in clause a dummy class id
    'was passed purely to obtain the parent lettercode
    'and we need to replace it with 0 when calling the
    'data creation methods
    Dim lClassID As Long
    If t_sClassIn = "" Then
        lClassID = t_lParentCatalogueID
    Else
        lClassID = 0
    End If
    
    If t_eIndexingLevel < t_eParentLevel Then
        RaiseError sSource, PEGeneralInternalError, "Cant index level " & t_eIndexingLevel & " with parent level " & t_eParentLevel
    End If
    
    '+only done for classes and lettercodes at the moment
    Select Case t_eParentLevel
    Case ctLetterCode
        GetBasicDataFromLive t_eIndexingLevel, lLettercodeID
        If m_ofwTitleData.RecordCount > 0 Then
            ' d20 tomh - removed
            'CreateLiveData COLBasic, t_eIndexingLevel, sLetterCode, t_sDate, t_sTime, t_eIndexAction, t_bForMove:=t_bForMove, t_sPart:=t_sPart
            GetAdvancedDataFromLive t_eIndexingLevel, lLettercodeID
            CreateLiveData COLAdvanced, t_eIndexingLevel, sLetterCode, t_sDate, t_sTime, t_eIndexAction, t_bForMove:=t_bForMove, t_sPart:=t_sPart
        End If
    Case ctClass
        GetBasicDataFromLive t_eIndexingLevel, lLettercodeID, lClassID, t_sClassIn:=t_sClassIn
        If m_ofwTitleData.RecordCount > 0 Then
            ' d20 tomh - removed
            'CreateLiveData COLBasic, t_eIndexingLevel, sLetterCode, t_sDate, t_sTime, t_eIndexAction, t_lParentCatalogueID, t_bForMove:=t_bForMove, t_sPart:=t_sPart
            GetAdvancedDataFromLive t_eIndexingLevel, lLettercodeID, lClassID, t_sClassIn:=t_sClassIn
            CreateLiveData COLAdvanced, t_eIndexingLevel, sLetterCode, t_sDate, t_sTime, t_eIndexAction, t_lParentCatalogueID, t_bForMove:=t_bForMove, t_sPart:=t_sPart
        End If
    End Select
    
    Exit Sub
ErrorHandler:
    ConvertError sSource
    
End Sub


Public Sub CreateSomeFromLetterCodes()
'   function to index the catalogue using selection of lettercodes.
'
'   To Use this function you need to do two things:
'   1) Alter the list of lettercodes below so that it contains a comma delimted list of
'      of LetterCode ID's
'   2) Alter the stored procedure sp_data_find_some_lettercode_pitem_sizes()
'      This also requires the same delimited list of lettercode ID's
'
'   d19 Tomh    16-Jul-2002     created
'   d20 Tomh    23-Aug-2002     removed all basic collection generation
On Error GoTo ErrorHandler

    Dim sSource As String
    sSource = "BOCatalogueIndexer.CreateSomeFromLetterCodes"
    
    
    Dim sLetterCodes As String
    Dim sLetterCodesExtra As String
    sLetterCodes = "243,205,114,156,258,184,249,44,235,244"
    sLetterCodesExtra = "(" & sLetterCodes & ")"
    
    
    If Not IndexStepExists(IXSTEPDataFileCreationCompleted) Then
        
        Dim sDate As String, sTime As String
        sDate = GetDateString()
        sTime = GetTimeString()
        
        'Create data files for lettercodes, divisions... down to subheaders
        ' d20 tomh - removed
        ' CreateSubheaderUpData sDate, sTime, COLBasic, sLetterCodesExtra
        CreateSubheaderUpData sDate, sTime, COLAdvanced, sLetterCodesExtra
        
        'Create class and above data files for small collection
        CreateAllClassUpData sDate, sTime, sLetterCodesExtra
                
        Dim sSQL As String
        Dim ofwHelper As New FWDBHelper
        
        LogIndexStep IXSTEPMessage, "Finding lettercode piece and item level sizes..."
        
        ' tomh changed the name of the stored procedure
        sSQL = " EXEC sp_data_find_some_lettercode_pitem_sizes "
        
        Set m_ofwLcodeSize = ofwHelper.FindBySql(sSQL, mk_lMaxRecords)
        m_lLcodeSizeRecordsetIndex = 1
        
        Dim lLettercodeID As Long
        Dim lEntries As Long
        Dim lItemEntries As Long
        
        While Not m_ofwLcodeSize.IsAtEnd
            lLettercodeID = m_ofwLcodeSize.GetFieldValue(gk_sDBCOLLetterCodeID)
            lEntries = m_ofwLcodeSize.GetFieldValue(gk_sDIEntries)
            lItemEntries = m_ofwLcodeSize.GetFieldValue(gk_sDIItemEntries)
            
                If lEntries < 10000 Then
                If Not IndexStepExists(IXSTEPSmallLettercodes, COLAdvanced, ctPiece) Then
                    CreateSmallLettercodeCatData sDate, sTime, COLAdvanced, ctPiece, IDXInsert
                    LogIndexStep IXSTEPSmallLettercodes, "Created advanced files for pieces in small lettercodes", COLAdvanced, ctPiece
                    'transfer the indexing actions to the master log
                    DumpLogToMasterLog LogFileName(sDate & sTime)
                End If
               If Not IndexStepExists(IXSTEPSmallLettercodes, COLAdvanced, ctItem) Then
                   CreateSmallLettercodeCatData sDate, sTime, COLAdvanced, ctItem, IDXInsert
                   LogIndexStep IXSTEPSmallLettercodes, "Created advanced files for items in small lettercodes", COLAdvanced, ctItem
                  'transfer the indexing actions to the master log
                   DumpLogToMasterLog LogFileName(sDate & sTime)
               End If
            
            End If
           
           'we may have moved on in the record list now so reset the parameters
            lLettercodeID = m_ofwLcodeSize.GetFieldValue(gk_sDBCOLLetterCodeID)
            lEntries = m_ofwLcodeSize.GetFieldValue(gk_sDIEntries)
            lItemEntries = m_ofwLcodeSize.GetFieldValue(gk_sDIItemEntries)
           
            'medium sized letttercodes are indexed individually and
            'large ones are split up
            If lEntries < mk_lMaxLettercodeSize And lEntries >= 10000 Then
                If Not IndexStepExists(IXSTEPMediumLettercode, t_eLevel:=ctPiece, t_lLettercodeID:=lLettercodeID) Then
                    CreateLiveIndexingData lLettercodeID, ctLetterCode, ctPiece, sDate, sTime, False, IDXInsert
                    LogIndexStep IXSTEPMediumLettercode, "Created files for pieces in medium sized lettercode", t_eLevel:=ctPiece, t_lLettercodeID:=lLettercodeID
                    'transfer the indexing actions to the master log
                    DumpLogToMasterLog LogFileName(sDate & sTime)
                End If
                If lItemEntries > 0 Then
                    If Not IndexStepExists(IXSTEPMediumLettercode, t_eLevel:=ctItem, t_lLettercodeID:=lLettercodeID) Then
                        CreateLiveIndexingData lLettercodeID, ctLetterCode, ctItem, sDate, sTime, False, IDXInsert
                        LogIndexStep IXSTEPMediumLettercode, "Created files for items in medium sized lettercode", t_eLevel:=ctItem, t_lLettercodeID:=lLettercodeID
                        'transfer the indexing actions to the master log
                        DumpLogToMasterLog LogFileName(sDate & sTime)
                    End If
                End If
            Else
                If lEntries >= mk_lMaxLettercodeSize Then
                    '+break up
                    CreateLargeLettercodeCatData lLettercodeID, sDate, sTime
                End If
            End If
            m_ofwLcodeSize.MoveNext
            m_lLcodeSizeRecordsetIndex = m_lLcodeSizeRecordsetIndex + 1
        Wend
        LogIndexStep IXSTEPDataFileCreationCompleted, "Index file generation completed"

    Else
        LogIndexStep IXSTEPDataFileCreationCompleted, "Data file generation not necessary - already performed fully"
    End If
    
    
    Exit Sub
ErrorHandler:
    LogIndexStep IXSTEPError, "Error occurred in indexing (see Error log for more detail)"
    ConvertError sSource
    
End Sub


Public Sub CreateAllCatalogueDataFromLive(Optional ByVal t_bIndexSubSet As Boolean)
'   Master function to index the catalogue
'
'   d8  WDP     13-Dec-1999     Created
'   d19 Tomh    16-Jul-2002     Altered so now hs optional parameter to index a subset
'   d20 Tomh    23-Aug-2002     Removed all basic collection generation
On Error GoTo ErrorHandler

    Dim sSource As String
    sSource = "BOCatalogueIndexer.CreateAllCatalogueDataFromLive"
    
    If t_bIndexSubSet Then
    
        ' Want to index subset. The subset to index is set in function below (CreateSomeFromLCodes)
        CreateSomeFromLetterCodes
        
    Else
    
    
    
    If Not IndexStepExists(IXSTEPDataFileCreationCompleted) Then
        
        Dim sDate As String, sTime As String
        sDate = GetDateString()
        sTime = GetTimeString()
        
        'Create data files for lettercodes, divisions... down to subheaders
        ' d20 tomh - removed
        ' CreateSubheaderUpData sDate, sTime, COLBasic
        CreateSubheaderUpData sDate, sTime, COLAdvanced
        
        'Create class and above data files for small collection
        CreateAllClassUpData sDate, sTime
                
        Dim sSQL As String
        Dim ofwHelper As New FWDBHelper
        
        LogIndexStep IXSTEPMessage, "Finding lettercode piece and item level sizes..."
        sSQL = " EXEC sp_data_find_lettercode_pitem_sizes "
        Set m_ofwLcodeSize = ofwHelper.FindBySql(sSQL, mk_lMaxRecords)
        m_lLcodeSizeRecordsetIndex = 1
        
        Dim lLettercodeID As Long
        Dim lEntries As Long
        Dim lItemEntries As Long
        
        While Not m_ofwLcodeSize.IsAtEnd
            lLettercodeID = m_ofwLcodeSize.GetFieldValue(gk_sDBCOLLetterCodeID)
            lEntries = m_ofwLcodeSize.GetFieldValue(gk_sDIEntries)
            lItemEntries = m_ofwLcodeSize.GetFieldValue(gk_sDIItemEntries)
            
            If lEntries < 10000 Then
                If Not IndexStepExists(IXSTEPSmallLettercodes, COLAdvanced, ctPiece) Then
                    CreateSmallLettercodeCatData sDate, sTime, COLAdvanced, ctPiece, IDXInsert
                    LogIndexStep IXSTEPSmallLettercodes, "Created advanced files for pieces in small lettercodes", COLAdvanced, ctPiece
                    'transfer the indexing actions to the master log
                    DumpLogToMasterLog LogFileName(sDate & sTime)
                End If
                If Not IndexStepExists(IXSTEPSmallLettercodes, COLAdvanced, ctItem) Then
                    CreateSmallLettercodeCatData sDate, sTime, COLAdvanced, ctItem, IDXInsert
                    LogIndexStep IXSTEPSmallLettercodes, "Created advanced files for items in small lettercodes", COLAdvanced, ctItem
                    'transfer the indexing actions to the master log
                    DumpLogToMasterLog LogFileName(sDate & sTime)
                End If
            
            End If
            
            'we may have moved on in the record list now so reset the parameters
            lLettercodeID = m_ofwLcodeSize.GetFieldValue(gk_sDBCOLLetterCodeID)
            lEntries = m_ofwLcodeSize.GetFieldValue(gk_sDIEntries)
            lItemEntries = m_ofwLcodeSize.GetFieldValue(gk_sDIItemEntries)
            
            'medium sized letttercodes are indexed individually and
            'large ones are split up
            If lEntries < mk_lMaxLettercodeSize And lEntries >= 10000 Then
                If Not IndexStepExists(IXSTEPMediumLettercode, t_eLevel:=ctPiece, t_lLettercodeID:=lLettercodeID) Then
                    CreateLiveIndexingData lLettercodeID, ctLetterCode, ctPiece, sDate, sTime, False, IDXInsert
                    LogIndexStep IXSTEPMediumLettercode, "Created files for pieces in medium sized lettercode", t_eLevel:=ctPiece, t_lLettercodeID:=lLettercodeID
                    'transfer the indexing actions to the master log
                    DumpLogToMasterLog LogFileName(sDate & sTime)
                End If
                If lItemEntries > 0 Then
                    If Not IndexStepExists(IXSTEPMediumLettercode, t_eLevel:=ctItem, t_lLettercodeID:=lLettercodeID) Then
                        CreateLiveIndexingData lLettercodeID, ctLetterCode, ctItem, sDate, sTime, False, IDXInsert
                        LogIndexStep IXSTEPMediumLettercode, "Created files for items in medium sized lettercode", t_eLevel:=ctItem, t_lLettercodeID:=lLettercodeID
                        'transfer the indexing actions to the master log
                        DumpLogToMasterLog LogFileName(sDate & sTime)
                    End If
                End If
            Else
                If lEntries >= mk_lMaxLettercodeSize Then
                    '+break up
                    CreateLargeLettercodeCatData lLettercodeID, sDate, sTime
                End If
            End If
            m_ofwLcodeSize.MoveNext
            m_lLcodeSizeRecordsetIndex = m_lLcodeSizeRecordsetIndex + 1
        Wend
        LogIndexStep IXSTEPDataFileCreationCompleted, "Index file generation completed"
    Else
        LogIndexStep IXSTEPDataFileCreationCompleted, "Data file generation not necessary - already performed fully"
    End If
    
    ' if is index subset
    End If
    
    Exit Sub
ErrorHandler:
    LogIndexStep IXSTEPError, "Error occurred in indexing (see Error log for more detail)"
    ConvertError sSource
    
End Sub


Public Sub CreateLargeLettercodeCatData(t_lLettercodeID As Long, t_sDate As String, t_sTime As String)
'   index large lettercodes
'
'   d8  WDP     13-Dec-1999     Created
On Error GoTo ErrorHandler

    Dim sSource As String
    sSource = "BOCatalogueIndexer.CreateLargeLettercodeCatData"
    
    Dim sSQL As String
    Dim ofwHelper As New FWDBHelper
    Dim ofwResults As FWRecordList
    
    'Find the sizes of the classes
    sSQL = " EXEC sp_data_find_class_sizes " & t_lLettercodeID
    Set ofwResults = ofwHelper.FindBySql(sSQL, mk_lMaxRecords)
    
    Dim lClassID As Long
    Dim lEntries As Long
    
    If Not ofwResults.IsAtEnd Then
        lEntries = ofwResults.GetFieldValue(gk_sDIEntries)
    End If
    
    'Results are ordered by # entries ascending
    'skip over empty classes
    While Not ofwResults.IsAtEnd And lEntries < 1
        ofwResults.MoveNext
        If Not ofwResults.IsAtEnd Then
            lEntries = ofwResults.GetFieldValue(gk_sDIEntries)
        End If
    Wend
    
    'get all the smaller classes and group them together in groups
    ' to index
    Dim lCount As Long
    lCount = 1
    While Not ofwResults.IsAtEnd And lEntries < mk_lMaxClassSize
        Dim lTotalGrouped As Long
        Dim sClassIn As String
        sClassIn = ""
        lTotalGrouped = lEntries
        While lTotalGrouped < mk_lMaxClassSize And Not ofwResults.IsAtEnd
            lClassID = ofwResults.GetFieldValue(gk_sDBCOLClassID)
        
            If sClassIn = "" Then
                sClassIn = "(" & lClassID
            Else
                sClassIn = sClassIn & ", " & lClassID
            End If
            ofwResults.MoveNext
            If Not ofwResults.IsAtEnd Then
                lEntries = ofwResults.GetFieldValue(gk_sDIEntries)
                lTotalGrouped = lTotalGrouped + lEntries
            End If
        Wend
        If Not sClassIn = "" Then
            sClassIn = sClassIn & ")"
            If Not IndexStepExists(IXSTEPLargeLettercode, t_eLevel:=ctPiece, t_lLettercodeID:=t_lLettercodeID, t_lClassID:=lClassID) Then
                CreateLiveIndexingData lClassID, ctClass, ctPiece, t_sDate, t_sTime, False, IDXInsert, sClassIn, "part" & lCount
                LogIndexStep IXSTEPLargeLettercode, "Created files for pieces from small classes in large lettercode from classID: " & lClassID, t_eLevel:=ctPiece, t_lLettercodeID:=t_lLettercodeID, t_lClassID:=lClassID
                'transfer the indexing actions to the master log
                DumpLogToMasterLog LogFileName(t_sDate & t_sTime)
            End If
            If Not IndexStepExists(IXSTEPLargeLettercode, t_eLevel:=ctItem, t_lLettercodeID:=t_lLettercodeID, t_lClassID:=lClassID) Then
                CreateLiveIndexingData lClassID, ctClass, ctItem, t_sDate, t_sTime, False, IDXInsert, sClassIn, "part" & lCount
                LogIndexStep IXSTEPLargeLettercode, "Created files for items from small classes in large lettercode from classID: " & lClassID, t_eLevel:=ctItem, t_lLettercodeID:=t_lLettercodeID, t_lClassID:=lClassID
                'transfer the indexing actions to the master log
                DumpLogToMasterLog LogFileName(t_sDate & t_sTime)
            End If
        End If
        lCount = lCount + 1
    Wend
      
    While Not ofwResults.IsAtEnd
        'do big classes individually
        lClassID = ofwResults.GetFieldValue(gk_sDBCOLClassID)
        If Not IndexStepExists(IXSTEPLargeLettercode, t_eLevel:=ctPiece, t_lLettercodeID:=t_lLettercodeID, t_lClassID:=lClassID) Then
            CreateLiveIndexingData lClassID, ctClass, ctPiece, t_sDate, t_sTime, False, IDXInsert
            LogIndexStep IXSTEPLargeLettercode, "Created files for pieces from large class in large lettercode ", t_eLevel:=ctPiece, t_lLettercodeID:=t_lLettercodeID, t_lClassID:=lClassID
            'transfer the indexing actions to the master log
            DumpLogToMasterLog LogFileName(t_sDate & t_sTime)
        End If
        If Not IndexStepExists(IXSTEPLargeLettercode, t_eLevel:=ctItem, t_lLettercodeID:=t_lLettercodeID, t_lClassID:=lClassID) Then
            CreateLiveIndexingData lClassID, ctClass, ctItem, t_sDate, t_sTime, False, IDXInsert
            LogIndexStep IXSTEPLargeLettercode, "Created files for items from large class in large lettercode ", t_eLevel:=ctPiece, t_lLettercodeID:=t_lLettercodeID, t_lClassID:=lClassID
            'transfer the indexing actions to the master log
            DumpLogToMasterLog LogFileName(t_sDate & t_sTime)
        End If
        
        ofwResults.MoveNext
    Wend
    
    Exit Sub
ErrorHandler:
    ConvertError sSource
    
End Sub

Public Sub CreateSmallLettercodeCatData( _
        t_sDate As String, t_sTime As String, _
        t_eCollectionType As ECollectionType, _
        t_eLevel As ECATALOGUELEVEL, t_eAction As EIndexingAction)
'   index small lettercodes by grouping them together
'
'   d8  WDP     23-Dec-1999     Created
On Error GoTo ErrorHandler

    Dim sSource As String
    sSource = "BOCatalogueIndexer.CreateSmallLettercodeCatData"
    
    Dim ofwAppData As New FWApplicationData
    Dim sFileDrive As String
    Dim sCollectionDrive As String
    sFileDrive = ofwAppData.GetApplicationConfigurationInfomation(pRootForIndexingFiles)
    sCollectionDrive = ofwAppData.GetApplicationConfigurationInfomation(pRootForCollections)
    
    Dim oboVerConfig As New BOVerityConfig
    
    Dim colCollectionClusters As New Collection
    Set colCollectionClusters = oboVerConfig.CollectionLettercodes(t_eCollectionType)
    
    Dim lEntries As Long
    Dim lTotalEntries As Long
    Dim sEntriesField
    
    'The recordset contains the number of pieces and items
    'for each lettercode, we need to select the right total.
    If t_eLevel = ctPiece Then
        sEntriesField = gk_sDIPieceEntries
    Else
        sEntriesField = gk_sDIItemEntries
    End If
    
    
    
    Dim lLettercodeID As Long
    
    Dim lCount As Long
    For lCount = 1 To colCollectionClusters.Count
        
        Dim sCollection As String
        Dim sData As String
        'start from the smallest
        m_ofwLcodeSize.MoveFirst
        lEntries = m_ofwLcodeSize.GetFieldValue(sEntriesField)
        lTotalEntries = m_ofwLcodeSize.GetFieldValue(gk_sDIEntries)
    
        'Look for the first non empty set of pieces/items
        While lEntries = 0 And Not m_ofwLcodeSize.IsAtEnd
            m_ofwLcodeSize.MoveNext
            lEntries = m_ofwLcodeSize.GetFieldValue(sEntriesField)
            lTotalEntries = m_ofwLcodeSize.GetFieldValue(gk_sDIEntries)
        Wend
        sData = colCollectionClusters.Item(lCount)
        sCollection = Left(sData, InStr(sData, gk_sClusterPathDelimiter) - 1)
        
        'For labelling the bulk and sgml files
        Dim lSection As Long
        lSection = 1
        
        While (Not m_ofwLcodeSize.IsAtEnd) And lTotalEntries < 10000
            Dim lTotal As Long
            lTotal = 0
            Dim sLettercodeIn As String
            sLettercodeIn = ""
            
            While lTotal < mk_lMedLettercodeSize And (Not m_ofwLcodeSize.IsAtEnd) And lTotalEntries < 10000
                lLettercodeID = m_ofwLcodeSize.GetFieldValue(gk_sDBCOLLetterCodeID)
                If Not lEntries = 0 Then
                    If InStr(sData, "," & CStr(lLettercodeID) & ",") > 0 _
                    Or InStr(sData, "," & CStr(lLettercodeID) & ")") > 0 _
                    Or InStr(sData, "(" & CStr(lLettercodeID) & ",") > 0 Then
                        lTotal = lTotal + lEntries
                        If sLettercodeIn = "" Then
                            sLettercodeIn = "(" & lLettercodeID
                        Else
                            sLettercodeIn = sLettercodeIn & ", " & lLettercodeID
                        End If
                    End If
                End If
                
                m_ofwLcodeSize.MoveNext
                
                If Not m_ofwLcodeSize.IsAtEnd Then
                    lEntries = m_ofwLcodeSize.GetFieldValue(sEntriesField)
                    lTotalEntries = m_ofwLcodeSize.GetFieldValue(gk_sDIEntries)
                End If
                
                
            Wend
            
            sLettercodeIn = sLettercodeIn & ")"
            
            If t_eCollectionType = COLBasic Then
                GetBasicDataFromLive t_eLevel, t_sLettercodeIn:=sLettercodeIn, t_bGetExtraExtnData:=False
            Else
                GetBasicDataFromLive t_eLevel, t_sLettercodeIn:=sLettercodeIn, t_bGetExtraExtnData:=True
                GetAdvancedDataFromLive t_eLevel, t_sLettercodeIn:=sLettercodeIn
            End If
            
            Dim sFilenamePrefix As String
            Dim sBLKFilepath As String
            Dim sLogFile As String
            sLogFile = LogFileName(t_sDate & t_sTime)
            
            sFilenamePrefix = GetFileName(t_sDate & "part" & lCount & "." & lSection, True, t_eCollectionType, t_eAction:=t_eAction, t_eLevel:=t_eLevel)
            sBLKFilepath = sFileDrive & mk_sFilepath & sFilenamePrefix & ".blk"
            
            If WriteToFile(sFilenamePrefix, t_eCollectionType:=t_eCollectionType) Then
                'The indexing actions required are logged to a file
                'which can later be read to perform the indexing itself
                'from the bulk and sgml files generated
                LogAction sLogFile, sBLKFilepath, sCollectionDrive & sCollection, t_eAction
            End If
            
            'increment the label number so we dont overwrite the
            'file next time round the loop
            lSection = lSection + 1
        Wend
        
    Next
    
    
    Exit Sub
ErrorHandler:
    ConvertError sSource
    
End Sub

Private Function LinkBulkData(t_iLevel As Integer, t_lCatalogueID As Long)
'   produces text for link data items in the bulk file
'
'   d2  WDP     25-Oct-1999     Created
'   d8  WDP     13-Dec-1999     Made more efficient

On Error GoTo ErrorHandler

    Dim sSource As String
    sSource = "BOCatalogueIndexer.LinkBulkData"
    
    'variables to store the field values
    Dim sCorpNameText As String
    Dim sCopiesText As String
    Dim sCreatorText As String
    Dim sLocOfOrigText As String
    Dim sPersonText As String
    Dim sPlaceText As String
    Dim sPlaceOfDepText As String
    Dim sSourceOfAcqText As String
    Dim sSubjectText As String
    
    'We look at the link data until we run past the
    'appropriate level or cat id, in this way we only need pass
    'over the record set once per set of indexing data as the data is
    'ordered by level then DB id.
    Do While Not m_ofwLinkData.IsAtEnd
        Dim iLevel As Integer
        iLevel = m_ofwLinkData.GetFieldValue(gk_sDICatalogueLevel)
        If iLevel > t_iLevel Then
            Exit Do
        End If
        'while we are below the desired level, skip
        While iLevel < t_iLevel
            m_ofwLinkData.MoveNext
            If m_ofwLinkData.IsAtEnd Then
                Exit Do
            End If
            iLevel = m_ofwLinkData.GetFieldValue(gk_sDICatalogueLevel)
        Wend
        Dim lCatID As Long
        lCatID = m_ofwLinkData.GetFieldValue(gk_sDICatalogueID)
        'while we are below the appropriate id, skip
        While lCatID < t_lCatalogueID
            m_ofwLinkData.MoveNext
            If m_ofwLinkData.IsAtEnd Then
                Exit Do
            End If
            lCatID = m_ofwLinkData.GetFieldValue(gk_sDICatalogueID)
            iLevel = m_ofwLinkData.GetFieldValue(gk_sDICatalogueLevel)
            'if we have entered a new level we need to break out
            ' - there is no data for this entry
            If iLevel > t_iLevel Then
                Exit Do
            End If
        Wend
        
        If lCatID > t_lCatalogueID Then
            Exit Do
        End If
        
        'All the links for an entry are grouped together so
        'we extract them into the appropriate variables
        Do While lCatID = t_lCatalogueID And iLevel = t_iLevel
            Dim lLinkTypeID As Long
            lLinkTypeID = m_ofwLinkData.GetFieldValue(gk_sDILinkTypeID)
            Select Case lLinkTypeID
                Case mk_lCorpBodyPropID
                    sCorpNameText = sCorpNameText & gk_sLinkDataDelimiter & _
                        m_ofwLinkData.GetFieldValue(gk_sDILinkDataID) & gk_sLinkDataDelimiter
                Case mk_lCopiesPropID
                    sCopiesText = sCopiesText & gk_sLinkDataDelimiter & _
                        m_ofwLinkData.GetFieldValue(gk_sDILinkDataID) & gk_sLinkDataDelimiter
                Case mk_lCreatorPropID
                    sCreatorText = sCreatorText & gk_sLinkDataDelimiter & _
                        m_ofwLinkData.GetFieldValue(gk_sDILinkDataID) & gk_sLinkDataDelimiter
                Case mk_lLocOfOrigPropID
                    sLocOfOrigText = sLocOfOrigText & gk_sLinkDataDelimiter & _
                        m_ofwLinkData.GetFieldValue(gk_sDILinkDataID) & gk_sLinkDataDelimiter
                Case mk_lPersonPropID
                    sPersonText = sPersonText & gk_sLinkDataDelimiter & _
                        m_ofwLinkData.GetFieldValue(gk_sDILinkDataID) & gk_sLinkDataDelimiter
                Case mk_lPlaceOfDepPropID
                    sPlaceOfDepText = sPlaceOfDepText & gk_sLinkDataDelimiter & _
                        m_ofwLinkData.GetFieldValue(gk_sDILinkDataID) & gk_sLinkDataDelimiter
                Case mk_lPlacePropID
                    sPlaceText = sPlaceText & gk_sLinkDataDelimiter & _
                        m_ofwLinkData.GetFieldValue(gk_sDILinkDataID) & gk_sLinkDataDelimiter
                Case mk_lSourceOfAcqPropID
                    sSourceOfAcqText = sSourceOfAcqText & gk_sLinkDataDelimiter & _
                        m_ofwLinkData.GetFieldValue(gk_sDILinkDataID) & gk_sLinkDataDelimiter
                Case mk_lSubjectPropID
                    sSubjectText = sSubjectText & gk_sLinkDataDelimiter & _
                        m_ofwLinkData.GetFieldValue(gk_sDILinkDataID) & gk_sLinkDataDelimiter
            End Select
            m_ofwLinkData.MoveNext
            If m_ofwLinkData.IsAtEnd Then
                Exit Do
            End If
            lCatID = m_ofwLinkData.GetFieldValue(gk_sDICatalogueID)
            iLevel = m_ofwLinkData.GetFieldValue(gk_sDICatalogueLevel)
        Loop
        
        Dim sText As String
        
        'Create the text for the bulk file
        If Not sCopiesText = "" Then
            sText = sText & gk_sVFIELDCopies & ": """ & sCopiesText & """" & vbCrLf
        End If
        If Not sCorpNameText = "" Then
            sText = sText & gk_sVFIELDCorporateName & ": """ & sCorpNameText & """" & vbCrLf
        End If
        If Not sCreatorText = "" Then
            sText = sText & gk_sVFIELDCreatorName & ": """ & sCreatorText & """" & vbCrLf
        End If
        If Not sLocOfOrigText = "" Then
            sText = sText & gk_sVFIELDLocationOfOriginals & ": """ & sLocOfOrigText & """" & vbCrLf
        End If
        If Not sPersonText = "" Then
            sText = sText & gk_sVFIELDPerson & ": """ & sPersonText & """" & vbCrLf
        End If
        If Not sPlaceText = "" Then
            sText = sText & gk_sVFIELDPlace & ": """ & sPlaceText & """" & vbCrLf
        End If
        If Not sPlaceOfDepText = "" Then
            sText = sText & gk_sVFIELDPlaceOfDeposit & ": """ & sPlaceOfDepText & """" & vbCrLf
        End If
        If Not sSourceOfAcqText = "" Then
            sText = sText & gk_sVFIELDSourceOfAcquisition & ": """ & sSourceOfAcqText & """" & vbCrLf
        End If
        If Not sSubjectText = "" Then
            sText = sText & gk_sVFIELDSubject & ": """ & sSubjectText & """" & vbCrLf
        End If
        
        'Break out after we have extracted all data
        Exit Do
    Loop
    
    LinkBulkData = sText
    
    Exit Function
ErrorHandler:
    ConvertError sSource
    
End Function


Private Function GetLettercodeSql( _
            Optional t_eEntryType As EEditSetEntryType, _
            Optional t_bFromLive As Boolean = False, _
            Optional t_lLettercodeID As Long = 0, _
            Optional t_sLettercodeIn As String = "", _
            Optional t_bGetExtraExtnData As Boolean = True) As String
'   Generate the sql to select lettercodes from either the edit set
'   or live tables.
'   The parameters define the following:
'   t_eEntryType        For selecting from edit sets, the edit set entry type
'   t_bFromLive         Flag to indicate we are selecting from live tables
'                       (default is False (ie, from edit sets)
'   t_sLettercode       For selecting from live, the lettercode we
'                       select from
'   d2  WDP     25-Oct-1999      Combined GetLettercodeSql and
'                                GetLettercodeSqlFromLIVE to form
'                                Generic method.
'   d5  WDP     19-Nov-1999      Removed collection type parameter
'                                we now select all data
'   d7  WDP     30-Nov-1999     Altered to use lettercode id col
'   d8  WDP     13-Dec-1999     Allowed lettercode id to be 0
'                               If it is then all lettercodes are selected
'                               Added lettercode in list to allow
'                               selection from multiple lettercodes
'   d9  WDP     14-Feb-2000     Select non-modified entries that have been accessioned or created
'                               when generating new editset data
'   d12 WDP     14-Mar-2000     Altered SQL generation to only select live records
'                                   when selecting from LIVE (live_flag = 0 or 1)
'   d18 ALM     25-Oct-2001     Supports a new type of entry - delete
'
On Error GoTo ErrorHandler

    Dim sSource As String
    sSource = "BOCatalogueIndexer.GetLettercodeSql"
    
    Dim sSQL As String
    'We select NULL values for the fields which do not exist at
    'this level to maintain the common shape of the record set
    sSQL = "SELECT " & _
            "1 AS " & gk_sDICatalogueLevel & ", " & _
            "lc." & gk_sDBCOLLetterCodeID & " AS " & gk_sDIID & ", " & _
            "lc." & gk_sDBCOLLetterCodeCode & " AS " & gk_sDILetterCodeRef & ", " & _
            " NULL AS " & gk_sDIDivisionRef & ", " & _
            " NULL AS " & gk_sDIClassRef & ", " & _
            " NULL AS " & gk_sDISubclassRef & ", " & _
            " NULL AS " & gk_sDIHeaderRef & ", " & _
            " NULL AS " & gk_sDISubheaderRef & ", " & _
            " NULL AS " & gk_sDIPieceRef & ", " & _
            " NULL AS " & gk_sDIItemRef & ", " & _
            " NULL AS " & gk_sDIPieceKeyOrder & ", " & _
            " NULL AS " & gk_sDIItemKeyOrder & ", " & _
            "x." & gk_sDBCOLCoveringFromDate & " AS " & gk_sDIFirstDate & ", " & _
            "x." & gk_sDBCOLCoveringToDate & " AS " & gk_sDILastDate & ", " & _
            "x." & gk_sDBCOLCoveringDate & " AS " & gk_sDICoveringDates & ", " & _
            "x." & gk_sDBCOLMapDesignation & " AS " & gk_sDIMapDesignation & ", " & _
            "x." & gk_sDBCOLFormerDepartmentReference & " AS " & gk_sDIFormerRef & ", " & _
            "x." & gk_sDBCOLFormerPROReference & " AS " & gk_sDIFormerPRORef
    'Extra data for advanced collections
    If t_bGetExtraExtnData = True Then
        sSQL = sSQL & ", " & _
             "x." & gk_sDBCOLMapScaleNumber & " AS " & gk_sDIMapScale & ", " & _
             "x." & gk_sDBCOLRestrictionsOnUse & " AS " & gk_sDIRestrictions & ", " & _
             "x." & gk_sDBCOLAccumulationdates & " AS " & gk_sDIAccumulationDates & ", " & _
             "x." & gk_sDBCOLRecordOpeningDate & " AS " & gk_sDIOpeningDates & ", " & _
             "x." & gk_sDBCOLLegalStatusID & " AS " & gk_sDILegalStatus & ", " & _
             "x." & gk_sDBCOLPhysicalDescriptionFormID & " AS " & gk_sDIPhysicalDescForm & ", " & _
             "x." & gk_sDBCOLPhysicalDescriptionRecordQuantity & " AS " & gk_sDIPhysDescQuantity & ", " & _
             "x." & gk_sDBCOLPhysicalDescriptionDimension & " AS " & gk_sDIPhysDescDimensions & ", " & _
             "x." & gk_sDBCOLAccessConditionID & " AS " & gk_sDIAccessCondition & ", " & _
             "x." & gk_sDBCOLLanguageID & " AS " & gk_sDILanguage & ", " & _
             "x." & gk_sDBCOLPhysicalCondition & " AS " & gk_sDIPhysicalCondition & ", " & _
             " NULL AS " & gk_sDIClosureCode & ", " & _
             " NULL AS " & gk_sDIClosureType & ", " & _
             "lc." & gk_sDBCOLClosureStatus & " AS " & gk_sDIClosureStatus
    End If
    'The tables we are selecting from
    sSQL = sSQL & " FROM "
    If t_bFromLive = True Then
        sSQL = sSQL & _
            gk_sDBTABLELCode & " lc, " & _
            gk_sDBTABLELetterCodeExtension & " x "
    Else
        sSQL = sSQL & _
            gk_sDBTABLEEditsetLettercode & " lc, " & _
            gk_sDBTABLEEditsetLettercodeExtension & " x "
    End If
    
    'The joining clauses
    sSQL = sSQL & _
            " WHERE " & _
            "lc." & gk_sDBCOLLetterCodeID & "=x." & gk_sDBCOLLetterCodeID
    If t_bFromLive = False Then
        sSQL = sSQL & " AND lc." & gk_sDBCOLEditSetID & "=" & m_lEditsetID & _
                  " AND x." & gk_sDBCOLEditSetID & "=" & m_lEditsetID
    Else
        If Not t_lLettercodeID = 0 Then
            If t_sLettercodeIn = "" Then
                sSQL = sSQL & " AND lc." & gk_sDBCOLLetterCodeID & "=" & t_lLettercodeID
            End If
        Else
            If Not t_sLettercodeIn = "" Then
                sSQL = sSQL & _
                    " AND lc." & gk_sDBCOLLetterCodeID & " IN " & t_sLettercodeIn
            End If
        End If
        sSQL = sSQL & " AND x." & gk_sDBCOLLiveFlag & " IN (0,1) "
        
    End If
        
    'Add filter to get the desired type of records (new or updated)
    If t_bFromLive = False Then
        Select Case t_eEntryType
            Case EDITNew
                sSQL = sSQL & " AND (x." & gk_sDBCOLEditSetEntryTypeID & " IN (" & _
                    cesImportedModified & ", " & cesPROCATCreatedModified & ", " & _
                    cesPROCATCreated & ", " & cesImported & ")) "
            Case EDITUpdated
                sSQL = sSQL & " AND x." & gk_sDBCOLEditSetEntryTypeID & "=" & cesAddedFromLiveModified
            Case EDITDelete 'd18
                sSQL = sSQL & " AND x." & gk_sDBCOLEditSetEntryTypeID & "=" & cesAddedFromLive
            Case Else
                RaiseError sSource, PEGeneralInternalError, "Invalid editset entry type"
        End Select
    End If
    
    GetLettercodeSql = sSQL
    
    Exit Function
    
ErrorHandler:
    ConvertError sSource
    
End Function

Private Function GetDivisionSql( _
            Optional t_eEntryType As EEditSetEntryType, _
            Optional t_bFromLive As Boolean = False, _
            Optional t_lLettercodeID As Long = 0, _
            Optional t_sLettercodeIn As String, _
            Optional t_bGetExtraExtnData As Boolean = True) As String
'   Generate the sql to select divisions from either the edit set
'   or live tables.
'   The parameters define the following:
'   t_eEntryType        For selecting from edit sets, the edit set entry type
'   t_bFromLive         Flag to indicate we are selecting from live tables
'                       (default is False (ie, from edit sets)
'   t_sLettercode       For selecting from live, the lettercode we
'                       select from
'   d2  WDP     25-Oct-1999      Combined GetDivisionSql and
'                                GetDivisionSqlFromLIVE to form
'                                Generic method.
'   d5  WDP     19-Nov-1999      Removed collection type parameter
'                                we now select all data
'   d7  WDP     30-Nov-1999      Use lettercode id col
'   d8  WDP     13-Dec-1999     Allowed lettercode id to be 0
'                               If it is then all divisions are selected
'                               Added lettercode in list to allow
'                               selection from multiple lettercodes
'   d9  WDP     14-Feb-2000     Select non-modified entries that have been accessioned or created
'                               when generating new editset data
'   d12 WDP     14-Mar-2000     Altered SQL generation to only select live records
'                                   when selecting from LIVE (live_flag = 0 or 1)
'   d18 ALM     25-Oct-2001     Supports a new type of entry - delete
'
On Error GoTo ErrorHandler

    Dim sSource As String
    sSource = "BOCatalogueIndexer.GetDivisionSql"
    
    Dim sSQL As String
    'We select NULL values for the fields which do not exist at
    'this level to maintain the cammon shape of the record set
    sSQL = "SELECT " & _
            "2 AS " & gk_sDICatalogueLevel & ", " & _
            "d." & gk_sDBCOLDivisionID & " AS " & gk_sDIID & ", " & _
            "lc." & gk_sDBCOLLetterCodeCode & " AS " & gk_sDILetterCodeRef & ", " & _
            "d." & gk_sDBCOLDivisionNo & " AS " & gk_sDIDivisionRef & ", " & _
            " NULL AS " & gk_sDIClassRef & ", " & _
            " NULL AS " & gk_sDISubclassRef & ", " & _
            " NULL AS " & gk_sDIHeaderRef & ", " & _
            " NULL AS " & gk_sDISubheaderRef & ", " & _
            " NULL AS " & gk_sDIPieceRef & ", " & _
            " NULL AS " & gk_sDIItemRef & ", " & _
            " NULL AS " & gk_sDIPieceKeyOrder & ", " & _
            " NULL AS " & gk_sDIItemKeyOrder & ", " & _
            "x." & gk_sDBCOLCoveringFromDate & " AS " & gk_sDIFirstDate & ", " & _
            "x." & gk_sDBCOLCoveringToDate & " AS " & gk_sDILastDate & ", " & _
            "x." & gk_sDBCOLCoveringDate & " AS " & gk_sDICoveringDates & ", " & _
            "x." & gk_sDBCOLMapDesignation & " AS " & gk_sDIMapDesignation & ", " & _
            "x." & gk_sDBCOLFormerDepartmentReference & " AS " & gk_sDIFormerRef & ", " & _
            "x." & gk_sDBCOLFormerPROReference & " AS " & gk_sDIFormerPRORef
    
    'Extra fields for advanced collections
    If t_bGetExtraExtnData = True Then
        sSQL = sSQL & ", " & _
             "x." & gk_sDBCOLMapScaleNumber & " AS " & gk_sDIMapScale & ", " & _
             "x." & gk_sDBCOLRestrictionsOnUse & " AS " & gk_sDIRestrictions & ", " & _
             "x." & gk_sDBCOLAccumulationdates & " AS " & gk_sDIAccumulationDates & ", " & _
             "x." & gk_sDBCOLRecordOpeningDate & " AS " & gk_sDIOpeningDates & ", " & _
             "x." & gk_sDBCOLLegalStatusID & " AS " & gk_sDILegalStatus & ", " & _
             "x." & gk_sDBCOLPhysicalDescriptionFormID & " AS " & gk_sDIPhysicalDescForm & ", " & _
             "x." & gk_sDBCOLPhysicalDescriptionRecordQuantity & " AS " & gk_sDIPhysDescQuantity & ", " & _
             "x." & gk_sDBCOLPhysicalDescriptionDimension & " AS " & gk_sDIPhysDescDimensions & ", " & _
             "x." & gk_sDBCOLAccessConditionID & " AS " & gk_sDIAccessCondition & ", " & _
             "x." & gk_sDBCOLLanguageID & " AS " & gk_sDILanguage & ", " & _
             "x." & gk_sDBCOLPhysicalCondition & " AS " & gk_sDIPhysicalCondition & ", " & _
             " NULL AS " & gk_sDIClosureCode & ", " & _
             " NULL AS " & gk_sDIClosureType & ", " & _
             "d." & gk_sDBCOLClosureStatus & " AS " & gk_sDIClosureStatus
    End If
    
    'The tables we are selecting from
    sSQL = sSQL & " FROM "
    If t_bFromLive = False Then
        sSQL = sSQL & _
            gk_sDBTABLEEditsetLettercode & " lc, " & _
            gk_sDBTABLEEditsetDivision & " d, " & _
            gk_sDBTABLEEditsetDivisionExtension & " x "
    Else
        sSQL = sSQL & _
            gk_sDBTABLELCode & " lc, " & _
            gk_sDBTABLEDivision & " d, " & _
            gk_sDBTableDivisionExtension & " x "
    End If
    
    'The joining clauses
    sSQL = sSQL & _
            " WHERE " & _
            "d." & gk_sDBCOLLetterCodeID & "=lc." & gk_sDBCOLLetterCodeID & " AND " & _
            "d." & gk_sDBCOLDivisionID & "=x." & gk_sDBCOLDivisionID
            
    If t_bFromLive = False Then
        sSQL = sSQL & " AND lc." & gk_sDBCOLEditSetID & "=" & m_lEditsetID & _
                  " AND d." & gk_sDBCOLEditSetID & "=" & m_lEditsetID & _
                  " AND x." & gk_sDBCOLEditSetID & "=" & m_lEditsetID
    Else
        If Not t_lLettercodeID = 0 Then
            If t_sLettercodeIn = "" Then
                sSQL = sSQL & " AND lc." & gk_sDBCOLLetterCodeID & "=" & t_lLettercodeID
            End If
        Else
            If Not t_sLettercodeIn = "" Then
                sSQL = sSQL & _
                    " AND lc." & gk_sDBCOLLetterCodeID & " IN " & t_sLettercodeIn
            End If
        End If
        sSQL = sSQL & " AND x." & gk_sDBCOLLiveFlag & " IN (0,1) "
    End If
    
    'Add filter to get the desired type of records (new or updated)
    If t_bFromLive = False Then
        Select Case t_eEntryType
            Case EDITNew
                sSQL = sSQL & " AND (x." & gk_sDBCOLEditSetEntryTypeID & " IN (" & _
                    cesImportedModified & ", " & cesPROCATCreatedModified & ", " & _
                    cesPROCATCreated & ", " & cesImported & ")) "
            Case EDITUpdated
                sSQL = sSQL & " AND x." & gk_sDBCOLEditSetEntryTypeID & "=" & cesAddedFromLiveModified
            Case EDITDelete
                sSQL = sSQL & " AND x." & gk_sDBCOLEditSetEntryTypeID & "=" & cesAddedFromLive
            Case Else
                RaiseError sSource, PEGeneralInternalError, "Invalid editset entry type"
        End Select
    End If
    
    GetDivisionSql = sSQL
    
    Exit Function
    
ErrorHandler:
    ConvertError sSource
    
End Function


Private Function GetClassSql(t_eCatStructure As ECatStructure, _
            Optional t_eEntryType As EEditSetEntryType, _
            Optional t_bFromLive As Boolean = False, _
            Optional t_lLettercodeID As Long, _
            Optional t_sLettercodeIn As String, _
            Optional t_bGetExtraExtnData As Boolean = True) As String
'   Generate the sql to select classes from either the edit set
'   or live tables.
'   The parameters define the following:
'   t_eCatStructure     The structure of catalogu entry we are selecting for
'                       Eg, classes with no divisions (CATPlain)
'   t_eEntryType        For selecting from edit sets, the edit set entry type
'   t_bFromLive         Flag to indicate we are selecting from live tables
'                       (default is False (ie, from edit sets)
'   t_sLettercode       For selecting from live, the lettercode we
'                       select from
'   d2  WDP     25-Oct-1999      Combined GetClassSql and
'                                GetClassSqlFromLIVE to form
'                                Generic method.
'   d5  WDP     19-Nov-1999      Removed collection type parameter
'                                we now select all data
'   d7  WDP     30-Nov-1999      Use Lettercode id col
'   d8  WDP     13-Dec-1999     Allowed lettercode id to be 0
'                               If it is then all classes are selected
'                               Added lettercode in list to allow
'                               selection from multiple lettercodes
'   d9  WDP     14-Feb-2000     Select non-modified entries that have been accessioned or created
'                               when generating new editset data
'   d12 WDP     14-Mar-2000     Altered SQL generation to only select live records
'                                   when selecting from LIVE (live_flag = 0 or 1)
'   d18 ALM     25-Oct-2001     Supports a new type of entry - delete
'
On Error GoTo ErrorHandler

    Dim sSource As String
    sSource = "BOCatalogueIndexer.GetClassSql"
    
    Dim sSQL As String
    'We select NULL values for the fields which do not exist at
    'this level to maintain the cammon shape of the record set
    sSQL = "SELECT " & _
            "3 AS " & gk_sDICatalogueLevel & ", " & _
            "c." & gk_sDBCOLClassID & " AS " & gk_sDIID & ", " & _
            "lc." & gk_sDBCOLLetterCodeCode & " AS " & gk_sDILetterCodeRef & ", "
    
    If t_eCatStructure = CATPlain Then
        sSQL = sSQL & " NULL "
    Else
        sSQL = sSQL & " d." & gk_sDBCOLDivisionNo
    End If
    
    sSQL = sSQL & " AS " & gk_sDIDivisionRef & ", " & _
            "c." & gk_sDBCOLClassNumber & " AS " & gk_sDIClassRef & ", " & _
            "c." & gk_sDBCOLSubClassNumber & " AS " & gk_sDISubclassRef & ", " & _
            " NULL AS " & gk_sDIHeaderRef & ", " & _
            " NULL AS " & gk_sDISubheaderRef & ", " & _
            " NULL AS " & gk_sDIPieceRef & ", " & _
            " NULL AS " & gk_sDIItemRef & ", " & _
            " NULL AS " & gk_sDIPieceKeyOrder & ", " & _
            " NULL AS " & gk_sDIItemKeyOrder & ", " & _
            "c." & gk_sDBCOLFirstDate & " AS " & gk_sDIFirstDate & ", " & _
            "c." & gk_sDBCOLLastDate & " AS " & gk_sDILastDate & ", " & _
            "c." & gk_sDBCOLDateText & " AS " & gk_sDICoveringDates & ", " & _
            "x." & gk_sDBCOLMapDesignation & " AS " & gk_sDIMapDesignation & ", " & _
            "x." & gk_sDBCOLFormerDepartmentReference & " AS " & gk_sDIFormerRef & ", " & _
            "x." & gk_sDBCOLFormerPROReference & " AS " & gk_sDIFormerPRORef
    
    'Extra fields for advanced collections
    If t_bGetExtraExtnData = True Then
        sSQL = sSQL & ", " & _
             "x." & gk_sDBCOLMapScaleNumber & " AS " & gk_sDIMapScale & ", " & _
             "x." & gk_sDBCOLRestrictionsOnUse & " AS " & gk_sDIRestrictions & ", " & _
             "x." & gk_sDBCOLAccumulationdates & " AS " & gk_sDIAccumulationDates & ", " & _
             "x." & gk_sDBCOLRecordOpeningDate & " AS " & gk_sDIOpeningDates & ", " & _
             "x." & gk_sDBCOLLegalStatusID & " AS " & gk_sDILegalStatus & ", " & _
             "x." & gk_sDBCOLPhysicalDescriptionFormID & " AS " & gk_sDIPhysicalDescForm & ", " & _
             "x." & gk_sDBCOLPhysicalDescriptionRecordQuantity & " AS " & gk_sDIPhysDescQuantity & ", " & _
             "x." & gk_sDBCOLPhysicalDescriptionDimension & " AS " & gk_sDIPhysDescDimensions & ", " & _
             "x." & gk_sDBCOLAccessConditionID & " AS " & gk_sDIAccessCondition & ", " & _
             "x." & gk_sDBCOLLanguageID & " AS " & gk_sDILanguage & ", " & _
             "x." & gk_sDBCOLPhysicalCondition & " AS " & gk_sDIPhysicalCondition & ", " & _
             " NULL AS " & gk_sDIClosureCode & ", " & _
             " NULL AS " & gk_sDIClosureType & ", " & _
             "c." & gk_sDBCOLClosureStatus & " AS " & gk_sDIClosureStatus
    End If
    
    'The tables we are selecting from
    sSQL = sSQL & " FROM "
    If t_bFromLive = False Then
        sSQL = sSQL & _
            gk_sDBTABLEEditsetLettercode & " lc, " & _
            gk_sDBTABLEEditsetClass & " c, " & _
            gk_sDBTABLEEditsetClassExtension & " x "
        If Not t_eCatStructure = CATPlain Then
            sSQL = sSQL & ", " & gk_sDBTABLEEditsetDivision & " d "
        End If
    Else
        sSQL = sSQL & _
            gk_sDBTABLELCode & " lc, " & _
            gk_sDBTABLEClass & " c, " & _
            gk_sDBTableClassExtension & " x "
        If Not t_eCatStructure = CATPlain Then
            sSQL = sSQL & ", " & gk_sDBTABLEDivision & " d "
        End If
    End If
    
    'The joining clauses
    sSQL = sSQL & _
            " WHERE " & _
            "c." & gk_sDBCOLLetterCodeID & "=lc." & gk_sDBCOLLetterCodeID & " AND " & _
            "c." & gk_sDBCOLClassID & "=x." & gk_sDBCOLClassID
            
    If t_eCatStructure = CATPlain Then
        sSQL = sSQL & " AND (c." & gk_sDBCOLDivisionID & " IS NULL OR c." & _
                        gk_sDBCOLDivisionID & "=0)"
    Else
        sSQL = sSQL & " AND (c." & gk_sDBCOLDivisionID & " IS NOT NULL AND c." & _
                        gk_sDBCOLDivisionID & "<>0)" & _
                        " AND c." & gk_sDBCOLDivisionID & "=d." & gk_sDBCOLDivisionID
                        
    End If
    
    If t_bFromLive = False Then
        sSQL = sSQL & " AND lc." & gk_sDBCOLEditSetID & "=" & m_lEditsetID & _
                  " AND c." & gk_sDBCOLEditSetID & "=" & m_lEditsetID & _
                  " AND x." & gk_sDBCOLEditSetID & "=" & m_lEditsetID
    
        If Not t_eCatStructure = CATPlain Then
            sSQL = sSQL & " AND d." & gk_sDBCOLEditSetID & "=" & m_lEditsetID
        End If
    Else
        If Not t_lLettercodeID = 0 Then
            If t_sLettercodeIn = "" Then
                sSQL = sSQL & " AND lc." & gk_sDBCOLLetterCodeID & "=" & t_lLettercodeID
            End If
        Else
            If Not t_sLettercodeIn = "" Then
                sSQL = sSQL & _
                    " AND lc." & gk_sDBCOLLetterCodeID & " IN " & t_sLettercodeIn
            End If
        End If
        sSQL = sSQL & " AND x." & gk_sDBCOLLiveFlag & " IN (0,1) "
    End If
    
    If t_bFromLive = False Then
        Select Case t_eEntryType
            Case EDITNew
                sSQL = sSQL & " AND (x." & gk_sDBCOLEditSetEntryTypeID & " IN (" & _
                    cesImportedModified & ", " & cesPROCATCreatedModified & ", " & _
                    cesPROCATCreated & ", " & cesImported & ")) "
            Case EDITUpdated
                sSQL = sSQL & " AND x." & gk_sDBCOLEditSetEntryTypeID & "=" & cesAddedFromLiveModified
            Case EDITDelete
                sSQL = sSQL & " AND x." & gk_sDBCOLEditSetEntryTypeID & "=" & cesAddedFromLive
            Case Else
                RaiseError sSource, PEGeneralInternalError, "Invalid editset entry type"
        End Select
    End If
    
    GetClassSql = sSQL
    
    Exit Function
    
ErrorHandler:
    ConvertError sSource
    
End Function


Private Function GetHeaderSql(t_eCatStructure As ECatStructure, _
            Optional t_eEntryType As EEditSetEntryType, _
            Optional t_bFromLive As Boolean = False, _
            Optional t_lLettercodeID As Long, _
            Optional t_lClassID As Long = 0, _
            Optional t_sLettercodeIn As String, _
            Optional t_bGetExtraExtnData As Boolean = True) As String
'   Generate the sql to select headers from either the edit set
'   or live tables.
'   The parameters define the following:
'   t_eCatStructure     The structure of catalogue entry we are selecting for
'                       Eg, headers with no divisions (CATPlain)
'   t_eEntryType        For selecting from edit sets, the edit set entry type
'   t_bFromLive         Flag to indicate we are selecting from live tables
'                       (default is False (ie, from edit sets)
'   t_sLettercode       For selecting from live, the lettercode we
'                       select from
'   d2  WDP     25-Oct-1999      Combined GetHeaderSql and
'                                GetHeaderSqlFromLIVE to form
'                                Generic method.
'   d5  WDP     19-Nov-1999      Removed collection type parameter
'                                we now select all data
'   d7  WDP     30-Nov-1999      Use Lettercode id col and added class id param
'   d8  WDP     13-Dec-1999     Allowed lettercode id to be 0
'                               If it is then all headers are selected
'                               Added lettercode in list to allow
'                               selection from multiple lettercodes
'   d9  WDP     14-Feb-2000     Select non-modified entries that have been accessioned or created
'                               when generating new editset data
'   d12 WDP     14-Mar-2000     Altered SQL generation to only select live records
'                                   when selecting from LIVE (live_flag = 0 or 1)
'   d18 ALM     25-Oct-2001     Supports a new type of entry - delete
'
On Error GoTo ErrorHandler

    Dim sSource As String
    sSource = "BOCatalogueIndexer.GetHeaderSql"
    
    Dim sSQL As String
    sSQL = "SELECT " & _
            "4 AS " & gk_sDICatalogueLevel & ", " & _
            "h." & gk_sDBCOLHeaderID & " AS " & gk_sDIID & ", " & _
            "lc." & gk_sDBCOLLetterCodeCode & " AS " & gk_sDILetterCodeRef & ", "
    
    If t_eCatStructure = CATPlain Then
        sSQL = sSQL & " NULL "
    Else
        sSQL = sSQL & " d." & gk_sDBCOLDivisionNo
    End If
    
    sSQL = sSQL & " AS " & gk_sDIDivisionRef & ", " & _
            "c." & gk_sDBCOLClassNumber & " AS " & gk_sDIClassRef & ", " & _
            "c." & gk_sDBCOLSubClassNumber & " AS " & gk_sDISubclassRef & ", " & _
            "h." & gk_sDBCOLClassHeaderNo & " AS " & gk_sDIHeaderRef & ", " & _
            " NULL AS " & gk_sDISubheaderRef & ", " & _
            " NULL AS " & gk_sDIPieceRef & ", " & _
            " NULL AS " & gk_sDIItemRef & ", " & _
            " NULL AS " & gk_sDIPieceKeyOrder & ", " & _
            " NULL AS " & gk_sDIItemKeyOrder & ", " & _
            "x." & gk_sDBCOLCoveringFromDate & " AS " & gk_sDIFirstDate & ", " & _
            "x." & gk_sDBCOLCoveringToDate & " AS " & gk_sDILastDate & ", " & _
            "h." & gk_sDBCOLDateText & " AS " & gk_sDICoveringDates & ", " & _
            "x." & gk_sDBCOLMapDesignation & " AS " & gk_sDIMapDesignation & ", " & _
            "x." & gk_sDBCOLFormerDepartmentReference & " AS " & gk_sDIFormerRef & ", " & _
            "x." & gk_sDBCOLFormerPROReference & " AS " & gk_sDIFormerPRORef
            
    'Extra fields for advanced collections
    If t_bGetExtraExtnData = True Then
        sSQL = sSQL & ", " & _
             "x." & gk_sDBCOLMapScaleNumber & " AS " & gk_sDIMapScale & ", " & _
             "x." & gk_sDBCOLRestrictionsOnUse & " AS " & gk_sDIRestrictions & ", " & _
             "x." & gk_sDBCOLAccumulationdates & " AS " & gk_sDIAccumulationDates & ", " & _
             "x." & gk_sDBCOLRecordOpeningDate & " AS " & gk_sDIOpeningDates & ", " & _
             "x." & gk_sDBCOLLegalStatusID & " AS " & gk_sDILegalStatus & ", " & _
             "x." & gk_sDBCOLPhysicalDescriptionFormID & " AS " & gk_sDIPhysicalDescForm & ", " & _
             "x." & gk_sDBCOLPhysicalDescriptionRecordQuantity & " AS " & gk_sDIPhysDescQuantity & ", " & _
             "x." & gk_sDBCOLPhysicalDescriptionDimension & " AS " & gk_sDIPhysDescDimensions & ", " & _
             "x." & gk_sDBCOLAccessConditionID & " AS " & gk_sDIAccessCondition & ", " & _
             "x." & gk_sDBCOLLanguageID & " AS " & gk_sDILanguage & ", " & _
             "x." & gk_sDBCOLPhysicalCondition & " AS " & gk_sDIPhysicalCondition & ", " & _
             " NULL AS " & gk_sDIClosureCode & ", " & _
             " NULL AS " & gk_sDIClosureType & ", " & _
             "NULL AS " & gk_sDIClosureStatus
    End If
    
    'The tables we are selecting from
    sSQL = sSQL & " FROM "
    If t_bFromLive = False Then
        sSQL = sSQL & _
            gk_sDBTABLEEditsetLettercode & " lc, " & _
            gk_sDBTABLEEditsetClass & " c, " & _
            gk_sDBTABLEEditsetHeader & " h, " & _
            gk_sDBTABLEEditsetHeaderExtension & " x "
        If Not t_eCatStructure = CATPlain Then
            sSQL = sSQL & ", " & gk_sDBTABLEEditsetDivision & " d "
        End If
    Else
        sSQL = sSQL & _
            gk_sDBTABLELCode & " lc, " & _
            gk_sDBTABLEClass & " c, " & _
            gk_sDBTABLEHeader & " h, " & _
            gk_sDBTableHeaderExtension & " x "
        If Not t_eCatStructure = CATPlain Then
            sSQL = sSQL & ", " & gk_sDBTABLEDivision & " d "
        End If
    End If
    
    'The joining clauses
    sSQL = sSQL & _
            " WHERE " & _
            "h." & gk_sDBCOLClassID & "=c." & gk_sDBCOLClassID & " AND " & _
            "c." & gk_sDBCOLLetterCodeID & "=lc." & gk_sDBCOLLetterCodeID & " AND " & _
            "h." & gk_sDBCOLHeaderID & "=x." & gk_sDBCOLHeaderID
            
    If t_eCatStructure = CATPlain Then
        sSQL = sSQL & " AND (c." & gk_sDBCOLDivisionID & " IS NULL OR c." & _
                        gk_sDBCOLDivisionID & "=0)"
    Else
        sSQL = sSQL & " AND (c." & gk_sDBCOLDivisionID & " IS NOT NULL AND c." & _
                        gk_sDBCOLDivisionID & "<>0)" & _
                        " AND c." & gk_sDBCOLDivisionID & "=d." & gk_sDBCOLDivisionID
    End If
    
    If t_bFromLive = False Then
        sSQL = sSQL & " AND lc." & gk_sDBCOLEditSetID & "=" & m_lEditsetID & _
                  " AND c." & gk_sDBCOLEditSetID & "=" & m_lEditsetID & _
                  " AND h." & gk_sDBCOLEditSetID & "=" & m_lEditsetID & _
                  " AND x." & gk_sDBCOLEditSetID & "=" & m_lEditsetID
    
        If Not t_eCatStructure = CATPlain Then
            sSQL = sSQL & " AND d." & gk_sDBCOLEditSetID & "=" & m_lEditsetID
        End If
    Else
        If Not t_lLettercodeID = 0 Then
            If t_sLettercodeIn = "" Then
                sSQL = sSQL & " AND lc." & gk_sDBCOLLetterCodeID & "=" & t_lLettercodeID
            End If
        Else
            If Not t_sLettercodeIn = "" Then
                sSQL = sSQL & _
                    " AND lc." & gk_sDBCOLLetterCodeID & " IN " & t_sLettercodeIn
            End If
        End If
        If Not t_lClassID = 0 Then
            sSQL = sSQL & " AND c." & gk_sDBCOLClassID & "=" & t_lClassID
        End If
        sSQL = sSQL & " AND x." & gk_sDBCOLLiveFlag & " IN (0,1) "
    End If
    
    If t_bFromLive = False Then
        Select Case t_eEntryType
            Case EDITNew
                sSQL = sSQL & " AND (x." & gk_sDBCOLEditSetEntryTypeID & " IN (" & _
                    cesImportedModified & ", " & cesPROCATCreatedModified & ", " & _
                    cesPROCATCreated & ", " & cesImported & ")) "
            Case EDITUpdated
                sSQL = sSQL & " AND x." & gk_sDBCOLEditSetEntryTypeID & "=" & cesAddedFromLiveModified
            Case EDITDelete 'd18
                sSQL = sSQL & " AND x." & gk_sDBCOLEditSetEntryTypeID & "=" & cesAddedFromLive
            Case Else
                RaiseError sSource, PEGeneralInternalError, "Invalid editset entry type"
        End Select
    End If
    
    GetHeaderSql = sSQL
    
    Exit Function
    
ErrorHandler:
    ConvertError sSource
    
End Function

Private Function GetSubheaderSql(t_eCatStructure As ECatStructure, _
            Optional t_eEntryType As EEditSetEntryType, _
            Optional t_bFromLive As Boolean = False, _
            Optional t_lLettercodeID As Long, _
            Optional t_lClassID As Long = 0, _
            Optional t_sLettercodeIn As String, _
            Optional t_bGetExtraExtnData As Boolean = True) As String
'   Generate the sql to select subheaders from either the edit set
'   or live tables.
'   The parameters define the following:
'   t_eCatStructure     The structure of catalogue entry we are selecting for
'                       Eg, subheaders with no divisions (CATPlain)
'   t_eEntryType        For selecting from edit sets, the edit set entry type
'   t_bFromLive         Flag to indicate we are selecting from live tables
'                       (default is False (ie, from edit sets)
'   t_sLettercode       For selecting from live, the lettercode we
'                       select from
'   d2  WDP     25-Oct-1999      Combined GetSubheaderSql and
'                                GetSubheaderSqlFromLIVE to form
'                                Generic method.
'   d5  WDP     19-Nov-1999      Removed collection type parameter
'                                we now select all data
'   d7  WDP     30-Nov-1999      Use lettercode id col
'   d8  WDP     13-Dec-1999     Allowed lettercode id to be 0
'                               If it is then all subheaders are selected
'                               Added lettercode in list to allow
'                               selection from multiple lettercodes
'   d9  WDP     14-Feb-2000     Select non-modified entries that have been accessioned or created
'                               when generating new editset data
'   d12 WDP     14-Mar-2000     Altered SQL generation to only select live records
'                                   when selecting from LIVE (live_flag = 0 or 1)
'   d18 ALM     25-Oct-2001     Supports a new type of entry - delete
'
On Error GoTo ErrorHandler

    Dim sSource As String
    sSource = "BOCatalogueIndexer.GetSubheaderSql"
    
    Dim sSQL As String
    sSQL = "SELECT " & _
            "5 AS " & gk_sDICatalogueLevel & ", " & _
            "sh." & gk_sDBCOLSubHeaderID & " AS " & gk_sDIID & ", " & _
            "lc." & gk_sDBCOLLetterCodeCode & " AS " & gk_sDILetterCodeRef & ", "

    If t_eCatStructure = CATPlain Then
        sSQL = sSQL & " NULL "
    Else
        sSQL = sSQL & " d." & gk_sDBCOLDivisionNo
    End If

    sSQL = sSQL & " AS " & gk_sDIDivisionRef & ", " & _
            "c." & gk_sDBCOLClassNumber & " AS " & gk_sDIClassRef & ", " & _
            "c." & gk_sDBCOLSubClassNumber & " AS " & gk_sDISubclassRef & ", " & _
            "h." & gk_sDBCOLClassHeaderNo & " AS " & gk_sDIHeaderRef & ", " & _
            "sh." & gk_sDBCOLSubHeaderNo & " AS " & gk_sDISubheaderRef & ", " & _
            " NULL AS " & gk_sDIPieceRef & ", " & _
            " NULL AS " & gk_sDIItemRef & ", " & _
            " NULL AS " & gk_sDIPieceKeyOrder & ", " & _
            " NULL AS " & gk_sDIItemKeyOrder & ", " & _
            "x." & gk_sDBCOLCoveringFromDate & " AS " & gk_sDIFirstDate & ", " & _
            "x." & gk_sDBCOLCoveringToDate & " AS " & gk_sDILastDate & ", " & _
            "sh." & gk_sDBCOLDateText & " AS " & gk_sDICoveringDates & ", " & _
            "x." & gk_sDBCOLMapDesignation & " AS " & gk_sDIMapDesignation & ", " & _
            "x." & gk_sDBCOLFormerDepartmentReference & " AS " & gk_sDIFormerRef & ", " & _
            "x." & gk_sDBCOLFormerPROReference & " AS " & gk_sDIFormerPRORef
    
    'Extra fields for advanced collections
    If t_bGetExtraExtnData = True Then
        sSQL = sSQL & ", " & _
             "x." & gk_sDBCOLMapScaleNumber & " AS " & gk_sDIMapScale & ", " & _
             "x." & gk_sDBCOLRestrictionsOnUse & " AS " & gk_sDIRestrictions & ", " & _
             "x." & gk_sDBCOLAccumulationdates & " AS " & gk_sDIAccumulationDates & ", " & _
             "x." & gk_sDBCOLRecordOpeningDate & " AS " & gk_sDIOpeningDates & ", " & _
             "x." & gk_sDBCOLLegalStatusID & " AS " & gk_sDILegalStatus & ", " & _
             "x." & gk_sDBCOLPhysicalDescriptionFormID & " AS " & gk_sDIPhysicalDescForm & ", " & _
             "x." & gk_sDBCOLPhysicalDescriptionRecordQuantity & " AS " & gk_sDIPhysDescQuantity & ", " & _
             "x." & gk_sDBCOLPhysicalDescriptionDimension & " AS " & gk_sDIPhysDescDimensions & ", " & _
             "x." & gk_sDBCOLAccessConditionID & " AS " & gk_sDIAccessCondition & ", " & _
             "x." & gk_sDBCOLLanguageID & " AS " & gk_sDILanguage & ", " & _
             "x." & gk_sDBCOLPhysicalCondition & " AS " & gk_sDIPhysicalCondition & ", " & _
             " NULL AS " & gk_sDIClosureCode & ", " & _
             " NULL AS " & gk_sDIClosureType & ", " & _
             "NULL AS " & gk_sDIClosureStatus
    End If
    
    'The tables we are selecting from
    sSQL = sSQL & " FROM "
    If t_bFromLive = False Then
        sSQL = sSQL & _
            gk_sDBTABLEEditsetLettercode & " lc, " & _
            gk_sDBTABLEEditsetClass & " c, " & _
            gk_sDBTABLEEditsetHeader & " h, " & _
            gk_sDBTABLEEditsetSubheader & " sh, " & _
            gk_sDBTABLEEditsetSubheaderExtension & " x "
        If Not t_eCatStructure = CATPlain Then
            sSQL = sSQL & ", " & gk_sDBTABLEEditsetDivision & " d "
        End If
    Else
        sSQL = sSQL & _
            gk_sDBTABLELCode & " lc, " & _
            gk_sDBTABLEClass & " c, " & _
            gk_sDBTABLEHeader & " h, " & _
            gk_sDBTABLESubHeader & " sh, " & _
            gk_sDBTableSubHeaderExtension & " x "
        If Not t_eCatStructure = CATPlain Then
            sSQL = sSQL & ", " & gk_sDBTABLEDivision & " d "
        End If
    End If
    
    'The joining clauses
    sSQL = sSQL & _
            " WHERE " & _
            "sh." & gk_sDBCOLHeaderID & "=h." & gk_sDBCOLHeaderID & " AND " & _
            "h." & gk_sDBCOLClassID & "=c." & gk_sDBCOLClassID & " AND " & _
            "c." & gk_sDBCOLLetterCodeID & "=lc." & gk_sDBCOLLetterCodeID & " AND " & _
            "sh." & gk_sDBCOLSubHeaderID & "=x." & gk_sDBCOLSubHeaderID
    
    If t_eCatStructure = CATPlain Then
        sSQL = sSQL & " AND (c." & gk_sDBCOLDivisionID & " IS NULL OR c." & _
                        gk_sDBCOLDivisionID & "=0)"
    Else
        sSQL = sSQL & " AND (c." & gk_sDBCOLDivisionID & " IS NOT NULL AND c." & _
                        gk_sDBCOLDivisionID & "<>0)" & _
                        " AND c." & gk_sDBCOLDivisionID & "=d." & gk_sDBCOLDivisionID
    End If
    
    If t_bFromLive = False Then
        sSQL = sSQL & " AND lc." & gk_sDBCOLEditSetID & "=" & m_lEditsetID & _
                  " AND c." & gk_sDBCOLEditSetID & "=" & m_lEditsetID & _
                  " AND h." & gk_sDBCOLEditSetID & "=" & m_lEditsetID & _
                  " AND sh." & gk_sDBCOLEditSetID & "=" & m_lEditsetID & _
                  " AND x." & gk_sDBCOLEditSetID & "=" & m_lEditsetID
    
        If Not t_eCatStructure = CATPlain Then
            sSQL = sSQL & " AND d." & gk_sDBCOLEditSetID & "=" & m_lEditsetID
        End If
    Else
        If Not t_lLettercodeID = 0 Then
            If t_sLettercodeIn = "" Then
                sSQL = sSQL & " AND lc." & gk_sDBCOLLetterCodeID & "=" & t_lLettercodeID
            End If
        Else
            If Not t_sLettercodeIn = "" Then
                sSQL = sSQL & _
                    " AND lc." & gk_sDBCOLLetterCodeID & " IN " & t_sLettercodeIn
            End If
        End If
        If Not t_lClassID = 0 Then
            sSQL = sSQL & " AND c." & gk_sDBCOLClassID & "=" & t_lClassID
        End If
        sSQL = sSQL & " AND x." & gk_sDBCOLLiveFlag & " IN (0,1) "
    End If
    
    If t_bFromLive = False Then
        Select Case t_eEntryType
            Case EDITNew
                sSQL = sSQL & " AND (x." & gk_sDBCOLEditSetEntryTypeID & " IN (" & _
                    cesImportedModified & ", " & cesPROCATCreatedModified & ", " & _
                    cesPROCATCreated & ", " & cesImported & ")) "
            Case EDITUpdated
                sSQL = sSQL & " AND x." & gk_sDBCOLEditSetEntryTypeID & "=" & cesAddedFromLiveModified
            Case EDITDelete 'd18
                sSQL = sSQL & " AND x." & gk_sDBCOLEditSetEntryTypeID & "=" & cesAddedFromLive
            Case Else
                RaiseError sSource, PEGeneralInternalError, "Invalid editset entry type"
        End Select
    End If
    
    GetSubheaderSql = sSQL
    
    Exit Function
    
ErrorHandler:
    ConvertError sSource
    
End Function


Private Function GetPieceSql(t_eCatStructure As ECatStructure, _
            Optional t_eEntryType As EEditSetEntryType, _
            Optional t_bFromLive As Boolean = False, _
            Optional t_lLettercodeID As Long, _
            Optional t_lClassID As Long = 0, _
            Optional t_lHeaderID As Long = 0, _
            Optional t_sClassIn As String = "", _
            Optional t_sLettercodeIn As String = "", _
            Optional t_bGetExtraExtnData As Boolean = True) As String
'   Generate the sql to select pieces from either the edit set
'   or live tables.
'   The parameters define the following:
'   t_eCatStructure     The structure of catalogue entry we are selecting for
'                       Eg, pieces with no headers but divisions(CATWithDiv)
'   t_eEntryType        For selecting from edit sets, the edit set entry type
'   t_bFromLive         Flag to indicate we are selecting from live tables
'                       (default is False (ie, from edit sets)
'   t_sLettercode       For selecting from live, the lettercode we
'                       select from
'   t_sClassRef         For selecting from live, the class we
'                       select from (optional)
'   t_sHeaderRef        For selecting from live, the header we
'                       select from (optional)
'
'   d2  WDP     25-Oct-1999      Combined GetPieceSql and
'                                GetPieceSqlFromLIVE to form
'                                Generic method.
'   d3  WDP     29-Oct-1999      Added class and header optional params
'   d5  WDP     19-Nov-1999      Removed collection type parameter
'                                we now select all data
'   d7  WDP     30-Nov-1999      Use id cols instead of ref parts
'   d9  WDP     14-Feb-2000     Select non-modified entries that have been accessioned or created
'                               when generating new editset data
'   d12 WDP     14-Mar-2000     Altered SQL generation to only select live records
'                                   when selecting from LIVE (live_flag = 0 or 1)
'   d18 ALM     25-Oct-2001     Supports a new type of entry - delete
'
On Error GoTo ErrorHandler

    Dim sSource As String
    sSource = "BOCatalogueIndexer.GetPieceSql"
    
    Dim sSQL As String
    sSQL = "SELECT " & _
            "6 AS " & gk_sDICatalogueLevel & ", " & _
            "p." & gk_sDBCOLPieceID & " AS " & gk_sDIID & ", " & _
            "lc." & gk_sDBCOLLetterCodeCode & " AS " & gk_sDILetterCodeRef & ", "
            
    Select Case t_eCatStructure
        Case CATPlain, CATWithHeader, CATWithSubheader
            sSQL = sSQL & " NULL "
        Case Else
            sSQL = sSQL & " d." & gk_sDBCOLDivisionNo
    End Select

    sSQL = sSQL & " AS " & gk_sDIDivisionRef & ", " & _
            "c." & gk_sDBCOLClassNumber & " AS " & gk_sDIClassRef & ", " & _
            "c." & gk_sDBCOLSubClassNumber & " AS " & gk_sDISubclassRef & ", "
    
    Select Case t_eCatStructure
    Case CATPlain, CATWithDiv
        sSQL = sSQL & " NULL "
    Case Else
        sSQL = sSQL & " h." & gk_sDBCOLClassHeaderNo
    End Select

    sSQL = sSQL & " AS " & gk_sDIHeaderRef & ", "
    
    Select Case t_eCatStructure
    Case CATPlain, CATWithDiv, CATWithHeader, CATWithHeaderAndDiv
        sSQL = sSQL & " NULL "
    Case Else
        sSQL = sSQL & " sh." & gk_sDBCOLSubHeaderNo
    End Select

    sSQL = sSQL & " AS " & gk_sDISubheaderRef & ", " & _
            "p." & gk_sDBCOLPieceRef & " AS " & gk_sDIPieceRef & ", " & _
            " NULL AS " & gk_sDIItemRef & ", " & _
            "p." & gk_sDBColPiecekeyOrder & " AS " & gk_sDIPieceKeyOrder & ", " & _
            " NULL AS " & gk_sDIItemKeyOrder & ", " & _
            "p." & gk_sDBCOLFirstDate & " AS " & gk_sDIFirstDate & ", " & _
            "p." & gk_sDBCOLLastDate & " AS " & gk_sDILastDate & ", " & _
            "p." & gk_sDBCOLDateText & " AS " & gk_sDICoveringDates & ", " & _
            "x." & gk_sDBCOLMapDesignation & " AS " & gk_sDIMapDesignation & ", " & _
            "x." & gk_sDBCOLFormerDepartmentReference & " AS " & gk_sDIFormerRef & ", " & _
            "x." & gk_sDBCOLFormerPROReference & " AS " & gk_sDIFormerPRORef
    'Extra fields for advanced collections
    
    If t_bGetExtraExtnData = True Then
        sSQL = sSQL & ", " & _
             "x." & gk_sDBCOLMapScaleNumber & " AS " & gk_sDIMapScale & ", " & _
             "x." & gk_sDBCOLRestrictionsOnUse & " AS " & gk_sDIRestrictions & ", " & _
             "x." & gk_sDBCOLAccumulationdates & " AS " & gk_sDIAccumulationDates & ", " & _
             "p." & gk_sDBCOLPieceOpenDate & " AS " & gk_sDIOpeningDates & ", " & _
             "x." & gk_sDBCOLLegalStatusID & " AS " & gk_sDILegalStatus & ", " & _
             "x." & gk_sDBCOLPhysicalDescriptionFormID & " AS " & gk_sDIPhysicalDescForm & ", " & _
             "x." & gk_sDBCOLPhysicalDescriptionRecordQuantity & " AS " & gk_sDIPhysDescQuantity & ", " & _
             "x." & gk_sDBCOLPhysicalDescriptionDimension & " AS " & gk_sDIPhysDescDimensions & ", " & _
             "x." & gk_sDBCOLAccessConditionID & " AS " & gk_sDIAccessCondition & ", " & _
             "p." & gk_sDBCOLLanguageID & " AS " & gk_sDILanguage & ", " & _
             "x." & gk_sDBCOLPhysicalCondition & " AS " & gk_sDIPhysicalCondition & ", " & _
             "p." & gk_sDBCOLClosureCode & " AS " & gk_sDIClosureCode & ", " & _
             "p." & gk_sDBCOLClosureType & " AS " & gk_sDIClosureType & ", " & _
             "p." & gk_sDBCOLClosureStatus & " AS " & gk_sDIClosureStatus
    End If
    
    'The tables we are selecting from
    sSQL = sSQL & " FROM "
    
    If t_bFromLive = False Then
        sSQL = sSQL & _
            gk_sDBTABLEEditsetLettercode & " lc, " & _
            gk_sDBTABLEEditsetClass & " c, " & _
            gk_sDBTABLEEditsetPiece & " p, " & _
            gk_sDBTABLEEditsetPieceExtension & " x "
        Select Case t_eCatStructure
            Case CATWithDiv
                sSQL = sSQL & ", " & gk_sDBTABLEEditsetDivision & " d "
            Case CATWithHeader
                sSQL = sSQL & ", " & gk_sDBTABLEEditsetHeader & " h "
            Case CATWithHeaderAndDiv
                sSQL = sSQL & ", " & gk_sDBTABLEEditsetDivision & " d "
                sSQL = sSQL & ", " & gk_sDBTABLEEditsetHeader & " h "
            Case CATWithSubheaderAndDiv
                sSQL = sSQL & ", " & gk_sDBTABLEEditsetDivision & " d "
                sSQL = sSQL & ", " & gk_sDBTABLEEditsetHeader & " h "
                sSQL = sSQL & ", " & gk_sDBTABLEEditsetSubheader & " sh "
            Case CATWithSubheader
                sSQL = sSQL & ", " & gk_sDBTABLEEditsetHeader & " h "
                sSQL = sSQL & ", " & gk_sDBTABLEEditsetSubheader & " sh "
            Case CATPlain
        End Select
    Else
        sSQL = sSQL & _
            gk_sDBTABLELCode & " lc, " & _
            gk_sDBTABLEClass & " c, " & _
            gk_sDBTABLEPiece & " p, " & _
            gk_sDBTablePieceExtension & " x "
        Select Case t_eCatStructure
            Case CATWithDiv
                sSQL = sSQL & ", " & gk_sDBTABLEDivision & " d "
            Case CATWithHeader
                sSQL = sSQL & ", " & gk_sDBTABLEHeader & " h "
            Case CATWithHeaderAndDiv
                sSQL = sSQL & ", " & gk_sDBTABLEDivision & " d "
                sSQL = sSQL & ", " & gk_sDBTABLEHeader & " h "
            Case CATWithSubheaderAndDiv
                sSQL = sSQL & ", " & gk_sDBTABLEDivision & " d "
                sSQL = sSQL & ", " & gk_sDBTABLEHeader & " h "
                sSQL = sSQL & ", " & gk_sDBTABLESubHeader & " sh "
            Case CATWithSubheader
                sSQL = sSQL & ", " & gk_sDBTABLEHeader & " h "
                sSQL = sSQL & ", " & gk_sDBTABLESubHeader & " sh "
            Case CATPlain
        End Select
    End If
    
    'The joining clauses
    sSQL = sSQL & _
            " WHERE " & _
            "p." & gk_sDBCOLClassID & "=c." & gk_sDBCOLClassID & " AND " & _
            "c." & gk_sDBCOLLetterCodeID & "=lc." & gk_sDBCOLLetterCodeID & " AND " & _
            "p." & gk_sDBCOLPieceID & "=x." & gk_sDBCOLPieceID
    
    Select Case t_eCatStructure
        Case CATWithDiv
            sSQL = sSQL & " AND (c." & gk_sDBCOLDivisionID & " IS NOT NULL AND c." & _
                        gk_sDBCOLDivisionID & "<>0)" & _
                        " AND c." & gk_sDBCOLDivisionID & "=d." & gk_sDBCOLDivisionID
            sSQL = sSQL & " AND (p." & gk_sDBCOLHeaderID & " IS NULL OR p." & _
                        gk_sDBCOLHeaderID & "=0)"
            sSQL = sSQL & " AND (p." & gk_sDBCOLSubHeaderID & " IS NULL OR p." & _
                        gk_sDBCOLSubHeaderID & "=0)"
            
        Case CATWithHeaderAndDiv
            sSQL = sSQL & " AND (c." & gk_sDBCOLDivisionID & " IS NOT NULL AND c." & _
                        gk_sDBCOLDivisionID & "<>0)" & _
                        " AND c." & gk_sDBCOLDivisionID & "=d." & gk_sDBCOLDivisionID
            sSQL = sSQL & " AND (p." & gk_sDBCOLHeaderID & " IS NOT NULL AND p." & _
                        gk_sDBCOLHeaderID & "<>0)" & _
                        " AND p." & gk_sDBCOLHeaderID & "=h." & gk_sDBCOLHeaderID
            sSQL = sSQL & " AND (p." & gk_sDBCOLSubHeaderID & " IS NULL OR p." & _
                        gk_sDBCOLSubHeaderID & "=0)"
            
        Case CATWithSubheaderAndDiv
            sSQL = sSQL & " AND (c." & gk_sDBCOLDivisionID & " IS NOT NULL AND c." & _
                        gk_sDBCOLDivisionID & "<>0)" & _
                        " AND c." & gk_sDBCOLDivisionID & "=d." & gk_sDBCOLDivisionID
            sSQL = sSQL & " AND (p." & gk_sDBCOLHeaderID & " IS NOT NULL AND p." & _
                        gk_sDBCOLHeaderID & "<>0)" & _
                        " AND p." & gk_sDBCOLHeaderID & "=h." & gk_sDBCOLHeaderID
            sSQL = sSQL & " AND (p." & gk_sDBCOLSubHeaderID & " IS NOT NULL AND p." & _
                        gk_sDBCOLSubHeaderID & "<>0)" & _
                        " AND p." & gk_sDBCOLSubHeaderID & "=sh." & gk_sDBCOLSubHeaderID
                        
        Case CATWithHeader
            sSQL = sSQL & " AND (c." & gk_sDBCOLDivisionID & " IS NULL OR c." & _
                        gk_sDBCOLDivisionID & "=0)"
            sSQL = sSQL & " AND (p." & gk_sDBCOLHeaderID & " IS NOT NULL AND p." & _
                        gk_sDBCOLHeaderID & "<>0)" & _
                        " AND p." & gk_sDBCOLHeaderID & "=h." & gk_sDBCOLHeaderID
            sSQL = sSQL & " AND (p." & gk_sDBCOLSubHeaderID & " IS NULL OR p." & _
                        gk_sDBCOLSubHeaderID & "=0)"
            
        Case CATWithSubheader
            sSQL = sSQL & " AND (c." & gk_sDBCOLDivisionID & " IS NULL OR c." & _
                        gk_sDBCOLDivisionID & "=0)"
            sSQL = sSQL & " AND (p." & gk_sDBCOLHeaderID & " IS NOT NULL AND p." & _
                        gk_sDBCOLHeaderID & "<>0)" & _
                        " AND p." & gk_sDBCOLHeaderID & "=h." & gk_sDBCOLHeaderID
            sSQL = sSQL & " AND (p." & gk_sDBCOLSubHeaderID & " IS NOT NULL AND p." & _
                        gk_sDBCOLSubHeaderID & "<>0)" & _
                        " AND p." & gk_sDBCOLSubHeaderID & "=sh." & gk_sDBCOLSubHeaderID
                        
        Case CATPlain
            sSQL = sSQL & " AND (c." & gk_sDBCOLDivisionID & " IS NULL OR c." & _
                        gk_sDBCOLDivisionID & "=0)"
            sSQL = sSQL & " AND (p." & gk_sDBCOLHeaderID & " IS NULL OR p." & _
                        gk_sDBCOLHeaderID & "=0)"
            sSQL = sSQL & " AND (p." & gk_sDBCOLSubHeaderID & " IS NULL OR p." & _
                        gk_sDBCOLSubHeaderID & "=0)"
                        
    End Select
    
    If t_bFromLive = False Then
        sSQL = sSQL & " AND lc." & gk_sDBCOLEditSetID & "=" & m_lEditsetID & _
                  " AND c." & gk_sDBCOLEditSetID & "=" & m_lEditsetID & _
                  " AND p." & gk_sDBCOLEditSetID & "=" & m_lEditsetID & _
                  " AND x." & gk_sDBCOLEditSetID & "=" & m_lEditsetID
        Select Case t_eCatStructure
            Case CATWithDiv
                sSQL = sSQL & " AND d." & gk_sDBCOLEditSetID & "=" & m_lEditsetID
            Case CATWithHeaderAndDiv
                sSQL = sSQL & " AND d." & gk_sDBCOLEditSetID & "=" & m_lEditsetID
                sSQL = sSQL & " AND h." & gk_sDBCOLEditSetID & "=" & m_lEditsetID
            Case CATWithSubheaderAndDiv
                sSQL = sSQL & " AND d." & gk_sDBCOLEditSetID & "=" & m_lEditsetID
                sSQL = sSQL & " AND h." & gk_sDBCOLEditSetID & "=" & m_lEditsetID
                sSQL = sSQL & " AND sh." & gk_sDBCOLEditSetID & "=" & m_lEditsetID
            Case CATWithHeader
                sSQL = sSQL & " AND h." & gk_sDBCOLEditSetID & "=" & m_lEditsetID
            Case CATWithSubheader
                sSQL = sSQL & " AND h." & gk_sDBCOLEditSetID & "=" & m_lEditsetID
                sSQL = sSQL & " AND sh." & gk_sDBCOLEditSetID & "=" & m_lEditsetID
            Case CATPlain
        End Select
    Else
        If t_sLettercodeIn = "" Then
            sSQL = sSQL & " AND lc." & gk_sDBCOLLetterCodeID & "=" & t_lLettercodeID
        Else
            sSQL = sSQL & " AND lc." & gk_sDBCOLLetterCodeID & " IN " & t_sLettercodeIn
        End If
        
        If t_sClassIn = "" Then
            If Not t_lClassID = 0 Then
                sSQL = sSQL & " AND c." & gk_sDBCOLClassID & "=" & t_lClassID
            End If
        Else
            sSQL = sSQL & " AND c." & gk_sDBCOLClassID & " IN " & t_sClassIn
        End If
        If Not (t_lHeaderID = 0 Or t_lHeaderID = mk_lNullRef) Then
            sSQL = sSQL & " AND h." & gk_sDBCOLHeaderID & "=" & t_lHeaderID
        End If
        sSQL = sSQL & " AND x." & gk_sDBCOLLiveFlag & " IN (0,1) "
    End If
    
    If t_bFromLive = False Then
        Select Case t_eEntryType
            Case EDITNew
                sSQL = sSQL & " AND (x." & gk_sDBCOLEditSetEntryTypeID & " IN (" & _
                    cesImportedModified & ", " & cesPROCATCreatedModified & ", " & _
                    cesPROCATCreated & ", " & cesImported & ")) "
            Case EDITUpdated
                sSQL = sSQL & " AND x." & gk_sDBCOLEditSetEntryTypeID & "=" & cesAddedFromLiveModified
            Case EDITDelete 'd18
                sSQL = sSQL & " AND x." & gk_sDBCOLEditSetEntryTypeID & "=" & cesAddedFromLive
            Case Else
                RaiseError sSource, PEGeneralInternalError, "Invalid editset entry type"
        End Select
    End If
    
    GetPieceSql = sSQL
    Exit Function
    
ErrorHandler:
    ConvertError sSource
    
End Function


Private Function GetItemSql(t_eCatStructure As ECatStructure, _
            Optional t_eEntryType As EEditSetEntryType, _
            Optional t_bFromLive As Boolean = False, _
            Optional t_lLettercodeID As Long, _
            Optional t_lClassID As Long = 0, _
            Optional t_lHeaderID As Long = 0, _
            Optional t_lPieceID As Long = 0, _
            Optional t_sClassIn As String = "", _
            Optional t_sLettercodeIn As String = "", _
            Optional t_bGetExtraExtnData As Boolean = True) As String
'   Generate the sql to select items from either the edit set
'   or live tables.
'   The parameters define the following:
'   t_eCatStructure     The structure of catalogue entry we are selecting for
'                       Eg, pieces with no headers but divisions(CATWithDiv)
'   t_eEntryType        For selecting from edit sets, the edit set entry type
'   t_bFromLive         Flag to indicate we are selecting from live tables
'                       (default is False (ie, from edit sets)
'   t_sLettercode       For selecting from live, the lettercode we
'                       select from
'   t_sClassRef         For selecting from live, the class we
'                       select from (optional)
'   t_sHeaderRef        For selecting from live, the header we
'                       select from (optional)
'   t_lPieceID          If this is specified we select only items
'                       within this piece (for moving pieces)
'
'   d2  WDP     25-Oct-1999      Combined GetItemSql and
'                                GetItemSqlFromLIVE to form
'                                Generic method.
'   d3  WDP     29-Oct-1999      Added class and header optional params
'   d5  WDP     19-Nov-1999      Removed collection type parameter
'                                we now select all data
'   d6  WDP     19-Nov-1999      Added piece ID param
'   d7  WDP     30-Nov-1999      Use id cols instead of ref parts
'   d9  WDP     14-Feb-2000     Select non-modified entries that have been accessioned or created
'                               when generating new editset data
'   d12 WDP     14-Mar-2000     Altered SQL generation to only select live records
'                                   when selecting from LIVE (live_flag = 0 or 1)
'   d18 ALM     25-Oct-2001     Supports a new type of entry - delete
'
On Error GoTo ErrorHandler

    Dim sSource As String
    sSource = "BOCatalogueIndexer.GetItemSql"
    
    Dim sSQL As String
    'We select NULL values for the fields which do not exist at
    'this level to maintain the cammon shape of the record set
    sSQL = "SELECT " & _
            "7 AS " & gk_sDICatalogueLevel & ", " & _
            "i." & gk_sDBCOLItemID & " AS " & gk_sDIID & ", " & _
            "lc." & gk_sDBCOLLetterCodeCode & " AS " & gk_sDILetterCodeRef & ", "
    
    Select Case t_eCatStructure
        Case CATPlain, CATWithHeader, CATWithSubheader
            sSQL = sSQL & " NULL "
        Case Else
            sSQL = sSQL & " d." & gk_sDBCOLDivisionNo
    End Select

    sSQL = sSQL & " AS " & gk_sDIDivisionRef & ", " & _
            "c." & gk_sDBCOLClassNumber & " AS " & gk_sDIClassRef & ", " & _
            "c." & gk_sDBCOLSubClassNumber & " AS " & gk_sDISubclassRef & ", "
    
    Select Case t_eCatStructure
    Case CATPlain, CATWithDiv
        sSQL = sSQL & " NULL "
    Case Else
        sSQL = sSQL & " h." & gk_sDBCOLClassHeaderNo
    End Select

    sSQL = sSQL & " AS " & gk_sDIHeaderRef & ", "
    
    Select Case t_eCatStructure
    Case CATPlain, CATWithDiv, CATWithHeader, CATWithHeaderAndDiv
        sSQL = sSQL & " NULL "
    Case Else
        sSQL = sSQL & " sh." & gk_sDBCOLSubHeaderNo
    End Select

    sSQL = sSQL & " AS " & gk_sDISubheaderRef & ", " & _
            "p." & gk_sDBCOLPieceRef & " AS " & gk_sDIPieceRef & ", " & _
            "i." & gk_sDBCOLItemRef & " AS " & gk_sDIItemRef & ", " & _
            "p." & gk_sDBColPiecekeyOrder & " AS " & gk_sDIPieceKeyOrder & ", " & _
            "i." & gk_sDBCOLItemKeyOrder & " AS " & gk_sDIItemKeyOrder & ", " & _
            "i." & gk_sDBCOLFirstDate & " AS " & gk_sDIFirstDate & ", " & _
            "i." & gk_sDBCOLLastDate & " AS " & gk_sDILastDate & ", " & _
            "i." & gk_sDBCOLDateText & " AS " & gk_sDICoveringDates & ", " & _
            "x." & gk_sDBCOLMapDesignation & " AS " & gk_sDIMapDesignation & ", " & _
            "x." & gk_sDBCOLFormerDepartmentReference & " AS " & gk_sDIFormerRef & ", " & _
            "x." & gk_sDBCOLFormerPROReference & " AS " & gk_sDIFormerPRORef
    'Extra fields for advanced collections
    If t_bGetExtraExtnData = True Then
        sSQL = sSQL & ", " & _
             "x." & gk_sDBCOLMapScaleNumber & " AS " & gk_sDIMapScale & ", " & _
             "x." & gk_sDBCOLRestrictionsOnUse & " AS " & gk_sDIRestrictions & ", " & _
             "x." & gk_sDBCOLAccumulationdates & " AS " & gk_sDIAccumulationDates & ", " & _
             "i." & gk_sDBCOLPieceOpenDate & " AS " & gk_sDIOpeningDates & ", " & _
             "x." & gk_sDBCOLLegalStatusID & " AS " & gk_sDILegalStatus & ", " & _
             "x." & gk_sDBCOLPhysicalDescriptionFormID & " AS " & gk_sDIPhysicalDescForm & ", " & _
             "x." & gk_sDBCOLPhysicalDescriptionRecordQuantity & " AS " & gk_sDIPhysDescQuantity & ", " & _
             "x." & gk_sDBCOLPhysicalDescriptionDimension & " AS " & gk_sDIPhysDescDimensions & ", " & _
             "x." & gk_sDBCOLAccessConditionID & " AS " & gk_sDIAccessCondition & ", " & _
             "i." & gk_sDBCOLLanguageID & " AS " & gk_sDILanguage & ", " & _
             "x." & gk_sDBCOLPhysicalCondition & " AS " & gk_sDIPhysicalCondition & ", " & _
             "i." & gk_sDBCOLClosureCode & " AS " & gk_sDIClosureCode & ", " & _
             "i." & gk_sDBCOLClosureType & " AS " & gk_sDIClosureType & ", " & _
             "i." & gk_sDBCOLClosureStatus & " AS " & gk_sDIClosureStatus
    End If
    
    'The tables we are selecting from
    sSQL = sSQL & " FROM "
    
    If t_bFromLive = False Then
        sSQL = sSQL & _
            gk_sDBTABLEEditsetLettercode & " lc, " & _
            gk_sDBTABLEEditsetClass & " c, " & _
            gk_sDBTABLEEditsetPiece & " p, " & _
            gk_sDBTABLEEditsetItem & " i, " & _
            gk_sDBTABLEEditsetItemExtension & " x "
        Select Case t_eCatStructure
            Case CATWithDiv
                sSQL = sSQL & ", " & gk_sDBTABLEEditsetDivision & " d "
            Case CATWithHeader
                sSQL = sSQL & ", " & gk_sDBTABLEEditsetHeader & " h "
            Case CATWithHeaderAndDiv
                sSQL = sSQL & ", " & gk_sDBTABLEEditsetDivision & " d "
                sSQL = sSQL & ", " & gk_sDBTABLEEditsetHeader & " h "
            Case CATWithSubheaderAndDiv
                sSQL = sSQL & ", " & gk_sDBTABLEEditsetDivision & " d "
                sSQL = sSQL & ", " & gk_sDBTABLEEditsetHeader & " h "
                sSQL = sSQL & ", " & gk_sDBTABLEEditsetSubheader & " sh "
            Case CATWithSubheader
                sSQL = sSQL & ", " & gk_sDBTABLEEditsetHeader & " h "
                sSQL = sSQL & ", " & gk_sDBTABLEEditsetSubheader & " sh "
            Case CATPlain
        End Select
    Else
        sSQL = sSQL & _
            gk_sDBTABLELCode & " lc, " & _
            gk_sDBTABLEClass & " c, " & _
            gk_sDBTABLEPiece & " p, " & _
            gk_sDBTABLEItem & " i, " & _
            gk_sDBTableItemExtension & " x "
        Select Case t_eCatStructure
            Case CATWithDiv
                sSQL = sSQL & ", " & gk_sDBTABLEDivision & " d "
            Case CATWithHeader
                sSQL = sSQL & ", " & gk_sDBTABLEHeader & " h "
            Case CATWithHeaderAndDiv
                sSQL = sSQL & ", " & gk_sDBTABLEDivision & " d "
                sSQL = sSQL & ", " & gk_sDBTABLEHeader & " h "
            Case CATWithSubheaderAndDiv
                sSQL = sSQL & ", " & gk_sDBTABLEDivision & " d "
                sSQL = sSQL & ", " & gk_sDBTABLEHeader & " h "
                sSQL = sSQL & ", " & gk_sDBTABLESubHeader & " sh "
            Case CATWithSubheader
                sSQL = sSQL & ", " & gk_sDBTABLEHeader & " h "
                sSQL = sSQL & ", " & gk_sDBTABLESubHeader & " sh "
            Case CATPlain
        End Select
    End If
    
    'The joining clauses
    sSQL = sSQL & _
            " WHERE " & _
            "i." & gk_sDBCOLPieceID & "=p." & gk_sDBCOLPieceID & " AND " & _
            "p." & gk_sDBCOLClassID & "=c." & gk_sDBCOLClassID & " AND " & _
            "c." & gk_sDBCOLLetterCodeID & "=lc." & gk_sDBCOLLetterCodeID & " AND " & _
            "i." & gk_sDBCOLItemID & "=x." & gk_sDBCOLItemID
            
    Select Case t_eCatStructure
        Case CATWithDiv
            sSQL = sSQL & " AND (c." & gk_sDBCOLDivisionID & " IS NOT NULL AND c." & _
                        gk_sDBCOLDivisionID & "<>0)" & _
                        " AND c." & gk_sDBCOLDivisionID & "=d." & gk_sDBCOLDivisionID
            sSQL = sSQL & " AND (p." & gk_sDBCOLHeaderID & " IS NULL OR p." & _
                        gk_sDBCOLHeaderID & "=0)"
            sSQL = sSQL & " AND (p." & gk_sDBCOLSubHeaderID & " IS NULL OR p." & _
                        gk_sDBCOLSubHeaderID & "=0)"
            
        Case CATWithHeaderAndDiv
            sSQL = sSQL & " AND (c." & gk_sDBCOLDivisionID & " IS NOT NULL AND c." & _
                        gk_sDBCOLDivisionID & "<>0)" & _
                        " AND c." & gk_sDBCOLDivisionID & "=d." & gk_sDBCOLDivisionID
            sSQL = sSQL & " AND (p." & gk_sDBCOLHeaderID & " IS NOT NULL AND p." & _
                        gk_sDBCOLHeaderID & "<>0)" & _
                        " AND p." & gk_sDBCOLHeaderID & "=h." & gk_sDBCOLHeaderID
            sSQL = sSQL & " AND (p." & gk_sDBCOLSubHeaderID & " IS NULL OR p." & _
                        gk_sDBCOLSubHeaderID & "=0)"
            
        Case CATWithSubheaderAndDiv
            sSQL = sSQL & " AND (c." & gk_sDBCOLDivisionID & " IS NOT NULL AND c." & _
                        gk_sDBCOLDivisionID & "<>0)" & _
                        " AND c." & gk_sDBCOLDivisionID & "=d." & gk_sDBCOLDivisionID
            sSQL = sSQL & " AND (p." & gk_sDBCOLHeaderID & " IS NOT NULL AND p." & _
                        gk_sDBCOLHeaderID & "<>0)" & _
                        " AND p." & gk_sDBCOLHeaderID & "=h." & gk_sDBCOLHeaderID
            sSQL = sSQL & " AND (p." & gk_sDBCOLSubHeaderID & " IS NOT NULL AND p." & _
                        gk_sDBCOLSubHeaderID & "<>0)" & _
                        " AND p." & gk_sDBCOLSubHeaderID & "=sh." & gk_sDBCOLSubHeaderID
                        
        Case CATWithHeader
            sSQL = sSQL & " AND (c." & gk_sDBCOLDivisionID & " IS NULL OR c." & _
                        gk_sDBCOLDivisionID & "=0)"
            sSQL = sSQL & " AND (p." & gk_sDBCOLHeaderID & " IS NOT NULL AND p." & _
                        gk_sDBCOLHeaderID & "<>0)" & _
                        " AND p." & gk_sDBCOLHeaderID & "=h." & gk_sDBCOLHeaderID
            sSQL = sSQL & " AND (p." & gk_sDBCOLSubHeaderID & " IS NULL OR p." & _
                        gk_sDBCOLSubHeaderID & "=0)"
            
        Case CATWithSubheader
            sSQL = sSQL & " AND (c." & gk_sDBCOLDivisionID & " IS NULL OR c." & _
                        gk_sDBCOLDivisionID & "=0)"
            sSQL = sSQL & " AND (p." & gk_sDBCOLHeaderID & " IS NOT NULL AND p." & _
                        gk_sDBCOLHeaderID & "<>0)" & _
                        " AND p." & gk_sDBCOLHeaderID & "=h." & gk_sDBCOLHeaderID
            sSQL = sSQL & " AND (p." & gk_sDBCOLSubHeaderID & " IS NOT NULL AND p." & _
                        gk_sDBCOLSubHeaderID & "<>0)" & _
                        " AND p." & gk_sDBCOLSubHeaderID & "=sh." & gk_sDBCOLSubHeaderID
                        
        Case CATPlain
            sSQL = sSQL & " AND (c." & gk_sDBCOLDivisionID & " IS NULL OR c." & _
                        gk_sDBCOLDivisionID & "=0)"
            sSQL = sSQL & " AND (p." & gk_sDBCOLHeaderID & " IS NULL OR p." & _
                        gk_sDBCOLHeaderID & "=0)"
            sSQL = sSQL & " AND (p." & gk_sDBCOLSubHeaderID & " IS NULL OR p." & _
                        gk_sDBCOLSubHeaderID & "=0)"
                        
    End Select
            
    If t_bFromLive = False Then
        sSQL = sSQL & " AND lc." & gk_sDBCOLEditSetID & "=" & m_lEditsetID & _
                  " AND c." & gk_sDBCOLEditSetID & "=" & m_lEditsetID & _
                  " AND p." & gk_sDBCOLEditSetID & "=" & m_lEditsetID & _
                  " AND i." & gk_sDBCOLEditSetID & "=" & m_lEditsetID & _
                  " AND x." & gk_sDBCOLEditSetID & "=" & m_lEditsetID
        Select Case t_eCatStructure
            Case CATWithDiv
                sSQL = sSQL & " AND d." & gk_sDBCOLEditSetID & "=" & m_lEditsetID
            Case CATWithHeaderAndDiv
                sSQL = sSQL & " AND d." & gk_sDBCOLEditSetID & "=" & m_lEditsetID
                sSQL = sSQL & " AND h." & gk_sDBCOLEditSetID & "=" & m_lEditsetID
            Case CATWithSubheaderAndDiv
                sSQL = sSQL & " AND d." & gk_sDBCOLEditSetID & "=" & m_lEditsetID
                sSQL = sSQL & " AND h." & gk_sDBCOLEditSetID & "=" & m_lEditsetID
                sSQL = sSQL & " AND sh." & gk_sDBCOLEditSetID & "=" & m_lEditsetID
            Case CATWithHeader
                sSQL = sSQL & " AND h." & gk_sDBCOLEditSetID & "=" & m_lEditsetID
            Case CATWithSubheader
                sSQL = sSQL & " AND h." & gk_sDBCOLEditSetID & "=" & m_lEditsetID
                sSQL = sSQL & " AND sh." & gk_sDBCOLEditSetID & "=" & m_lEditsetID
            Case CATPlain
        End Select
    Else
        If t_lPieceID = 0 Then
            If t_sLettercodeIn = "" Then
                sSQL = sSQL & " AND lc." & gk_sDBCOLLetterCodeID & "=" & t_lLettercodeID
            Else
                sSQL = sSQL & " AND lc." & gk_sDBCOLLetterCodeID & " IN " & t_sLettercodeIn
            End If
            If t_sClassIn = "" Then
                If Not t_lClassID = 0 Then
                    sSQL = sSQL & " AND c." & gk_sDBCOLClassID & "=" & t_lClassID
                End If
            Else
                sSQL = sSQL & " AND c." & gk_sDBCOLClassID & " IN " & t_sClassIn
            End If
            If Not (t_lHeaderID = 0 Or t_lHeaderID = mk_lNullRef) Then
                sSQL = sSQL & " AND h." & gk_sDBCOLHeaderID & "=" & t_lHeaderID
            End If
        Else
            sSQL = sSQL & " AND i." & gk_sDBCOLPieceID & "=" & t_lPieceID
        End If
        sSQL = sSQL & " AND x." & gk_sDBCOLLiveFlag & " IN (0,1) "
    End If
    
    If t_bFromLive = False Then
        Select Case t_eEntryType
            Case EDITNew
                sSQL = sSQL & " AND (x." & gk_sDBCOLEditSetEntryTypeID & " IN (" & _
                    cesImportedModified & ", " & cesPROCATCreatedModified & ", " & _
                    cesPROCATCreated & ", " & cesImported & ")) "
            Case EDITUpdated
                sSQL = sSQL & " AND x." & gk_sDBCOLEditSetEntryTypeID & "=" & cesAddedFromLiveModified
            Case EDITDelete 'd18
                sSQL = sSQL & " AND x." & gk_sDBCOLEditSetEntryTypeID & "=" & cesAddedFromLive
            Case Else
                RaiseError sSource, PEGeneralInternalError, "Invalid editset entry type"
        End Select
    End If
    
    GetItemSql = sSQL
    
    Exit Function
    
ErrorHandler:
    ConvertError sSource
    
End Function



Private Function GetLargeFieldSql( _
            t_eField As ELargeField, _
            Optional t_eEntryType As EEditSetEntryType, _
            Optional t_bFromLive As Boolean = False, _
            Optional t_eLevel As ECATALOGUELEVEL, _
            Optional t_lLettercodeID As Long = 0, _
            Optional t_lClassID As Long = 0, _
            Optional t_lHeaderID As Long = 0, _
            Optional t_sLettercodeIn As String = "", _
            Optional t_sClassIn As String = "") As String
'   Generate the sql to select fields from the edit set
'   or live tables that are large so require a separate record set
'   The parameters are as follows:
'   t_eField            The field we are selecting for
'   t_entryType         When selecting from edit sets, the entry type we want
'   t_bFromLive         Flag to indicate if we want to select from edit set
'                       or live tables (default is edit set)
'   t_eLevel            When selecting from live this gives the level
'                       to select
'   t_sLettercode       When selecting from live, this gives the lettercode
'                       select entries from.
'   t_sClassRef         For selecting from live, the class we
'                       select from (optional)
'   t_sHeaderRef        For selecting from live, the header we
'                       select from (optional)
'
'   d2  WDP     26-Oct-1999      Created from GetLargeFieldSql
'                                and GetLargeFieldSqlFromLIVE
'   d3  WDP     29-Oct-1999      Added class and header optional params
'   d8  WDP     13-Dec-1999      Allowed lettercode id to be optional
'                                if it is omitted then all entries are selected
'                                for that level
'                                Added lettercode in list to allow
'                                selection from multiple lettercodes
'
'   d9  WDP     14-Feb-2000     Select non-modified entries that have been accessioned or created
'                               when generating new editset data
'   d12 WDP     14-Mar-2000     Altered SQL generation to only select live records
'                                   when selecting from LIVE (live_flag = 0 or 1)
'   d18 ALM     25-Oct-2001     Supports a new type of entry - delete
'
On Error GoTo ErrorHandler

    Dim sSource As String
    sSource = "BOCatalogueIndexer.GetLargeFieldSql"
    
    Dim sLCTableName As String, sDivTableName As String, sClsTableName As String
    Dim sHdrTableName As String, sSubTableName As String
    Dim sPcTableName As String, sItmTableName As String
    Dim sZoneDI As String, sZoneFieldName As String
    
    'Set table, data item and column names depending on the zone
    Select Case t_eField
    Case LARGETitle
        If t_bFromLive = False Then
            sLCTableName = gk_sDBTABLEEditsetLettercodeTitle
            sDivTableName = gk_sDBTABLEEditsetDivisionTitle
            sClsTableName = gk_sDBTABLEEditsetClassTitle
            sHdrTableName = gk_sDBTABLEEditsetHeaderTitle
            sSubTableName = gk_sDBTABLEEditsetSubheaderTitle
            sPcTableName = gk_sDBTABLEEditsetPieceTitle
            sItmTableName = gk_sDBTABLEEditsetItemTitle
        Else
            sLCTableName = gk_sDBTableLettercodeTitle
            sDivTableName = gk_sDBTableDivisionTitle
            sClsTableName = gk_sDBTableClassTitle
            sHdrTableName = gk_sDBTableHeaderTitle
            sSubTableName = gk_sDBTableSubHeaderTitle
            sPcTableName = gk_sDBTablePieceTitle
            sItmTableName = gk_sDBTableItemTitle
        End If
        sZoneDI = gk_sDITitle
        sZoneFieldName = gk_sDBCOLTitle
    Case LARGEScopeContent
        If t_bFromLive = False Then
            sLCTableName = gk_sDBTABLEEditsetLettercodeScopeAndContent
            sDivTableName = gk_sDBTABLEEditsetDivisionScopeAndContent
            sClsTableName = gk_sDBTABLEEditsetClassScopeAndContent
            sHdrTableName = gk_sDBTABLEEditsetHeaderScopeAndContent
            sSubTableName = gk_sDBTABLEEditsetSubheaderScopeAndContent
            sPcTableName = gk_sDBTABLEEditsetPieceScopeAndContent
            sItmTableName = gk_sDBTABLEEditsetItemScopeAndContent
        Else
            sLCTableName = gk_sDBTableLettercodeScopeAndContent
            sDivTableName = gk_sDBTableDivisionScopeAndContent
            sClsTableName = gk_sDBTableClassScopeAndContent
            sHdrTableName = gk_sDBTableHeaderScopeAndContent
            sSubTableName = gk_sDBTableSubHeaderScopeAndContent
            sPcTableName = gk_sDBTablePieceScopeAndContent
            sItmTableName = gk_sDBTableItemScopeAndContent
        End If
        sZoneDI = gk_sDIScopeContent
        sZoneFieldName = gk_sDBCOLScopeAndContent
    Case LARGEAdminHistory
        If t_bFromLive = False Then
            sLCTableName = gk_sDBTABLEEditsetLettercodeAdministrativeHistory
            sDivTableName = gk_sDBTABLEEditsetDivisionAdministrativeHistory
            sClsTableName = gk_sDBTABLEEditsetClassAdministrativeHistory
            sHdrTableName = gk_sDBTABLEEditsetHeaderAdministrativeHistory
            sSubTableName = gk_sDBTABLEEditsetSubheaderAdministrativeHistory
            sPcTableName = gk_sDBTABLEEditsetPieceAdministrativeHistory
            sItmTableName = gk_sDBTABLEEditsetItemAdministrativeHistory
        Else
            sLCTableName = gk_sDBTABLELettercodeAdminBiogHist
            sDivTableName = gk_sDBTABLEDivisionAdminBiogHist
            sClsTableName = gk_sDBTABLEClassAdminBiogHist
            sHdrTableName = gk_sDBTABLEHeaderAdminBiogHist
            sSubTableName = gk_sDBTABLESubHeaderAdminBiogHist
            sPcTableName = gk_sDBTABLEPieceAdminBiogHist
            sItmTableName = gk_sDBTABLEItemAdminBiogHist
        End If
        sZoneDI = gk_sDIAdminHistory
        sZoneFieldName = gk_sDBCOLAdministrativeHistory
    Case LARGEAccruals
        If t_bFromLive = False Then
            sLCTableName = gk_sDBTABLEEditsetLettercodeAccruals
            sDivTableName = gk_sDBTABLEEditsetDivisionAccruals
            sClsTableName = gk_sDBTABLEEditsetClassAccruals
            sHdrTableName = gk_sDBTABLEEditsetHeaderAccruals
            sSubTableName = gk_sDBTABLEEditsetSubheaderAccruals
            sPcTableName = gk_sDBTABLEEditsetPieceAccruals
            sItmTableName = gk_sDBTABLEEditsetItemAccruals
        Else
            sLCTableName = gk_sDBTABLELettercodeAccruals
            sDivTableName = gk_sDBTABLEDivisionAccruals
            sClsTableName = gk_sDBTABLEClassAccruals
            sHdrTableName = gk_sDBTABLEHeaderAccruals
            sSubTableName = gk_sDBTABLESubHeaderAccruals
            sPcTableName = gk_sDBTABLEPieceAccruals
            sItmTableName = gk_sDBTABLEItemAccruals
        End If
        sZoneDI = gk_sDIAccruals
        sZoneFieldName = gk_sDBCOLAccruals
    Case LARGEArrangement
        If t_bFromLive = False Then
            sLCTableName = gk_sDBTABLEEditsetLettercodeArrangement
            sDivTableName = gk_sDBTABLEEditsetDivisionArrangement
            sClsTableName = gk_sDBTABLEEditsetClassArrangement
            sHdrTableName = gk_sDBTABLEEditsetHeaderArrangement
            sSubTableName = gk_sDBTABLEEditsetSubheaderArrangement
            sPcTableName = gk_sDBTABLEEditsetPieceArrangement
            sItmTableName = gk_sDBTABLEEditsetItemArrangement
        Else
            sLCTableName = gk_sDBTABLELettercodeArrangement
            sDivTableName = gk_sDBTABLEDivisionArrangement
            sClsTableName = gk_sDBTABLEClassArrangement
            sHdrTableName = gk_sDBTABLEHeaderArrangement
            sSubTableName = gk_sDBTABLESubHeaderArrangement
            sPcTableName = gk_sDBTABLEPieceArrangement
            sItmTableName = gk_sDBTABLEItemArrangement
        End If
        sZoneDI = gk_sDIArrangement
        sZoneFieldName = gk_sDBCOLArrangement
    Case LARGECustodialHistory
        If t_bFromLive = False Then
            sLCTableName = gk_sDBTABLEEditsetLettercodeCustodialHistory
            sDivTableName = gk_sDBTABLEEditsetDivisionCustodialHistory
            sClsTableName = gk_sDBTABLEEditsetClassCustodialHistory
            sHdrTableName = gk_sDBTABLEEditsetHeaderCustodialHistory
            sSubTableName = gk_sDBTABLEEditsetSubheaderCustodialHistory
            sPcTableName = gk_sDBTABLEEditsetPieceCustodialHistory
            sItmTableName = gk_sDBTABLEEditsetItemCustodialHistory
        Else
            sLCTableName = gk_sDBTABLELettercodeCustodialHist
            sDivTableName = gk_sDBTABLEDivisionCustodialHist
            sClsTableName = gk_sDBTABLEClassCustodialHist
            sHdrTableName = gk_sDBTABLEHeaderCustodialHist
            sSubTableName = gk_sDBTABLESubHeaderCustodialHist
            sPcTableName = gk_sDBTABLEPieceCustodialHist
            sItmTableName = gk_sDBTABLEItemCustodialHist
        End If
        sZoneDI = gk_sDICustodialHist
        sZoneFieldName = gk_sDBCOLCustodialHistory
    Case LARGEAppraisalInfo
        If t_bFromLive = False Then
            sLCTableName = gk_sDBTABLEEditsetLettercodeAppraisalDestructionInformation
            sDivTableName = gk_sDBTABLEEditsetDivisionAppraisalDestructionInformation
            sClsTableName = gk_sDBTABLEEditsetClassAppraisalDestructionInformation
            sHdrTableName = gk_sDBTABLEEditsetHeaderAppraisalDestructionInformation
            sSubTableName = gk_sDBTABLEEditsetSubheaderAppraisalDestructionInformation
            sPcTableName = gk_sDBTABLEEditsetPieceAppraisalDestructionInformation
            sItmTableName = gk_sDBTABLEEditsetItemAppraisalDestructionInformation
        Else
            sLCTableName = gk_sDBTABLELettercodeAppDestInfo
            sDivTableName = gk_sDBTABLEDivisionAppDestInfo
            sClsTableName = gk_sDBTABLEClassAppDestInfo
            sHdrTableName = gk_sDBTABLEHeaderAppDestInfo
            sSubTableName = gk_sDBTABLESubHeaderAppDestInfo
            sPcTableName = gk_sDBTABLEPieceAppDestInfo
            sItmTableName = gk_sDBTABLEItemAppDestInfo
        End If
        sZoneDI = gk_sDIAppraisalInfo
        sZoneFieldName = gk_sDBCOLAppraisalDestructionInformation
    Case LARGENote
        If t_bFromLive = False Then
            sLCTableName = gk_sDBTABLEEditsetLetterCodeNote
            sDivTableName = gk_sDBTABLEEditsetDivisionNote
            sClsTableName = gk_sDBTABLEEditsetClassNote
            sHdrTableName = gk_sDBTABLEEditsetHeaderNote
            sSubTableName = gk_sDBTABLEEditsetSubHeaderNote
            sPcTableName = gk_sDBTABLEEditsetPieceNote
            sItmTableName = gk_sDBTABLEEditsetItemNote
        Else
            sLCTableName = gk_sDBTABLELettercodeNote
            sDivTableName = gk_sDBTABLEDivisionNote
            sClsTableName = gk_sDBTABLEClassNote
            sHdrTableName = gk_sDBTABLEHeaderNote
            sSubTableName = gk_sDBTABLESubHeaderNote
            sPcTableName = gk_sDBTABLEPieceNote
            sItmTableName = gk_sDBTABLEItemNote
        End If
        sZoneDI = gk_sDINoteText
        sZoneFieldName = gk_sDBCOLNoteText
    Case Else
        RaiseError sSource, PEGeneralInternalError, "Invalid zone enum."
    End Select
    
    Dim sSQL As String, sLCodeSQL As String, sDivSQL As String, sClsSQL As String
    Dim sHdrSQL As String, sSubSQL As String, sPcSQL As String, sItmSQL As String
    'filter to get the desired type of records (new or updated)
    Dim sFilterSQL As String
    
    If t_bFromLive = False Then
        Select Case t_eEntryType
            Case EDITNew
                sFilterSQL = " AND (x." & gk_sDBCOLEditSetEntryTypeID & " IN (" & _
                    cesImportedModified & ", " & cesPROCATCreatedModified & ", " & _
                    cesPROCATCreated & ", " & cesImported & ")) "
            Case EDITUpdated
                sFilterSQL = " AND x." & gk_sDBCOLEditSetEntryTypeID & "=" & cesAddedFromLiveModified
            Case EDITDelete 'd18
                sFilterSQL = " AND x." & gk_sDBCOLEditSetEntryTypeID & "=" & cesAddedFromLive
            Case Else
                RaiseError sSource, PEGeneralInternalError, "Invalid editset entry type"
        End Select
    End If
    'We select from each catalogue level separately
    
    'Lettercodes
    sLCodeSQL = "SELECT " & _
            "1 AS " & gk_sDICatalogueLevel & ", " & _
            "lc." & gk_sDBCOLLetterCodeID & " AS " & gk_sDIID & ", " & _
            "a." & sZoneFieldName & " AS " & sZoneDI & _
            " FROM "
    If t_bFromLive = False Then
        sLCodeSQL = sLCodeSQL & _
            gk_sDBTABLEEditsetLettercode & " lc, " & _
            gk_sDBTABLEEditsetLettercodeExtension & " x, "
    Else
        sLCodeSQL = sLCodeSQL & _
            gk_sDBTABLELCode & " lc, " & _
            gk_sDBTABLELetterCodeExtension & " x, "
    End If
    sLCodeSQL = sLCodeSQL & _
            sLCTableName & " a " & _
            " WHERE " & _
            "lc." & gk_sDBCOLLetterCodeID & "=a." & gk_sDBCOLLetterCodeID & _
            " AND lc." & gk_sDBCOLLetterCodeID & "=x." & gk_sDBCOLLetterCodeID
    
    If t_bFromLive = False Then
        sLCodeSQL = sLCodeSQL & _
            " AND lc." & gk_sDBCOLEditSetID & "=" & m_lEditsetID & _
            " AND x." & gk_sDBCOLEditSetID & "=" & m_lEditsetID & _
            " AND a." & gk_sDBCOLEditSetID & "=" & m_lEditsetID & _
            sFilterSQL
    Else
        If Not t_lLettercodeID = 0 Then
            If t_sLettercodeIn = "" Then
                sLCodeSQL = sLCodeSQL & _
                    " AND lc." & gk_sDBCOLLetterCodeID & "=" & t_lLettercodeID
            End If
        Else
            If Not t_sLettercodeIn = "" Then
                sLCodeSQL = sLCodeSQL & _
                    " AND lc." & gk_sDBCOLLetterCodeID & " IN " & t_sLettercodeIn
            End If
        End If
        sLCodeSQL = sLCodeSQL & " AND x." & gk_sDBCOLLiveFlag & " IN (0,1) "
    End If
    
    'Divisions
    sDivSQL = "SELECT " & _
            "2 AS " & gk_sDICatalogueLevel & ", " & _
            "d." & gk_sDBCOLDivisionID & " AS " & gk_sDIID & ", " & _
            "a." & sZoneFieldName & " AS " & sZoneDI & _
            " FROM "
    If t_bFromLive = False Then
        sDivSQL = sDivSQL & _
            gk_sDBTABLEEditsetDivision & " d, " & _
            gk_sDBTABLEEditsetDivisionExtension & " x, "
    Else
        sDivSQL = sDivSQL & _
            gk_sDBTABLEDivision & " d, " & _
            gk_sDBTableDivisionExtension & " x, " & _
            gk_sDBTABLELCode & " lc, "
    End If
    sDivSQL = sDivSQL & _
            sDivTableName & " a " & _
            " WHERE " & _
            "d." & gk_sDBCOLDivisionID & "=a." & gk_sDBCOLDivisionID & _
            " AND d." & gk_sDBCOLDivisionID & "=x." & gk_sDBCOLDivisionID
    
    If t_bFromLive = False Then
        sDivSQL = sDivSQL & _
            " AND d." & gk_sDBCOLEditSetID & "=" & m_lEditsetID & _
            " AND x." & gk_sDBCOLEditSetID & "=" & m_lEditsetID & _
            " AND a." & gk_sDBCOLEditSetID & "=" & m_lEditsetID & _
            sFilterSQL
    Else
        If Not t_lLettercodeID = 0 Then
            If t_sLettercodeIn = "" Then
                sDivSQL = sDivSQL & _
                    " AND lc." & gk_sDBCOLLetterCodeID & "=" & t_lLettercodeID
            End If
        Else
            If Not t_sLettercodeIn = "" Then
                sDivSQL = sDivSQL & _
                    " AND lc." & gk_sDBCOLLetterCodeID & " IN " & t_sLettercodeIn
            End If
        End If
        sDivSQL = sDivSQL & _
            " AND lc." & gk_sDBCOLLetterCodeID & "=d." & gk_sDBCOLLetterCodeID
        sDivSQL = sDivSQL & " AND x." & gk_sDBCOLLiveFlag & " IN (0,1) "
    End If
            
    'Classes
    sClsSQL = "SELECT " & _
            "3 AS " & gk_sDICatalogueLevel & ", " & _
            "c." & gk_sDBCOLClassID & " AS " & gk_sDIID & ", " & _
            "a." & sZoneFieldName & " AS " & sZoneDI & _
            " FROM "
    If t_bFromLive = False Then
        sClsSQL = sClsSQL & _
            gk_sDBTABLEEditsetClass & " c, " & _
            gk_sDBTABLEEditsetClassExtension & " x, "
    Else
        sClsSQL = sClsSQL & _
            gk_sDBTABLEClass & " c, " & _
            gk_sDBTableClassExtension & " x, " & _
            gk_sDBTABLELCode & " lc, "
    End If
    sClsSQL = sClsSQL & _
            sClsTableName & " a " & _
            " WHERE " & _
            "c." & gk_sDBCOLClassID & "=a." & gk_sDBCOLClassID & _
            " AND c." & gk_sDBCOLClassID & "=x." & gk_sDBCOLClassID
    If t_bFromLive = False Then
        sClsSQL = sClsSQL & _
            " AND c." & gk_sDBCOLEditSetID & "=" & m_lEditsetID & _
            " AND x." & gk_sDBCOLEditSetID & "=" & m_lEditsetID & _
            " AND a." & gk_sDBCOLEditSetID & "=" & m_lEditsetID & _
            sFilterSQL
    Else
        If Not t_lLettercodeID = 0 Then
            If t_sLettercodeIn = "" Then
                sClsSQL = sClsSQL & _
                    " AND lc." & gk_sDBCOLLetterCodeID & "=" & t_lLettercodeID
            End If
        Else
            If Not t_sLettercodeIn = "" Then
                sClsSQL = sClsSQL & _
                    " AND lc." & gk_sDBCOLLetterCodeID & " IN " & t_sLettercodeIn
            End If
        End If
        sClsSQL = sClsSQL & _
            " AND lc." & gk_sDBCOLLetterCodeID & "=c." & gk_sDBCOLLetterCodeID
        sClsSQL = sClsSQL & " AND x." & gk_sDBCOLLiveFlag & " IN (0,1) "
    End If
    
    'Headers
    sHdrSQL = "SELECT " & _
            "4 AS " & gk_sDICatalogueLevel & ", " & _
            "h." & gk_sDBCOLHeaderID & " AS " & gk_sDIID & ", " & _
            "a." & sZoneFieldName & " AS " & sZoneDI & _
            " FROM "
    If t_bFromLive = False Then
        sHdrSQL = sHdrSQL & _
            gk_sDBTABLEEditsetHeader & " h, " & _
            gk_sDBTABLEEditsetHeaderExtension & " x, "
    Else
        sHdrSQL = sHdrSQL & _
            gk_sDBTABLEHeader & " h, " & _
            gk_sDBTableHeaderExtension & " x, " & _
            gk_sDBTABLEClass & " c, " & _
            gk_sDBTABLELCode & " lc, "
    End If
    sHdrSQL = sHdrSQL & _
            sHdrTableName & " a " & _
            " WHERE " & _
            "h." & gk_sDBCOLHeaderID & "=a." & gk_sDBCOLHeaderID & _
            " AND h." & gk_sDBCOLHeaderID & "=x." & gk_sDBCOLHeaderID
    If t_bFromLive = False Then
        sHdrSQL = sHdrSQL & _
            " AND h." & gk_sDBCOLEditSetID & "=" & m_lEditsetID & _
            " AND x." & gk_sDBCOLEditSetID & "=" & m_lEditsetID & _
            " AND a." & gk_sDBCOLEditSetID & "=" & m_lEditsetID & _
            sFilterSQL
    Else
        If Not t_lLettercodeID = 0 Then
            If t_sLettercodeIn = "" Then
                sHdrSQL = sHdrSQL & _
                    " AND lc." & gk_sDBCOLLetterCodeID & "=" & t_lLettercodeID
            End If
        Else
            If Not t_sLettercodeIn = "" Then
                sHdrSQL = sHdrSQL & _
                    " AND lc." & gk_sDBCOLLetterCodeID & " IN " & t_sLettercodeIn
            End If
        End If
        sHdrSQL = sHdrSQL & _
            " AND c." & gk_sDBCOLClassID & "=h." & gk_sDBCOLClassID & _
            " AND lc." & gk_sDBCOLLetterCodeID & "=c." & gk_sDBCOLLetterCodeID
        sHdrSQL = sHdrSQL & " AND x." & gk_sDBCOLLiveFlag & " IN (0,1) "
    End If
    
    
    'Subheaders
    sSubSQL = "SELECT " & _
            "5 AS " & gk_sDICatalogueLevel & ", " & _
            "sh." & gk_sDBCOLSubHeaderID & " AS " & gk_sDIID & ", " & _
            "a." & sZoneFieldName & " AS " & sZoneDI & _
            " FROM "
    If t_bFromLive = False Then
        sSubSQL = sSubSQL & _
            gk_sDBTABLEEditsetSubheader & " sh, " & _
            gk_sDBTABLEEditsetSubheaderExtension & " x, "
    Else
        sSubSQL = sSubSQL & _
            gk_sDBTABLESubHeader & " sh, " & _
            gk_sDBTableSubHeaderExtension & " x, " & _
            gk_sDBTABLEHeader & " h, " & _
            gk_sDBTABLEClass & " c, " & _
            gk_sDBTABLELCode & " lc, "
    End If
    sSubSQL = sSubSQL & _
            sSubTableName & " a " & _
            " WHERE " & _
            "sh." & gk_sDBCOLSubHeaderID & "=a." & gk_sDBCOLSubHeaderID & _
            " AND sh." & gk_sDBCOLSubHeaderID & "=x." & gk_sDBCOLSubHeaderID
    If t_bFromLive = False Then
        sSubSQL = sSubSQL & _
            " AND sh." & gk_sDBCOLEditSetID & "=" & m_lEditsetID & _
            " AND x." & gk_sDBCOLEditSetID & "=" & m_lEditsetID & _
            " AND a." & gk_sDBCOLEditSetID & "=" & m_lEditsetID & _
            sFilterSQL
    Else
        If Not t_lLettercodeID = 0 Then
            If t_sLettercodeIn = "" Then
                sSubSQL = sSubSQL & _
                    " AND lc." & gk_sDBCOLLetterCodeID & "=" & t_lLettercodeID
            End If
        Else
            If Not t_sLettercodeIn = "" Then
                sSubSQL = sSubSQL & _
                    " AND lc." & gk_sDBCOLLetterCodeID & " IN " & t_sLettercodeIn
            End If
        End If
        sSubSQL = sSubSQL & _
            " AND h." & gk_sDBCOLHeaderID & "=sh." & gk_sDBCOLHeaderID & _
            " AND c." & gk_sDBCOLClassID & "=h." & gk_sDBCOLClassID & _
            " AND lc." & gk_sDBCOLLetterCodeID & "=c." & gk_sDBCOLLetterCodeID
        sSubSQL = sSubSQL & " AND x." & gk_sDBCOLLiveFlag & " IN (0,1) "
    End If
    
    
    'Pieces
    sPcSQL = "SELECT " & _
            "6 AS " & gk_sDICatalogueLevel & ", " & _
            "p." & gk_sDBCOLPieceID & " AS " & gk_sDIID & ", " & _
            "a." & sZoneFieldName & " AS " & sZoneDI & _
            " FROM "
    If t_bFromLive = False Then
        sPcSQL = sPcSQL & _
            gk_sDBTABLEEditsetPiece & " p, " & _
            gk_sDBTABLEEditsetPieceExtension & " x, "
    Else
        sPcSQL = sPcSQL & _
            gk_sDBTABLEPiece & " p, " & _
            gk_sDBTablePieceExtension & " x, " & _
            gk_sDBTABLEClass & " c, " & _
            gk_sDBTABLELCode & " lc, "
        If Not (t_lHeaderID = 0 Or t_lHeaderID = mk_lNullRef) Then
            sPcSQL = sPcSQL & _
                gk_sDBTABLEHeader & " h, "
        End If
    End If
    sPcSQL = sPcSQL & _
            sPcTableName & " a " & _
            " WHERE " & _
            "p." & gk_sDBCOLPieceID & "=a." & gk_sDBCOLPieceID & _
            " AND p." & gk_sDBCOLPieceID & "=x." & gk_sDBCOLPieceID
    If t_bFromLive = False Then
        sPcSQL = sPcSQL & _
            " AND p." & gk_sDBCOLEditSetID & "=" & m_lEditsetID & _
            " AND x." & gk_sDBCOLEditSetID & "=" & m_lEditsetID & _
            " AND a." & gk_sDBCOLEditSetID & "=" & m_lEditsetID & _
            sFilterSQL
    Else
        sPcSQL = sPcSQL & _
            " AND lc." & gk_sDBCOLLetterCodeID & "=c." & gk_sDBCOLLetterCodeID & _
            " AND c." & gk_sDBCOLClassID & "=p." & gk_sDBCOLClassID
        If t_sLettercodeIn = "" Then
            sPcSQL = sPcSQL & " AND lc." & gk_sDBCOLLetterCodeID & "=" & t_lLettercodeID
        Else
            sPcSQL = sPcSQL & " AND lc." & gk_sDBCOLLetterCodeID & " IN " & t_sLettercodeIn
        End If
        If t_sClassIn = "" Then
            If Not t_lClassID = 0 Then
                sPcSQL = sPcSQL & " AND c." & gk_sDBCOLClassID & "=" & t_lClassID
            End If
        Else
            sPcSQL = sPcSQL & " AND c." & gk_sDBCOLClassID & " IN " & t_sClassIn
        End If
        If Not t_lHeaderID = 0 Then
            If t_lHeaderID = mk_lNullRef Then
                sPcSQL = sPcSQL & " AND p." & gk_sDBCOLHeaderID & " IS NULL"
            Else
                sPcSQL = sPcSQL & " AND h." & gk_sDBCOLHeaderID & "=p." & gk_sDBCOLHeaderID
                sPcSQL = sPcSQL & " AND h." & gk_sDBCOLHeaderID & "=" & t_lHeaderID
            End If
        End If
        sPcSQL = sPcSQL & " AND x." & gk_sDBCOLLiveFlag & " IN (0,1) "
    End If
    
    'Items
    sItmSQL = "SELECT " & _
            "7 AS " & gk_sDICatalogueLevel & ", " & _
            "i." & gk_sDBCOLItemID & " AS " & gk_sDIID & ", " & _
            "a." & sZoneFieldName & " AS " & sZoneDI & _
            " FROM "
    If t_bFromLive = False Then
        sItmSQL = sItmSQL & _
            gk_sDBTABLEEditsetItem & " i, " & _
            gk_sDBTABLEEditsetItemExtension & " x, "
    Else
        sItmSQL = sItmSQL & _
            gk_sDBTABLEItem & " i, " & _
            gk_sDBTableItemExtension & " x, " & _
            gk_sDBTABLEPiece & " p, " & _
            gk_sDBTABLEClass & " c, " & _
            gk_sDBTABLELCode & " lc, "
        If Not t_lHeaderID = 0 Then
            sItmSQL = sItmSQL & _
                gk_sDBTABLEHeader & " h, "
        End If
    End If
    
    sItmSQL = sItmSQL & _
            sItmTableName & " a " & _
            " WHERE " & _
            "i." & gk_sDBCOLItemID & "=a." & gk_sDBCOLItemID & _
            " AND i." & gk_sDBCOLItemID & "=x." & gk_sDBCOLItemID
    If t_bFromLive = False Then
        sItmSQL = sItmSQL & _
            " AND i." & gk_sDBCOLEditSetID & "=" & m_lEditsetID & _
            " AND x." & gk_sDBCOLEditSetID & "=" & m_lEditsetID & _
            " AND a." & gk_sDBCOLEditSetID & "=" & m_lEditsetID & _
            sFilterSQL
    Else
        sItmSQL = sItmSQL & _
            " AND lc." & gk_sDBCOLLetterCodeID & "=c." & gk_sDBCOLLetterCodeID & _
            " AND p." & gk_sDBCOLPieceID & "=i." & gk_sDBCOLPieceID & _
            " AND c." & gk_sDBCOLClassID & "=p." & gk_sDBCOLClassID
        If t_sLettercodeIn = "" Then
            sItmSQL = sItmSQL & " AND lc." & gk_sDBCOLLetterCodeID & "=" & t_lLettercodeID
        Else
            sItmSQL = sItmSQL & " AND lc." & gk_sDBCOLLetterCodeID & " IN " & t_sLettercodeIn
        End If
        
        If t_sClassIn = "" Then
            If Not t_lClassID = 0 Then
                sItmSQL = sItmSQL & " AND c." & gk_sDBCOLClassID & "=" & t_lClassID
            End If
        Else
            sItmSQL = sItmSQL & " AND c." & gk_sDBCOLClassID & " IN " & t_sClassIn
        End If
        If Not t_lHeaderID = 0 Then
            If t_lHeaderID = mk_lNullRef Then
                sItmSQL = sItmSQL & " AND p." & gk_sDBCOLHeaderID & " IS NULL"
            Else
                sItmSQL = sItmSQL & " AND h." & gk_sDBCOLHeaderID & "=p." & gk_sDBCOLHeaderID
                sItmSQL = sItmSQL & " AND h." & gk_sDBCOLHeaderID & "=" & t_lHeaderID
            End If
        End If
        sItmSQL = sItmSQL & " AND x." & gk_sDBCOLLiveFlag & " IN (0,1) "
    End If
    
    'If we are selecting from an edit set then we
    'union the statements to get a combined record set
    'otherwise we only select from the desired level
    
    If t_bFromLive = False Then
        GetLargeFieldSql = sLCodeSQL & " UNION " & _
                        sDivSQL & " UNION " & _
                        sClsSQL & " UNION " & _
                        sHdrSQL & " UNION " & _
                        sSubSQL & " UNION " & _
                        sPcSQL & " UNION " & _
                        sItmSQL
    Else
        Select Case t_eLevel
            Case ctLetterCode
                GetLargeFieldSql = sLCodeSQL
            Case ctDivision
                GetLargeFieldSql = sDivSQL
            Case ctClass
                GetLargeFieldSql = sClsSQL
            Case ctHeaders
                GetLargeFieldSql = sHdrSQL
            Case ctSubHeaders
                GetLargeFieldSql = sSubSQL
            Case ctPiece
                GetLargeFieldSql = sPcSQL
            Case ctItem
                GetLargeFieldSql = sItmSQL
        End Select
        
    End If
    
    'We order by catalogue level, then reference ID as
    'this is the same ordering used for the main recordset
    'and we need to keep this in synch with that
    GetLargeFieldSql = GetLargeFieldSql & " ORDER BY " & _
            gk_sDICatalogueLevel & ", " & gk_sDIID
            
    Exit Function
    
ErrorHandler:
    ConvertError sSource
    
End Function


Private Function GetLargeFieldItemsInPieceSQL(t_lPieceID As Long, t_eField As ELargeField) As String
'   Generate the sql to select fields from live
'   or live tables that are large so require a separate record set
'   This is done for items below the given piece
'
'   d6  WDP     19-Nov-1999     Created
On Error GoTo ErrorHandler

    Dim sSource As String
    sSource = "BOCatalogueIndexer.GetLargeFieldItemsInPieceSQL"
    
    Dim sItmTableName As String
    Dim sZoneDI As String, sZoneFieldName As String
    
    'Set table, data item and column names depending on the zone
    Select Case t_eField
    Case LARGETitle
        sItmTableName = gk_sDBTableItemTitle
        sZoneDI = gk_sDITitle
        sZoneFieldName = gk_sDBCOLTitle
    Case LARGEScopeContent
        sItmTableName = gk_sDBTableItemScopeAndContent
        sZoneDI = gk_sDIScopeContent
        sZoneFieldName = gk_sDBCOLScopeAndContent
    Case LARGEAdminHistory
        sItmTableName = gk_sDBTABLEItemAdminBiogHist
        sZoneDI = gk_sDIAdminHistory
        sZoneFieldName = gk_sDBCOLAdministrativeHistory
    Case LARGEAccruals
        sItmTableName = gk_sDBTABLEItemAccruals
        sZoneDI = gk_sDIAccruals
        sZoneFieldName = gk_sDBCOLAccruals
    Case LARGEArrangement
        sItmTableName = gk_sDBTABLEItemArrangement
        sZoneDI = gk_sDIArrangement
        sZoneFieldName = gk_sDBCOLArrangement
    Case LARGECustodialHistory
        sItmTableName = gk_sDBTABLEItemCustodialHist
        sZoneDI = gk_sDICustodialHist
        sZoneFieldName = gk_sDBCOLCustodialHistory
    Case LARGEAppraisalInfo
        sItmTableName = gk_sDBTABLEItemAppDestInfo
        sZoneDI = gk_sDIAppraisalInfo
        sZoneFieldName = gk_sDBCOLAppraisalDestructionInformation
    Case LARGENote
        sItmTableName = gk_sDBTABLEItemNote
        sZoneDI = gk_sDINoteText
        sZoneFieldName = gk_sDBCOLNoteText
    Case Else
        RaiseError sSource, PEGeneralInternalError, "Invalid zone enum."
    End Select
    
    Dim sSQL As String
    
    'Lettercodes
    
    sSQL = "SELECT " & _
            "7 AS " & gk_sDICatalogueLevel & ", " & _
            "i." & gk_sDBCOLItemID & " AS " & gk_sDIID & ", " & _
            "a." & sZoneFieldName & " AS " & sZoneDI & _
            " FROM " & _
            gk_sDBTABLEItem & " i, " & _
            sItmTableName & " a " & _
            " WHERE " & _
            "i." & gk_sDBCOLItemID & "=a." & gk_sDBCOLItemID & _
            " AND i." & gk_sDBCOLPieceID & "=" & t_lPieceID
            

    'We order by catalogue level, then reference ID as
    'this is the same ordering used for the main recordset
    'and we need to keep this in synch with that
    GetLargeFieldItemsInPieceSQL = sSQL & " ORDER BY " & _
            gk_sDICatalogueLevel & ", " & gk_sDIID
            
    Exit Function
    
ErrorHandler:
    ConvertError sSource
    
End Function

Private Function LinksSQL( _
            Optional t_bFromLive As Boolean = False, _
            Optional t_eLevel As ECATALOGUELEVEL, _
            Optional t_lLettercodeID As Long = 0, _
            Optional t_lClassID As Long = 0, _
            Optional t_lHeaderID As Long = 0, _
            Optional t_lPieceID As Long = 0, _
            Optional t_sLettercodeIn As String = "", _
            Optional t_sClassIn As String = "") As String
'   Function to generate SQL to get Link data for an edit set
'   or from live tables. The optional parmaters indicate if the data is
'   to be selected from live tables (default is from edit sets)
'   and which level and lettercode to select from.
'   When used with a piece ID specified this method will
'   return the sql to get all the child items's links for that piece
'
'   d2  WDP     25-Oct-1999     Created
'   d3  WDP     29-Oct-1999     Added class and header optional params
'   d6  WDP     19-Nov-1999     Added piece ID parameter for moving pieces
'   d8  WDP     13-Dec-1999     Allow lettercode id to be optional
'                               if none is given all lettercodes are selected from
'                               Added lettercodeIn parameter to allow
'                               selection from multiple lettercodes
On Error GoTo ErrorHandler

    Dim sSource As String
    sSource = "BOCatalogueIndexer.LinksSQL"
    Dim sSQL As String
    
    Dim sPersonTable As String, sPlaceTable As String
    Dim sCorpbodyTable As String, sSubjTable As String
    If t_bFromLive = True Then
        sPersonTable = gk_sDBTABLEPCPerson
        sPlaceTable = gk_sDBTABLEPCPlace
        sCorpbodyTable = gk_sDBTABLEPCCorporateBody
        sSubjTable = gk_sDBTABLEPCSubject
    Else
        sPersonTable = gk_sDBTABLEEditSetPCPerson
        sPlaceTable = gk_sDBTABLEEditsetPCPlace
        sCorpbodyTable = gk_sDBTABLEEditSetPCCorporateBody
        sSubjTable = gk_sDBTABLEEditsetPCSubject
    End If
        
    Dim sExtraLiveTables As String
    If t_bFromLive = True Then
        If t_lPieceID = 0 Then
            Select Case t_eLevel
                Case ctLetterCode
                    sExtraLiveTables = ", " & gk_sDBTABLELCode & " lc "
                Case ctDivision
                    sExtraLiveTables = ", " & gk_sDBTABLELCode & " lc " & _
                            ", " & gk_sDBTABLEDivision & " d "
                Case ctClass
                    sExtraLiveTables = ", " & gk_sDBTABLELCode & " lc " & _
                            ", " & gk_sDBTABLEClass & " c "
                Case ctHeaders
                    sExtraLiveTables = ", " & gk_sDBTABLELCode & " lc " & _
                            ", " & gk_sDBTABLEClass & " c " & _
                            ", " & gk_sDBTABLEHeader & " h "
                Case ctSubHeaders
                    sExtraLiveTables = ", " & gk_sDBTABLELCode & " lc " & _
                            ", " & gk_sDBTABLEClass & " c " & _
                            ", " & gk_sDBTABLEHeader & " h " & _
                            ", " & gk_sDBTABLESubHeader & " sh "
                Case ctPiece
                    sExtraLiveTables = ", " & gk_sDBTABLELCode & " lc " & _
                            ", " & gk_sDBTABLEClass & " c " & _
                            ", " & gk_sDBTABLEPiece & " p "
                    If Not t_lHeaderID = 0 Then
                        sExtraLiveTables = sExtraLiveTables & _
                            ", " & gk_sDBTABLEHeader & " h "
                    End If
                Case ctItem
                    sExtraLiveTables = ", " & gk_sDBTABLELCode & " lc " & _
                            ", " & gk_sDBTABLEClass & " c " & _
                            ", " & gk_sDBTABLEPiece & " p " & _
                            ", " & gk_sDBTABLEItem & " i "
                    If Not t_lHeaderID = 0 Then
                        sExtraLiveTables = sExtraLiveTables & _
                            ", " & gk_sDBTABLEHeader & " h "
                    End If
                Case Else
                    RaiseError sSource, PEGeneralInternalError, "invalid cat. level enum."
            End Select
        Else
            sExtraLiveTables = ", " & gk_sDBTABLEItem & " i "
        End If
    End If
    
    Dim sWhereClause As String
    If t_bFromLive = True Then
        If t_lPieceID = 0 Then
            If t_sLettercodeIn = "" Then
                If Not t_lLettercodeID = 0 Then
                    sWhereClause = " lc." & gk_sDBCOLLetterCodeID & _
                                   "=" & t_lLettercodeID
                Else
                'IGNORE: Dummy clause to keep shape of query
                    sWhereClause = " lc." & gk_sDBCOLLetterCodeID & _
                                   "<>0"
                End If
            Else
                sWhereClause = " lc." & gk_sDBCOLLetterCodeID & _
                               " IN " & t_sLettercodeIn
           
            End If
            
            Select Case t_eLevel
                Case ctLetterCode
                    sWhereClause = sWhereClause & _
                        " AND t." & gk_sDBCOLCatalogueLevel & "=1" & _
                        " AND t." & gk_sDBCOLCatalogueID & "=lc." & gk_sDBCOLLetterCodeID
                Case ctDivision
                    sWhereClause = sWhereClause & _
                        " AND t." & gk_sDBCOLCatalogueLevel & "=2" & _
                        " AND t." & gk_sDBCOLCatalogueID & "=d." & gk_sDBCOLDivisionID & _
                        " AND d." & gk_sDBCOLLetterCodeID & "=lc." & gk_sDBCOLLetterCodeID
                Case ctClass
                    sWhereClause = sWhereClause & _
                        " AND t." & gk_sDBCOLCatalogueLevel & "=3" & _
                        " AND t." & gk_sDBCOLCatalogueID & "=c." & gk_sDBCOLClassID & _
                        " AND c." & gk_sDBCOLLetterCodeID & "=lc." & gk_sDBCOLLetterCodeID
                Case ctHeaders
                    sWhereClause = sWhereClause & _
                        " AND t." & gk_sDBCOLCatalogueLevel & "=4" & _
                        " AND t." & gk_sDBCOLCatalogueID & "=h." & gk_sDBCOLHeaderID & _
                        " AND c." & gk_sDBCOLLetterCodeID & "=lc." & gk_sDBCOLLetterCodeID & _
                        " AND h." & gk_sDBCOLClassID & "=c." & gk_sDBCOLClassID
                Case ctSubHeaders
                    sWhereClause = sWhereClause & _
                        " AND t." & gk_sDBCOLCatalogueLevel & "=5" & _
                        " AND t." & gk_sDBCOLCatalogueID & "=sh." & gk_sDBCOLSubHeaderID & _
                        " AND c." & gk_sDBCOLLetterCodeID & "=lc." & gk_sDBCOLLetterCodeID & _
                        " AND h." & gk_sDBCOLClassID & "=c." & gk_sDBCOLClassID & _
                        " AND sh." & gk_sDBCOLHeaderID & "=h." & gk_sDBCOLHeaderID
                Case ctPiece
                    sWhereClause = sWhereClause & _
                        " AND t." & gk_sDBCOLCatalogueLevel & "=6" & _
                        " AND t." & gk_sDBCOLCatalogueID & "=p." & gk_sDBCOLPieceID & _
                        " AND c." & gk_sDBCOLLetterCodeID & "=lc." & gk_sDBCOLLetterCodeID & _
                        " AND p." & gk_sDBCOLClassID & "=c." & gk_sDBCOLClassID
                Case ctItem
                    sWhereClause = sWhereClause & _
                        " AND t." & gk_sDBCOLCatalogueLevel & "=7" & _
                        " AND t." & gk_sDBCOLCatalogueID & "=i." & gk_sDBCOLItemID & _
                        " AND c." & gk_sDBCOLLetterCodeID & "=lc." & gk_sDBCOLLetterCodeID & _
                        " AND p." & gk_sDBCOLClassID & "=c." & gk_sDBCOLClassID & _
                        " AND i." & gk_sDBCOLPieceID & "=p." & gk_sDBCOLPieceID
                Case Else
                    RaiseError sSource, PEGeneralInternalError, "invalid cat. level enum."
            End Select
        Else
            sWhereClause = "t." & gk_sDBCOLCatalogueLevel & "=7" & _
                        " AND t." & gk_sDBCOLCatalogueID & "=i." & gk_sDBCOLItemID & _
                        " AND i." & gk_sDBCOLPieceID & "=" & t_lPieceID
                        
        End If
        If t_sClassIn = "" Then
            If Not t_lClassID = 0 Then
                sWhereClause = sWhereClause & _
                        " AND c." & gk_sDBCOLClassID & "=" & t_lClassID
            End If
        Else
            sWhereClause = sWhereClause & _
                    " AND c." & gk_sDBCOLClassID & " IN " & t_sClassIn
        End If
        
    Else
        sWhereClause = " t." & gk_sDBCOLEditSetID & "=" & m_lEditsetID
    End If
    
    'Person links
    sSQL = sSQL & " SELECT " & Quotes(gk_sPersonTermPrefix) & "+" & _
        " CONVERT(VARCHAR," & _
        gk_sDBCOLPersonReferenceID & ") AS " & gk_sDILinkDataID & ", " & _
        gk_sDBCOLCatalogueID & " AS " & gk_sDICatalogueID & ", " & _
        gk_sDBCOLCatalogueLevel & " AS " & gk_sDICatalogueLevel & ", " & _
        gk_sDBCOLPropertyID & " AS " & gk_sDILinkTypeID & _
        " FROM " & sPersonTable & " t " & sExtraLiveTables & _
        " WHERE " & sWhereClause & _
        " AND (" & _
        gk_sDBCOLPropertyID & "=" & mk_lCreatorPropID & _
        " OR " & gk_sDBCOLPropertyID & "=" & mk_lPersonPropID & _
        " OR " & gk_sDBCOLPropertyID & "=" & mk_lSourceOfAcqPropID & ")"
        
        
    'Corp bodies links
    sSQL = sSQL & " UNION SELECT " & Quotes(gk_sCorpBodyTermPrefix) & "+" & _
        " CONVERT(VARCHAR," & _
        gk_sDBCOLCorporateBodyReferenceID & ") AS " & gk_sDILinkDataID & ", " & _
        gk_sDBCOLCatalogueID & " AS " & gk_sDICatalogueID & ", " & _
        gk_sDBCOLCatalogueLevel & " AS " & gk_sDICatalogueLevel & ", " & _
        gk_sDBCOLPropertyID & " AS " & gk_sDILinkTypeID & _
        " FROM " & sCorpbodyTable & " t " & sExtraLiveTables & _
        " WHERE " & sWhereClause & _
        " AND (" & _
        gk_sDBCOLPropertyID & "=" & mk_lCreatorPropID & _
        " OR " & gk_sDBCOLPropertyID & "=" & mk_lCorpBodyPropID & _
        " OR " & gk_sDBCOLPropertyID & "=" & mk_lPlaceOfDepPropID & _
        " OR " & gk_sDBCOLPropertyID & "=" & mk_lLocOfOrigPropID & _
        " OR " & gk_sDBCOLPropertyID & "=" & mk_lCopiesPropID & _
        " OR " & gk_sDBCOLPropertyID & "=" & mk_lSourceOfAcqPropID & ")"
        
    'Place links
    sSQL = sSQL & " UNION SELECT " & Quotes(gk_sPlaceTermPrefix) & "+" & _
        " CONVERT(VARCHAR," & _
        gk_sDBCOLPlaceReferenceID & ") AS " & gk_sDILinkDataID & ", " & _
        gk_sDBCOLCatalogueID & " AS " & gk_sDICatalogueID & ", " & _
        gk_sDBCOLCatalogueLevel & " AS " & gk_sDICatalogueLevel & ", " & _
        gk_sDBCOLPropertyID & " AS " & gk_sDILinkTypeID & _
        " FROM " & sPlaceTable & " t " & sExtraLiveTables & _
        " WHERE " & sWhereClause & _
        " AND " & gk_sDBCOLPropertyID & "=" & mk_lPlacePropID
    
    'Subject links
    sSQL = sSQL & " UNION SELECT " & Quotes(gk_sSubjectTermPrefix) & "+" & _
        " CONVERT(VARCHAR," & _
        gk_sDBCOLSubjectReferenceID & ") AS " & gk_sDILinkDataID & ", " & _
        gk_sDBCOLCatalogueID & " AS " & gk_sDICatalogueID & ", " & _
        gk_sDBCOLCatalogueLevel & " AS " & gk_sDICatalogueLevel & ", " & _
        gk_sDBCOLPropertyID & " AS " & gk_sDILinkTypeID & _
        " FROM " & sSubjTable & " t " & sExtraLiveTables & _
        " WHERE " & sWhereClause & _
        " AND " & gk_sDBCOLPropertyID & "=" & mk_lSubjectPropID
    
    'We order the recordset in such a way to aid in extracting the data
    sSQL = sSQL & " ORDER BY " & gk_sDICatalogueLevel & ", " & _
            gk_sDICatalogueID & ", " & gk_sDILinkTypeID
    
    LinksSQL = sSQL

    Exit Function
ErrorHandler:
    ConvertError sSource
    
End Function

Private Function WriteToFile( _
            t_sFilePrefix As String, _
            Optional t_sLettercode As String = "", _
            Optional t_eCollectionType As ECollectionType) As Boolean
            
'   Process the data in the record sets to write to bulk
'   and sgml files. The format for
'   the non-basic collections is the same as for basic but with some
'   additional fields and zones
'   The data is written to different files for edit sets and
'   live data these two cases are distinguished by the parameters:
'   t_sFilePrefix       The file name prefix to use in creating files
'   t_sLettercode       The lettercode we are indexing
'   t_eCollectionType   The collection type we are indexing for
'
'   If the function does not find any appropriate data to write to
'   the file it will return False otherwise True
'
'   d2  WDP     26-Oct-1999     Created from WriteToFile and WriteLiveDataToFile
'   d3  WDP     29-Oct-1999      Added class and header optional params
'   d5  WDP     19-Nov-1999     Only add classes and above for ClassUp collection
'   d8  WDP     19-Nov-1999     Allowed t_sLettercode to be optional if it
'                               is omitted then all records are written
'   d10 WDP     2-Mar-2000      Convert DB NULLs into defined values for numeric fields
'                                   in WriteToFile (Verity does not have NULLs)
'   d20 TomH    15-Aug-2002     Altered so now it includes <basic> in SGML
'                               This allows us to do basic searches over advanced catalogue
On Error GoTo ErrorHandler

    Dim sSource As String
    sSource = "BOCatalogueIndexer.WriteToFile"
    Dim lSGMLFileNum As Long, lBLKFileNum As Long
    Dim bHaveWrittenData As Boolean
    bHaveWrittenData = False
    Dim sIndexFileDrive As String
    Dim sCollectionDrive As String
    
    Dim ofwAppData As New FWApplicationData
    sIndexFileDrive = ofwAppData.GetApplicationConfigurationInfomation(pRootForIndexingFiles)
    sCollectionDrive = ofwAppData.GetApplicationConfigurationInfomation(pRootForCollections)
    Set ofwAppData = Nothing
    
    'If we have no data we dont want to write to file
    If m_ofwMainData.RecordCount = 0 Then
        WriteToFile = False
        Exit Function
    End If
    
    'create filenames
    Dim sFilename As String, sSGMLFilename As String, sBLKFilename As String
    
    'sFilename = GetFileName(t_sFileDatePrefix, t_bFromLive, _
            t_eCollectionType, t_sLettercode, t_eAction, t_eLevel, t_sClassRef, t_sHeaderRef)
    
    sSGMLFilename = t_sFilePrefix & ".sgml"
    sBLKFilename = t_sFilePrefix & ".blk"
    
    'Open the files, if they already exist they will be overwritten
    '+make read only
    lSGMLFileNum = FreeFile
    Open (sIndexFileDrive & mk_sFilepath & sSGMLFilename) For Output As lSGMLFileNum
    If lSGMLFileNum = 0 Then
        RaiseError sSource, PEGeneralInternalError, "Unable to open sgml file"
    End If
    
    lBLKFileNum = FreeFile
    Open (sIndexFileDrive & mk_sFilepath & sBLKFilename) For Output As lBLKFileNum
    If lBLKFileNum = 0 Then
        RaiseError sSource, PEGeneralInternalError, "Unable to open blk file"
    End If
    
    'We need to calculate offsets and lengths when producing the
    'sgml and blk files
    Dim lOffset As Long
    Dim lDocSize As Long
    lOffset = 0
    lDocSize = 0
    
    If Not m_ofwMainData.RecordCount = 0 Then
        m_ofwMainData.MoveFirst
        m_ofwTitleData.MoveFirst
        m_ofwScopeData.MoveFirst
        m_ofwAdminData.MoveFirst
        
        If t_eCollectionType = COLAdvanced Or t_eCollectionType = COLClassUp Then
            m_ofwArrangeData.MoveFirst
            m_ofwCustHistData.MoveFirst
            m_ofwAppInfoData.MoveFirst
            m_ofwAccrualData.MoveFirst
            m_ofwNoteData.MoveFirst
            If Not m_ofwLinkData.RecordCount = 0 Then
                m_ofwLinkData.MoveFirst
            End If
        End If
    End If
    
    Do While Not m_ofwMainData.IsAtEnd
        
        Dim sLetterCode As String, eLevel As ECATALOGUELEVEL
        sLetterCode = VariantStringAsString(m_ofwMainData.GetFieldValue(gk_sDILetterCodeRef))
        eLevel = CInt(m_ofwMainData.GetFieldValue(gk_sDICatalogueLevel))
        'skip records until we find those with the lettercode
        'we want and is we want classes upwards only then skip
        'records which are below class level
        While (Not (sLetterCode = t_sLettercode Or t_sLettercode = "")) Or (t_eCollectionType = COLClassUp And eLevel > ctClass)
            m_ofwMainData.MoveNext
            m_ofwScopeData.MoveNext
            m_ofwTitleData.MoveNext
            m_ofwAdminData.MoveNext
            If t_eCollectionType = COLAdvanced Or t_eCollectionType = COLClassUp Then
                m_ofwArrangeData.MoveNext
                m_ofwCustHistData.MoveNext
                m_ofwAppInfoData.MoveNext
                m_ofwAccrualData.MoveNext
                m_ofwNoteData.MoveNext
            End If
            If m_ofwMainData.IsAtEnd Then
                Exit Do
            End If
            
            eLevel = CInt(m_ofwMainData.GetFieldValue(gk_sDICatalogueLevel))
            sLetterCode = VariantStringAsString(m_ofwMainData.GetFieldValue(gk_sDILetterCodeRef))
        Wend
        
        ' Get the fields into variables
        Dim sDivisionno As String
        Dim sClassNo As String
        Dim sSubclassNo As String
        Dim sHeaderNo As String
        Dim sSubheaderNo As String
        Dim sPieceRef As String
        Dim sItemRef As String
        Dim sScope As String
        Dim sTitle As String
        Dim sHistory As String
        Dim sDateText As String
        Dim sFirstDate As String
        Dim sLastDate As String
        Dim sCatLevel As String
        Dim sItemKeyOrder As String
        Dim sPieceKeyOrder As String
        Dim sFormerRef As String
        Dim sFormerPRORef As String
        Dim sStrippedFormerRef As String
        Dim sStrippedFormerPRORef As String
        Dim sStrippedRef As String
        Dim sTitleScope As String
        Dim sRef As String
        Dim sVdkVgwKey As String
        Dim iLevel As Integer, lCatID As Long
        Dim sMapDesig As String
                
        '+TO DO Add check to see record sets are in synch (using ID field)
        sLetterCode = VariantStringAsString(m_ofwMainData.GetFieldValue(gk_sDILetterCodeRef))
        sDivisionno = VariantStringAsString(m_ofwMainData.GetFieldValue(gk_sDIDivisionRef))
        'For numeric fields that are NULL in the DB we set them to the defined
        'Verity equivalent (there are no nulls in Verity) if this
        'is different to 0 (which is what verity uses as a default numeric value)
        sClassNo = VariantStringAsString(m_ofwMainData.GetFieldValue(gk_sDIClassRef))
        If sClassNo = "" And gk_lVerityNumericNullValue <> 0 Then
            sClassNo = gk_lVerityNumericNullValue
        End If
        sSubclassNo = VariantStringAsString(m_ofwMainData.GetFieldValue(gk_sDISubclassRef))
        If sSubclassNo = "" And gk_lVerityNumericNullValue <> 0 Then
            sSubclassNo = gk_lVerityNumericNullValue
        End If
        sHeaderNo = VariantStringAsString(m_ofwMainData.GetFieldValue(gk_sDIHeaderRef))
        sSubheaderNo = VariantStringAsString(m_ofwMainData.GetFieldValue(gk_sDISubheaderRef))
        sPieceRef = VariantStringAsString(m_ofwMainData.GetFieldValue(gk_sDIPieceRef))
        sItemRef = VariantStringAsString(m_ofwMainData.GetFieldValue(gk_sDIItemRef))
        sItemKeyOrder = VariantStringAsString(m_ofwMainData.GetFieldValue(gk_sDIItemKeyOrder))
        sPieceKeyOrder = VariantStringAsString(m_ofwMainData.GetFieldValue(gk_sDIPieceKeyOrder))
        sCatLevel = VariantStringAsString(m_ofwMainData.GetFieldValue(gk_sDICatalogueLevel))
        sFormerRef = VariantStringAsString(m_ofwMainData.GetFieldValue(gk_sDIFormerRef))
        sFormerPRORef = VariantStringAsString(m_ofwMainData.GetFieldValue(gk_sDIFormerPRORef))
        eLevel = CInt(sCatLevel)
        iLevel = m_ofwMainData.GetFieldValue(gk_sDICatalogueLevel)
        lCatID = m_ofwMainData.GetFieldValue(gk_sDIID)
            
        'The titlescope field is generated from the scope of pieces
        'and items or the title of records from other levels
        Select Case eLevel
            Case ctItem, ctPiece
                sTitleScope = TruncateTitle(VariantStringAsString( _
                    m_ofwScopeData.GetFieldValue(gk_sDIScopeContent)))
            Case Else
                sTitleScope = TruncateTitle(VariantStringAsString( _
                    m_ofwTitleData.GetFieldValue(gk_sDITitle)))
        End Select
        
        sScope = ProcessZone(VariantStringAsString( _
            m_ofwScopeData.GetFieldValue(gk_sDIScopeContent)), gk_sVZONEScopeContent)
        sTitle = ProcessZone(VariantStringAsString( _
            m_ofwTitleData.GetFieldValue(gk_sDITitle)), gk_sVZONETitle)
        sHistory = ProcessZone(VariantStringAsString( _
            m_ofwAdminData.GetFieldValue(gk_sDIAdminHistory)), gk_sVZONEAdminHistory)
        sMapDesig = ProcessZone(VariantStringAsString( _
                m_ofwMainData.GetFieldValue(gk_sDIMapDesignation)), gk_sVZONEMapDesignation)
            
        sDateText = VariantStringAsString(m_ofwMainData.GetFieldValue(gk_sDICoveringDates))
        sFirstDate = PadDate(VariantStringAsString(m_ofwMainData.GetFieldValue(gk_sDIFirstDate)))
        sLastDate = PadDate(VariantStringAsString(m_ofwMainData.GetFieldValue(gk_sDILastDate)))
        
        sRef = BuildReference(sLetterCode, sClassNo, sSubclassNo, sPieceRef, sItemRef, eLevel)
        
        'Divisions, subheaders and headers do not have references
        'so we set the stripped_ref field to be empty for them
        'so they are not hit by mistake when searching for references
        'For example searching for 'E' would return the lettercode E and
        'all divisions within E as their stripped references would all
        'be 'E'
        Select Case eLevel
            Case ctDivision, ctHeaders, ctSubHeaders
                sStrippedRef = ""
            Case Else
                sStrippedRef = Replace(sRef, " ", "")
                sStrippedRef = Replace(sStrippedRef, vbTab, "")
                sStrippedRef = Replace(sStrippedRef, vbNewLine, "")
        End Select
        
        'Put into global function
        sStrippedFormerPRORef = Replace(sFormerPRORef, " ", "")
        sStrippedFormerPRORef = Replace(sStrippedFormerPRORef, vbTab, "")
        sStrippedFormerPRORef = Replace(sStrippedFormerPRORef, vbNewLine, "")
        sStrippedFormerRef = Replace(sFormerRef, " ", "")
        sStrippedFormerRef = Replace(sStrippedFormerRef, vbTab, "")
        sStrippedFormerRef = Replace(sStrippedFormerRef, vbNewLine, "")
        
        sVdkVgwKey = VerityDocKey(lCatID, eLevel)

        Dim sSGMLText As String, sBulkText As String
        
        sSGMLText = "<entry>" & vbCrLf & "<basic>" & vbCrLf & sTitle & vbCrLf & sScope & _
                    vbCrLf & sHistory & vbCrLf & sMapDesig & vbCrLf & "</basic>" & vbCrLf
                    
        If t_eCollectionType = COLAdvanced Or t_eCollectionType = COLClassUp Then
            Dim sArrangement As String
            Dim sPhysicalCondition As String
            Dim sRestrictions As String
            Dim sCustHistory As String
            Dim sAppraisalInfo As String
            Dim sAccruals As String
            Dim sNote As String
            
            'Get the extra zone data for the advanced collection
            
            sArrangement = ProcessZone(VariantStringAsString( _
                m_ofwArrangeData.GetFieldValue(gk_sDIArrangement)), gk_sVZONEArrangement)
            sCustHistory = ProcessZone(VariantStringAsString( _
                m_ofwCustHistData.GetFieldValue(gk_sDICustodialHist)), gk_sVZONECustodialHistory)
            sNote = ProcessZone(VariantStringAsString( _
                m_ofwNoteData.GetFieldValue(gk_sDINoteText)), gk_sVZONENote)
            sAccruals = ProcessZone(VariantStringAsString( _
                m_ofwAccrualData.GetFieldValue(gk_sDIAccruals)), gk_sVZONEAccruals)
            sAppraisalInfo = ProcessZone(VariantStringAsString( _
                m_ofwAppInfoData.GetFieldValue(gk_sDIAppraisalInfo)), gk_sVZONEAppDestInfo)
            sPhysicalCondition = ProcessZone(VariantStringAsString( _
                m_ofwMainData.GetFieldValue(gk_sDIPhysicalCondition)), gk_sVZONEPhysicalCondition)
            sRestrictions = ProcessZone(VariantStringAsString( _
                m_ofwMainData.GetFieldValue(gk_sDIRestrictions)), gk_sVZONERestrictions)
            sSGMLText = sSGMLText & sArrangement & vbCrLf & _
                    sPhysicalCondition & vbCrLf & _
                    sRestrictions & vbCrLf & sCustHistory & vbCrLf & _
                    sAppraisalInfo & vbCrLf & sAccruals & vbCrLf & sNote & vbCrLf
        End If
        
        sSGMLText = sSGMLText & "</entry>" & vbCrLf
        
        'Write to the sgml file
        Print #lSGMLFileNum, sSGMLText
        lDocSize = Len(sSGMLText)
        bHaveWrittenData = True
        
        sBulkText = _
            FieldLine(gk_sVFIELDVerityKey, DoubleQuotes(sVdkVgwKey)) & _
            FieldLine("DOC_FN", Replace(DoubleQuotes(sIndexFileDrive & mk_sFilepathForVerity & sSGMLFilename), "\", "/")) & _
            "DOC_OF: " & lOffset & vbCrLf & _
            "DOC_SZ: " & lDocSize & vbCrLf & _
            FieldLine(gk_sVFIELDCatalogueLevel, sCatLevel) & _
            FieldLine(gk_sVFIELDItemRef, Replace(DoubleQuotes(sItemRef), vbNewLine, " ")) & _
            FieldLine(gk_sVFIELDItemKeyOrder, sItemKeyOrder) & _
            FieldLine(gk_sVFIELDPieceRef, DoubleQuotes(sPieceRef)) & _
            FieldLine(gk_sVFIELDPieceKeyOrder, sPieceKeyOrder) & _
            FieldLine(gk_sVFIELDSubheaderRef, sSubheaderNo) & _
            FieldLine(gk_sVFIELDHeaderRef, sHeaderNo) & _
            FieldLine(gk_sVFIELDSubclassRef, sSubclassNo) & _
            FieldLine(gk_sVFIELDClassRef, sClassNo) & _
            FieldLine(gk_sVFIELDDivisionRef, sDivisionno) & _
            FieldLine(gk_sVFIELDLetterCodeRef, DoubleQuotes(sLetterCode)) & _
            FieldLine(gk_sVFIELDFirstDate, sFirstDate) & _
            FieldLine(gk_sVFIELDLastDate, sLastDate) & _
            FieldLine(gk_sVFIELDCoveringDates, Replace(DoubleQuotes(sDateText), vbNewLine, " ")) & _
            FieldLine(gk_sVFIELDTitleOrScope, DoubleQuotes(sTitleScope)) & _
            FieldLine(gk_sVFIELDFormerRef, Replace(DoubleQuotes(sFormerRef), vbNewLine, " ")) & _
            FieldLine(gk_sVFIELDFormerPRORef, Replace(DoubleQuotes(sFormerPRORef), vbNewLine, " ")) & _
            FieldLine(gk_sVFIELDStrippedFormerRef, DoubleQuotes(sStrippedFormerRef)) & _
            FieldLine(gk_sVFIELDStrippedFormerPRORef, DoubleQuotes(sStrippedFormerPRORef)) & _
            FieldLine(gk_sVFIELDStrippedRef, DoubleQuotes(sStrippedRef))
            
        If t_eCollectionType = COLAdvanced Or t_eCollectionType = COLClassUp Then
        
            Dim sMapScale As String
            Dim sAccumDates As String
            Dim sOpenDate As String
            Dim sLegalStatus As String
            Dim sPhysDescForm As String
            Dim sPhysDescQuan As String
            Dim sPhysDescDims As String
            Dim sAccessCond As String
            Dim sClosureStat As String
            Dim sClosureType As String
            Dim sClosureCode As String
            Dim sLanguage As String
            
            sMapScale = VariantStringAsString( _
                m_ofwMainData.GetFieldValue(gk_sDIMapScale))
            If sMapScale = "" And gk_lVerityNumericNullValue <> 0 Then
                sMapScale = gk_lVerityNumericNullValue
            End If
            sAccumDates = VariantStringAsString( _
                m_ofwMainData.GetFieldValue(gk_sDIAccumulationDates))
            sOpenDate = VariantDateAsYYYYMMDD( _
                m_ofwMainData.GetFieldValue(gk_sDIOpeningDates))
            sLegalStatus = VariantStringAsString( _
                m_ofwMainData.GetFieldValue(gk_sDILegalStatus))
            If sLegalStatus = "" And gk_lVerityLookupNullValue <> 0 Then
                sLegalStatus = gk_lVerityLookupNullValue
            End If
            sPhysDescForm = VariantStringAsString( _
                m_ofwMainData.GetFieldValue(gk_sDIPhysicalDescForm))
            If sPhysDescForm = "" And gk_lVerityLookupNullValue <> 0 Then
                sPhysDescForm = gk_lVerityLookupNullValue
            End If
            sPhysDescQuan = VariantStringAsString( _
                m_ofwMainData.GetFieldValue(gk_sDIPhysDescQuantity))
            If sPhysDescQuan = "" And gk_lVerityNumericNullValue <> 0 Then
                sPhysDescQuan = gk_lVerityNumericNullValue
            End If
            sPhysDescDims = VariantStringAsString( _
                m_ofwMainData.GetFieldValue(gk_sDIPhysDescDimensions))
            sAccessCond = VariantStringAsString( _
                m_ofwMainData.GetFieldValue(gk_sDIAccessCondition))
            If sAccessCond = "" And gk_lVerityLookupNullValue <> 0 Then
                sAccessCond = gk_lVerityLookupNullValue
            End If
            sClosureStat = VariantStringAsString( _
                m_ofwMainData.GetFieldValue(gk_sDIClosureStatus))
            
            'Closure status can be " " as a valid value!
            'so we need to use an alternative for NULLs
            If IsNull(m_ofwMainData.GetFieldValue(gk_sDIClosureType)) Then
                sClosureType = mk_sNeverUsedClosureType
            Else
                sClosureType = Trim(VariantStringAsString(m_ofwMainData.GetFieldValue(gk_sDIClosureType)))
            End If
            
            sClosureCode = VariantStringAsString( _
                m_ofwMainData.GetFieldValue(gk_sDIClosureCode))
            If sClosureCode = "" And gk_lVerityLookupNullValue <> 0 Then
                sClosureCode = gk_lVerityLookupNullValue
            End If
            sLanguage = VariantStringAsString( _
                m_ofwMainData.GetFieldValue(gk_sDILanguage))
            If sLanguage = "" And gk_lVerityLookupNullValue <> 0 Then
                sLanguage = gk_lVerityLookupNullValue
            End If
            
            sBulkText = sBulkText & _
                FieldLine(gk_sVFIELDMapScale, sMapScale) & _
                FieldLine(gk_sVFIELDAccumulationDates, Replace(DoubleQuotes(sAccumDates), vbNewLine, " ")) & _
                FieldLine(gk_sVFIELDOpenDate, sOpenDate) & _
                FieldLine(gk_sVFIELDLegalStatus, sLegalStatus) & _
                FieldLine(gk_sVFIELDPhysicalDescForm, sPhysDescForm) & _
                FieldLine(gk_sVFIELDPhysDescQuantity, sPhysDescQuan) & _
                FieldLine(gk_sVFIELDPhysDescDimensions, Replace(DoubleQuotes(sPhysDescDims), vbNewLine, " ")) & _
                FieldLine(gk_sVFIELDAccessConditions, sAccessCond) & _
                FieldLine(gk_sVFIELDClosureStatus, Replace(DoubleQuotes(sClosureStat), vbNewLine, " ")) & _
                FieldLine(gk_sVFIELDClosureType, Replace(DoubleQuotes(sClosureType), vbNewLine, " ")) & _
                FieldLine(gk_sVFIELDClosureCode, sClosureCode) & _
                FieldLine(gk_sVFIELDLanguage, sLanguage)
            'Add link data ids
            sBulkText = sBulkText & LinkBulkData(iLevel, lCatID)
            'sBulkText = sBulkText & LinkBulkData(AFLINKCorporateName, iLevel, lCatID)
            'sBulkText = sBulkText & LinkBulkData(AFLINKCreatorName, iLevel, lCatID)
            'sBulkText = sBulkText & LinkBulkData(AFLINKLocOfOriginals, iLevel, lCatID)
            'sBulkText = sBulkText & LinkBulkData(AFLINKPerson, iLevel, lCatID)
            'sBulkText = sBulkText & LinkBulkData(AFLINKPlace, iLevel, lCatID)
            'sBulkText = sBulkText & LinkBulkData(AFLINKPlaceOfDeposit, iLevel, lCatID)
            'sBulkText = sBulkText & LinkBulkData(AFLINKSourceOfAcq, iLevel, lCatID)
            'sBulkText = sBulkText & LinkBulkData(AFLINKSubject, iLevel, lCatID)

        End If
        
        sBulkText = sBulkText & "<<EOD>>" & vbCrLf & vbCrLf
        
        'print to the blk file
        Print #lBLKFileNum, sBulkText
        
        '+for testing completeness of indexing - every entry written
        'to file is logged in a DB table
        '!remove this call after testing as it will effect move (&? release)
        'and reduce performance of indexing
        
        'If m_lEditsetID = 0 Then
        '    LogEntryWritten iLevel, lCatID
        'End If
        
        ' Update offset
        '! Don't know why I have to add 2!
        lOffset = lOffset + lDocSize + 2
        
        m_ofwMainData.MoveNext
        m_ofwScopeData.MoveNext
        m_ofwTitleData.MoveNext
        m_ofwAdminData.MoveNext
        If t_eCollectionType = COLAdvanced Or t_eCollectionType = COLClassUp Then
            m_ofwArrangeData.MoveNext
            m_ofwCustHistData.MoveNext
            m_ofwAppInfoData.MoveNext
            m_ofwAccrualData.MoveNext
            m_ofwNoteData.MoveNext
        End If
    Loop
    
    'Close the files
    Close #lSGMLFileNum
    Close #lBLKFileNum
        
    WriteToFile = bHaveWrittenData
        
    Exit Function
    
ErrorHandler:
    ConvertError sSource
    
End Function

Private Function GetFileName( _
            t_sDatePrefix As String, _
            t_bFromLive As Boolean, _
            t_eCollectionType As ECollectionType, _
            Optional t_sLettercode As String, _
            Optional t_eAction As EIndexingAction, _
            Optional t_eLevel As ECATALOGUELEVEL, _
            Optional t_lClassID As Long = 0, _
            Optional t_lHeaderID As Long = 0, _
            Optional t_lPieceID As Long = 0) As String
'   Create filenames for indexing files
'
'   d2  WDP     26-Oct-1999      Created
'   d3  WDP     29-Oct-1999      Added class and header optional params
'   d6  WDP     19-Nov-1999     Added PieceID param
On Error GoTo ErrorHandler

    Dim sSource As String
    sSource = "BOCatalogueIndexer.GetFilename"
    
    Dim sRes As String
    sRes = Trim(t_sDatePrefix) & "_"
    If t_bFromLive = True Then
        sRes = sRes & "LIVE_" & _
        Trim(GetCollectionTypeString(t_eCollectionType)) & _
        "_" & Trim(t_sLettercode)
        If t_lPieceID = 0 Then
            If Not t_lClassID = 0 Then
                sRes = sRes & "_" & t_lClassID
            End If
            If Not t_lHeaderID = 0 Then
                sRes = sRes & "_" & t_lHeaderID
            End If
            sRes = sRes & "_Level" & t_eLevel
        Else
            sRes = sRes & "MovePiece_" & t_lPieceID
        End If
        
        sRes = sRes & "_" & GetIndexActionString(t_eAction)
    Else
        sRes = sRes & "EDIT_" & _
        Trim(GetCollectionTypeString(t_eCollectionType)) & _
        "_" & Trim(t_sLettercode) & "_" & m_lEditsetID & _
        "_" & GetIndexActionString(t_eAction)
    End If
    
    GetFileName = sRes
    
    Exit Function
ErrorHandler:
    ConvertError sSource
    
End Function


Private Function GetMovedEntryFileName( _
            t_sDatePrefix As String, _
            t_eLevel As ECATALOGUELEVEL, _
            t_lCatID As Long, t_eCollectionType As ECollectionType) As String
'   Create filenames for indexing files for moved cat entries
'
'   d7  WDP     26-Nov-1999     Created
On Error GoTo ErrorHandler

    Dim sSource As String
    sSource = "BOCatalogueIndexer.GetMovedEntryFilename"
    
    Dim sRes As String
    sRes = Trim(t_sDatePrefix) & "_" & "Editset_" & m_lEditsetID & _
        "_" & GetCollectionTypeString(t_eCollectionType) & "_MOVE_CatID_" & t_lCatID & "_Level_" & t_eLevel
    
    GetMovedEntryFileName = sRes
    
    Exit Function
    
ErrorHandler:
    ConvertError sSource
    
End Function


Private Function GetLettercodeID(t_sLettercode As String) As Long
'   Get DB id of a lettercode
'
'   d7  WDP     30-Nov-1999     Created
On Error GoTo ErrorHandler

    Dim sSource As String
    sSource = "BOCatalogueIndexer.GetLettercodeID"
    
    Dim sSQL As String, ofwHelper As New FWDBHelper, ofwResult As FWRecordList
    sSQL = " SELECT " & gk_sDBCOLLetterCodeID & " FROM " & gk_sDBTABLELCode & _
            " WHERE " & gk_sDBCOLLetterCodeCode & "=" & Quotes(t_sLettercode)
    
    Set ofwResult = ofwHelper.FindBySql(sSQL)
    GetLettercodeID = ofwResult.GetFieldValue(gk_sDBCOLLetterCodeID)
    
    Exit Function
    
ErrorHandler:
    ConvertError sSource
    
End Function




Public Function LogFileName( _
    t_sDate As String, _
    Optional t_lEditSetID As Long = -1, _
    Optional t_bIsMove As Boolean)
'   Generate a logfile name & path given date string and editset id
'
'   d5  WDP     18-Nov-1999     Created
On Error GoTo ErrorHandler

    Dim sSource As String
    sSource = "BOCatalogueIndexer.LogFileName"
    Dim sIndexFileDrive As String
    
    Dim ofwAppData As New FWApplicationData
    sIndexFileDrive = ofwAppData.GetApplicationConfigurationInfomation(pRootForIndexingFiles)
    Set ofwAppData = Nothing
    
    LogFileName = sIndexFileDrive & mk_sFilepath & t_sDate & "_"
    If t_lEditSetID = -1 Or t_lEditSetID = 0 Then
        LogFileName = LogFileName & "LIVE.log"
    Else
        LogFileName = LogFileName & "Editset_" & t_lEditSetID
        If t_bIsMove = True Then
            LogFileName = LogFileName & "_MOVES"
        End If
        LogFileName = LogFileName & ".log"
    End If
    
    Exit Function
ErrorHandler:
    ConvertError sSource
End Function

Private Function FieldLine(t_sFieldName As String, t_sValue As String)
'   Create the bulk file field line (if there is a value)
'
'   d6  WDP     19-Nov-1999     Created
On Error GoTo ErrorHandler

    Dim sSource As String
    sSource = "BOCatalogueIndexer.FieldLine"
    
    If Not (t_sValue = "" Or t_sValue = "0" Or t_sValue = """""") Then
        FieldLine = t_sFieldName & ": " & t_sValue & vbCrLf
    Else
        FieldLine = ""
    End If
    
    Exit Function
ErrorHandler:
    ConvertError sSource
End Function

Private Function GetParentLettercodeID(t_lCatID As Long, t_eCatLevel As ECATALOGUELEVEL) As Long
'   Get the parent lettercode of the cat entry
'
'   d7  WDP     30-Nov-1999     Created
On Error GoTo ErrorHandler

    Dim sSource As String
    sSource = "BOCatalogueIndexer.GetParentLettercodeID(" & t_lCatID & ", " & t_eCatLevel & ")"
    
    If t_eCatLevel > ctLetterCode Then
    
        Dim sSQL As String, ofwHelper As New FWDBHelper, ofwResult As FWRecordList
        sSQL = " SELECT " & gk_sDBCOLLetterCodeID & " AS " & gk_sDIID
        
        Select Case t_eCatLevel
            Case ctDivision
                sSQL = sSQL & " FROM " & gk_sDBTABLEDivision & " d "
                sSQL = sSQL & " WHERE d." & gk_sDBCOLDivisionID & "=" & t_lCatID
            Case ctClass
                sSQL = sSQL & " FROM " & gk_sDBTABLEClass & " c "
                sSQL = sSQL & " WHERE c." & gk_sDBCOLClassID & "=" & t_lCatID
            Case ctHeaders
                sSQL = sSQL & " FROM " & gk_sDBTABLEClass & " c, " & _
                            gk_sDBTABLEHeader & " h "
                sSQL = sSQL & " WHERE h." & gk_sDBCOLHeaderID & "=" & t_lCatID & _
                            " AND h." & gk_sDBCOLClassID & "=c." & gk_sDBCOLClassID
            Case ctSubHeaders
                sSQL = sSQL & " FROM " & gk_sDBTABLEClass & " c, " & _
                            gk_sDBTABLEHeader & " h, " & _
                            gk_sDBTABLESubHeader & " sh "
                sSQL = sSQL & " WHERE sh." & gk_sDBCOLSubHeaderID & "=" & t_lCatID & _
                            " AND h." & gk_sDBCOLClassID & "=c." & gk_sDBCOLClassID & _
                            " AND sh." & gk_sDBCOLHeaderID & "=h." & gk_sDBCOLHeaderID
            Case ctPiece
                sSQL = sSQL & " FROM " & gk_sDBTABLEClass & " c, " & _
                            gk_sDBTABLEPiece & " p "
                sSQL = sSQL & " WHERE p." & gk_sDBCOLPieceID & "=" & t_lCatID & _
                            " AND p." & gk_sDBCOLClassID & "=c." & gk_sDBCOLClassID
            Case ctItem
                sSQL = sSQL & " FROM " & gk_sDBTABLEClass & " c, " & _
                            gk_sDBTABLEPiece & " p, " & _
                            gk_sDBTABLEItem & " i "
                sSQL = sSQL & " WHERE i." & gk_sDBCOLItemID & "=" & t_lCatID & _
                            " AND p." & gk_sDBCOLClassID & "=c." & gk_sDBCOLClassID & _
                            " AND p." & gk_sDBCOLPieceID & "=i." & gk_sDBCOLPieceID
        End Select
    
        Set ofwResult = ofwHelper.FindBySql(sSQL)
        If ofwResult.RecordCount = 0 Then
            RaiseError sSource, PEGeneralInternalError, "Could not find parent ID"
        Else
            GetParentLettercodeID = ofwResult.GetFieldValue(gk_sDIID)
        End If
    Else
        GetParentLettercodeID = t_lCatID
    End If
    Exit Function
ErrorHandler:
    ConvertError sSource
End Function


Private Function GetParentLettercode(t_lCatID As Long, t_eCatLevel As ECATALOGUELEVEL) As String
'   Get the parent lettercode of the cat entry
'
'   d7  WDP     30-Nov-1999     Created
On Error GoTo ErrorHandler

    Dim sSource As String
    sSource = "BOCatalogueIndexer.GetParentLettercode"
    
    
    Dim sSQL As String, ofwHelper As New FWDBHelper, ofwResult As FWRecordList
    
    sSQL = " SELECT l." & gk_sDBCOLLetterCodeCode & " AS " & gk_sDILetterCodeRef & _
            " FROM " & gk_sDBTABLELCode & " l "
    
    Select Case t_eCatLevel
        Case ctLetterCode
            sSQL = sSQL & " WHERE l." & gk_sDBCOLLetterCodeID & "=" & t_lCatID
        Case ctDivision
            sSQL = sSQL & ", " & gk_sDBTABLEDivision & " d "
            sSQL = sSQL & " WHERE d." & gk_sDBCOLDivisionID & "=" & t_lCatID & _
                            " AND l." & gk_sDBCOLLetterCodeID & "=d." & gk_sDBCOLLetterCodeID
        Case ctClass
            sSQL = sSQL & ", " & gk_sDBTABLEClass & " c "
            sSQL = sSQL & " WHERE c." & gk_sDBCOLClassID & "=" & t_lCatID & _
                            " AND l." & gk_sDBCOLLetterCodeID & "=c." & gk_sDBCOLLetterCodeID
        Case ctHeaders
            sSQL = sSQL & ", " & gk_sDBTABLEClass & " c " & _
                        ", " & gk_sDBTABLEHeader & " h "
            sSQL = sSQL & " WHERE h." & gk_sDBCOLHeaderID & "=" & t_lCatID & _
                        " AND h." & gk_sDBCOLClassID & "=c." & gk_sDBCOLClassID & _
                        " AND l." & gk_sDBCOLLetterCodeID & "=c." & gk_sDBCOLLetterCodeID
        Case ctSubHeaders
            sSQL = sSQL & ", " & gk_sDBTABLEClass & " c " & _
                        ", " & gk_sDBTABLEHeader & " h " & _
                        ", " & gk_sDBTABLESubHeader & " sh "
            sSQL = sSQL & " WHERE sh." & gk_sDBCOLSubHeaderID & "=" & t_lCatID & _
                        " AND h." & gk_sDBCOLClassID & "=c." & gk_sDBCOLClassID & _
                        " AND sh." & gk_sDBCOLHeaderID & "=h." & gk_sDBCOLHeaderID & _
                        " AND l." & gk_sDBCOLLetterCodeID & "=c." & gk_sDBCOLLetterCodeID
        Case ctPiece
            sSQL = sSQL & ", " & gk_sDBTABLEClass & " c " & _
                        ", " & gk_sDBTABLEPiece & " p "
            sSQL = sSQL & " WHERE p." & gk_sDBCOLPieceID & "=" & t_lCatID & _
                        " AND p." & gk_sDBCOLClassID & "=c." & gk_sDBCOLClassID & _
                        " AND l." & gk_sDBCOLLetterCodeID & "=c." & gk_sDBCOLLetterCodeID
        Case ctItem
             sSQL = sSQL & ", " & gk_sDBTABLEClass & " c " & _
                        ", " & gk_sDBTABLEPiece & " p, " & _
                        gk_sDBTABLEItem & " i "
            sSQL = sSQL & " WHERE i." & gk_sDBCOLItemID & "=" & t_lCatID & _
                        " AND p." & gk_sDBCOLClassID & "=c." & gk_sDBCOLClassID & _
                        " AND p." & gk_sDBCOLPieceID & "=i." & gk_sDBCOLPieceID & _
                        " AND l." & gk_sDBCOLLetterCodeID & "=c." & gk_sDBCOLLetterCodeID
    End Select
    
    Set ofwResult = ofwHelper.FindBySql(sSQL)
    GetParentLettercode = ofwResult.GetFieldValue(gk_sDILetterCodeRef)
    
    Exit Function
ErrorHandler:
    ConvertError sSource
End Function


Public Function IndexStepExists(t_eStep As EIndexingStep, Optional t_eCollectionType As Variant, Optional t_eLevel As ECATALOGUELEVEL = ctNone, _
                            Optional t_lLettercodeID As Long = 0, Optional t_lClassID As Long = 0) As Boolean
'   Returns true if a matching step found
'
'   d11 WDP     6-Mar-2000      Created
On Error GoTo ErrorHandler
    
    Dim sSource As String
    sSource = "BOCatDataExtractor.IndexStepExists"
    
    Dim sSQL As String
    Dim ofwHelper As New FWDBHelper
    Dim ofwResult As FWRecordList
    
    sSQL = " SELECT COUNT(*) AS " & gk_sDICount & " FROM " & _
            gk_sDBTABLEIndexingProgress & " WHERE " & _
            gk_sDBCOLIndexingProgressStepNumber & "=" & t_eStep
    
    If Not t_eLevel = ctNone Then
        sSQL = sSQL & " AND " & gk_sDBCOLIndexingProgressLevel & "=" & t_eLevel
    End If
    If Not t_lLettercodeID = 0 Then
        sSQL = sSQL & " AND " & gk_sDBCOLIndexingProgressLettercodeID & "=" & t_lLettercodeID
    End If
    If Not t_lClassID = 0 Then
        sSQL = sSQL & " AND " & gk_sDBCOLIndexingProgressClassID & "=" & t_lClassID
    End If
    If Not IsMissing(t_eCollectionType) Then
        sSQL = sSQL & " AND " & gk_sDBCOLIndexingProgressCollectionType & "=" & t_eCollectionType
    End If
    
    Set ofwResult = ofwHelper.FindBySql(sSQL)
    
    If ofwResult.GetFieldValue(gk_sDICount) > 0 Then
        IndexStepExists = True
    Else
        IndexStepExists = False
    End If
    
    Exit Function
ErrorHandler:
    ConvertError sSource
    
End Function


Public Sub LogIndexStep(t_eStep As EIndexingStep, t_sMessage As String, _
                            Optional t_eCollectionType As ECollectionType = COLBasic, _
                            Optional t_eLevel As ECATALOGUELEVEL = ctNone, _
                            Optional t_lLettercodeID As Long = 0, Optional t_lClassID As Long = 0)
'   Log a step in the indexing progress table
'
'   d11 WDP     6-Mar-2000      Created
On Error GoTo ErrorHandler
    
    Dim sSource As String
    sSource = "BOCatDataExtractor.LogIndexStep"
    
    Dim sSQL As String
    Dim ofwHelper As New FWDBHelper
    
    sSQL = " INSERT INTO " & _
            gk_sDBTABLEIndexingProgress & "(" & _
                gk_sDBCOLIndexingProgressStepNumber & ", " & _
                gk_sDBCOLIndexingProgressLettercodeID & ", " & _
                gk_sDBCOLIndexingProgressCollectionType & ", " & _
                gk_sDBCOLIndexingProgressLevel & ", " & _
                gk_sDBCOLIndexingProgressClassID & ", " & _
                gk_sDBCOLIndexingProgressMessage & ")"
            
    sSQL = sSQL & " VALUES(" & t_eStep & ", " & t_lLettercodeID & ", " & _
                        t_eCollectionType & ", " & t_eLevel & ", " & t_lClassID & ", " & QuoteValueWithQuotes(t_sMessage) & ")"
    
    ofwHelper.Execute sSQL
    
    Exit Sub
ErrorHandler:
    ConvertError sSource
    
End Sub

Private Sub LogEntryWritten(t_iLevel As Integer, t_lCatID As Long)
'   test method for logging indexed entries when doing
'   a complete indexing of the catalogue
'   !NOTE: remove calls to this on release as it reduces performance
'           of indexing and could mess up release and move.
'
'   d12 WDP     15-Mar-2000      Created
On Error GoTo ErrorHandler
    
    Dim sSource As String
    sSource = "BOCatDataExtractor.LogEntryWritten"
    
    Dim sSQL As String
    Dim ofwHelper As New FWDBHelper
    
    sSQL = " INSERT INTO tempIndexedEntries(catalogue_id, level_no) " & _
            " VALUES(" & t_lCatID & ", " & t_iLevel & ")"
            
    ofwHelper.Execute sSQL
    
    Exit Sub
ErrorHandler:
    ConvertError sSource
    
End Sub

Private Sub DumpLogToMasterLog(t_sLogFile As String)
'   copy and truncate the temporary log of indexed actions passed
'   into the master log file
'
'   d12 WDP     15-Mar-2000      Created
On Error GoTo ErrorHandler
    
    Dim sSource As String
    sSource = "BOCatDataExtractor.DumpLogToMasterLog"
    
    Dim lFilenum As Long
    Dim sLogText As String
    
    lFilenum = FreeFile
    Open t_sLogFile For Input Access Read As lFilenum
    If lFilenum > 0 Then
        'Read the actions from the temp log
        sLogText = StrConv(InputB(LOF(lFilenum), lFilenum), vbUnicode)
        Close #lFilenum
        lFilenum = FreeFile
        'open the master log
        Dim sMasterLog As String
        Dim ofwAppData As New FWApplicationData
        sMasterLog = ofwAppData.GetApplicationConfigurationInfomation(pRootForIndexingFiles) & gk_sMasterIDXLogFile
        Open sMasterLog For Append As lFilenum
        If lFilenum > 0 Then
            'add the actions to the master log
            Print #lFilenum, sLogText
        Else
            RaiseError sSource, PEGeneralInternalError, "couldnt open master logfile: " & sMasterLog
        End If
        Close #lFilenum
        'Truncate the temporary log (just open it not for append)
        lFilenum = FreeFile
        Open t_sLogFile For Output As lFilenum
        Close #lFilenum
    Else
        RaiseError sSource, PEGeneralInternalError, "couldnt open temp logfile: " & t_sLogFile
    End If
    
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




