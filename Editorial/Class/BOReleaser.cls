VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "BOReleaser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    BOReleaser.cls
' System:       PRO, PROCAT
' Copyright:    (C) Quidnunc Limited
'
' Description:  Business object to perform various stages of
'               releasing of edit sets
'
' Amendment history:
'   d1  WDP 12-Nov-1999     Created
'   d2  WDP 24-Nov-1999     Added TestTransaction
'   d3  WDP 26-Nov-1999     Added DeleteEditset and CreateMovedDataFiles
'   d4  WDP 25-Feb-2000     Fixed bug in UpdateTitleScope
'   d5  DAJ 03-Jul-2000     Added extra logging info, changed maximum return size for LogUnreleasableEditSets
'   d6  DAJ 10-Aug-2000     Changed UpdateTitleScope to update Accessioned entries too
'   d7  DAJ 08-Jan-2001     Changed UpdateTitleScope to correctly join up the piece tables
'   d8  DAJ 05-Feb-2001     Changed UpdateTitleScope to remove extra tbl_description entries
'                           if the title/scope & content has been edited to drop below 25 characters
'   d9  MAP 24-Oct-2001     Amended to delete entries from the catalogue upon release
'   d10 ALM 29-Oct-2001     Removed creation of deletable editset collection and
'                           function GetDeletedEntries
'   d11 TMD 25-Jun-2003     Changed UpdateTitleScope to correctly join up the item tables
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Implements ObjectControl
Private m_objContext As Object
Private m_aspResponse As Response
Private m_aspRequest As Request
Private m_aspServer As Server
Private m_aspSession As Session

' Class constants
Private Const mk_lMaxRecords As Long = 50000

'
' Public Methods
'

Private Function IsDeletionEditSet(ByVal t_lEditSetID As Long) As Boolean
'   Determines if edit set passed in is a deletion one
'
'   d9  MAP     24-Oct-2001     Created

On Error GoTo ErrorHandler

    Dim sSource As String
    Dim sSQL As String
    Dim ofwDBHelper As New FWDBHelper
    Dim ofwResults As FWRecordList
    
    sSource = "BOReleaser.IsDeletionEditSet"
    
    sSQL = "SELECT " & gk_sDBCOLEditSetType & _
            " FROM " & gk_sDBTABLEEditset & _
            " WHERE " & gk_sDBCOLEditSetID & "=" & t_lEditSetID
    
    Set ofwResults = ofwDBHelper.FindBySql(sSQL)
    
    ' Enums for edit set type does not match database values, so have to use hard coded value
    If ofwResults.GetFieldValue(gk_sDBCOLEditSetType) = 4 Then
        IsDeletionEditSet = True
    Else
        IsDeletionEditSet = False
    End If
    
    Exit Function
    
ErrorHandler:
    ConvertError sSource
    
End Function

Public Function ReleaseEditsetInDB(t_lEditSetID As Long, t_sDate As String, ByRef t_sTime As String, _
                                   ByRef t_colMovedEntries As Collection) As Boolean
'   Releasing an edit set. We bundle up the database processes
'   into a single transaction to ensure we can rollback
'   if something goes wrong. BOReleaser objects must be created
'   with CreateObject for these transactions to be enforced.
'   separately. It returns the time stamp used to identify the
'   indexing log (which is returned from the indexing object when
'   indexing files are created)
'
'   d1  WDP     16-Nov-1999   Created
'   d9  MAP     24-Oct-2001     Amended to deal with deletion edit sets
'   d10 ALM     29-Oct-2001   Removed creation of deletable editset collection
On Error GoTo ErrorHandler
    
    Dim sSource As String
    sSource = "VWReleaseEditSet.ReleaseEditset"
    
    Dim sSQL As String
    Dim ofwHelper As New FWDBHelper
    Dim oboDataExtractor As New BOCatDataExtractor
    
    ' d9 - Need to behave differently if the edit set is a deletion one
    If IsDeletionEditSet(t_lEditSetID) Then
                
        t_sTime = oboDataExtractor.CreateEditSetIndexData(t_lEditSetID, t_sDate, True)
        Set oboDataExtractor = Nothing
        
        sSQL = "sp_Edit_Release_Deletion_EditSet " & t_lEditSetID
        ofwHelper.Execute sSQL
        Set ofwHelper = Nothing
        
        LogReleaseMessage "Entries in deletion edit set " & t_lEditSetID & " removed from database", RELInfo
        
    Else
    
        Set t_colMovedEntries = GetMovedEntries(t_lEditSetID)
    
        sSQL = "sp_edit_release_editset " & t_lEditSetID
        ofwHelper.Execute sSQL
        Set ofwHelper = Nothing
        
        AdditionalEditsetProcessing t_lEditSetID
        
        LogReleaseMessage "Editset " & t_lEditSetID & " released in database", RELInfo
        
        t_sTime = oboDataExtractor.CreateEditSetIndexData(t_lEditSetID, t_sDate)
        Set oboDataExtractor = Nothing
        
    End If
    
    'If the processes have succeeded then commit.
    ReleaseEditsetInDB = True
    m_objContext.SetComplete
    
Exit Function
ErrorHandler:
    'If any of the above processes fails we want to
    'rollback
    ReleaseEditsetInDB = False
    LogReleaseMessage "Error releasing editset " & t_lEditSetID & ": " & Err.Description, RELError
    m_objContext.SetAbort
    'ConvertError sSource
End Function


Public Function GetMovedEntries(t_lEditSetID As Long) As Collection
'   Finds which (if any) edit set entries have been
'   moved
'
'   d1  WDP     16-Nov-1999   Created

On Error GoTo ErrorHandler
    
    Dim sSource As String
    sSource = "BOReleaser.GetMovedEntries"
    
    Dim sSQL As String
    Dim ofwDBHelper As New FWDBHelper
    Dim ofwResults As FWRecordList
    
    sSQL = "sp_edit_find_moved_entries " & t_lEditSetID
    Set ofwResults = ofwDBHelper.FindBySql(sSQL, mk_lMaxRecords)
    Dim colMovedCatalogueEntries As New Collection
    'Read all the moved entries into the results
    While Not ofwResults.IsAtEnd
        'Use the doc key format of <cat level>*<cat id>
        colMovedCatalogueEntries.Add VerityDocKey(ofwResults.GetFieldValue(gk_sDBCOLCatalogueID), _
                    ofwResults.GetFieldValue(gk_sDBCOLCatalogueLevel))
        ofwResults.MoveNext
    Wend
    Set GetMovedEntries = colMovedCatalogueEntries
    
Exit Function
ErrorHandler:
    ConvertError sSource
End Function

Public Function FindReleasableEditSets() As Collection
'   Reads list of releasable editsets into collection
'   Also logs which edit sets are in editorial release but not
'   releasable and why.
'
'   d1  WDP     16-Nov-1999   Created

On Error GoTo ErrorHandler
    
    Dim sSource As String
    sSource = "BOReleaser.FindReleasableEditSets"
                       
    Dim ofwDBHelper As New FWDBHelper
    Dim ofwRecords As FWRecordList
    Dim sSQL As String
    Dim lFileNum As Long
    Dim colEditsets As New Collection
    
    sSQL = "sp_edit_find_ready_editsets"
    
    Set ofwRecords = ofwDBHelper.FindBySql(sSQL, mk_lMaxRecords)
    
    While Not ofwRecords.IsAtEnd
        colEditsets.Add ofwRecords.GetFieldValue(gk_sDBCOLEditSetID)
        ofwRecords.MoveNext
    Wend
    
    LogUnreleasableEditSets
    
    Set FindReleasableEditSets = colEditsets
    
Exit Function
ErrorHandler:
    ConvertError sSource
End Function


Public Function DeleteEditset(t_lEditSetID) As Boolean
'   Deletes all trace of an edit set from the ES tables
'
'   d3  WDP     26-Nov-1999   Created

On Error GoTo ErrorHandler
    
    Dim sSource As String
    sSource = "BOReleaser.DeleteEditset"
                       
    Dim ofwDBHelper As New FWDBHelper
    Dim sSQL As String
    
    'SP to do the deletion
    sSQL = "sp_edit_DeleteEditset " & t_lEditSetID
    
    ofwDBHelper.Execute sSQL
    
    'If no error has been raised, we return True and commit
    DeleteEditset = True
    m_objContext.SetComplete
    
Exit Function
ErrorHandler:
    
    'return false and rollback
    DeleteEditset = False
    m_objContext.SetAbort
    
End Function


Public Function CreateMovedDataFiles(t_colMovedEntries As Collection, t_sDate As String, ByRef t_sTime As String, t_lEditSetID As Long) As Boolean
'   Creates data files for re-indexing the children of
'   catalogue entries
'
'   d3  WDP     26-Nov-1999   Created

On Error GoTo ErrorHandler
    
    Dim sSource As String
    sSource = "BOReleaser.CreateMovedDataFiles"
                       
    Dim oboDataExtractor As New BOCatDataExtractor
    
    t_sTime = GetTimeString()
    
    Dim lCount As Long
    For lCount = 1 To t_colMovedEntries.Count
        oboDataExtractor.CreateIndexDataForEntry t_colMovedEntries.Item(lCount), t_sDate, t_sTime, t_lEditSetID
    Next
        
    'If no error has been raised, we return True and commit
    CreateMovedDataFiles = True
    m_objContext.SetComplete
    
Exit Function
ErrorHandler:
    
    'return false and rollback
    CreateMovedDataFiles = False
    LogReleaseMessage "Error creating data files for moving catalogue entry: " & t_colMovedEntries.Item(lCount) & " : " & Err.Description, RELError
    m_objContext.SetAbort
    
End Function


Private Sub LogUnreleasableEditSets()
'   Reads list of releasable editsets into member collection
'   Also logs which edit sets are in editorial release but not
'   releasable and why.
'
'   d1  WDP     16-Nov-1999   Created
'   d5  DAJ 03-Jul-2000     Added extra logging info, changed maximum return size for LogUnreleasableEditSets

On Error GoTo ErrorHandler
    
    Dim sSource As String
    sSource = "VWReleaseEditSet.LogUnreleasableEditSets"
                       
    Dim sSQL As String
    Dim lFileNum As Long
    Dim sText As String
    Dim ofwDBHelper As New FWDBHelper
    Dim ofwRecords As FWRecordList
    
    'Select logging data from the log table
    sSQL = "SELECT " & gk_sDBCOLEditSetID & ", " & _
            gk_sDBCOLNotReadyEditSetReason & ", " & _
            gk_sDBCOLNotReadyEditSetDate & ", " & _
            gk_sDBCOLEditsetName & ", " & _
            gk_sDBCOLName & ", " & _
            gk_sDBCOLCatalogueID & _
            " FROM " & gk_sDBTABLENotReadyEditset
    
    ' d5 - Make sure we can get them all back....
    Set ofwRecords = ofwDBHelper.FindBySql(sSQL, mk_lMaxRecords)
            
    'If there is logging data write it to file
    If ofwRecords.RecordCount > 0 Then

        sText = sText & "At the times indicated, the following edit sets were in editorial release" & _
                " but could not be released to live for the given reasons." & vbCrLf
        While Not ofwRecords.IsAtEnd
            sText = sText & ofwRecords.GetFieldValue(gk_sDBCOLNotReadyEditSetDate) & ": " & _
                    "Editset : " & ofwRecords.GetFieldValue(gk_sDBCOLEditsetName) & _
                    " (ID = " & ofwRecords.GetFieldValue(gk_sDBCOLEditSetID) & ")" & _
                    "  Entry: '" & ofwRecords.GetFieldValue(gk_sDBCOLName) & "'" & _
                    " (ID = " & ofwRecords.GetFieldValue(gk_sDBCOLCatalogueID) & ")" & _
                    "  Reason: " & ofwRecords.GetFieldValue(gk_sDBCOLNotReadyEditSetReason) & vbCrLf
            ofwRecords.MoveNext
        Wend
        
        LogReleaseMessage sText, RELInfo
        
    End If
        
Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub


Private Sub AdditionalEditsetProcessing(t_lEditSetID As Long)
'   Performs additional editset processing such as
'   updating the ILDB title/scope fields and tbl_description
'
'   d1  WDP     16-Nov-1999   Created

On Error GoTo ErrorHandler
    
    Dim sSource As String
    sSource = "VWReleaseEditSet.AdditionalEditsetProcessing"
                       
    UpdateTitleScope t_lEditSetID, ctLetterCode
    UpdateTitleScope t_lEditSetID, ctDivision
    UpdateTitleScope t_lEditSetID, ctClass
    UpdateTitleScope t_lEditSetID, ctHeaders
    UpdateTitleScope t_lEditSetID, ctSubHeaders
    UpdateTitleScope t_lEditSetID, ctPiece
    UpdateTitleScope t_lEditSetID, ctItem
     
Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub



Private Sub UpdateTitleScope(t_lEditSetID As Long, t_eLevel As ECATALOGUELEVEL)
'   Performs additional editset processing such as
'   updating the ILDB title/scope fields and tbl_description
'
'   d1  WDP     16-Nov-1999   Created
'   d4  WDP     25-Feb-2000     Fixed bug in UpdateTitleScope
'   d5  TMJ     30-Mar-2000   Added SpecialCharacterDecode call
'   d6  DAJ 10-Aug-2000     Changed UpdateTitleScope to update Accessioned entries too
'   d7  DAJ 08-Jan-2001     Changed UpdateTitleScope to correctly join up the piece tables
'   d8  DAJ 05-Feb-2001     Changed UpdateTitleScope to remove extra tbl_description entries
'                           if the title/scope & content has been edited to drop below 25 characters
'   d11 TMD 25-Jun-2003     Changed UpdateTitleScope to correctly join up the item tables

On Error GoTo ErrorHandler
    
    Dim sSource As String
    sSource = "VWReleaseEditSet.UpdateTitleScope"
                       
    Dim sSQL As String
    Dim ofwDBHelper As New FWDBHelper
    Dim ofwRecords As FWRecordList
    Dim sTitleTable As String, sIDCol As String, sExtnTable As String
    Dim sUpdateTitleProc As String, lTitleDescType As Long
    Dim sExtraTables As String, sExtraWhere As String
    Dim sClassIDSelect As String
    Dim sTitleScopeCol As String
    
    Select Case t_eLevel
        Case ctLetterCode
            sTitleTable = gk_sDBTABLEEditsetLettercodeTitle
            sIDCol = gk_sDBCOLLetterCodeID
            sExtnTable = gk_sDBTABLEEditsetLettercodeExtension
            sUpdateTitleProc = " sp_edit_update_lettercode_title "
            lTitleDescType = 1
            sClassIDSelect = ", NULL AS " & gk_sDIClassRef
            sTitleScopeCol = gk_sDBCOLTitle
        Case ctDivision
            sTitleTable = gk_sDBTABLEEditsetDivisionTitle
            sIDCol = gk_sDBCOLDivisionID
            sExtnTable = gk_sDBTABLEEditsetDivisionExtension
            sUpdateTitleProc = " sp_edit_update_division_title "
            lTitleDescType = 1
            sClassIDSelect = ", NULL AS " & gk_sDIClassRef
            sTitleScopeCol = gk_sDBCOLTitle
        Case ctClass
            sTitleTable = gk_sDBTABLEEditsetClassTitle
            sIDCol = gk_sDBCOLClassID
            sExtnTable = gk_sDBTABLEEditsetClassExtension
            sUpdateTitleProc = " sp_edit_update_class_title "
            lTitleDescType = 1
            sClassIDSelect = ", t." & gk_sDBCOLClassID & " AS " & gk_sDIClassRef
            sTitleScopeCol = gk_sDBCOLTitle
        Case ctHeaders
            sTitleTable = gk_sDBTABLEEditsetHeaderTitle
            sIDCol = gk_sDBCOLHeaderID
            sExtnTable = gk_sDBTABLEEditsetHeaderExtension
            sUpdateTitleProc = " sp_edit_update_header_title "
            lTitleDescType = 1
            sClassIDSelect = ", h." & gk_sDBCOLClassID & " AS " & gk_sDIClassRef
            sExtraTables = ", " & gk_sDBTABLEEditsetHeader & " h "
            sExtraWhere = " AND h." & gk_sDBCOLHeaderID & "=t." & gk_sDBCOLHeaderID
            sTitleScopeCol = gk_sDBCOLTitle
        Case ctSubHeaders
            sTitleTable = gk_sDBTABLEEditsetSubheaderTitle
            sIDCol = gk_sDBCOLSubHeaderID
            sExtnTable = gk_sDBTABLEEditsetSubheaderExtension
            sUpdateTitleProc = " sp_edit_update_subheader_title "
            lTitleDescType = 1
            sExtraTables = ", " & gk_sDBTABLEEditsetHeader & " h " & _
                            ", " & gk_sDBTABLEEditsetSubheader & " sh "
            sClassIDSelect = ", h." & gk_sDBCOLClassID & " AS " & gk_sDIClassRef
            sExtraWhere = " AND sh." & gk_sDBCOLSubHeaderID & "=t." & gk_sDBCOLSubHeaderID & _
                          " AND sh." & gk_sDBCOLHeaderID & "=h." & gk_sDBCOLHeaderID
            sTitleScopeCol = gk_sDBCOLTitle
        Case ctPiece
            ' d7 - Fixed sExtraWhere by adding in extra join between p and e tables
            sTitleTable = gk_sDBTABLEEditsetPieceScopeAndContent
            sIDCol = gk_sDBCOLPieceID
            sExtnTable = gk_sDBTABLEEditsetPieceExtension
            sUpdateTitleProc = " sp_edit_update_piece_scope "
            lTitleDescType = 2
            sClassIDSelect = ", p." & gk_sDBCOLClassID & " AS " & gk_sDIClassRef
            sExtraTables = ", " & gk_sDBTABLEEditsetPiece & " p "
            sExtraWhere = " AND p." & gk_sDBCOLPieceID & "= t." & gk_sDBCOLPieceID & _
                " AND p." & gk_sDBCOLPieceID & " = e." & gk_sDBCOLPieceID
            sTitleScopeCol = gk_sDBCOLScopeAndContent
        Case ctItem
            ' d11 - Fixed sExtraWhere by adding in extra join between i and e tables
            sTitleTable = gk_sDBTABLEEditsetItemScopeAndContent
            sIDCol = gk_sDBCOLItemID
            sExtnTable = gk_sDBTABLEEditsetItemExtension
            sUpdateTitleProc = " sp_edit_update_item_scope "
            lTitleDescType = 2
            sExtraTables = ", " & gk_sDBTABLEEditsetPiece & " p " & _
                            ", " & gk_sDBTABLEEditsetItem & " i "
            sClassIDSelect = ", p." & gk_sDBCOLClassID & " AS " & gk_sDIClassRef
            sExtraWhere = " AND i." & gk_sDBCOLItemID & "=t." & gk_sDBCOLItemID & _
                          " AND i." & gk_sDBCOLPieceID & "=p." & gk_sDBCOLPieceID & _
                          " AND i." & gk_sDBCOLItemID & " = e." & gk_sDBCOLItemID
            sTitleScopeCol = gk_sDBCOLScopeAndContent
    End Select
    
    ' d6 - The type list should include accessioned
    sSQL = " SELECT t." & sIDCol & " AS " & gk_sDIID & ", t." & _
            sTitleScopeCol & " AS " & gk_sDITitleOrScope & sClassIDSelect & _
           " FROM " & sTitleTable & " t, " & sExtnTable & " e " & _
           sExtraTables & _
           " WHERE e." & gk_sDBCOLEditSetID & "=" & t_lEditSetID & _
           " AND e." & gk_sDBCOLEditSetID & "=t." & gk_sDBCOLEditSetID & _
           " AND e." & gk_sDBCOLEditSetEntryTypeID & " IN (" & _
           cesPROCATCreatedModified & ", " & _
           cesImportedModified & ", " & _
           cesAddedFromLiveModified & ", " & _
           cesImported & ")" & sExtraWhere
           
    Set ofwRecords = ofwDBHelper.FindBySql(sSQL, mk_lMaxRecords)
    If ofwRecords.RecordCount > 0 Then
        ofwRecords.MoveFirst
        While Not ofwRecords.IsAtEnd
            Dim sStrippedTitle As String, sSortedTitle As String, lID As Long
            Dim sExtraTitle As String
            sStrippedTitle = VariantStringAsString(ofwRecords.GetFieldValue(gk_sDITitleOrScope))
            lID = ofwRecords.GetFieldValue(gk_sDIID)
            If t_eLevel = ctLetterCode Then
                'Update lettercode_sorted_title
                sSortedTitle = GetSortedTitle(sStrippedTitle)
                sSQL = " sp_edit_update_sorted_title " & lID & ", " & QuoteValueWithQuotes(sSortedTitle)
                ofwDBHelper.Execute sSQL
                
            End If
            
            ' Put any overflow in the tbl_description table for headers down
            ' d5 TMJ Removed special characters
            ' d8 - Removed all the values from tbl_description anyway
            sStrippedTitle = RemoveTagsFromString(sStrippedTitle)
            sStrippedTitle = SpecialCharacterDecode(sStrippedTitle)
            
            If t_eLevel > ctClass Then
                ' We need to remove all entries for this entry, as the title may have been edited
                ' to be under 255 characters, and no longer requires an overflow
                Dim lClassID As Long
                lClassID = ofwRecords.GetFieldValue(gk_sDIClassRef)
                sSQL = "delete " & gk_sDBTABLEDescription & _
                    " where " & gk_sDBCOLListReferenceCatalogueID & " = " & lID & _
                    " and " & gk_sDBCOLCatalogueLevel & " = " & t_eLevel & _
                    " and " & gk_sDBCOLPieceClassID & " = " & lClassID
                ofwDBHelper.Execute sSQL
                
                ' If the entry is still over 255 characters after editing, then we put it back
                If Len(sStrippedTitle) > 255 Then
                    sExtraTitle = Mid(sStrippedTitle, 256)
                    sSQL = "sp_edit_update_tbl_description " & _
                        "@t_iCatID = " & lID & _
                        ", @t_sText = " & QuoteValueWithQuotes(sExtraTitle) & _
                        ", @t_iCatLevel = " & t_eLevel & _
                        ", @t_iDescType = " & lTitleDescType & _
                        ", @t_iClassID = " & lClassID
                    ofwDBHelper.Execute sSQL
                End If
            End If
            sStrippedTitle = Left(sStrippedTitle, 255)
            'update ILDB title/scope
            sSQL = sUpdateTitleProc & " " & lID & ", " & QuoteValueWithQuotes(sStrippedTitle)
            ofwDBHelper.Execute sSQL
                
            ofwRecords.MoveNext
        Wend
    End If
        
Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub


'
' Properties
'


''''''''''''''''' Implementation ''''''''''''''''''''''

'
' Private functions
'


'
' Class initialisation
'



Private Sub ObjectControl_Activate()
    Initialise
End Sub

Private Sub ObjectControl_Deactivate()
    Set m_objContext = Nothing
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = True
End Function

Private Sub Initialise()
    Set m_objContext = GetObjectContext()
    
    
    ' Specific initialisation
    ' <Specific initialisation here>
    
End Sub

Public Sub TestTransaction(t_bRaiseError As Boolean)
'   Test commit and rollback of transactions
'   If the parameter passed is true an error will
'   be raised and the transaction should be rolled back
'
'   d2  WDP     24-Nov-1999   Created

On Error GoTo ErrorHandler
    
    Dim sSource As String
    sSource = "VWReleaseEditSet.TestTransaction"
    
    Dim sSQL As String
    Dim ofwHelper As New FWDBHelper

    sSQL = "sp_test_transaction "
    ofwHelper.Execute sSQL
    
    If t_bRaiseError = True Then
        RaiseError sSource, PEGeneralInternalError, "An error!"
    End If
    
    m_objContext.SetComplete
Exit Sub
ErrorHandler:
    'If any of the above processes fails we want to
    'rollback
    m_objContext.SetAbort
    
End Sub

Public Sub TestMethod(t_bAbort As Boolean)
'Test method for scheduled task
On Error GoTo ErrorHandler

    Dim sSource As String
    sSource = "BOReleaser.TestMethod"
    
    Dim obj As New FWDBHelper
    obj.Execute ("INSERT INTO zTempTable(id) VALUES (1)")
    
    If t_bAbort Then
        m_objContext.SetAbort
    Else
        m_objContext.SetComplete
    End If
    
Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub
