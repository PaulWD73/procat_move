VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BODeleteCatalogue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    BODeleteCatalogue.cls
' System:       PROCAT, PRO
' Copyright:    (C) Quidnunc Limited
'
' Description:  Business object to handle the deletion of editorial catalogue entries
'
' Amendment history:
'   d1  ndk     22Oct1999       Created
'   d2  ndk     15Nov1999       Completed lots of incomplete functionality
'   d3  ndk     22Nov1999       Raising error if entry to be deleted is not found in live tables
'   d4  ndk     10Dec1999       Deleting children of created entry when it is deleted (as it could now have children)
'   d5  wdp     10Dec1999       Added reset record status calls on deletion
'   d6  ndk     29Dec1999       Rechecking releasing of variables, especially if an error is deliberately raised
'   d7  David Manning  02-feb-2000  Fixed defect 283 (sp_Edit_AddADivisioncodeToEditSet should be called sp_Edit_AddADivisionToEditSet)
'   d8  D Johnston  03Mar2000   Changed GetRecordListFromTable to check for division columns to be zero
'                               instead of null.  Also fixed recursion error in DeleteCatalogueChildren...
'   d9  D Johnston  08Mar2000   Fixed the delete method on classes to cope with a division ID of zero
'                               instead of null
'   d10 D Johnston  15Mar2000   Fixed the delete method to delete PC_Link entries pointing TO current entry as well as away from it
'   d11 tmj     13Apr2000       Added deletion of leaflet links
'   d12 M Phanides  23May2001   Added function UpdateParentIdInLiveTableForMovedEntries and
'                               amended DeleteCatalogueChildrenAndReadOnlyParentsFromLiveAndEditorial
'   d13 M Phanides  23May2001   Amended bug with recursive loop in DeleteCatalogueChildrenAndReadOnlyParentsFromLiveAndEditorial
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
'Property property variables
Private m_lEditsetID As Long
Private m_lCatalogueID As Long
Private m_eCatalogueLevelID As ECATALOGUELEVEL

'private framework and business objects
Private m_oFWDBHelper As FWDBHelper

Private Sub Initialise()
    Set m_oFWDBHelper = New FWDBHelper
End Sub
Private Sub Terminate()
    Set m_oFWDBHelper = Nothing
End Sub
Private Sub Class_Initialize()
    Initialise
End Sub
Private Sub Class_Terminate()
    Terminate
End Sub
Public Property Get Name()
    Name = "BODeleteCatalogue"
End Property
Public Property Let EditsetID(ByVal t_lEditSetID As Long)
    m_lEditsetID = t_lEditSetID
End Property
Public Property Let CatalogueID(ByVal t_lCatalogueID As Long)
    m_lCatalogueID = t_lCatalogueID
End Property
Public Property Let CatalogueLevelID(ByVal t_eCatalogueLevelID As ECATALOGUELEVEL)
    m_eCatalogueLevelID = t_eCatalogueLevelID
End Property

Public Sub DeleteEditSetCatalogue()

' Description:  function to delete the catalogue entry from editorial tables
'               This function also takes care of deleting its non dependant readonly parents
'               and its children in certain cases.
'
' Amendment history:
'   d1  ndk     22Oct1999       Created
'   d4  ndk     10Dec1999       Deleting children of created entry when it is deleted (as it could now have children)
'   d5  wdp     10Dec1999       Added reset record status calls on deletion
'
    On Error GoTo ErrorHandler
        
    Dim sSource As String, sSql As String
    Dim iEditSetEntryType As Integer
    
    sSource = Me.Name & ".DeleteEditSetCatalogue"
    
    'Get the edit_set_entrytype value to determine how this entry was created
    sSql = _
        " SELECT " & gk_sDBCOLEntryTypeID & _
        " FROM " & GetTableName(cttExtensionTable, ESectionType.SECeditorial, m_eCatalogueLevelID) & _
        " WHERE " & GetColumnName(cctIDColumn, m_eCatalogueLevelID) & "=" & m_lCatalogueID & " AND " & _
                    gk_sDBCOLEditSetID & "=" & m_lEditsetID
        
    iEditSetEntryType = m_oFWDBHelper.FindBySql(sSql).GetFieldValue(gk_sDBCOLEntryTypeID)
    
    If iEditSetEntryType = ECATALOGUEENTRYSTATUS.cesPROCATCreated Or _
            iEditSetEntryType = ECATALOGUEENTRYSTATUS.cesPROCATCreatedModified Then
        
        '1. Deletes the entry and its childless readonly parents from editorial
        '2. Deletes the entry and all its children from live and editorial
        
            DeleteCatalogueChildrenAndReadOnlyParentsFromLiveAndEditorial _
                m_lCatalogueID, m_eCatalogueLevelID, m_lEditsetID, True

    ElseIf iEditSetEntryType = ECATALOGUEENTRYSTATUS.cesAddedFromLive Or _
            iEditSetEntryType = ECATALOGUEENTRYSTATUS.cesAddedFromLiveModified Then
        
        'This entry was added from live tables
        
        '1. If the entry has children (editable/noneditable) in editorial set the entry as readonly
            '(this done by deleting the entry from editorial and recopying it from live)
            
            If CatalogueHasChildrenInEditSet( _
                m_lCatalogueID, m_eCatalogueLevelID, m_lEditsetID) Then
                    
                'delete
                DeleteCatalogue ESectionType.SECeditorial, m_lCatalogueID, m_eCatalogueLevelID, m_lEditsetID
                
                'reset the record status to its state before entering the ES
                If m_eCatalogueLevelID >= ctPiece Then
                    ResetRecordStatus
                End If
                
                'recreate it and change the status in live_tables to live
                AddCatlogueToEditSetAsReadOnly m_lCatalogueID, m_eCatalogueLevelID, m_lEditsetID
                
                
            Else        'if it has no children (editable/noneditable)
                '2. Delete the entry from editorial
                'and all its readonly parents if they are not referenced
                'by any other entry
                
                DeleteFromEditorialTheEntryAndReadOnlyParentsWithNoOtherChildren _
                    m_lCatalogueID, m_eCatalogueLevelID, m_lEditsetID, True
                
                'reset the record status to its state before entering the ES
                If m_eCatalogueLevelID >= ctPiece Then
                    ResetRecordStatus
                End If
                
                'Set the live_flag in the live tables as live
                UpdateLiveStatusFlag m_lCatalogueID, m_eCatalogueLevelID, _
                    ELIVEFLAGSTATUS.lfsLive
                
                
            End If
    ElseIf iEditSetEntryType = ECATALOGUEENTRYSTATUS.cesImported Or _
            iEditSetEntryType = ECATALOGUEENTRYSTATUS.cesImportedModified Then
            
            
            '1. Deletes the entry and its childless readonly parents from editorial
            '2. Deletes the entry and all its children from live and editorial
            
                DeleteCatalogueChildrenAndReadOnlyParentsFromLiveAndEditorial _
                    m_lCatalogueID, m_eCatalogueLevelID, m_lEditsetID, True
            
                    
    ElseIf iEditSetEntryType = ECATALOGUEENTRYSTATUS.cesReadeOnly Then
        'If status was not read only then the move method does not
        'call this method at all
        'This is used  when we want to delete catalogue becuase its only child catalogue was made a child
        'of another catalogue (this can happen only for division,header and subheader)
        'The catalogue and its parents (if required) are deleted only from editorial
        
        If Not CatalogueHasChildrenInEditSet( _
            m_lCatalogueID, m_eCatalogueLevelID, m_lEditsetID) Then
                'There are no children in the editset, delete this
                'and its readonly parents if they do not have other children
                
                DeleteFromEditorialTheEntryAndReadOnlyParentsWithNoOtherChildren _
                    m_lCatalogueID, m_eCatalogueLevelID, m_lEditsetID, True
                
        End If
        
    End If
    
    Exit Sub
    
ErrorHandler:
    ConvertError sSource
End Sub
Private Sub DeleteFromEditorialTheEntryAndReadOnlyParentsWithNoOtherChildren( _
    ByVal t_lCatalogueID As Long, _
    ByVal t_eCatalogueLevel As ECATALOGUELEVEL, _
    ByVal t_lEditSetID As Long, _
    ByVal t_bCalledToDeleteMainEntry As Boolean)

' Description:
' Deletes from the editorial the main entry (for which this class was instantiated) and all its
' read only parents  if they do not have any other editable entries under them
'
' Amendment history:
'   d1  ndk     22Oct1999       Created
'   d9  D Johnston  08Mar2000   Fixed the delete method on classes to cope with a division ID of zero
'                               instead of null
    
    On Error GoTo ErrorHandler

    Dim sSource As String, sSql As String
    Dim oFWRecordlist As FWRecordList, iEditSetEntryType As Integer
    Dim lCatalogueParentId As Long, eCatalogueParentLevel As ECATALOGUELEVEL
    Dim sMainTableName As String, sExtensionTableName As String, sIDColumn As String
    Dim bEntryDeleted As Boolean
    
    sSource = Me.Name & ".DeleteFromEditorialTheEntryAndReadOnlyParentsWithNoOtherChildren"
    
    'the current entry is not yet deleted
    bEntryDeleted = False
    
    'To find the parents of the deleted catalogue
    sIDColumn = GetColumnName(cctIDColumn, t_eCatalogueLevel)
    sMainTableName = GetTableName(cttMainTable, ESectionType.SECeditorial, t_eCatalogueLevel)
    sExtensionTableName = GetTableName(cttExtensionTable, ESectionType.SECeditorial, t_eCatalogueLevel)
    
    'get the details of the entry to be deleted
    sSql = _
        " SELECT " & sMainTableName & ".* " & "," & sExtensionTableName & "." & gk_sDBCOLEntryTypeID & _
        " FROM " & _
            sMainTableName & ", " & sExtensionTableName & _
        " WHERE " & _
            sMainTableName & "." & sIDColumn & "=" & t_lCatalogueID & " AND " & _
            sExtensionTableName & "." & sIDColumn & "=" & t_lCatalogueID & " AND " & _
            sMainTableName & "." & gk_sDBCOLEditSetID & "=" & t_lEditSetID & " AND " & _
            sExtensionTableName & "." & gk_sDBCOLEditSetID & "=" & t_lEditSetID
                    
    Set oFWRecordlist = m_oFWDBHelper.FindBySql(sSql)
    
    If oFWRecordlist.RecordCount = 0 Then
        RaiseError sSource, PEGeneralInternalError, "This query should have returned atleast on row"
    End If
    Select Case t_eCatalogueLevel
        Case ECATALOGUELEVEL.ctLetterCode
            'Lettercode would not have any parents, just delete the entry
            'The recursion would stop if a lettercode is encountered
        
        Case ECATALOGUELEVEL.ctDivision
            'parent would be a lettercode
            lCatalogueParentId = oFWRecordlist.GetFieldValue(gk_sDBCOLLetterCodeID)
            eCatalogueParentLevel = ECATALOGUELEVEL.ctLetterCode
            
        Case ECATALOGUELEVEL.ctClass
            'parent could be a division or lettercode
            ' d9 - Division ID is zero to indicate no division
            If oFWRecordlist.GetFieldValue(gk_sDBCOLDivisionID) <> 0 Then
                'parent is a divison
                lCatalogueParentId = oFWRecordlist.GetFieldValue(gk_sDBCOLDivisionID)
                eCatalogueParentLevel = ECATALOGUELEVEL.ctDivision
            Else 'parent is a lettercode
                lCatalogueParentId = oFWRecordlist.GetFieldValue(gk_sDBCOLLetterCodeID)
                eCatalogueParentLevel = ECATALOGUELEVEL.ctLetterCode
            End If
        
        Case ECATALOGUELEVEL.ctHeaders
            'parent would be a class
            lCatalogueParentId = oFWRecordlist.GetFieldValue(gk_sDBCOLClassID)
            eCatalogueParentLevel = ECATALOGUELEVEL.ctClass
        
        Case ECATALOGUELEVEL.ctSubHeaders
            'parent would be a header
            lCatalogueParentId = oFWRecordlist.GetFieldValue(gk_sDBCOLHeaderID)
            eCatalogueParentLevel = ECATALOGUELEVEL.ctHeaders
        
        Case ECATALOGUELEVEL.ctPiece
            'parent could be a subheader, header, class
            If Not IsNull(oFWRecordlist.GetFieldValue(gk_sDBCOLSubHeaderID)) Then
                'parent is a sub header
                lCatalogueParentId = oFWRecordlist.GetFieldValue(gk_sDBCOLSubHeaderID)
                eCatalogueParentLevel = ECATALOGUELEVEL.ctSubHeaders
            
            ElseIf Not IsNull(oFWRecordlist.GetFieldValue(gk_sDBCOLHeaderID)) Then
                'parent is a header
                lCatalogueParentId = oFWRecordlist.GetFieldValue(gk_sDBCOLHeaderID)
                eCatalogueParentLevel = ECATALOGUELEVEL.ctHeaders
            
            Else
                'parent is a class
                lCatalogueParentId = oFWRecordlist.GetFieldValue(gk_sDBCOLClassID)
                eCatalogueParentLevel = ECATALOGUELEVEL.ctClass
            End If
            
        Case ECATALOGUELEVEL.ctItem
            'parent would be a piece
            lCatalogueParentId = oFWRecordlist.GetFieldValue(gk_sDBCOLPieceID)
            eCatalogueParentLevel = ECATALOGUELEVEL.ctPiece
    End Select
    
    If t_bCalledToDeleteMainEntry Then
        'This is true when the function is called the first time
        'We have to delete the main entry, no need to check
        'whether it is read only or has children
        'Just delete it
        
        DeleteCatalogue _
            ESectionType.SECeditorial, t_lCatalogueID, t_eCatalogueLevel, t_lEditSetID
            
        'the current was deleted
        bEntryDeleted = True
    Else
        'Delete the catalogue if it is read only and does not have any children
        If oFWRecordlist.GetFieldValue(gk_sDBCOLEntryTypeID) = ECATALOGUEENTRYSTATUS.cesReadeOnly Then
            If Not CatalogueHasChildrenInEditSet( _
                t_lCatalogueID, t_eCatalogueLevel, t_lEditSetID) Then
                
                'If there are no editable children in the editset and the catalogue to be
                'deleted is read only then delete it
        
                DeleteCatalogue _
                    ESectionType.SECeditorial, t_lCatalogueID, t_eCatalogueLevel, t_lEditSetID
                
                'the current was deleted
                bEntryDeleted = True
            End If
        End If
    End If
    'Only if the catalogue is deleted then go ahead and check for its parents and delete them
    If bEntryDeleted Then
        If Not t_eCatalogueLevel = ECATALOGUELEVEL.ctLetterCode Then
            'Call the current procedure again to delete the parent of the current catalogue
            DeleteFromEditorialTheEntryAndReadOnlyParentsWithNoOtherChildren _
                lCatalogueParentId, eCatalogueParentLevel, t_lEditSetID, False
        End If
    End If
    
    Set oFWRecordlist = Nothing
    
    Exit Sub
    
ErrorHandler:
    Set oFWRecordlist = Nothing
    
    ConvertError sSource
End Sub
Private Sub DeleteCatalogueChildrenAndReadOnlyParentsFromLiveAndEditorial( _
    ByVal t_lCatalogueID As Long, _
    ByVal t_eCatalogueLevel As Long, _
    ByVal t_lEditSetID As Long, _
    ByVal t_bCalledToDeleteMainEntry As Boolean)
    
'
' Description:
' Deletes the catalogue, children and parentss of the entry passed as parameter
' from editorial and live
' Used to call itself recursively when a catalogue created by importing is deleted
' t_bCalledToDeleteMainEntry is set to true the first time so as to delete only the
' parents of that
'
' Amendment history:
'   d1  ndk     22Oct1999       Created
'   d8  D Johnston  03Mar2000   Changed GetRecordListFromTable to check for division columns to be zero
'                               instead of null.  Also fixed recursion error in DeleteCatalogueChildren...
'   d12  M Phanides  22May2001  Fixed defect where you could not delete an entry if it previously had a child but
'                               has now moved to a different parent. This was because the child no longer has the
'                               parent id in the editorial tables, but when deleting the parent the constraint
'                               would be hit in the childs live table.
'   d13 M Phanides  23May2001   Fixed bug in recursive loop as in d8
    
    On Error GoTo ErrorHandler

    Dim sSource As String, sSql As String
    Dim oFWRecordlist1 As FWRecordList, sIDColumn1 As String, eCatalogueLevelID1 As ECATALOGUELEVEL
    Dim oFWRecordList2 As FWRecordList, sIDColumn2 As String, eCatalogueLevelID2 As ECATALOGUELEVEL
    
    Set oFWRecordlist1 = Nothing
    Set oFWRecordList2 = Nothing
    
    ' d12 MAP Call function to update the live table for a child entry that has moved - otherwise would hit
    '        foreign key constraints when deleting parent.
    '        Only need to call function for parents that can have their children move
    If t_eCatalogueLevel = ECATALOGUELEVEL.ctDivision Or _
       t_eCatalogueLevel = ECATALOGUELEVEL.ctHeaders Or _
       t_eCatalogueLevel = ECATALOGUELEVEL.ctSubHeaders Then
        UpdateParentIdInLiveTableForMovedEntries t_eCatalogueLevel, t_lCatalogueID, t_lEditSetID
    End If
        
    Select Case t_eCatalogueLevel
    
        Case ECATALOGUELEVEL.ctLetterCode
        'lettercode could have children in division and class tables
    
            'list of divisions under the letterocode
            Set oFWRecordlist1 = GetRecordListFromTable _
                (GetTableName(cttMainTable, ESectionType.SECeditorial, ECATALOGUELEVEL.ctDivision), _
                t_lEditSetID, gk_sDBCOLLetterCodeID, t_lCatalogueID)
            sIDColumn1 = gk_sDBCOLDivisionID
            eCatalogueLevelID1 = ECATALOGUELEVEL.ctDivision
            
            'list of classes directly under the letterocode (which have division id as null)
            Set oFWRecordList2 = GetRecordListFromTable _
                (GetTableName(cttMainTable, ESectionType.SECeditorial, ECATALOGUELEVEL.ctClass), _
                t_lEditSetID, gk_sDBCOLLetterCodeID, t_lCatalogueID, gk_sDBCOLDivisionID)
            sIDColumn2 = gk_sDBCOLClassID
            eCatalogueLevelID2 = ECATALOGUELEVEL.ctClass
            
        Case ECATALOGUELEVEL.ctDivision
        'division could have children in class table

            'list of classes under the division
            Set oFWRecordlist1 = GetRecordListFromTable _
                (GetTableName(cttMainTable, ESectionType.SECeditorial, ECATALOGUELEVEL.ctClass), _
                t_lEditSetID, gk_sDBCOLDivisionID, t_lCatalogueID)
            sIDColumn1 = gk_sDBCOLClassID
            eCatalogueLevelID1 = ECATALOGUELEVEL.ctClass

        Case ECATALOGUELEVEL.ctClass
        'class could have children in header and piece tables

            'list of headers under class
            Set oFWRecordlist1 = GetRecordListFromTable _
                (GetTableName(cttMainTable, ESectionType.SECeditorial, ECATALOGUELEVEL.ctHeaders), _
                t_lEditSetID, gk_sDBCOLClassID, t_lCatalogueID)
            sIDColumn1 = gk_sDBCOLHeaderID
            eCatalogueLevelID1 = ECATALOGUELEVEL.ctHeaders
            
            'list of pieces directly under the class (which have header and subheader as null)
            Set oFWRecordList2 = GetRecordListFromTable _
                (GetTableName(cttMainTable, ESectionType.SECeditorial, ECATALOGUELEVEL.ctPiece), _
                t_lEditSetID, gk_sDBCOLClassID, t_lCatalogueID, gk_sDBCOLHeaderID, gk_sDBCOLSubHeaderID)
            sIDColumn2 = gk_sDBCOLPieceID
            eCatalogueLevelID2 = ECATALOGUELEVEL.ctPiece

        Case ECATALOGUELEVEL.ctHeaders
        'header could have children in sub header and piece

            'list of sub headers under header
            Set oFWRecordlist1 = GetRecordListFromTable _
                (GetTableName(cttMainTable, ESectionType.SECeditorial, ECATALOGUELEVEL.ctSubHeaders), _
                t_lEditSetID, gk_sDBCOLHeaderID, t_lCatalogueID)
            sIDColumn1 = gk_sDBCOLSubHeaderID
            eCatalogueLevelID1 = ECATALOGUELEVEL.ctSubHeaders

            'list of pieces directly under header (which have subheader id as null)
            Set oFWRecordList2 = GetRecordListFromTable _
                (GetTableName(cttMainTable, ESectionType.SECeditorial, ECATALOGUELEVEL.ctPiece), _
                t_lEditSetID, gk_sDBCOLHeaderID, t_lCatalogueID, gk_sDBCOLSubHeaderID)
            sIDColumn2 = gk_sDBCOLPieceID
            eCatalogueLevelID2 = ECATALOGUELEVEL.ctPiece

        Case ECATALOGUELEVEL.ctSubHeaders
        'subheader could have children in piece

            'list of pieces under sub header
            Set oFWRecordlist1 = GetRecordListFromTable _
                (GetTableName(cttMainTable, ESectionType.SECeditorial, ECATALOGUELEVEL.ctPiece), _
                t_lEditSetID, gk_sDBCOLSubHeaderID, t_lCatalogueID)
            sIDColumn1 = gk_sDBCOLPieceID
            eCatalogueLevelID1 = ECATALOGUELEVEL.ctPiece
        
        Case ECATALOGUELEVEL.ctPiece
        'piece could have children in item

            'list of items under piece
            Set oFWRecordlist1 = GetRecordListFromTable _
                (GetTableName(cttMainTable, ESectionType.SECeditorial, ECATALOGUELEVEL.ctItem), _
                t_lEditSetID, gk_sDBCOLPieceID, t_lCatalogueID)
            sIDColumn1 = gk_sDBCOLItemID
            eCatalogueLevelID1 = ECATALOGUELEVEL.ctItem
            
        Case ECATALOGUELEVEL.ctItem
            'there would be no children for item
    End Select
    
    'Process the first recordlist
    If Not oFWRecordlist1 Is Nothing Then
        If Not oFWRecordlist1.RecordCount = 0 Then  'If there are any children
            ' d8 - Moved the MoveFirst *outside* the loop!
            oFWRecordlist1.MoveFirst
            Do While Not oFWRecordlist1.IsAtEnd
                
                'call the current routine recursively
                DeleteCatalogueChildrenAndReadOnlyParentsFromLiveAndEditorial _
                    oFWRecordlist1.GetFieldValue(sIDColumn1), _
                    eCatalogueLevelID1, _
                    t_lEditSetID, False
                
                oFWRecordlist1.MoveNext
            Loop
        End If
    End If
                
    'Process the first recordlist
    If Not oFWRecordList2 Is Nothing Then
        If Not oFWRecordList2.RecordCount = 0 Then  'If there are any children
            ' d13 MAP Moved the MoveFirst *outside* the loop as in d8
            oFWRecordList2.MoveFirst
            Do While Not oFWRecordList2.IsAtEnd
                        
                'call the current routine recursively
                DeleteCatalogueChildrenAndReadOnlyParentsFromLiveAndEditorial _
                    oFWRecordList2.GetFieldValue(sIDColumn2), _
                    eCatalogueLevelID2, _
                    t_lEditSetID, False

                oFWRecordList2.MoveNext
            Loop
        End If
    End If
                    
    'delete from editorial
    If t_bCalledToDeleteMainEntry Then
        'the  main entry is to be deletedd
        'also find its readonly parents and delete them if they have no
        'other children
        DeleteFromEditorialTheEntryAndReadOnlyParentsWithNoOtherChildren _
            t_lCatalogueID, t_eCatalogueLevel, t_lEditSetID, True
    Else
        'since the parents would have be already in the stack of
        'to-be-deleted entries just delete this
        
        DeleteCatalogue ESectionType.SECeditorial, t_lCatalogueID, _
            t_eCatalogueLevel, t_lEditSetID
    
    End If
        
    'delete from live
    DeleteCatalogue ESectionType.SECReader, t_lCatalogueID, t_eCatalogueLevel
    
    Set oFWRecordlist1 = Nothing
    Set oFWRecordList2 = Nothing
    
    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub

Private Sub UpdateParentIdInLiveTableForMovedEntries( _
    ByVal t_eCatalogueLevel As ECATALOGUELEVEL, _
    ByVal t_lCatalogueID As Long, _
    ByVal t_lEditSetID As Long)

' Description:  Note: BE CAREFUL using or updating this function as it updates live tables without using the
'               normal method of releasing an edit set.
'               Updates the value of the parent id in the live tables for any children the entry used to have.
'               This is usually done by releasing the edit set, but need a function to do it
'               seperately to use in cases of trying to delete accessioned/created entries that used to have a
'               child in the edit set at some point, but the child has now moved. Need to update the child's
'               parent id in the live table to avoid foreign key constraints on deleting the (parent) entry.
'               (See 5008QCS for reasons for this function.)
'
' Parameters:
'               t_eCatalogueLevl - Level of catalogue to check for children under
'               t_lCatalogueID - Id of catalogue to check for children under
'               t_lEditSetId - Id of edit set the entry belongs to
'
' Amendment history:
'   d12  M Phanides  22May2001   Created

    On Error GoTo ErrorHandler
    
    Dim sSource As String, sSql As String
    Dim sChildLiveTable As String, sChildEditorialTable As String
    Dim sParentIdColumn As String, sChildIdColumn As String
    Dim lIncrement As Integer
        
    sSource = Me.Name & ".DeleteCatalogue"
    
    ' Only have to worry about children moving at certain parent levels...
    ' If current level is header then checking for moved child at piece level - eg, 2 levels down catalogue
    ' Else division and subheader will be checking at child level just 1 level down
    Select Case t_eCatalogueLevel
        
        Case ECATALOGUELEVEL.ctHeaders
            
            lIncrement = 2
            
        Case ECATALOGUELEVEL.ctDivision, ECATALOGUELEVEL.ctSubHeaders
            
            lIncrement = 1
            
        Case ECATALOGUELEVEL.ctNone, ECATALOGUELEVEL.ctLetterCode, ECATALOGUELEVEL.ctClass, _
            ECATALOGUELEVEL.ctPiece, ECATALOGUELEVEL.ctItem
            
            RaiseError sSource, PEGeneralInternalError, "Function should not be used with current catalogue level"
            
    End Select
    
    ' Get table and column names relevant to catalogue level passed to function
    sChildLiveTable = GetTableName(cttMainTable, SECReader, t_eCatalogueLevel + lIncrement)
    sChildEditorialTable = GetTableName(cttMainTable, SECeditorial, t_eCatalogueLevel + lIncrement)
    sParentIdColumn = GetColumnName(cctIDColumn, t_eCatalogueLevel)
    sChildIdColumn = GetColumnName(cctIDColumn, t_eCatalogueLevel + lIncrement)


    ' Update parent ids in LIVE tables for entries that *used* to be a child of the entry passed into the function
    sSql = "UPDATE t" & _
           " SET t." & sParentIdColumn & " = e." & sParentIdColumn & _
           " FROM " & sChildLiveTable & " t, " & sChildEditorialTable & " e" & _
           " WHERE t." & sChildIdColumn & " = e." & sChildIdColumn & _
           " AND t." & sParentIdColumn & " = " & t_lCatalogueID & _
           " AND ISNULL(e." & sParentIdColumn & ", 0) <> " & t_lCatalogueID & _
           " AND e." & gk_sDBCOLEditSetID & " = " & t_lEditSetID
           
    m_oFWDBHelper.Execute sSql
    
    
    Exit Sub
    
ErrorHandler:
    ConvertError sSource
    
End Sub

Private Sub DeleteCatalogue( _
    ByVal t_eDeleteFromSection As ESectionType, _
    ByVal t_lCatalogueID As Long, _
    ByVal t_eCatalogueLevel As ECATALOGUELEVEL, _
    Optional ByVal t_lEditSetID As Long)
'
' Description:  Deletes a catalogue entry from editorial  or live depending on parameter t_eDeleteFromSection
'               There is an added pre-caution not to delete live entries from reader, even though this function is never
'               called to delete live entries from reader
'
' Amendment history:
'   d1  ndk     22Oct1999       Created
'   d3  ndk     22Nov1999       Raising error if entry to be deleted is not found in live tables
'   d10 D Johnston  15Mar2000   Fixed the delete method to delete PC_Link entries pointing TO current entry as well as away from it
'   d11 tmj     13Apr2000       Added deletion of leaflet links

    On Error GoTo ErrorHandler

    Dim sSource As String, sSql As String
    Dim sConditionForMainAndExtnTables As String, sConditionForLinkTables As String
    Dim sConditionForCatalogueLinkTable As String
    Dim oFWRecordlist As FWRecordList
    
    sSource = Me.Name & ".DeleteCatalogue"
    
    'condtion for main/extension tables
    sConditionForMainAndExtnTables = _
            GetColumnName(ECATALOGUECOLUMNTYPE.cctIDColumn, t_eCatalogueLevel) & "=" & t_lCatalogueID
            
    'for af term and sfa tables
    sConditionForLinkTables = _
        gk_sDBCOLCatalogueID & "=" & t_lCatalogueID & " AND " & _
        gk_sDBCOLCatalogueLevel & "=" & t_eCatalogueLevel
    
    'for pc_link and es_pc_link
    ' d10 - Use (catID1 OR catID2 is current entry)
    sConditionForCatalogueLinkTable = "(" & _
        "(" & gk_sDBCOLCatalogueID1 & "=" & t_lCatalogueID & " AND " & _
        gk_sDBCOLCatalogueLevel1 & "=" & t_eCatalogueLevel & ")" & _
        " OR " & _
        "(" & gk_sDBCOLCatalogueID2 & "=" & t_lCatalogueID & " AND " & _
        gk_sDBCOLCatalogueLevel2 & "=" & t_eCatalogueLevel & ")" & _
        ")"
    
    If t_eDeleteFromSection = ESectionType.SECeditorial Then
        'Deleting from editorial - have editset id in condition
        
        sConditionForMainAndExtnTables = sConditionForMainAndExtnTables & _
            " AND " & gk_sDBCOLEditSetID & "=" & t_lEditSetID
        
        sConditionForLinkTables = sConditionForLinkTables & _
            " AND " & gk_sDBCOLEditSetID & "=" & t_lEditSetID
            
        sConditionForCatalogueLinkTable = sConditionForCatalogueLinkTable & _
            " AND " & gk_sDBCOLEditSetID & "=" & t_lEditSetID
        
    End If
    
    'An added pre-caution not to delete live entries
    
    If t_eDeleteFromSection = ESectionType.SECReader Then
            sSql = _
                " SELECT " & gk_sDBCOLLiveFlag & _
                " FROM " & GetTableName(cttExtensionTable, t_eDeleteFromSection, t_eCatalogueLevel) & _
                " WHERE " & GetColumnName(cctIDColumn, t_eCatalogueLevel) & _
                    "=" & t_lCatalogueID
            
            Set oFWRecordlist = m_oFWDBHelper.FindBySql(sSql)
            
            If oFWRecordlist.RecordCount = 0 Then
                RaiseError sSource, peEntryNotFoundInLiveTables, "Catalogue you are trying to delete is missing from live, can not delete this"
            End If
            
            If Not oFWRecordlist.GetFieldValue(gk_sDBCOLLiveFlag) = ELIVEFLAGSTATUS.lfsNotLive Then
                'Trying to delete a live entry
                RaiseError sSource, pETryingToDeleteNotToBeDeletedEntries, "Trying to delete not to be deleted entry"
            End If
    End If
    
    'Delete from accruals
    DeleteFromMainExtensionLinkTables _
        GetTableName(cttAccrualsTable, t_eDeleteFromSection, t_eCatalogueLevel), sConditionForMainAndExtnTables

    'Delete from admin history
    DeleteFromMainExtensionLinkTables _
        GetTableName(cttAdminHistoryTable, t_eDeleteFromSection, t_eCatalogueLevel), sConditionForMainAndExtnTables

    'Delete from app dest info
    DeleteFromMainExtensionLinkTables _
        GetTableName(cttAppDestructionTable, t_eDeleteFromSection, t_eCatalogueLevel), sConditionForMainAndExtnTables

    'Delete from arrangement
    DeleteFromMainExtensionLinkTables _
        GetTableName(cttArrangementTable, t_eDeleteFromSection, t_eCatalogueLevel), sConditionForMainAndExtnTables
    
    'Delete from custodial history
    DeleteFromMainExtensionLinkTables _
        GetTableName(cttCustodialHistoryTable, t_eDeleteFromSection, t_eCatalogueLevel), sConditionForMainAndExtnTables
    
    'Delete from note
    DeleteFromMainExtensionLinkTables _
        GetTableName(cttNoteTable, t_eDeleteFromSection, t_eCatalogueLevel), sConditionForMainAndExtnTables

    'Delete from scope and content
    DeleteFromMainExtensionLinkTables _
        GetTableName(cttScopeAndContentTable, t_eDeleteFromSection, t_eCatalogueLevel), sConditionForMainAndExtnTables

    'Delete from title
    DeleteFromMainExtensionLinkTables _
        GetTableName(cttTitleTable, t_eDeleteFromSection, t_eCatalogueLevel), sConditionForMainAndExtnTables
    
    If t_eCatalogueLevel = ECATALOGUELEVEL.ctLetterCode Then
        'This table is present only for lettercodes
        If t_eDeleteFromSection = ESectionType.SECeditorial Then
            DeleteFromMainExtensionLinkTables gk_sDBTABLEEditSetLetterCodeSortedTitle, sConditionForMainAndExtnTables
        Else
            DeleteFromMainExtensionLinkTables gk_sDBTABLELetterCodeSortedTitle, sConditionForMainAndExtnTables
        End If
    End If
    
    'Delete from extension table
    DeleteFromMainExtensionLinkTables _
        GetTableName(cttExtensionTable, t_eDeleteFromSection, t_eCatalogueLevel), _
        sConditionForMainAndExtnTables
    
    'Delete from main table
    DeleteFromMainExtensionLinkTables _
        GetTableName(cttMainTable, t_eDeleteFromSection, t_eCatalogueLevel), sConditionForMainAndExtnTables
    
    
    'Delete from the link tables
    
    'Person
    DeleteFromMainExtensionLinkTables GetTableName(cttPerson, t_eDeleteFromSection), sConditionForLinkTables
    
    'Place
    DeleteFromMainExtensionLinkTables GetTableName(cttPlace, t_eDeleteFromSection), sConditionForLinkTables
    
    'Subject
    DeleteFromMainExtensionLinkTables GetTableName(cttSubject, t_eDeleteFromSection), sConditionForLinkTables
    
    'Corporate body
    DeleteFromMainExtensionLinkTables GetTableName(cttCorporateBody, t_eDeleteFromSection), sConditionForLinkTables
    
    'SFA
    DeleteFromMainExtensionLinkTables GetTableName(cttSFA, t_eDeleteFromSection), sConditionForLinkTables
    
    'd1l added
    'Leaflets
    DeleteFromMainExtensionLinkTables GetTableName(cttLeaflet, t_eDeleteFromSection), sConditionForLinkTables
    
    'Link data item
    DeleteFromMainExtensionLinkTables GetTableName(cttLinkDataElement, t_eDeleteFromSection), sConditionForLinkTables
    
    'PC_link,es_pc_link
    DeleteFromMainExtensionLinkTables GetTableName(cttLink, t_eDeleteFromSection), sConditionForCatalogueLinkTable
    
    'Release objects
    Set oFWRecordlist = Nothing
    
    Exit Sub
ErrorHandler:
    Set oFWRecordlist = Nothing
    
    ConvertError sSource
End Sub
Private Sub DeleteFromMainExtensionLinkTables( _
    t_sTableName As String, _
    t_sConditionString As String)
'
' Description:  Deletes the records from the table t_sTableName satisfying the condition t_sConditionString
'
' Amendment history:
'   d1  ndk     22Oct1999       Created
    
    On Error GoTo ErrorHandler

    Dim sSource As String, sSql As String
    
    sSource = Me.Name & ".DeleteFromMainExtensionLinkTables"
    
    'construct the sql
    sSql = "DELETE " & t_sTableName & " WHERE " & t_sConditionString
    
    'delete the entries
    m_oFWDBHelper.Execute sSql
    
    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub


Private Function CatalogueHasChildrenInEditSet( _
    t_lCatalogueID As Long, _
    t_eCatalogueLevel As ECATALOGUELEVEL, _
    t_lEditSetID As Long) As Boolean
'
' Description: Checks whether the said catalogue has any children in the editset
'
' Amendment history:
'   d1  ndk     22Oct1999       Created
    
    On Error GoTo ErrorHandler

    Dim sSource As String, sConditionString As String
    
    sSource = Me.Name & ".CatalogueHasChildrenInEditSet"
    
    'condtion common for all queries
    
    sConditionString = _
        gk_sDBCOLEditSetID & "=" & t_lEditSetID
        
    Select Case t_eCatalogueLevel
    
        Case ECATALOGUELEVEL.ctLetterCode
        'lettercode could have children in division and class tables
    
            'division table
            If RecordsPresentInTable _
                (GetTableName(cttMainTable, ESectionType.SECeditorial, ECATALOGUELEVEL.ctDivision), _
                sConditionString, gk_sDBCOLLetterCodeID, t_lCatalogueID) Then
                
                CatalogueHasChildrenInEditSet = True
            
            'class table
            ElseIf RecordsPresentInTable( _
                GetTableName(cttMainTable, ESectionType.SECeditorial, ECATALOGUELEVEL.ctClass), _
                sConditionString, gk_sDBCOLLetterCodeID, t_lCatalogueID) Then
                
                CatalogueHasChildrenInEditSet = True
            End If
        
        Case ECATALOGUELEVEL.ctDivision
        'division could have children in class table

            If RecordsPresentInTable( _
                GetTableName(cttMainTable, ESectionType.SECeditorial, ECATALOGUELEVEL.ctClass), _
                sConditionString, gk_sDBCOLDivisionID, t_lCatalogueID) Then
                
                CatalogueHasChildrenInEditSet = True
            End If

        Case ECATALOGUELEVEL.ctClass
        'class could have children in header and piece tables

            'header table
            If RecordsPresentInTable( _
                GetTableName(cttMainTable, ESectionType.SECeditorial, ECATALOGUELEVEL.ctHeaders), _
                sConditionString, gk_sDBCOLClassID, t_lCatalogueID) Then
                
                CatalogueHasChildrenInEditSet = True
            
            'piece table
            ElseIf RecordsPresentInTable( _
                GetTableName(cttMainTable, ESectionType.SECeditorial, ECATALOGUELEVEL.ctPiece), _
                sConditionString, gk_sDBCOLClassID, t_lCatalogueID) Then
                
                CatalogueHasChildrenInEditSet = True
            End If

        Case ECATALOGUELEVEL.ctHeaders
        'header could have children in sub header or piece

            If RecordsPresentInTable( _
                GetTableName(cttMainTable, ESectionType.SECeditorial, ECATALOGUELEVEL.ctSubHeaders), _
                sConditionString, gk_sDBCOLHeaderID, t_lCatalogueID) Then
                
                CatalogueHasChildrenInEditSet = True
                
            ElseIf RecordsPresentInTable( _
                GetTableName(cttMainTable, ESectionType.SECeditorial, ECATALOGUELEVEL.ctPiece), _
                sConditionString, gk_sDBCOLHeaderID, t_lCatalogueID) Then
                
                CatalogueHasChildrenInEditSet = True
            End If

        Case ECATALOGUELEVEL.ctSubHeaders
        'subheader could have children in piece

            If RecordsPresentInTable( _
                GetTableName(cttMainTable, ESectionType.SECeditorial, ECATALOGUELEVEL.ctPiece), _
                sConditionString, gk_sDBCOLSubHeaderID, t_lCatalogueID) Then
                
                CatalogueHasChildrenInEditSet = True
            End If
        Case ECATALOGUELEVEL.ctPiece
        'piece could have children in item

            If RecordsPresentInTable( _
                GetTableName(cttMainTable, ESectionType.SECeditorial, ECATALOGUELEVEL.ctItem), _
                sConditionString, gk_sDBCOLPieceID, t_lCatalogueID) Then
                
                CatalogueHasChildrenInEditSet = True
            End If
            
        Case ECATALOGUELEVEL.ctItem
            'there would be no children for item
            CatalogueHasChildrenInEditSet = False
    End Select
    
    Exit Function
ErrorHandler:
    ConvertError sSource
End Function


Private Function RecordsPresentInTable( _
    ByVal t_sTableName As String, _
    ByVal t_sConditionString As String, _
    ByVal t_sColumnName As String, _
    ByVal t_vColumnValue As Variant) As Boolean
    
'
' Description: Checks whether any records are present in the table which satisfy the
'              condition passed as parameter and the value in column t_sColumnName is equal to t_vColumnValue
'
' Amendment history:
'   d1  ndk     22Oct1999       Created
    
    On Error GoTo ErrorHandler

    Dim sSource As String, sSql As String, oFWRecordlist As FWRecordList
    
    sSource = Me.Name & ".RecordsPresentInTable"
    
    'sql to get the records
    sSql = _
        " SELECT * FROM " & t_sTableName & _
        " WHERE " & t_sConditionString & " AND " & _
        t_sColumnName & "=" & t_vColumnValue
        
    Set oFWRecordlist = m_oFWDBHelper.FindBySql(sSql)
    
    If oFWRecordlist.RecordCount = 0 Then   'No records matching the condition
        RecordsPresentInTable = False
    Else
        RecordsPresentInTable = True
    End If
    
    Set oFWRecordlist = Nothing
    
    Exit Function
ErrorHandler:
    ConvertError sSource
End Function

Private Sub AddCatlogueToEditSetAsReadOnly( _
    ByVal t_lCatalogueID As Long, _
    ByVal t_eCatalogueLevel As ECATALOGUELEVEL, _
    ByVal t_lEditSetID As Long)
'
' Description: Copies a catalogue as readonly from live to editorial
'              (only the catalogue and not its children from live to editorial)
'
' Amendment history:
'   d1  ndk     22Oct1999       Created
'   d7  David Manning  02-feb-2000  Fixed defect 283 (sp_Edit_AddADivisioncodeToEditSet should be called sp_Edit_AddADivisionToEditSet)
    
    On Error GoTo ErrorHandler

    Dim sSource As String, sSql As String
    
    sSource = Me.Name & ".AddCatlogueToEditSetAsReadOnly"

    'get the appropirate sp name depending on level
    Select Case t_eCatalogueLevel
        Case ECATALOGUELEVEL.ctLetterCode
            sSql = "sp_Edit_AddALettercodeToEditSet"
        Case ECATALOGUELEVEL.ctDivision
            sSql = "sp_Edit_AddADivisionToEditSet"
        Case ECATALOGUELEVEL.ctClass
            sSql = "sp_Edit_AddAClassToEditSet"
        Case ECATALOGUELEVEL.ctHeaders
            sSql = "sp_Edit_AddAHeaderToEditSet"
        Case ECATALOGUELEVEL.ctSubHeaders
            sSql = "sp_Edit_AddASubHeaderToEditSet"
        Case ECATALOGUELEVEL.ctPiece
            sSql = "sp_Edit_AddAPieceToEditSet"
        Case ECATALOGUELEVEL.ctItem
            sSql = "sp_Edit_AddAItemToEditSet"
    End Select

    'Parameters for the sp
    sSql = _
        sSql & " " & _
        "@t_iEditsetID  = " & t_lEditSetID & "," & _
        "@t_iCatalogueID = " & t_lCatalogueID & "," & _
        "@t_iEditsetEntryType = " & ECATALOGUEENTRYSTATUS.cesReadeOnly & "," & _
        "@t_iLiveFlagValueToSet = " & ELIVEFLAGSTATUS.lfsLive

    'Add the catalogue
    m_oFWDBHelper.Execute sSql
    
    'Add the links of the catalogue
    sSql = "sp_Edit_AddAllPCLinks"
    
    'Parameters for the sp
    sSql = _
        sSql & " " & _
        "@t_iEditsetID = " & t_lEditSetID & "," & _
        "@t_iCatalogueID = " & t_lCatalogueID & "," & _
        "@t_iCatalogueLevel = " & GetCatalogueLevelCode(t_eCatalogueLevel)
    
    m_oFWDBHelper.Execute sSql
    
    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub




Private Sub UpdateLiveStatusFlag( _
    t_lCatalogueID As Long, _
    t_eCatalogueLevel As ECATALOGUELEVEL, _
    t_eFlagValue As ELIVEFLAGSTATUS)
    
'
' Description:     Updates the live_flag column in live tables to the paremeter passed
'
' Amendment history:
'   d1  ndk     22Oct1999       Created
    
    On Error GoTo ErrorHandler

    Dim sSource As String, sSql As String
    
    sSource = Me.Name & ".UpdateLiveStatusFlag"
    
    'constuct the sql
    sSql = _
        " UPDATE " & GetTableName(cttExtensionTable, ESectionType.SECReader, t_eCatalogueLevel) & _
        " SET " & gk_sDBCOLLiveFlag & "=" & t_eFlagValue & _
        " WHERE " & GetColumnName(cctIDColumn, t_eCatalogueLevel) & "=" & t_lCatalogueID
    
    'exeute the query to update the table
    m_oFWDBHelper.Execute sSql
    
    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub

Private Function GetRecordListFromTable( _
    ByVal t_sTableName As String, _
    ByVal t_lEditSetID As Long, _
    ByVal t_sIDColumnName As String, _
    ByVal t_lIDColumnValue As Long, _
    Optional ByVal t_sColumn1ToBeNull As String = "", _
    Optional ByVal t_sColumn2ToBeNull As String = "") As FWRecordList
    
'
' Description:     Returns a record list from the editorial table where t_sIDColumnName  = t_vIDColumnValue
'                  If the t_sColumn1ToBeNull and t_sColumn2ToBeNull are not empty then the the records returned
'                  should have these columns set as null
'
' Amendment history:
'   d1  ndk     22Oct1999       Created
'   d8  D Johnston  03Mar2000   Changed GetRecordListFromTable to check for division columns to be zero
'                               instead of null.  Also fixed recursion error in DeleteCatalogueChildren...
    
    On Error GoTo ErrorHandler

    Dim sSource As String, sSql As String
    
    sSource = Me.Name & ".GetRecordListFromTable"
    
    'basic query
    sSql = _
        " SELECT * FROM " & t_sTableName & _
        " WHERE " & _
        gk_sDBCOLEditSetID & "=" & t_lEditSetID & " AND " & _
        t_sIDColumnName & "=" & t_lIDColumnValue
    
    If Not t_sColumn1ToBeNull = "" Then
        ' d8 - Check for division
        If t_sColumn1ToBeNull = gk_sDBCOLDivisionID Then
            sSql = sSql & " AND " & t_sColumn1ToBeNull & " = 0"
        Else
            sSql = sSql & " AND " & t_sColumn1ToBeNull & " IS  Null"
        End If
    End If
    
    If Not t_sColumn2ToBeNull = "" Then
        ' d8 - Check for division
        If t_sColumn2ToBeNull = gk_sDBCOLDivisionID Then
            sSql = sSql & " AND " & t_sColumn2ToBeNull & " = 0"
        Else
            sSql = sSql & " AND " & t_sColumn2ToBeNull & " IS  Null"
        End If
    End If
    
    Set GetRecordListFromTable = m_oFWDBHelper.FindBySql(sSql)
    
    Exit Function
ErrorHandler:
    ConvertError sSource
End Function

Private Sub ResetRecordStatus()
'   Resets the record status of the catalogue entry to what it
'   was when it was added from live
'
' Amendment history:
'   d5  wdp     10Dec1999       Created
    
    On Error GoTo ErrorHandler

    Dim sSource As String, sSql As String
    
    sSource = Me.Name & ".ResetRecordStatus"
    
    'basic query
    sSql = " sp_edit_reset_record_status " & m_lEditsetID & ", " & m_lCatalogueID & ", " & m_eCatalogueLevelID
    
    m_oFWDBHelper.Execute sSql
    
    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub

