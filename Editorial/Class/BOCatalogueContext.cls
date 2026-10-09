VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BOCatalogueContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    BOCatalogueContext.cls
' System:       PROCAT, PRO
' Copyright:    (C) Quidnunc Limited
'
' Description:  Manages the moving the catalogue to a different parent
'               The catalogues that can be moved are class and piece
'
' Amendment history:
'   d1   ndk        31Oct1999      Created
'   d2   ndk        04Nov1999      Editset entry status updated when catalogue is moved
'   d3   ndk        15Nov1999      Completed lots of incomplete functionality
'   d4   alm        19Oct2001      Added function MoveMultipleCatalogueToNewParent and
'                                  changed the error msgs so that the cat ref and level
'                                  are shown
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'Property property variables
Private m_lEditsetID As Long
Private m_lCatalogueID As Long
Private m_eCatalogueLevelID As ECATALOGUELEVEL
Private m_lProposedImmedieteParentID As Long
Private m_eProposedImmedieteParentLevelID As ECATALOGUELEVEL

'framework and business objects
Private m_oFWDBHelper As FWDBHelper
Private m_ofwRecordlist As FWRecordList

'local variables
Private m_lHeaderIDOfSubHeader As Long
Private m_lCurrentImmedieteParentID As Long, m_eCurrentImmedieteParentLevel As ECATALOGUELEVEL  'these can be lcode,div,cls,head,sub
Private m_lCurrentReferentialParentID As Long, m_eCurrentReferentialParentLevel As ECATALOGUELEVEL 'these can be class or piece
Private m_lProposedReferentialParentID As Long, m_eProposedReferentialParentLevel As ECATALOGUELEVEL  'these can be class or piece

Private Sub Initialise()
    Set m_oFWDBHelper = New FWDBHelper
End Sub
Private Sub Terminate()
    Set m_oFWDBHelper = Nothing
    Set m_ofwRecordlist = Nothing
End Sub
Private Sub Class_Initialize()
    Initialise
End Sub
Private Sub Class_Terminate()
    Terminate
End Sub
Public Property Get Name()
    Name = "BOCatalogueContext"
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
Public Property Let ProposedImmedieteParentID(ByVal t_lProposedImmedieteParentID As Long)
    m_lProposedImmedieteParentID = t_lProposedImmedieteParentID
End Property
Public Property Let ProposedImmedieteParentLevelID(ByVal t_eProposedImmedieteParentLevelID As ECATALOGUELEVEL)
    m_eProposedImmedieteParentLevelID = t_eProposedImmedieteParentLevelID
End Property

Public Sub MoveMultipleCatalogueToNewParent(t_lCatalogueID() As Long, _
                                            t_eCatalogueLevel() As ECATALOGUELEVEL)
'
' Description: Moves multiple catalogue entries from one parent to another parent
'              Based on MoveCatalogueToNewParent()
'
' Amendment history:
'   d4   alm      16Oct2001      Created (based on MoveCatalogueToNewParent)
'
On Error GoTo ErrorHandler
    
    Dim oBOCatalogueDBHelper As BOCatalogueDBHelper
    Dim sSource              As String
    Dim i                    As Long
    
    sSource = Me.Name & ".MoveMultipleCatalogueToNewParent"
    
    'Make sure that the number of catalogueID elements match Catalogue levels
    If UBound(t_lCatalogueID) <> UBound(t_eCatalogueLevel) Then
        RaiseError sSource, peEditSet, "Catalogue ID and Catalogue Level List do not match"
    End If
    
    ' Check that all catalogue elements are movable
    For i = 1 To UBound(t_lCatalogueID)
    
        ' Set internal properties
        m_lCatalogueID = t_lCatalogueID(i)
        m_eCatalogueLevelID = t_eCatalogueLevel(i)
        
        ' See whether the entry being moved is a movable catalalogue
        ' Moveable catalogues are class and piece and whether the target parent
        ' is of the valid level. The method raises an error if not so
        CheckWhetherCatalogueIsMovable
        
        ' Set value to variables holding referential parent of the current catalogue and
        ' the referential parent of proposed parent.
        ' after setting these values are validated
        ValidateAndSetReferentialParentIDAndLevel
        
        ' Check that you are not trying to move the entry to the same parent
        If m_lCurrentImmedieteParentID = m_lProposedImmedieteParentID And _
            m_eCurrentImmedieteParentLevel = m_eProposedImmedieteParentLevelID Then
            
            RaiseError sSource, pECanNotMoveCatalogueToExisintParent, _
            " You are trying to move the catalogue to its existing parent " & _
            " Catalogue Ref - " & GenerateCatalogueReferenceFromID(m_lCatalogueID, m_eCatalogueLevelID) & _
            " Catalogue Level - " & GetCatalogueLevelDescription( _
                                    GetCatalogueLevelCode(m_eCatalogueLevelID))
        
        End If
        
        ' Now every thing is validated
    
        ' If you are moving under a catalogue which does not exist in editorial, copy
        ' the catalogue as readonly (in case of subheader, also copy its header)
        CopyNewParentsIfNotPresentInEditset
        
        ' To move the catalogue
        MoveCatalogue
    
        ' Update the status that catalogue was modified
        Set oBOCatalogueDBHelper = New BOCatalogueDBHelper
        
        oBOCatalogueDBHelper.SetCatalogueStatusAsModified _
        GetTableName(cttExtensionTable, SECeditorial, m_eCatalogueLevelID), _
        gk_sDBCOLEditSetID & "=" & m_lEditsetID & " AND " & _
        GetColumnName(cctIDColumn, m_eCatalogueLevelID) & "=" & m_lCatalogueID
    
        Set oBOCatalogueDBHelper = Nothing
    
        ' Delete the old parent of the catalogue if it is a read only
        ' and has no dependents
        DeleteOldParentCatalogueIfNoChildren
    
    Next i
    
    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub
                        
Public Sub MoveCatalogueToNewParent()
'
'Description: Moves the catalogue entry from one parent to another parent
'
' Amendment history:
'   d1   ndk      25Aug1999      Created
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String, sSql As String
    
    Dim iEditSetEntryType As Integer
    Dim oBOCatalogueDBHelper As BOCatalogueDBHelper

    sSource = Me.Name & ".MoveCatalogueToNewParent"
    
    'initialise variables
    Set oBOCatalogueDBHelper = New BOCatalogueDBHelper
    
    'see whether the entry being moved is a movable catalalogue
    'moveable catalogues are class and piece and whether the target parent
    'is of the valid level.
    'The method raises an error if not so
    
    CheckWhetherCatalogueIsMovable
    
    
    'set value to variables holding referential parent of the current catalogue and]
    'the referential parent of proposed parent.
    'after setting these values are validated
    
    ValidateAndSetReferentialParentIDAndLevel

    If m_lCurrentImmedieteParentID = m_lProposedImmedieteParentID And _
        m_eCurrentImmedieteParentLevel = m_eProposedImmedieteParentLevelID Then
        'trying to put the catalogue under the same parent
        
        RaiseError sSource, pECanNotMoveCatalogueToExisintParent, _
            " You are trying to move the catalogue to its existing parent"
        
    End If
    
    'now every thing is validated
    
    'if you are moving under a catalogue which does not exist in editorial, copy
    'the catalogue as readonly (in case of subheader, also copy its header)
    CopyNewParentsIfNotPresentInEditset
    
    'to move the catalogue
    MoveCatalogue
    
    'update the status that catalogue was modified
    oBOCatalogueDBHelper.SetCatalogueStatusAsModified _
        GetTableName(cttExtensionTable, SECeditorial, m_eCatalogueLevelID), _
        gk_sDBCOLEditSetID & "=" & m_lEditsetID & " AND " & _
        GetColumnName(cctIDColumn, m_eCatalogueLevelID) & "=" & m_lCatalogueID
    
    'to delete the old parent of the catalogue if it is a read only
    'and has no dependents
        
    DeleteOldParentCatalogueIfNoChildren
        
    'release objects
    
    Set oBOCatalogueDBHelper = Nothing
    
    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub



Private Sub MoveCatalogue()
'
'Actually runs the sql to move the catalogue to a different parent
'
' Amendment history:
'   d3   ndk        15Nov1999
'

    On Error GoTo ErrorHandler
    
    Dim sSource As String, sSql As String
    
    sSource = Me.Name & "." & "MoveCatalogue"
    
    Select Case m_eProposedImmedieteParentLevelID
        
        Case ECATALOGUELEVEL.ctLetterCode
        
            'just set division to null as lettercode would already be present
            sSql = _
                " UPDATE " & GetTableName(cttMainTable, ESectionType.SECeditorial, ECATALOGUELEVEL.ctClass) & _
                " SET " & _
                    gk_sDBCOLDivisionID & "= 0" & _
                " WHERE " & _
                    gk_sDBCOLClassID & "=" & m_lCatalogueID & " AND " & _
                    gk_sDBCOLEditSetID & "=" & m_lEditsetID
        
        Case ECATALOGUELEVEL.ctDivision
    
            'update division, leave lettercode untouched
            sSql = _
                " UPDATE " & GetTableName(cttMainTable, ESectionType.SECeditorial, ECATALOGUELEVEL.ctClass) & _
                " SET " & _
                    gk_sDBCOLDivisionID & "=" & m_lProposedImmedieteParentID & _
                " WHERE " & _
                    gk_sDBCOLClassID & "=" & m_lCatalogueID & " AND " & _
                    gk_sDBCOLEditSetID & "=" & m_lEditsetID
        
        Case ECATALOGUELEVEL.ctClass
    
            'just set header and subheader to null as classid would already be present
            sSql = _
                " UPDATE " & GetTableName(cttMainTable, ESectionType.SECeditorial, ECATALOGUELEVEL.ctPiece) & _
                " SET " & _
                    gk_sDBCOLHeaderID & "= Null" & "," & _
                    gk_sDBCOLSubHeaderID & "= Null" & _
                " WHERE " & _
                    gk_sDBCOLPieceID & "=" & m_lCatalogueID & " AND " & _
                    gk_sDBCOLEditSetID & "=" & m_lEditsetID
        
        Case ECATALOGUELEVEL.ctHeaders
    
            'update header, set the subheader to NULL,leave alone class
            sSql = _
                " UPDATE " & GetTableName(cttMainTable, ESectionType.SECeditorial, ECATALOGUELEVEL.ctPiece) & _
                " SET " & _
                    gk_sDBCOLHeaderID & "=" & m_lProposedImmedieteParentID & "," & _
                    gk_sDBCOLSubHeaderID & "= Null" & _
                " WHERE " & _
                    gk_sDBCOLPieceID & "=" & m_lCatalogueID & " AND " & _
                    gk_sDBCOLEditSetID & "=" & m_lEditsetID
                
        Case ECATALOGUELEVEL.ctSubHeaders
            'update subheader and header, leave alone class
            sSql = _
            " UPDATE " & GetTableName(cttMainTable, ESectionType.SECeditorial, ECATALOGUELEVEL.ctPiece) & _
            " SET " & _
                gk_sDBCOLSubHeaderID & "=" & m_lProposedImmedieteParentID & "," & _
                gk_sDBCOLHeaderID & "=" & m_lHeaderIDOfSubHeader & _
            " WHERE " & _
                gk_sDBCOLPieceID & "=" & m_lCatalogueID & " AND " & _
                gk_sDBCOLEditSetID & "=" & m_lEditsetID
    End Select
    
    m_oFWDBHelper.Execute sSql

    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub

Private Sub DeleteOldParentCatalogueIfNoChildren()
'
'   Deletes the current parent of the catalogue which was moved to a new parent
'   The deletion is done only if the current parent is readonly and has no other chidlren
'   This method uses to bodeletecatalogue object to accomplish this task
'
' Amendment history:
'   d3   ndk        15Nov1999   Created
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String, sSql As String, oBODeleteCatalogue As BODeleteCatalogue
    
    sSource = Me.Name & ".DeleteOldParentCatalogueIfNoChildren"
    
    Set oBODeleteCatalogue = New BODeleteCatalogue
    
    sSql = _
        " SELECT " & gk_sDBCOLEntryTypeID & _
        " FROM " & GetTableName(cttExtensionTable, ESectionType.SECeditorial, m_eCurrentImmedieteParentLevel) & _
        " WHERE " & GetColumnName(cctIDColumn, m_eCurrentImmedieteParentLevel) & "=" & m_lCurrentImmedieteParentID & " AND " & _
                    gk_sDBCOLEditSetID & "=" & m_lEditsetID
        
    'delete it if it is read only
    If m_oFWDBHelper.FindBySql(sSql).GetFieldValue(gk_sDBCOLEntryTypeID) = _
        ECATALOGUEENTRYSTATUS.cesReadeOnly Then
        
        'set the properties
        oBODeleteCatalogue.EditsetID = m_lEditsetID
        oBODeleteCatalogue.CatalogueID = m_lCurrentImmedieteParentID
        oBODeleteCatalogue.CatalogueLevelID = m_eCurrentImmedieteParentLevel
        
        'delete
        oBODeleteCatalogue.DeleteEditSetCatalogue
            
    End If
    
    Set oBODeleteCatalogue = Nothing
    
    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub





Private Sub CopyNewParentsIfNotPresentInEditset()
'
'   Copies the new parent and grandparents into the editset if they are not already
'   present.
'   This method calls a utitliy stored procedures which does this
'
' Amendment history:
'   d3   ndk        15Nov1999   Created
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String, sSql As String
    
    sSource = Me.Name & ".CopyNewParentsIfNotPresentInEditset"
    
    sSql = _
        "sp_Edit_AddCatalogueEntryDetailsWithParents " & _
            " @t_iEditsetID = " & m_lEditsetID & _
            ",@t_iCatalogueID = " & m_lProposedImmedieteParentID & _
            ",@t_iCatalogueLevel =" & GetCatalogueLevelCode(m_eProposedImmedieteParentLevelID) & _
            ",@t_iEditsetEntryTypeValueToSet = " & ECATALOGUEENTRYSTATUS.cesReadeOnly & _
            ",@t_iLiveFlagValueToSet = " & ELIVEFLAGSTATUS.lfsLive & _
            ",@t_iReadOnlyEditsetEntryType = " & ECATALOGUEENTRYSTATUS.cesReadeOnly
                
    m_oFWDBHelper.Execute sSql
    
    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub


Private Sub CheckWhetherCatalogueIsMovable()
'
'   Checks whether the catalogue is movable, and the target parent is of
'   valid level
'
' Amendment history:
'   d3   ndk        15Nov1999   Created
'   d4   alm        16Oct2001   Error messages now display the cat ref and
'                               level
    On Error GoTo ErrorHandler
    
    Dim sSource   As String
    Dim sSql      As String
    Dim sCatRef   As String
    Dim sCatLevel As String
    
    sSource = Me.Name & ".CheckWhetherCatalogueIsMovable"
    
    ' d4 Get the Catalogue Ref and Level
    sCatRef = GenerateCatalogueReferenceFromID(m_lCatalogueID, m_eCatalogueLevelID)
    sCatLevel = GetCatalogueLevelDescription(GetCatalogueLevelCode(m_eCatalogueLevelID))
    
    'only a class or piece can be moved
    If m_eCatalogueLevelID <> ECATALOGUELEVEL.ctClass And m_eCatalogueLevelID <> ECATALOGUELEVEL.ctPiece Then
        RaiseError sSource, PECanNotMoveCataloguesOtherThanClassOrPiece, _
        " The catalogue you are trying to move is not a class or a piece, it was not moved " & _
        " Catalogue Ref - " & sCatRef & _
        " Catalogue Level - " & sCatLevel
    End If
    
    'check whether the level of the new parent selected is OK
    If m_eCatalogueLevelID = ECATALOGUELEVEL.ctClass Then
        'current catalogue level is class
        'new level has to be a division
        If m_eProposedImmedieteParentLevelID <> ECATALOGUELEVEL.ctDivision And _
            m_eProposedImmedieteParentLevelID <> ECATALOGUELEVEL.ctLetterCode Then
            
            RaiseError sSource, PECanNotMoveCatalogueToInvalidParent, _
            " You are trying to move the catalogue to in invalid parent, it was not moved " & _
            " Catalogue Ref - " & sCatRef & _
            " Catalogue Level - " & sCatLevel

        End If
    Else    'piece
        'current catalogue is a piece
        'new level has to be a header or subheader
        If m_eProposedImmedieteParentLevelID <> ECATALOGUELEVEL.ctHeaders And _
            m_eProposedImmedieteParentLevelID <> ECATALOGUELEVEL.ctSubHeaders And _
            m_eProposedImmedieteParentLevelID <> ECATALOGUELEVEL.ctClass Then
            
            RaiseError sSource, PECanNotMoveCatalogueToInvalidParent, _
            " You are trying to move the catalogue to in invalid parent, it was not moved " & _
            " Catalogue Ref - " & sCatRef & _
            " Catalogue Level - " & sCatLevel

        End If
    End If
    
    'see if the catalogue is present in any other editset
    'if so , raise an error - can move only those entries which are not present in any editset
    
    sSql = _
        " SELECT " & gk_sDBCOLEditSetID & _
        " FROM " & GetTableName(cttMainTable, SECeditorial, m_eCatalogueLevelID) & _
        " WHERE " & GetColumnName(cctIDColumn, m_eCatalogueLevelID) & "=" & m_lCatalogueID & _
            " AND NOT " & gk_sDBCOLEditSetID & " = " & m_lEditsetID
    
    If m_oFWDBHelper.FindBySql(sSql).RecordCount <> 0 Then
        RaiseError sSource, PECanNotMoveCataloguePresentInMoreThanOneEditset, _
            " The catalogue you are trying to move is present in more than one editset, it was not moved " & _
            " Catalogue Ref - " & sCatRef & _
            " Catalogue Level - " & sCatLevel
    End If
    
    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub

    
Private Sub ValidateAndSetReferentialParentIDAndLevel()
'
'   Sets the referential parent of the current catalogue and the proposed
'   parent
'
' Amendment history:
'   d3   ndk        15Nov1999   Created
'   d4   alm        16Oct2001   Error message shows the Cat Ref and Level
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String, sSql As String
    
    sSource = Me.Name & ".ValidateAndSetReferentialParentIDAndLevel"

    'get the Referential parent of the current catalogue (lettercode in case of class, class in case of piece)
    'Check in editorial
    sSql = _
        " SELECT * FROM " & GetTableName(cttMainTable, ESectionType.SECeditorial, m_eCatalogueLevelID) & _
        " WHERE " & _
            GetColumnName(cctIDColumn, m_eCatalogueLevelID) & "=" & m_lCatalogueID & "  AND " & _
            gk_sDBCOLEditSetID & "=" & m_lEditsetID

    Set m_ofwRecordlist = m_oFWDBHelper.FindBySql(sSql)
        
    If m_eCatalogueLevelID = ECATALOGUELEVEL.ctClass Then
    
        'Referential parents
        m_lCurrentReferentialParentID = m_ofwRecordlist.GetFieldValue(gk_sDBCOLLetterCodeID)
        m_eCurrentReferentialParentLevel = ECATALOGUELEVEL.ctLetterCode

        'current immediete parent id
        If IsNull(m_ofwRecordlist.GetFieldValue(gk_sDBCOLDivisionID)) Or _
            m_ofwRecordlist.GetFieldValue(gk_sDBCOLDivisionID) = 0 Then
            
            m_lCurrentImmedieteParentID = m_ofwRecordlist.GetFieldValue(gk_sDBCOLLetterCodeID)
            m_eCurrentImmedieteParentLevel = ECATALOGUELEVEL.ctLetterCode
        Else
            m_lCurrentImmedieteParentID = m_ofwRecordlist.GetFieldValue(gk_sDBCOLDivisionID)
            m_eCurrentImmedieteParentLevel = ECATALOGUELEVEL.ctDivision
        End If
    Else    'for piece
    
        'Referential parents
        m_lCurrentReferentialParentID = m_ofwRecordlist.GetFieldValue(gk_sDBCOLClassID)
        m_eCurrentReferentialParentLevel = ECATALOGUELEVEL.ctClass
    
        'current immediete parent id
        If (IsNull(m_ofwRecordlist.GetFieldValue(gk_sDBCOLSubHeaderID)) Or _
            m_ofwRecordlist.GetFieldValue(gk_sDBCOLSubHeaderID) = 0) And _
            (IsNull(m_ofwRecordlist.GetFieldValue(gk_sDBCOLHeaderID)) Or _
            m_ofwRecordlist.GetFieldValue(gk_sDBCOLHeaderID) = 0) Then
    
            m_lCurrentImmedieteParentID = m_ofwRecordlist.GetFieldValue(gk_sDBCOLClassID)
            m_eCurrentImmedieteParentLevel = ECATALOGUELEVEL.ctClass
            
        ElseIf (IsNull(m_ofwRecordlist.GetFieldValue(gk_sDBCOLSubHeaderID)) Or _
            m_ofwRecordlist.GetFieldValue(gk_sDBCOLSubHeaderID) = 0) Then
            
            m_lCurrentImmedieteParentID = m_ofwRecordlist.GetFieldValue(gk_sDBCOLHeaderID)
            m_eCurrentImmedieteParentLevel = ECATALOGUELEVEL.ctHeaders
        Else
            m_lCurrentImmedieteParentID = m_ofwRecordlist.GetFieldValue(gk_sDBCOLSubHeaderID)
            m_eCurrentImmedieteParentLevel = ECATALOGUELEVEL.ctSubHeaders
        End If
    End If
    
    'get the Referential parent of the proposed parent (lettercode in case of divisoin, class in case of header,subheader)
    'check in reader
    sSql = _
        " SELECT * FROM " & GetTableName(cttMainTable, ESectionType.SECReader, m_eProposedImmedieteParentLevelID) & _
        " WHERE " & _
            GetColumnName(cctIDColumn, m_eProposedImmedieteParentLevelID) & "=" & m_lProposedImmedieteParentID

    Set m_ofwRecordlist = m_oFWDBHelper.FindBySql(sSql)
        
    If m_eCatalogueLevelID = ECATALOGUELEVEL.ctClass Then
        'Referential parent level and id
        m_eProposedReferentialParentLevel = ECATALOGUELEVEL.ctLetterCode
        m_lProposedReferentialParentID = m_ofwRecordlist.GetFieldValue(gk_sDBCOLLetterCodeID)
        
    Else    'for piece
        'Referential parent level and id
        m_eProposedReferentialParentLevel = ECATALOGUELEVEL.ctClass
        
        
        'for id
        If m_eProposedImmedieteParentLevelID = ECATALOGUELEVEL.ctSubHeaders Then
            'if the parent is a subheader, get its headerid also
            m_lHeaderIDOfSubHeader = m_ofwRecordlist.GetFieldValue(gk_sDBCOLHeaderID)
            
            'to find the class id
            sSql = _
                " SELECT * FROM " & GetTableName(cttMainTable, ESectionType.SECReader, ECATALOGUELEVEL.ctHeaders) & _
                " WHERE " & _
                GetColumnName(cctIDColumn, ECATALOGUELEVEL.ctHeaders) & "=" & m_lHeaderIDOfSubHeader
                
            Set m_ofwRecordlist = m_oFWDBHelper.FindBySql(sSql)
            
            m_lProposedReferentialParentID = m_ofwRecordlist.GetFieldValue(gk_sDBCOLClassID)
        Else    'for class and header
            m_lProposedReferentialParentID = m_ofwRecordlist.GetFieldValue(gk_sDBCOLClassID)
        End If
    End If

    'the referential parent of the catalogue to be moved and the referential parent of
    'the proposed parent have to be the same
    
    If m_lCurrentReferentialParentID <> m_lProposedReferentialParentID Or _
        m_eCurrentReferentialParentLevel <> m_eProposedReferentialParentLevel Then
        
        RaiseError sSource, PECanNotMoveCatalogueToParentOfDifferentLevel, _
            " The reference of the catalogue would change if moved to this parent , it was not moved " & _
            " Catalogue Ref - " & GenerateCatalogueReferenceFromID(m_lCatalogueID, m_eCatalogueLevelID) & _
            " Catalogue Level - " & GetCatalogueLevelDescription( _
                                    GetCatalogueLevelCode(m_eCatalogueLevelID))
    End If
    
    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub

