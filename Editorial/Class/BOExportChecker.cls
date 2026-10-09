VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BOExportChecker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    BOExportChecker.cls
' System:       PRO, PROCAT
' Copyright:    (C) Quidnunc Limited
'
' Description:  Class used to search for catalogue references and validate them for export.
'               It is similar in content to BORefChecker, although it contains
'               more structure checking e.g checking that a series reference falls within
'               a specified lettercode
'
'
'
' Amendment history:
'   Vers    Auth            Date            Comments
'   d1      A. Farrow       14-July-2000    Created
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Const mk_lMaxRestrictions As Long = 20

' The member variables that contian the IDs we get when searching
Dim m_SeriesRef() As Long
Dim m_RefCount As Integer
Dim m_iLetterCodeID As Long
Dim m_iPieceKey As Long

'Dim m_iMinimumPieceKey As Integer
'Dim m_iMaximumPiecekey As Integer

Dim m_oiResults As ISearchResults
Dim m_oboSearch As BOSearch


Public Function VerifyReference(ByRef t_sRefs As Collection, Optional t_eRefType As Variant, Optional t_sLettercode As String) As Boolean
' Amendment history:
'   Vers    Auth            Date            Comments
'   d1      A. Farrow       14-July-2000    Created'
'
'
'   Method used by the export handler to use verity in verifying a referece.
'   t_eRefType indicates the type of reference
'   we are searching for (pro ref, former department ref etc.)
'   Essentially if verity makes no hits then this entry does not exist and is invalid.
'
'   d1  ADF     14-July-2000     Created
'
On Error GoTo ErrorHandler
    
    '[ In the process of building BO to hold IDS ]
    ''Dim obExportID As BOExportIDs
    Dim sSource As String
    'Dim m_SeriesRef() As Long
    sSource = "BOExportChk.VerifyReference"
    Dim dockey As String
    Dim dataid As Long
    
    
    m_RefCount = 0
    
    Dim sStrippedRef As String
    Dim sField As String
    Dim bIsBreadthRestriction As Boolean
    Dim eRefType As EReferenceType
    
    If IsMissing(t_eRefType) Then
        bIsBreadthRestriction = True
        eRefType = REFPRO
    Else
        bIsBreadthRestriction = False
        eRefType = t_eRefType
    End If
    
    Select Case eRefType
        Case REFPRO
            sField = gk_sDIStrippedRef
        Case REFFormerDepartmental
            sField = gk_sDIStrippedFormerRef
        Case REFFormerPRO, REFAllFormerRefs
            sField = gk_sDIStrippedFormerPRORef
        Case Else
            RaiseError sSource, PEGeneralInternalError, "invalid reference type enum."
    End Select
    
    'Set up the new verity search
    Set m_oboSearch = New BOSearch
    m_oboSearch.SearchType = Reference
    
    If bIsBreadthRestriction = False Then
        m_oboSearch.Context = " a reference search for """ & t_sRefs.Item(1) & """"
    End If
    
    Dim oboTerms() As New BOCriteriaTerm
    Dim iCount As Integer
    For iCount = 0 To t_sRefs.Count - 1
        'Trim and strip the ref of whitespace
        sStrippedRef = Replace(Trim(t_sRefs.Item(iCount + 1)), " ", "")
        sStrippedRef = Replace(sStrippedRef, vbTab, "")
        ReDim Preserve oboTerms(iCount + 1)
        oboTerms(iCount).Operator = CTOEquality
        oboTerms(iCount).TermType = CTTValue
        oboTerms(iCount).Value = sStrippedRef
        oboTerms(iCount).Item = sField
    Next
    
    If eRefType = REFAllFormerRefs Then
        'If we are searcing over all former refs
        'which we only do for "go to" we need
        'to add another search term
        Dim oboTerm2 As New BOCriteriaTerm
        Dim oboComb As New BOCombinationTerm
        
        oboTerm2.Operator = CTOEquality
        oboTerm2.TermType = CTTValue
        oboTerm2.Value = sStrippedRef
        oboTerm2.Item = gk_sDIStrippedFormerRef
        
        oboComb.TermType = CTTOR
        oboComb.AddTerm oboTerms(0)
        oboComb.AddTerm oboTerm2
        
        m_oboSearch.RootSearchTerm = oboComb
        
    Else
        If t_sRefs.Count = 1 Then
            m_oboSearch.RootSearchTerm = oboTerms(0)
        Else
            Dim oboCombo As New BOCombinationTerm
            oboCombo.TermType = CTTOR
            Dim l As Long
            For l = 0 To t_sRefs.Count - 1
                oboCombo.AddTerm oboTerms(l)
            Next
            m_oboSearch.RootSearchTerm = oboCombo
        End If
    End If
    
    'Add depth restrictions is we are verifying breadth restrictions
    If bIsBreadthRestriction = True Then
        m_oboSearch.AddDepthRestriction ctLetterCode
        m_oboSearch.AddDepthRestriction ctClass
    End If
    
    'If there is no breadth rstriction we can specify a depth restriciton of piece
    If bIsBreadthRestriction = False Then
        m_oboSearch.AddDepthRestriction ctPiece
    End If
    
    Dim oboSearchExec As New BOSearchExecutive
    oboSearchExec.Search = m_oboSearch
    oboSearchExec.RangeMin = 1
    oboSearchExec.RangeMax = 20
    
    oboSearchExec.RunSearch
    Set m_oiResults = New ISearchResults
    Set m_oiResults = oboSearchExec.Results
    
    If m_oiResults.ResultCount = 0 And Not m_oiResults.ResultType = RESULTTooMany Then
        VerifyReference = False
    Else
        
        If bIsBreadthRestriction = False Then
            'For "go to" that is all the information we need for now
            dockey = m_oiResults.GetData(gk_sDIVerityDocKey)
            dataid = DBIDFromDocKey(dockey, ctPiece)
            
            'Get the piece key order
            m_iPieceKey = GetPieceOrder(dataid)
            
            
            
            VerifyReference = True
        Else
            'For breadth restrictions we need to check that each
            'of the restrictions has returned a match.
            Dim colBadRefs As New Collection
            Dim i As Integer
            For i = 1 To t_sRefs.Count
                Dim j As Integer
                Dim bHit As Boolean
                
                bHit = False
                
                m_oiResults.MoveFirst
                
                sStrippedRef = Replace(Trim(t_sRefs.Item(i)), " ", "")
                sStrippedRef = Replace(sStrippedRef, vbTab, "")
                
                ' If the lettercode parameter has been ommited we assume that we are parsing a
                ' a lettercode, we check that it is in fact a lettercode
                If t_sLettercode = "" Then
                    Do While Not m_oiResults.IsAtEnd
                        If m_oiResults.GetData(gk_sDIStrippedRef) = sStrippedRef And _
                        m_oiResults.GetData(gk_sDILetterCodeRef) = sStrippedRef Then
                        
                         'Get theID from the Dockey
                        dockey = m_oiResults.GetData(gk_sDIVerityDocKey)
                        dataid = DBIDFromDocKey(dockey, ctLetterCode)
                        
                        'Store the lettercode ID to use in getting the summary
                        m_iLetterCodeID = dataid
                        
                        bHit = True
                             Exit Do
                        End If
                        m_oiResults.MoveNext
                    Loop
                Else
                    '  if the lettercode has been included then we make sure that all the series
                    '  series reference fall within this lettercode
                    Do While Not m_oiResults.IsAtEnd
                        If m_oiResults.GetData(gk_sDIStrippedRef) = sStrippedRef And _
                        m_oiResults.GetData(gk_sDILetterCodeRef) = t_sLettercode Then
                        
                        'Get theID from the Dockey
                        dockey = m_oiResults.GetData(gk_sDIVerityDocKey)
                        
                        ' Create an array that holds all the series refs
                        dataid = DBIDFromDocKey(dockey, ctClass)
                        ReDim Preserve m_SeriesRef(m_RefCount + 1)
                        m_SeriesRef(m_RefCount) = dataid
                        m_RefCount = m_RefCount + 1
                    
                         bHit = True
                             Exit Do
                        End If
                        m_oiResults.MoveNext
                    Loop
                End If
                
                If bHit = False Then
                    colBadRefs.Add (t_sRefs.Item(i))
                End If
                
            Next
            If colBadRefs.Count = 0 Then
                VerifyReference = True
            Else
                VerifyReference = False
                Set t_sRefs = colBadRefs
            End If
            
        End If
    End If
    
    Exit Function
    
ErrorHandler:
    ConvertError sSource
End Function


Public Function GoToCatEntry(t_sRef As String) As Boolean
'   Method used in "go to cat entry"
'   Checks reference against PRO reference. If this is unsuccesful
'   then we check against former dept ref and former PRO ref
'
'   d9  WDP     27-Oct-1999     Created
On Error GoTo ErrorHandler
    
    Dim sSource As String
    sSource = "BORefChecker.GoToCatEntry"
    Dim colRefs As New Collection
    colRefs.Add t_sRef
    
    If VerifyReference(colRefs, REFPRO) Then
        GoToCatEntry = True
    Else
        If VerifyReference(colRefs, REFAllFormerRefs) Then
            GoToCatEntry = True
        Else
            GoToCatEntry = False
        End If
    End If
    
    Exit Function
    
ErrorHandler:
    ConvertError sSource
End Function

''''''''''''''''' Properties '''''''''''''''''

Public Property Get Results() As ISearchResults
    Set Results = m_oiResults
End Property

Public Property Get Search() As BOSearch
    Set Search = m_oboSearch
End Property

''''''''''''''''' Implementation ''''''''''''''''''''''


' Class initialisation and MTS event handlers
Private Sub Initialise()

End Sub

Private Sub Class_Initialize()

    Initialise
End Sub



Public Property Get LetterCode() As Long
    LetterCode = m_iLetterCodeID
End Property



Public Property Get SeriesRefs() As Long()
    SeriesRefs = m_SeriesRef()
End Property


Public Function GetPieceOrder(t_iID As Long)
'   Given a pieces ID get its piece key order
'
'   d1    ADF     14-July- 2000     Created
On Error GoTo ErrorHandler
    
    Dim sSource As String
    sSource = "GetPieceOrder"
    
    Dim sSql As String
    
    sSql = "SELECT " & gk_sDBCOLPieceKeyOrder & " AS " & gk_sDIID & " FROM " & gk_sDBTABLEPiece & " WHERE " & _
            gk_sDBCOLPieceID & "=" & t_iID
            
            
    Dim ofwHelper As New FWDBHelper
    Dim ofwResult As FWRecordList
    Set ofwResult = ofwHelper.FindBySql(sSql)
    
    If Not ofwResult.RecordCount = 1 Then
        RaiseError sSource, PEGeneralInternalError, _
            "Could not find any/unique ID for pitem "
    End If
    GetPieceOrder = ofwResult.GetFieldValue(gk_sDIID)
    
Exit Function
ErrorHandler:
    ConvertError sSource

End Function



Public Property Get PieceKey() As Long
     PieceKey = m_iPieceKey
End Property

