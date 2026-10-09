VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BOExportHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    BOExportHandler.cls
' System:       PRO, PROCAT
' Copyright:    (C) Quidnunc Limited
'
' Description:  This class is the BO that is repsonsible for parsing references and
'               the creation of the summary count.
'
' Amendment history:
'   Vers    Auth    Date            Comment
'   d1      ADF     4-Jul-2000      Created
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit


Private m_objContext As Object
Private m_aspResponse As Response
Private m_aspRequest As Request
Private m_aspServer As Server

'Search instance number
Private m_sSearchNumber As String

'Search object
Private m_boSearch As BOSearch

Private m_oFWAppData As FWApplicationData

Private m_lLoggedUserID  As Long
Private m_oBOEditorialPermissions As BOEditorialPermissions

'Used as part of form validation
Private m_iNumberOfSeries As Integer

'The Current level being checked (Dept, Series or Piece)
Private m_sLevelBeingChecked As String

' Member variable to make sure that all the series references are under the lettercode
Private m_sCurrentLetterCode As String
Private m_sCurrentSeriesRef As String

' The member variables that contain the IDs need to execute the summary
Private m_iLetterCodeID As Long

' The min and maximum piece key order
Private m_lPieceKeyMin As Long
Private m_lPieceKeyMax As Long


' The member variable that contain the series IDs
Private m_iSeriesIDS() As Long

'User stats message
Private m_boMessage As BOMessage

Public Sub Initialise()
    ReDim m_iSeriesIDS(0)
End Sub

Public Sub ParseRestrictions(t_sRestrictions)
'   Parses the restrictions string (from the basic search form) and
'   adds breadth restrictions to the search. Note this operates under
'   the assumption that the restrictions are a comma separated list
'   of strings which themselves do not contain any commas.
'
'   Vers    Auth    Date            Comment
'   d1      ADF     4-Jul-2000      Created

On Error GoTo ErrorHandler
    
    Dim sSource As String
    sSource = "VWValidateExport.ParseRestrictions"
        
    Dim sRestrictions As String
    sRestrictions = t_sRestrictions
    Dim colRestrs As New Collection
    
    
    ' Loop round removing the next restriction from the string, providing it is not a lettercode
    ' in which case we trim it and leave it, as it should be just one entry, any commas in the
    ' the lettercode are invalid and will produce an error
  
    While Not sRestrictions = "" And Not (m_sLevelBeingChecked = "department" Or m_sLevelBeingChecked = "piece")
        
        Dim lPos As Long
        Dim sRestr As String
        
        'Find the next comma
        lPos = InStr(sRestrictions, ",")
        If lPos > 0 Then
            'remove the next restriction from the string
            sRestr = UCase(Trim(Left$(sRestrictions, lPos - 1)))
            sRestrictions = Mid$(sRestrictions, lPos + 1)
        Else
            'If there are no commas left just use what remains
            sRestr = UCase(Trim(sRestrictions))
            sRestrictions = ""
        End If
        
        'Check that the user has not entered two comma's ( ',,' ) or other invalid entries
        If Not sRestr = "" Then
            sRestr = m_sCurrentLetterCode & sRestr
            m_sCurrentSeriesRef = sRestr
            colRestrs.Add sRestr
        End If

    
    Wend
    
    ' If this is a lettercode just get it all
    If m_sLevelBeingChecked = "department" Then
        Dim sDeptRestr As String
        sDeptRestr = UCase(Trim(sRestrictions))
        colRestrs.Add sDeptRestr
    End If
    
    ' We have no built the collection we are going to send to verity
    
    ' Create the new export checker object
    Dim oboExpChecker As New BOExportChecker
    
    ' If we are checking a piece object then we check it and get and retrive its piece key order
    If m_sLevelBeingChecked = "piece" Then
        Dim sPiece As String
        sPiece = UCase(Trim(sRestrictions))
        sPiece = m_sCurrentSeriesRef & "/" & sPiece
        colRestrs.Add sPiece
        If oboExpChecker.VerifyReference(colRestrs, REFPRO) Then
        
            '  See which of the piece keys we are retrieving
            If m_lPieceKeyMin = 0 Then
                m_lPieceKeyMin = oboExpChecker.PieceKey
            Else
                m_lPieceKeyMax = oboExpChecker.PieceKey
            End If
            
            
        Else
            RaiseError sSource, PEUSERInvalidBreadthRestrictions, "Invalid " & m_sLevelBeingChecked _
                & " restrictions: " & sRestrictions
        End If
    End If
    
    If Not m_sLevelBeingChecked = "piece" Then
        If oboExpChecker.VerifyReference(colRestrs, , m_sCurrentLetterCode) Then
            'If we have parsed the entry succesfully then we get the ID back
            If m_sLevelBeingChecked = "department" Then
                m_iLetterCodeID = oboExpChecker.LetterCode
            End If
       
            ' if we have verified a number of series then put all those into an array
            If m_sLevelBeingChecked = "series" Then
                m_iSeriesIDS() = oboExpChecker.SeriesRefs
            End If
            
        Else
            Dim sInvalidRestrs As String
            Dim m As Long
            For m = 1 To colRestrs.Count
            'The invalid restrs are returned in colRestrs
                If Not sInvalidRestrs = "" Then
                    sInvalidRestrs = sInvalidRestrs & ", "
                End If
                sInvalidRestrs = sInvalidRestrs & colRestrs(m)
            Next
            
            RaiseError sSource, PEUSERInvalidBreadthRestrictions, "Invalid " & m_sLevelBeingChecked & " restriction: " & sInvalidRestrs, sInvalidRestrs
        End If
    End If
    
    'Store the number of series referenced
    m_iNumberOfSeries = colRestrs.Count
    
Exit Sub

ErrorHandler:
    'm_objContext.SetComplete
    ConvertError sSource
End Sub

Public Property Let LevelBeingChecked(t_sLevelBeingChecked As String)
    m_sLevelBeingChecked = t_sLevelBeingChecked
End Property

Public Property Let CurrentLetterCode(t_sLettercode As String)
    m_sCurrentLetterCode = t_sLettercode
End Property


Public Function RetrieveIDs() As BOExportIDs
'   Description : Put all the member variables of this class into ExportIDs object
'
'   Vers    Auth    Date            Comment
'   d1      ADF     4-Jul-2000      Created

    ' Put all this classes member variables into an ID business object to be returned
    Dim oboExportIds As New BOExportIDs
    oboExportIds.LetterCode = m_iLetterCodeID
    oboExportIds.SeriesIDS = ArrayToRecordList(m_iSeriesIDS())
    oboExportIds.MinimumPieceKeyOrder = m_lPieceKeyMin
    oboExportIds.MaximumPieceKeyOrder = m_lPieceKeyMax
    
    'Return this object
    Set RetrieveIDs = oboExportIds
    
End Function


Public Property Get NumberOfSeries() As Integer
    NumberOfSeries = m_iNumberOfSeries
End Property

Private Function ArrayToRecordList(t_lSeriesIDs As Variant) As fwRecordList

    Dim lValue As Variant
    Dim rsRecordSet As ADODB.Recordset
    Dim fwRecordList As fwRecordList
    
    Set rsRecordSet = New Recordset
    rsRecordSet.CursorLocation = adUseClient
    rsRecordSet.Fields.Append gk_sDBCOLSeriesIDs, adBSTR
    rsRecordSet.Open
    
    If UBound(t_lSeriesIDs) > 0 Then
        For Each lValue In t_lSeriesIDs
            If Not lValue = 0 Then
                rsRecordSet.AddNew
                rsRecordSet.Fields(gk_sDBCOLSeriesIDs).Value = lValue
                rsRecordSet.UPDATE
            End If
        Next
    End If
    
    Set fwRecordList = New fwRecordList
    fwRecordList.SetData rsRecordSet
    
    Set ArrayToRecordList = fwRecordList
End Function

Public Property Get PieceKeyMinimum() As Long
    PieceKeyMinimum = m_lPieceKeyMin
End Property

Public Property Get PieceKeyMaximum() As Long
    PieceKeyMaximum = m_lPieceKeyMax
End Property

