VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BOExportSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    BOExportHandler.cls
' System:       PRO, PROCAT
' Copyright:    (C) Quidnunc Limited
'
' Description:  This class is the BO that is responsible for getting a summary count of
'               entries to export based on the list of IDs supplied.
'
' Amendment history:
'   Vers    Auth    Date            Comment
'   d1      ADF     4-Jul-2000      Created
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'The member variables holding the level counts
Dim m_iLetterCount As Long
Dim m_iDivisionCount As Long
Dim m_iClassCount As Long
Dim m_iHeaderCount As Long
Dim m_iSubheaderCount As Long
Dim m_iPieceCount As Long
Dim m_iItemCount As Long


'The member variables to hold the export ids
Dim m_oboExportIds As BOExportIDs








Public Sub GetSummary(t_eLowestLevel As ECATALOGUELEVEL)
'
'   Description:    This uses the IDs passed to call the SP and get back a count for each level.
'                   The SP has to be called recursively if there is more than one Series reference,
'                   and a check is made that there no duplicate divisions are counted.
'
'   Vers    Auth    Date                Created
'   d1      ADF     18-July- 2000       Created

On Error GoTo ErrorHandler
    
    ' Define the source in case of error
    Dim sSource As String
    sSource = "BOExportSummary.GetSummary"
    
    Dim fwEvaluateSeries As fwRecordList
    Dim sBatchStringID As String
    
    ' Get the IDs from the passed object
    Set fwEvaluateSeries = m_oboExportIds.SeriesIDS
    
    ' Set the table up in the database
    With m_oboExportIds
        sBatchStringID = CStr(.LetterCode) & "_" & CStr(.MaximumPieceKeyOrder) & "_" & _
            CStr(.MinimumPieceKeyOrder) & "_" & CStr(Date) & "_" & CStr(Time)
    End With
    
    'See if the Series references are empty
    If Not fwEvaluateSeries.RecordCount = 0 Then
        ' If it is not then recursively call the Execute to get a total list
        fwEvaluateSeries.MoveFirst
        While Not fwEvaluateSeries.IsAtEnd
            SetSeriesRestriction sBatchStringID, fwEvaluateSeries.GetFieldValue(gk_sDBCOLSeriesIDs)
            fwEvaluateSeries.MoveNext
        Wend
    End If
            
    ' Now execute the summary with the identifier
    ExecuteSummary sBatchStringID, t_eLowestLevel
    
    ClearSeriesRestrictions sBatchStringID
    
    Exit Sub

ErrorHandler:
    ConvertError sSource
End Sub


Private Sub SetSeriesRestriction(t_sBatch As String, t_lSeriesID As Long)
' Sets the series restrictions for this summary in the database so that the
' summary stored procedures can join all the restrictions together simultaneously
'
    Dim sSql As String
    Dim ofwQuery As FWDBHelper
    
    Set ofwQuery = New FWDBHelper
    
    sSql = "INSERT INTO " & gk_sDBTABLEExportSeriesRestrictions & _
        " (" & gk_sDBCOLBatchID & ", " & gk_sDBCOLClassID & ") VALUES ('" & _
        t_sBatch & "', " & CStr(t_lSeriesID) & ")"
    
    ofwQuery.Execute sSql
    
End Sub

Private Sub ClearSeriesRestrictions(t_sBatch As String)
' Sets the series restrictions for this summary in the database so that the
' summary stored procedures can join all the restrictions together simultaneously
'
    Dim sSql As String
    Dim ofwQuery As FWDBHelper
    
    Set ofwQuery = New FWDBHelper
    
    sSql = "DELETE " & gk_sDBTABLEExportSeriesRestrictions & _
        " WHERE " & gk_sDBCOLBatchID & " = '" & t_sBatch & "'"
    
    ofwQuery.Execute sSql
    
End Sub

''''''''''''''''' Implementation ''''''''''''''''''''''


' Class initialisation and MTS event handlers
Private Sub Initialise()

End Sub

Private Sub Class_Initialize()

    Initialise
End Sub


Public Property Let CatalogueIDS(t_oboExportIDs As BOExportIDs)
        Set m_oboExportIds = New BOExportIDs
        Set m_oboExportIds = t_oboExportIDs
End Property
'   Description:    Check if the array containing the series is empty
'
'   Vers    Auth    Date                Created
'   d1      ADF     18-July- 2000       Created
Public Function IsArrayEmpty(t_iArray As Variant) As Boolean

    Dim lngUBound As Long
    
    On Error Resume Next
    
    lngUBound = UBound(t_iArray)
    If Err.Number <> 0 Then
        IsArrayEmpty = True
    Else
        IsArrayEmpty = False
    End If
        
End Function

Public Property Get LetterCount() As Variant
    LetterCount = m_iLetterCount
End Property

Public Property Get DivisionCount() As Variant
    DivisionCount = m_iDivisionCount
End Property
Public Property Get ClassCount() As Variant
    ClassCount = m_iClassCount
End Property
Public Property Get HeaderCount() As Variant
    HeaderCount = m_iHeaderCount
End Property
Public Property Get SubheaderCount() As Variant
    SubheaderCount = m_iSubheaderCount
End Property
Public Property Get PieceCount() As Variant
    PieceCount = m_iPieceCount
End Property
Public Property Get ItemCount() As Variant
    ItemCount = m_iItemCount
End Property

Private Function ExecuteSummary(t_sBatchID As String, t_eLowestLevel As ECATALOGUELEVEL)
'
'   Description: Execute Export summary stored procedure to get a count
'                   of the number of entries to be exported
'
'
'   Vers    Auth    Date                Created
'   d1      ADF     19-July- 2000       Created


    Dim sSource As String
    sSource = "BOExportSummary.ExecuteSummary"

    ' Get the IDs and piece key order from the passed object
    Dim lLetterCodeID As Long
    Dim lPieceKeyMin As Long
    Dim lPieceKeyMax As Long
    
    lLetterCodeID = m_oboExportIds.LetterCode
    lPieceKeyMin = m_oboExportIds.MinimumPieceKeyOrder
    lPieceKeyMax = m_oboExportIds.MaximumPieceKeyOrder
    
    ' Here we check that piece key order is correct and if not reverse it for the user
    If lPieceKeyMin > lPieceKeyMax Then
        Dim temp As Long
        temp = lPieceKeyMax
        lPieceKeyMax = lPieceKeyMin
        lPieceKeyMin = temp
    End If

    ' The variables used to retrieve the count
    Dim sSql As String
    Dim ofwResult As fwRecordList
    Dim ofwQuery As FWDBHelper
    
    Set ofwQuery = New FWDBHelper

    'if there are no class restrictions
    
    'Execute the query to get the summary counts
    sSql = "Execute sp_Edit_GetExportSummary @t_iLettercodeID = " & lLetterCodeID & _
            " , @t_vBatchID = '" & t_sBatchID & "' ,  @t_iPieceKeyMin = " & lPieceKeyMin & _
            " , @t_iPieceKeyMax = " & lPieceKeyMax & " , @t_iLowestLevel = " & t_eLowestLevel
            
    Set ofwResult = ofwQuery.FindBySql(sSql)

    If ofwResult.RecordCount = 0 Then
        RaiseError sSource, PEGeneralInternalError, "Summary stored procedure did not return any rows"
    End If
    
    ' Take the result set and depending on the catalogue level get a count
    While Not ofwResult.IsAtEnd
       Select Case ofwResult.GetFieldValue(gk_sDBCOLCatalogueLevel)
            Case ECATALOGUELEVEL.ctLetterCode
                m_iLetterCount = ofwResult.GetFieldValue("entry_count")
            Case ECATALOGUELEVEL.ctDivision
                m_iDivisionCount = ofwResult.GetFieldValue("entry_count")
            Case ECATALOGUELEVEL.ctClass
                m_iClassCount = ofwResult.GetFieldValue("entry_count")
            Case ECATALOGUELEVEL.ctHeaders
                m_iHeaderCount = ofwResult.GetFieldValue("entry_count")
            Case ECATALOGUELEVEL.ctSubheaders
                m_iSubheaderCount = ofwResult.GetFieldValue("entry_count")
            Case ECATALOGUELEVEL.ctPiece
                m_iPieceCount = ofwResult.GetFieldValue("entry_count")
            Case ECATALOGUELEVEL.ctItem
                m_iItemCount = ofwResult.GetFieldValue("entry_count")
        End Select
        ofwResult.MoveNext
    Wend
    
Exit Function
ErrorHandler:
    ConvertError sSource
End Function
