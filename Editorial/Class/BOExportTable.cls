VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BOExportTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    BOAccessionBatch.cls
' System:       PRO, PROCAT
' Copyright:    (C) Quidnunc Limited
'
' Description:  This class contains interactions with DRUID database
'   The following tables will be used
'       PROCAT_export
'
'   This table will then contain ALL the information on the export form, so we can decide
'   what depth and level restrictions to apply.
'
' Amendment history:
'   Vers    Auth        Date            Comment
'   d1      A.Farrow    24-Jul-2000     Created
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' These are member variables that are not actually put into the table but are used
'   in the error handling
Private m_lBatchID As Long
Private m_sBatchDescription As Variant

' Member variables that are set by the calling method
Private m_bIsDepartmentWanted As Boolean
Private m_bIsDivisionWanted As Boolean
Private m_bIsSeriesWanted As Boolean
Private m_bIsSubSeriesWanted As Boolean
Private m_bIsSubSubSeriesWanted As Boolean
Private m_bIsPieceWanted As Boolean
Private m_bIsItemWanted As Boolean

' The filename
Private m_sFileName As Variant

' The Lettercode Ref and ID
Private m_sLettercode As String
Private m_lLetterCodeID As Long

'The Piecekey order maximum and minimum
'Private m_lMinPKeyOrder As Long
'Private m_lMaxPKeyOrder As Long

Private m_bPROCATAccessionRecordPresent As Boolean

Private m_oFWDBHelper As FWDBHelper

Private m_objContext As ObjectContext


' Class initialisation and MTS event handlers
Private Sub Initialise()
    Set m_oFWDBHelper = New FWDBHelper
End Sub

Private Sub Class_Initialize()
    Initialise
End Sub

Public Function UpdateExportTable(t_sFilename As String, t_oboExportIDs As BOExportIDs, t_sSeriesIDs As String, t_lBatchID As Long)
'   Description:    Populate the export table with data
'
'   Vers    Auth    Date                Created
'   d1      ADF     24-July- 2000       Created

Dim sSource As String

    Dim sSql As String
    Dim sSQLInsertValues As String
    Set m_oFWDBHelper = New FWDBHelper
    
    sSource = "BOExportTable.UpdateExportTable"
    
    
    sSql = "INSERT INTO " & gk_sDBTABLEExport & " ( " & _
                gk_sDBCOLDeptWanted & "," & _
                gk_sDBCOLDivWanted & "," & _
                gk_sDBCOLClassWanted & "," & _
                gk_sDBCOLHeaderWanted & "," & _
                gk_sDBCOLSubheaderWanted & "," & _
                gk_sDBCOLPieceWanted & "," & _
                gk_sDBCOLItemWanted & "," & _
                gk_sDBCOLExportFilename & "," & _
                gk_sDBCOLLetterCodeID & "," & _
                gk_sDBCOLMinPieceOrder & "," & _
                gk_sDBCOLMaxPieceOrder & "," & _
                gk_sDBCOLBatchID & ")"

    sSQLInsertValues = " VALUES (" & _
                QuoteValueWithQuotes(m_bIsDepartmentWanted) & "," & _
                QuoteValueWithQuotes(m_bIsDivisionWanted) & "," & _
                QuoteValueWithQuotes(m_bIsSeriesWanted) & "," & _
                QuoteValueWithQuotes(m_bIsSubSeriesWanted) & "," & _
                QuoteValueWithQuotes(m_bIsSubSubSeriesWanted) & "," & _
                QuoteValueWithQuotes(m_bIsPieceWanted) & "," & _
                QuoteValueWithQuotes(m_bIsItemWanted) & "," & _
                QuoteValueWithQuotes(t_sFilename) & "," & _
                t_oboExportIDs.LetterCode & "," & _
                t_oboExportIDs.MinimumPieceKeyOrder & "," & _
                t_oboExportIDs.MaximumPieceKeyOrder & "," & _
                t_lBatchID & ")"
                
                '  Add the sql helper execute code here
    m_oFWDBHelper.Execute (sSql & sSQLInsertValues)
    
    Dim dummy As String
    dummy = InsertSeriesIDs(t_sSeriesIDs, t_lBatchID)
                
End Function



Public Property Let IsDepartmentWanted(t_bBoolean As Boolean)
     m_bIsDepartmentWanted = t_bBoolean
End Property

Public Property Let IsDivWanted(t_bBoolean As Boolean)
    m_bIsDivisionWanted = t_bBoolean
End Property

Public Property Let IsSeriesWanted(t_bBoolean As Boolean)
    m_bIsSeriesWanted = t_bBoolean
End Property

Public Property Let IsSubSeriesWanted(t_bBoolean As Boolean)
    m_bIsSubSeriesWanted = t_bBoolean
End Property

Public Property Let IsSubSubSeriesWanted(t_bBoolean As Boolean)
    m_bIsSubSubSeriesWanted = t_bBoolean
End Property

Public Property Let IsPieceWanted(t_bBoolean As Boolean)
    m_bIsPieceWanted = t_bBoolean
End Property
Public Property Let IsItemWanted(t_bBoolean As Boolean)
    m_bIsItemWanted = t_bBoolean
End Property



Public Function InsertSeriesIDs(t_sSeries As String, t_lBatchID As Long)
'   Description:    Convert the passed ASP sting into series IDs and put them in the database
'
'   Vers    Auth            Date                Created
'   d1      A.Farrow        24-July- 2000       Created

    Dim lPos As Long
    Dim sRestr As String
    Dim sSql As String
    Dim sSQLInsertValues As String
    
    sSql = "INSERT INTO " & gk_sDBTABLESeriesIDS & " ( " & _
                        gk_sDBCOLSeriesIDs & "," & _
                        gk_sDBCOLBatchID & ")"

    While Not t_sSeries = ""
        If Not t_sSeries = "" Then
            lPos = InStr(t_sSeries, ",")
            If lPos > 0 Then
                'remove the next restriction from the string
                sRestr = UCase(Trim(Left$(t_sSeries, lPos - 1)))
                t_sSeries = Mid$(t_sSeries, lPos + 1)
                
                sSQLInsertValues = " VALUES (" & _
                    CLng(sRestr) & "," & _
                    t_lBatchID & ")"
                m_oFWDBHelper.Execute (sSql & sSQLInsertValues)
            Else
                'If there are no commas left just use what remains
                sRestr = UCase(Trim(t_sSeries))
                t_sSeries = ""
                
                 sSQLInsertValues = " VALUES (" & _
                    CLng(sRestr) & "," & _
                    t_lBatchID & ")"
                    
                m_oFWDBHelper.Execute (sSql & sSQLInsertValues)
            End If
        End If
    Wend
        
        
    
End Function

Public Function GetExportInformation(t_lBatchID As Long) As FWRecordList
'   Description:    This function provides a means of getting all the export information back
'                   from the export tables using the batch ID passed via the AsyncLoader
'
'   Vers    Auth            Date                Comment
'   d1      A.Farrow        24-July-2000        Created


    ' Define the source
    Dim sSource As String
    sSource = "VWExport.GetExportInformation"
    
    Dim sSql As String
    
    sSql = "SELECT " & gk_sDBCOLDeptWanted & "," & _
                gk_sDBCOLDivWanted & "," & _
                gk_sDBCOLClassWanted & "," & _
                gk_sDBCOLHeaderWanted & "," & _
                gk_sDBCOLSubheaderWanted & "," & _
                gk_sDBCOLPieceWanted & "," & _
                gk_sDBCOLItemWanted & "," & _
                gk_sDBCOLExportFilename & "," & _
                gk_sDBCOLLetterCodeID & "," & _
                gk_sDBCOLMinPieceOrder & "," & _
                gk_sDBCOLMaxPieceOrder & "," & _
                gk_sDBCOLBatchID & " FROM " & _
                gk_sDBTABLEExport & " WHERE " & _
                gk_sDBCOLBatchID & " = " & t_lBatchID

                
                
    Dim ofwHelper As New FWDBHelper
    Dim ofwResult As FWRecordList
    Set ofwResult = ofwHelper.FindBySql(sSql)
    
    
    
    'We return the whole record list to be used
    Set GetExportInformation = ofwResult
    
End Function

Public Function GetSeriesIDs(t_lBatchID As Long) As FWRecordList
'   Description:    This function provides a means of getting all the series ids from the table
'                   series_ids using the same batch_id as the one used to retrive all other export
'                   info ( GetExportInformation())
'
'   Vers    Auth            Date                Comment
'   d1      A.Farrow        24-July-2000        Created

     ' Define the source
    Dim sSource As String
    
    sSource = "VWExport.GetSeriesIDs"
    
    Dim sSql As String
    
    sSql = "SELECT s." & gk_sDBCOLSeriesIDs & " FROM " & _
                gk_sDBTABLESeriesIDS & " s, " & _
                gk_sDBTABLEClass & " c" & _
                " WHERE s." & gk_sDBCOLSeriesIDs & " = c." & gk_sDBCOLClassID & _
                " AND s." & gk_sDBCOLBatchID & " = " & t_lBatchID & _
                " ORDER BY c." & gk_sDBCOLClassNumber

    Set GetSeriesIDs = m_oFWDBHelper.FindBySql(sSql)
    
End Function

Public Function GetChildren(t_lCatalogueID As Long, t_iCatalogueLevel As Integer, _
                            t_lLetterCodeID As Long, t_lBatchID As Long, _
                            t_lMinPOrder As Long, t_lMaxPOrder As Long) As FWRecordList
'
'   Description:    This function is the VB code that triggers the GetChildren SP
'
'   Vers    Auth            Date                Comment
'   d1      A.Farrow        25-July-2000        Created

    
     ' Define the source
    Dim sSource As String

    sSource = "BOExportTable.GetChildren"
    
    Dim sSql As String
    Dim ofwHelper As New FWDBHelper
    Dim ofwResult As FWRecordList
    
    sSql = "Execute sp_Edit_GetDirectChildrenList @t_iCatalogueID = " & t_lCatalogueID & _
            " , @t_iCatalogueLevel = " & t_iCatalogueLevel & _
            " , @t_iLetterCodeID = " & t_lLetterCodeID & _
            " , @t_iBatchID = " & t_lBatchID & _
            " , @t_iPieceKeyMin = " & t_lMinPOrder & _
            " , @t_iPiecekeyMax = " & t_lMaxPOrder
    
    Set ofwResult = ofwHelper.FindBySql(sSql)
    
    Set GetChildren = ofwResult


End Function
