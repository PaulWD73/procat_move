VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BOExportIDs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    BOExportIDs
' System:       PRO, PROCAT
' Copyright:    (C) Quidnunc Limited
'
' Description:  This class is a small class to hold the IDs of the lettercode series IDs
'               and piecekey order needed to export EAD files
'
' Amendment history:
'  Vers     Auth        Created         Comments
'   d1      A.Farrow    14-July-2000    Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'The Member variables to store the IDs in

' The lettercode ID
Dim m_iLetterCodeID As Long

' An array of the series References
Dim m_fwSeriesIDS As FWRecordList

'and the count of the number of series references that we have stored so far
Dim m_iSeriesCount As Long


'The minimum and maximum piecekey order
Dim m_iMinimumPieceKey As Long
Dim m_iMaximumPieceKey As Long

Public Property Get LetterCode() As Long
    LetterCode = m_iLetterCodeID
End Property

Public Property Let LetterCode(t_iLetterCodeID As Long)
    m_iLetterCodeID = t_iLetterCodeID
End Property


Public Property Get SeriesIDS() As FWRecordList
    Set SeriesIDS = m_fwSeriesIDS
End Property

Public Property Let SeriesIDS(t_fwSeriesReferences As FWRecordList)
    Set m_fwSeriesIDS = t_fwSeriesReferences
End Property

Public Property Get MinimumPieceKeyOrder() As Long
    MinimumPieceKeyOrder = m_iMinimumPieceKey
End Property

Public Property Let MinimumPieceKeyOrder(t_iMinOrder As Long)
    m_iMinimumPieceKey = t_iMinOrder
End Property


Public Static Property Get MaximumPieceKeyOrder() As Long
    MaximumPieceKeyOrder = m_iMaximumPieceKey
End Property

Public Static Property Let MaximumPieceKeyOrder(t_iMaxOrder As Long)
    m_iMaximumPieceKey = t_iMaxOrder
End Property
