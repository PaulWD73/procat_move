VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BOEditorialDBUtility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    BOEditorialDBUtility.cls
' System:       PROCAT, PRO
' Copyright:    (C) Quidnunc Limited
' Instancing : Private
'
' Description:  Small Database utitlity methods
'
' Amendment history:
'
' d1    NDK     22Nov1999       Created.
' d2    NDK     10Dec1999       Added method CataloguePresentAndEditableInAnyOtherEditset
'

Public Property Get Name()
    Name = "BOEditorialDBUtility"
End Property
Public Function EditSetEntryTypeID( _
    ByVal t_lEditSetID As Long, _
    ByVal t_lCatalogueID As Long, _
    ByVal t_eCatalogueLevelID As ECATALOGUELEVEL) As ECATALOGUEENTRYSTATUS
'
' Returns the entry type of the editset entry
'
' d1    NDK     22Nov1999       Created.
'

    On Error GoTo ErrorHandler

    Dim sSQL As String, sSource As String
    Dim ofwDBHelper As FWDBHelper
    
    sSource = Me.Name & ".EditSetEntryTypeID"
    
    'initialise
    Set ofwDBHelper = New FWDBHelper

    'construct sql
    sSQL = _
        " SELECT " & gk_sDBCOLEditSetEntryTypeID & _
        " FROM " & GetTableName(cttExtensionTable, SECEditorial, t_eCatalogueLevelID) & _
        " WHERE " & GetColumnName(cctIDColumn, t_eCatalogueLevelID) & "=" & t_lCatalogueID & _
        " AND " & gk_sDBCOLEditSetID & "=" & t_lEditSetID
    
    EditSetEntryTypeID = _
        ofwDBHelper.FindBySql(sSQL).GetFieldValue(gk_sDBCOLEditSetEntryTypeID)
    
    'release
    Set ofwDBHelper = Nothing
    
    Exit Function
    
ErrorHandler:
    ConvertError sSource
End Function

Public Function CataloguePresentAndEditableInAnyOtherEditset( _
    ByVal t_lEditSetID As Long, _
    ByVal t_lCatalogueID As Long, _
    ByVal t_eCatalogueLevelID As ECATALOGUELEVEL) As ECATALOGUEENTRYSTATUS
'
' Checkes whether the catalogue is present in any editset other than t_lEditSetID and
' is not readony if present
'
' d2    NDK     10Dec1999       Created
'
    On Error GoTo ErrorHandler

    Dim sSQL As String, sSource As String
    Dim ofwDBHelper As FWDBHelper
    
    sSource = Me.Name & ".CataloguePresentAndEditableInAnyOtherEditset"
    
    'initialise
    Set ofwDBHelper = New FWDBHelper

    'construct sql
    sSQL = _
        " SELECT " & gk_sDBCOLEditSetID & _
        " FROM " & GetTableName(cttExtensionTable, SECEditorial, t_eCatalogueLevelID) & _
        " WHERE " & GetColumnName(cctIDColumn, t_eCatalogueLevelID) & "=" & t_lCatalogueID & _
            " AND NOT " & gk_sDBCOLEditSetEntryTypeID & " = " & ECATALOGUEENTRYSTATUS.cesReadeOnly & _
            " AND NOT " & gk_sDBCOLEditSetID & " = " & t_lEditSetID

    If ofwDBHelper.FindBySql(sSQL).RecordCount = 0 Then
        CataloguePresentAndEditableInAnyOtherEditset = False
    Else
        CataloguePresentAndEditableInAnyOtherEditset = True
    End If

    'release
    Set ofwDBHelper = Nothing
    
    Exit Function
    
ErrorHandler:
    ConvertError sSource
End Function



