VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "VWEditSetExt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    VWEditsetExt.cls
' System:       PROCAT, PRO
' Copyright:    (C) Quidnunc Limited
'
' Description:  View catalogue browser class with all display items for catalogues
'
' Amendment history:
'   d1   ksk     08Jun99      Created by copying from VWCatalogueBrowser.
'   d2   ksk     12Jun99      Removed the copied code and started a fresh.
'   d3   ksk     02Aug99      Moved View code to VWEditset
'   d4   ksk     02Aug99      Checked out for temp use for creation of catalogue
'                               list display
'   d5   ksk     22Oct99      Changed Catalogue_level_no to Level_No
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private Const ksNameCatalogueSearch As String = "FCS"
Private Const ksNameCheckBoxCount As String = "CBN"
Private Const ksNameCheckBoxes As String = "CHKBX"



'This width is absolute because the text wraps with in the frame size if not absolute
Private Const ksBrowserTableWidth As String = "800"

'Error message for the summary details not found. This will arise if the query failed
' of the query returned an empty recordlist.
Private Const ksCatalogueSummaryNotFoundHelp As String = "Catalogue summary not found"

Private Const ksCatalogueLevelColumnWidth As String = "3%"

'Constants for Parameters for values passes to and from asp
Private Const ksParamCatalogueID As String = "CID"
Private Const ksParamCatalogueLevel As String = "CL"
Private Const ksParamSearchDirection As String = "SD"
Private Const ksParamReference As String = "REF"
Private Const ksParamCatalogueTitle As String = "CT"
Private Const ksParamSortOrder As String = "SOD"

Private Const ksNextPageCode As String = "NPC"
Private Const ksNameOpenImage As String = "OCAT"
Private Const ksNameButtonMore As String = "btnMore"




Private m_oFWAppData As FWApplicationData
Private m_ofwRefData As FWReferenceData
Private m_oFWHTML As FWHTML
Private m_objContext As ObjectContext
Private m_aspResponse As Response
Private m_aspRequest As Request
Private m_aspServer As Server
Private m_aspSession As Session

Implements ObjectControl



'NOTE THIS IS A TEMPORARY FUNCTION STUB THAT IS TO BE REPLACED LATER
Public Sub StartCatalogueSearchForm()
    Dim sEditsetParam As String
    Dim sSource As String
    Dim lEditsetID  As Long
    
    sSource = "VWEditSetExt.StartCatalogueSearchForm"
    
    If m_aspRequest(gk_sParamEditSetID) = "" Then
        RaiseError sSource, EPROCATINTERNALERROR.peEditSetIDInvalue, _
                    "Editset ID not found"
    End If
    lEditsetID = CLng(m_aspRequest(gk_sParamEditSetID))
    If lEditsetID <= 0 Then
        RaiseError sSource, EPROCATINTERNALERROR.peEditSetIDInvalue, _
                            "Editset ID not found"
    End If
    sEditsetParam = "?" & gk_sParamEditSetID & "=" & lEditsetID
    
    m_aspResponse.Write m_oFWHTML.FormStart(gk_sLinkDisplayCatalogueSearchResult & _
                sEditsetParam, gk_sMainFrame, ksNameCatalogueSearch, MPost)
End Sub

'NOTE THIS IS A TEMPORARY FUNCTION STUB THAT IS TO BE REPLACED LATER
Public Sub EndCatalogueSearchForm()
    m_aspResponse.Write m_oFWHTML.EndTag(gk_sHTMLForm)
End Sub


'NOTE THIS IS A TEMPORARY FUNCTION STUB THAT IS TO BE REPLACED LATER
Public Sub DisplayCatalogueSearchResult()
'This function is to be remove and rewritten once the search screen and search result screen are received
'
On Error GoTo ErrHandler
    Dim sTitle As String
    Dim bL As Boolean, bD As Boolean, bC As Boolean, bH As Boolean, bS As Boolean, bP As Boolean, bI As Boolean
    Dim sSql As String
    Dim ofwQuery As FWDBHelper
    Dim ofwResultList As FWRecordList
    Dim iCurrentNumber As Integer
    Dim ofwDispTable As FWTable
    Dim typCellStyle As CellStyle
    Dim sEditsetParam  As String
    Dim lEditsetID  As Long
    Dim sSource As String
    
    sSource = "VWEditsetExt.DisplayCatalogueSearchResult"
    
    Set ofwDispTable = New FWTable
    
    If m_aspRequest(gk_sParamEditSetID) = "" Then
        RaiseError sSource, EPROCATINTERNALERROR.peEditSetIDInvalue, _
                    "Editset ID not found"
    End If
    lEditsetID = CLng(m_aspRequest(gk_sParamEditSetID))
    If lEditsetID <= 0 Then
        RaiseError sSource, EPROCATINTERNALERROR.peEditSetIDInvalue, _
                            "Editset ID not found"
    End If
    sEditsetParam = "?" & gk_sParamEditSetID & "=" & lEditsetID
    
    
    iCurrentNumber = 1
    sTitle = m_aspRequest("txttitle")
    bL = IIf(IsEmpty(m_aspRequest("L")), False, True)
    bD = IIf(IsEmpty(m_aspRequest("D")), False, True)
    bC = IIf(IsEmpty(m_aspRequest("C")), False, True)
    bH = IIf(IsEmpty(m_aspRequest("H")), False, True)
    bS = IIf(IsEmpty(m_aspRequest("S")), False, True)
    bP = IIf(IsEmpty(m_aspRequest("P")), False, True)
    bI = IIf(IsEmpty(m_aspRequest("I")), False, True)

    Set ofwQuery = New FWDBHelper
    
    m_aspResponse.Write m_oFWHTML.FormStart(gk_sLinkAddCataloguesToEditset & sEditsetParam _
                                                , gk_sMainFrame, "Test", MPost)
    
    'Start adding the Header row
    ofwDispTable.StartRow
    typCellStyle.eStyle = CELLCustomStyle
    typCellStyle.eBGCOLOR = COLOURRed
    typCellStyle.eALIGN = HALLeft
    typCellStyle.sWidth = "10%"
    ofwDispTable.AddHCellAsHTML m_oFWHTML.Font("Reference", COLOURWhite), typCellStyle
    typCellStyle.sWidth = "7%"
    ofwDispTable.AddHCellAsHTML m_oFWHTML.Font("Levels", COLOURWhite), typCellStyle
    typCellStyle.sWidth = "80%"
    ofwDispTable.AddHCellAsHTML m_oFWHTML.Font("Title", COLOURWhite), typCellStyle
    typCellStyle.sWidth = "3%"
    ofwDispTable.AddHCellAsHTML m_oFWHTML.Font("Add", COLOURWhite), typCellStyle
    ofwDispTable.EndRow
    
    
    If bL Then
        sSql = "Select Lettercode_id as catalogue_ID, 1 as Level_No, lettercode_Title as Title_text, Letter_code as Reference "
        sSql = sSql & " from tbl_lettercode(nolock) Where lettercode_Title like '%" & Trim(sTitle) & "%'"
        
        Set ofwResultList = ofwQuery.FindBySql(sSql, 25)
    End If
    DisplayResultInRecordList ofwResultList, iCurrentNumber, ofwDispTable
    
    If bD Then
        sSql = "Select division_id as catalogue_ID, 2 as Level_No, division_Title as Title_text, division_id as Reference "
        sSql = sSql & " from tbl_division(nolock) Where division_Title like '%" & Trim(sTitle) & "%'"
        Set ofwResultList = ofwQuery.FindBySql(sSql, 25)
    End If
    DisplayResultInRecordList ofwResultList, iCurrentNumber, ofwDispTable
    
    If bC Then
        sSql = "Select class_id as catalogue_ID, 3 as Level_No, class_Title as Title_text, class_id as Reference "
        sSql = sSql & " from tbl_class(nolock) Where class_Title like '%" & Trim(sTitle) & "%'"
        Set ofwResultList = ofwQuery.FindBySql(sSql, 25)
    End If
    DisplayResultInRecordList ofwResultList, iCurrentNumber, ofwDispTable
    
    If bH Then
        sSql = "Select header_id as catalogue_ID, 4 as Level_No, header_Title as Title_text, header_id as Reference "
        sSql = sSql & " from tbl_header(nolock) Where header_Title like '%" & Trim(sTitle) & "%'"
        Set ofwResultList = ofwQuery.FindBySql(sSql, 25)
    End If
    DisplayResultInRecordList ofwResultList, iCurrentNumber, ofwDispTable
    
    If bS Then
        sSql = "Select subheader_id as catalogue_ID, 5 as Level_No, subheader_Title as Title_text, subheader_id as Reference "
        sSql = sSql & " from tbl_subheader(nolock) Where subheader_Title like '%" & Trim(sTitle) & "%'"
        Set ofwResultList = ofwQuery.FindBySql(sSql, 25)
    End If
    DisplayResultInRecordList ofwResultList, iCurrentNumber, ofwDispTable
    
    If bP Then
        sSql = "Select piece_id as catalogue_ID, 6 as Level_No, piece_scope as Title_text, piece_id as Reference "
        sSql = sSql & " from tbl_piece(nolock) Where piece_scope like '%" & Trim(sTitle) & "%'"
        Set ofwResultList = ofwQuery.FindBySql(sSql, 25)
    End If
    DisplayResultInRecordList ofwResultList, iCurrentNumber, ofwDispTable
    If bI Then
        sSql = "Select item_id as catalogue_ID, 7 as Level_No, item_scope as Title_text, item_id as Reference "
        sSql = sSql & " from tbl_item(nolock) Where item_scope like '%" & Trim(sTitle) & "%'"
        Set ofwResultList = ofwQuery.FindBySql(sSql, 25)
    End If
    DisplayResultInRecordList ofwResultList, iCurrentNumber, ofwDispTable
    
    m_aspResponse.Write ofwDispTable.GetHTML
    
    m_aspResponse.Write m_oFWHTML.HiddenField(ksNameCheckBoxCount, CStr(iCurrentNumber))
    m_aspResponse.Write m_oFWHTML.SubmitImage(gk_sImagesPath & gk_sImageOK, "OK", "OK")
    m_aspResponse.Write m_oFWHTML.EndTag(gk_sHTMLForm)
    
    m_objContext.SetComplete
    Exit Sub
ErrHandler:
    m_objContext.SetComplete
End Sub


'NOTE THIS IS A TEMPORARY FUNCTION STUB THAT IS TO BE REPLACED LATER
Private Sub DisplayResultInRecordList(t_ResultList As FWRecordList, ByRef t_iStartNumber As Integer, _
                            ByRef t_ofwDisplayTable As FWTable)
    
    Dim sValue As String
    Dim iNumber As Integer
        
    If Not t_ResultList Is Nothing Then
        If t_ResultList.RecordCount > 0 Then
            
            While Not t_ResultList.IsAtEnd
                t_ofwDisplayTable.StartRow
                
                
                t_ofwDisplayTable.AddCellAsString GenerateCatalogueReferenceFromID( _
                                    t_ResultList.GetFieldValue(gk_sDBCOLCatalogueID), _
                                    t_ResultList.GetFieldValue(gk_sDBCOLCatalogueLevel))
                
                t_ofwDisplayTable.AddCellAsString GetCatalogueLevelDescription( _
                                        t_ResultList.GetFieldValue(gk_sDBCOLCatalogueLevel))
                
                t_ofwDisplayTable.AddCellAsString Left(t_ResultList.GetFieldValue(gk_sDBCOLTitle), _
                                                gk_iMaxTitleDisplayChars)
                
                'Store in the value field the catalogue ID and Catalogue level values
                sValue = CStr(t_ResultList.GetFieldValue(gk_sDBCOLCatalogueID)) & gk_sDataDelimiter & _
                        CStr(t_ResultList.GetFieldValue(gk_sDBCOLCatalogueLevel))
                
                sValue = m_oFWHTML.Checkbox("", ksNameCheckBoxes & Trim(CStr(t_iStartNumber)), False, sValue)
                t_iStartNumber = t_iStartNumber + 1
                
                t_ofwDisplayTable.AddCellAsHTML sValue
                
                t_ofwDisplayTable.EndRow
                
                t_ResultList.MoveNext
            Wend
        End If
    End If
    
    
End Sub





















''''''''''''''''' Implementation ''''''''''''''''''''''

'
' Private functions
'


'
' Class initialisation and MTS event handlers
'

Private Sub Initialise()
    Set m_objContext = GetObjectContext()
    If Not (m_objContext Is Nothing) And m_objContext.Count <> 0 Then
        Set m_aspResponse = m_objContext.Item("Response")
        Set m_aspRequest = m_objContext.Item("Request")
        Set m_aspServer = m_objContext.Item("Server")
        Set m_aspSession = m_objContext.Item("Session")
    End If
    Set m_oFWHTML = New FWHTML
    Set m_oFWAppData = New FWApplicationData
    Set m_ofwRefData = New FWReferenceData
End Sub

Private Sub Terminate()
    Set m_objContext = Nothing
    Set m_aspResponse = Nothing
    Set m_aspRequest = Nothing
    Set m_aspServer = Nothing
    Set m_aspSession = Nothing
    Set m_oFWHTML = Nothing
    Set m_oFWAppData = Nothing
    Set m_ofwRefData = Nothing
End Sub


Private Sub ObjectControl_Activate()
    Initialise
End Sub

Private Sub ObjectControl_Deactivate()
    Terminate
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = True
End Function










