VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "VWPopularSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    VWPopularSearch.cls
' System:       PROCAT, PRO
' Copyright:    (C) Quidnunc Limited
'
' Description:  The Popular Search View Class
'
' Amendment history:
'
'   d1      crm     09Dec1999      Created.
'   d2      crm     13Dec1999      Modified some Screens..added confirmation page for delete
'   d2      crm     13Dec1999      Description is a required field
'   d4      wdp     15Dec1999       Added call to index in verity
'   d5      crm     27Dec1999      Changed tool tips to be fetched from RefData
'   d6      aga     12Jan2000      Fixing defects
'   d7      wdp     01Feb2000      Dont pass serialised search as http param, use
'                                  saved_search id instead.
'   d8      crm     04Feb2000      Made  IsReaderSearch field editable
'   d9      crm     07Feb2000      Put limit on description length
'   d10     tmj     17Mar2000      Made it so cannot edit the name of a Popular Search
'                                  as it is the primary key.
'   d11     tmj     21Mar2000      Made desciption a mandatory field
'   d12     wdp     29Mar2000      Added verity deletion on delete
'   d13     tmj     29Mar2000      Changed so data in session is scoped correctly and added
'                                  Difference between update and delete when changing search
Implements ObjectControl

'To hold ASP object references
Private m_oResponse As Response
Private m_oRequest As Request
Private m_oSession As Session
Private m_oObjectContext  As ObjectContext
Private m_oboPermission As BOEditorialPermissions
Private m_lLoggedUserID As Long
Private m_ofwRefData As FWReferenceData

'QUeryString constants
Private Const k_sSearchString As String = "SearchString"
Private Const k_sSearchName As String = "SearchName"
Private Const k_sSearchDescription As String = "SearchDescription"
Private Const k_sChangeSearchString As String = "ChangeSearchString"
Private Const k_sDeleteConfirm As String = "DeleteConfirm"
Private Const k_sUpdate As String = "Update"



Public Sub DisplaySavedSearches()
'Description        :       Displays a list of Saved Searches  for a user
'
'
'Amendment History
'   d1      crm     09Dec1999      Created
'   d7      wdp     01Feb2000      Dont pass serialised search as http param, use
'                                  saved_search id instead.
'   d10     crm     11Feb2000      Saved searches now point to leafletdetail page instead
'                                  of editLeafletdetail
'   d12     tmj     29Mar2000      Added additional parameter when changing the source
On Error GoTo ErrorHandler
    Dim sSource As String
    Dim oboPopSearch As New BOPopSearch
    Dim sHTML As String
    Dim ofwHTML As New FWHTML
    Dim ofwTable As New FWTable
    Dim ofwRecList As FWRecordList
    Dim bChangeSearchString As Boolean
    Dim sTempSearchString As String
    Dim lSearchID As Long
    
    Dim typTableStyle  As TableStyle
    Dim typRowStyle As RowStyle
    Dim typCellStyle As CellStyle
            
    sSource = "VWPopularSearch.DisplaySavedSearches()"
    
    bChangeSearchString = False
    
    'Check if the user has right to perform the operation
    CheckForRightOfUser
            
    If m_oSession(k_sChangeSearchString) = "TRUE" Then
        bChangeSearchString = True
    End If
    
    If bChangeSearchString Then
    sHTML = ofwHTML.FontWithStyle("Change the Saved search for popular search ", ofwHTML.GetFontStyle(FONTPageHeadingStyle))
    Else
    sHTML = ofwHTML.FontWithStyle("Create a Popular Search", ofwHTML.GetFontStyle(FONTPageHeadingStyle))
    End If
    
    sHTML = sHTML & "<P>" & ofwHTML.FontWithStyle("The following are the Searches saved by you. Click on the one you want to use.", ofwHTML.GetFontStyle(FONTGeneralContentStyle)) & "<P>"
                   
    'set table style properties and assign it to table object
    With typTableStyle
        .eALIGN = HALCenter
        .eStyle = TABCustomStyle
        .sWidth = "100%"
        .lCELLSPACING = 2
        .lCELLPADDING = 1
    End With
    ofwTable.TableStyle = typTableStyle
        
    'set row style and start a new row
    typRowStyle = ofwHTML.GetRowStyle(RowBackgroundPink)
    
        
    'set cell style
    typCellStyle = ofwHTML.GetCellStyle(CELLColumnDataStyle)
    typCellStyle.sWidth = "100%"
    
    'fetch the saved searches for this user
    Set ofwRecList = oboPopSearch.SavedSearches(m_lLoggedUserID)
    
   Do Until ofwRecList.IsAtEnd
    
        'let each file be in its own row
        ofwTable.StartRow typRowStyle
        'if this page is for choosing a different search string from
        'the current one
        If IsNull(ofwRecList.GetFieldValue( _
                                    gk_sDBCOLSavedSearchString)) Then
                sTempSearchString = ""
        Else
                sTempSearchString = ofwRecList.GetFieldValue( _
                                    gk_sDBCOLSavedSearchString)
        End If
            
        lSearchID = ofwRecList.GetFieldValue(gk_sDBCOLSavedSearchID)
            
        If bChangeSearchString Then
            'each file in a single row,so only one cell
            ofwTable.AddCellAsHTML ofwHTML.Hyperlink(gk_sDisplayPopularSearchDetail & _
                                   "?" & gk_sFORMPSSearchResultsName & "=" & _
                                   URLEncode(m_oRequest.QueryString( _
                                   gk_sFORMPSSearchResultsName)) & "&" & _
                                   gk_sFORMSavedSearchID & "=" & lSearchID & "&" & _
                                   k_sUpdate & "=1", _
                                   ofwRecList.GetFieldValue(gk_sDBCOLSavedSearchDescription), _
                                   gk_sMainFrame), typCellStyle
        Else
            'each file in a single row,so only one cell
            ofwTable.AddCellAsHTML ofwHTML.Hyperlink(gk_sDisplayCreatePopularSearch & "?" & _
                                    gk_sFORMSavedSearchID & "=" & lSearchID, _
                                    ofwRecList.GetFieldValue(gk_sDBCOLSavedSearchDescription), _
                                    gk_sMainFrame), typCellStyle
        End If
        
        ofwTable.EndRow
        ofwRecList.MoveNext
    Loop
    
    ofwTable.AddBlankRow
    
    'typCellStyle.eBGCOLOR = COLOURWhite
    typCellStyle = ofwHTML.GetCellStyle(CELLNormalStyle)
    
    'cancel button
    ofwTable.StartRow
    If bChangeSearchString Then
        ofwTable.AddCellAsHTML ofwHTML.Hyperlink(gk_sDisplayPopularSearchDetail & "?" & gk_sFORMPSSearchResultsName & _
                                "=" & URLEncode(m_oRequest.QueryString( _
                                gk_sFORMPSSearchResultsName)), ofwHTML.Image("", gk_sImagesPath & _
                                gk_sImageCancel, m_ofwRefData.Search(gToolTip, ETOOLTIP.ttButtonCancelHelp)), gk_sMainFrame), typCellStyle
    Else
        ofwTable.AddCellAsHTML ofwHTML.Hyperlink(gk_sDisplayPopularSearches & "?" & gk_sFORMPSListUsage & _
                             "=" & EPSBrowserUsage.EPSMaintenance, ofwHTML.Image("", gk_sImagesPath & _
                                gk_sImageCancel, m_ofwRefData.Search(gToolTip, ETOOLTIP.ttButtonCancelHelp)), gk_sMainFrame), typCellStyle
    End If
    ofwTable.EndRow
    
    'get the table HTML
    sHTML = sHTML & ofwTable.GetHTML
    
    sHTML = sHTML & HTMLToAddHelpButton(gk_sHelpLinkToDisplaySavedSearches)
    
    'send the HTML out
    m_oResponse.Write sHTML
    
    'destroy object references
    Set ofwHTML = Nothing
    Set ofwTable = Nothing
    Set ofwRecList = Nothing
    Set oboPopSearch = Nothing
    
    'release object context
    m_oObjectContext.SetComplete
    
Exit Sub
ErrorHandler:
    m_oSession(k_sChangeSearchString) = ""
    m_oObjectContext.SetAbort
    ShowError sSource
    Set ofwHTML = Nothing
    Set ofwTable = Nothing
    Set ofwRecList = Nothing
    Set oboPopSearch = Nothing
End Sub


Public Sub DisplayCreatePopularSearch()
'Description        :       SCreen to create a new Popular search and edit an old one.
'                           user comes here on choosing a Saved search if its a new Popular
'                           search or on editing an old popular search
'
'Amendment History
'   d1      crm     09Dec1999       Created
'   d7      wdp     01Feb2000       Get search string from DB rather than
'                                   http parameter
'   d8      crm     04Feb2000       Made IsReaderSearch field editable
'   d9      crm     07Feb2000       Put limit on description length
'   d11     tmj     21Mar2000       Made description mandatory
On Error GoTo ErrorHandler
    Dim sSource As String
    Dim sLeafletName As String
    Dim sHTML As String
    Dim ofwHTML As FWHTML
    Dim ofwTable As FWTable
    Dim oboPopSearch As New BOPopSearchLoader
           
    Dim typTableStyle  As TableStyle
    Dim typRowStyle As RowStyle
    Dim typCellStyleL As CellStyle
    Dim typCellStyleR As CellStyle
   
    Dim sSearchString As String
    Dim sPopSearchString As String
    Dim bNewPopSearch As Boolean
    Dim lSavedSearchID As Long
    
    Dim sParam As String
    
    sSource = "VWPopularSearch.DisplayCreatePopularSearch()"
    
    'Check if the user has right to perform the operation
    CheckForRightOfUser
    
    'create instances of the helper objects
    Set ofwHTML = New FWHTML
    Set ofwTable = New FWTable
              
    'retrieve the name of the leaflet choosen,if there is one
    lSavedSearchID = CLng(m_oRequest.QueryString(gk_sFORMSavedSearchID))
    sPopSearchString = m_oRequest.QueryString(gk_sFORMPSSearchResultsName)
    
    'init to false
    If m_oSession(k_sChangeSearchString) = "TRUE" Or Len(Trim(sPopSearchString)) > 0 Then
    bNewPopSearch = False
    sSearchString = m_oRequest.QueryString(k_sSearchString)
    Else
    bNewPopSearch = True
    sSearchString = oboPopSearch.GetSavedSearchString(lSavedSearchID)
    
    End If
        
    If bNewPopSearch Then
        sHTML = ofwHTML.FontWithStyle("Create a new Popular search", ofwHTML.GetFontStyle(FONTPageHeadingStyle)) & "<P>"
    Else
        'the user has come here for an edit operation
            sHTML = ofwHTML.FontWithStyle("Edit popular search", ofwHTML.GetFontStyle(FONTPageHeadingStyle)) & "<P>"
            oboPopSearch.LoadData sPopSearchString
            If m_oSession(k_sChangeSearchString) <> "TRUE" Then
                sSearchString = oboPopSearch.SearchString
            End If
    End If
    
    sHTML = sHTML & ofwHTML.FormStart(gk_sSavePopularSearch, _
                                      "", "PopularSearchform", , True)
    
    sHTML = sHTML & ofwHTML.HiddenField(k_sUpdate, IIf(bNewPopSearch, "0", "1"))
        
    'set up the table properties
    With typTableStyle
        .eALIGN = HALNone
        .eBGCOLOR = COLOURWhite
        .lCELLSPACING = 2
        .lCELLPADDING = 1
        .sWidth = "100%"
        .lCELLSPACING = 1
    End With
    ofwTable.TableStyle = typTableStyle
    
    'now the row properties
    typRowStyle = ofwHTML.GetRowStyle(ROWCustomStyle)
    
    'set cell style properties for the two columns 'Left & Right'
    typCellStyleL = ofwHTML.GetCellStyle(CELLReadWriteLabelStyle)
    With typCellStyleL
          .sWidth = "30%"
    End With
    typCellStyleR = ofwHTML.GetCellStyle(CELLReadWriteDataStyle)
    With typCellStyleR
        .sWidth = "70%"
        .lCOLSPAN = 2
    End With
    
    'start a row
    ofwTable.StartRow typRowStyle
    'add CellLeft
    ofwTable.AddCellAsHTML "Name", typCellStyleL
    'add cellright
    'd10 Added If to prevent changing the name when editing
    If bNewPopSearch Then
        ofwTable.AddCellAsHTML ofwHTML.InputField(ITText, _
                               gk_sDBCOLPSSearchName, 50, , oboPopSearch.Name, , _
                               True), typCellStyleR
    Else
        ofwTable.AddCellAsHTML ofwHTML.EncodeAsHTML(oboPopSearch.Name) & _
                               ofwHTML.HiddenField(gk_sDBCOLPSSearchName, _
                                                   oboPopSearch.Name) _
                               , typCellStyleR
    End If
    
    'end row
    ofwTable.EndRow
    
    'start row
    ofwTable.StartRow typRowStyle
    'cell Left
    ofwTable.AddCellAsHTML "Description", typCellStyleL
    'cell Right
    '   d9      crm     07Feb2000
    '   d11     tmj     21Mar2000   Made Description Mandatory
    ofwTable.AddCellAsHTML ofwHTML.TextArea(gk_sDBCOLPSDescription, oboPopSearch.Description, 50, 3, 50, True), typCellStyleR
    'end row
    ofwTable.EndRow

    'start row
    ofwTable.StartRow typRowStyle
    'cell left
    ofwTable.AddCellAsHTML "Is Reader Search", typCellStyleL
    'cell right
    ofwTable.AddCellAsHTML ofwHTML.Checkbox("", gk_sDBCOLPSIsReaderSearch, oboPopSearch.IsReaderSearch) _
                           , typCellStyleR
    'end row
    ofwTable.EndRow
    
    ofwTable.AddBlankRow
    
    'change the cell color
    typCellStyleR = ofwHTML.GetCellStyle(CELLNormalStyle)
    typCellStyleR.sWidth = "70%"
    
    ofwTable.StartRow
    
    ofwTable.AddCellAsHTML "", typCellStyleL
    If bNewPopSearch Then
    ofwTable.AddCellAsHTML ofwHTML.SubmitImage(gk_sImagesPath & _
                            gk_sImageOK, "OK", "OK") & _
                            ofwHTML.ImageButton(gk_sImagesPath & gk_sImageCancel, _
                            "Cancel", gk_sDisplaySavedSearches), typCellStyleR
    Else
    ofwTable.AddCellAsHTML ofwHTML.SubmitImage(gk_sImagesPath & _
                            gk_sImageOK, "OK", "OK") & _
                            ofwHTML.ImageButton(gk_sImagesPath & gk_sImageCancel, _
                            "Cancel", gk_sDisplayPopularSearchDetail & "?" & _
                            gk_sFORMPSSearchResultsName & "=" & EncodeQueryStringValue(oboPopSearch.Name)), typCellStyleR
    End If
    
    ofwTable.EndRow
    
    sHTML = sHTML & ofwHTML.HiddenField(gk_sDBCOLPSSearchString, sSearchString)
    
    sHTML = sHTML & ofwTable.GetHTML
    
    sHTML = sHTML & "</FORM>"
    
    sHTML = sHTML & HTMLToAddHelpButton(gk_sHelpLinkToDisplayCreatePopularSearch)
    
    m_oResponse.Write sHTML
    
    'destroy object references
    Set ofwHTML = Nothing
    Set ofwTable = Nothing
    Set oboPopSearch = Nothing
    
    m_oObjectContext.SetComplete
    
Exit Sub
ErrorHandler:
    m_oSession(k_sChangeSearchString) = ""
    ShowError sSource
    m_oObjectContext.SetAbort
    Set ofwHTML = Nothing
    Set ofwTable = Nothing
    Set oboPopSearch = Nothing
End Sub

Public Sub SavePopularSearch()
'Description        :       Screen to save a Popular search
'                           user comes here on choosing OK button from the
'                           DisplayCreatePopularSearch page
'
'Amendment History
'   d1      crm     09Dec1999       Created
'   d6      aga     12Jan2000       Changed the redirection to details page instead of browser
'   d10     tmj     17Mar2000       Added optional parameter to SavePopularSearch
'

On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim oboPopSearch As New BOPopSearch
    
    sSource = "VWPopularSearch.SavePopularSearch()"
    
    'Check if the user has right to perform the operation
    CheckForRightOfUser
    'load data from form and save
    With oboPopSearch
        .Name = m_oRequest.Form(gk_sDBCOLPSSearchName)
        .Description = m_oRequest.Form(gk_sDBCOLPSDescription)
        .IsReaderSearch = IIf(Len(m_oRequest.Form(gk_sDBCOLPSIsReaderSearch)), True, False)
        .SearchString = m_oRequest.Form(gk_sDBCOLPSSearchString)
        .Save CBool(m_oRequest(k_sUpdate))
        .IndexInVerity
    End With
    'reset session variable
    m_oSession(k_sChangeSearchString) = ""
    'd6
    m_oResponse.Redirect gk_sDisplayPopularSearchDetail & "?" & gk_sFORMPSSearchResultsName & _
                         "=" & EncodeQueryStringValue(oboPopSearch.Name)
    'release object references
    Set oboPopSearch = Nothing
    
    m_oObjectContext.SetComplete
Exit Sub
ErrorHandler:
    m_oSession(k_sChangeSearchString) = ""
    ShowError sSource
    m_oObjectContext.SetAbort
    Set oboPopSearch = Nothing
End Sub

Public Sub DisplayPopularSearchDetail()
'Description        :       SCreen to create a new Popular search and edit an old one.
'                           user comes here on choosing a Saved search if its a new Popular
'                           search or on editing an old popular search
'
'Amendment History
'   d1      crm     09Dec1999       Created
'   d2      crm     13Dec1999       Aligned buttons
'   d6      aga     12Jan2000       Formatting buttons alignment
'   d10     crm     11Feb2000       Now can save a savedsearch when user changes them for a
'                                   popular search
'   d13     tmj     29Mar2000       Set the Changing to null at bottom of page and added choice
'                                   when saving between updating and adding
On Error GoTo ErrorHandler
    Dim sSource As String
    Dim sHTML As String
    Dim ofwHTML As FWHTML
    Dim ofwTable As FWTable
    Dim oboPopSearch As New BOPopSearchLoader
    Dim oboPopularSearch As New BOPopSearch
    Dim typTableStyle  As TableStyle
    Dim typRowStyle As RowStyle
    Dim typCellStyleL As CellStyle
    Dim typCellStyleR As CellStyle
    Dim sPopSearchName As String
    Dim lSavedSearchID As Long
    
    
    sSource = "VWPopularSearch.DisplayPopularSearchDetail()"
    
    'Check if the user has right to perform the operation
    CheckForRightOfUser
    
    'create instances of the helper objects
    Set ofwHTML = New FWHTML
    Set ofwTable = New FWTable
          
    ' arg name sent by the
    'Pop search browser
    sPopSearchName = m_oRequest.QueryString(gk_sFORMPSSearchResultsName)
    lSavedSearchID = m_oRequest.QueryString(gk_sFORMSavedSearchID)
    
    'user has come to change a saved search
    If lSavedSearchID > 0 Then
         With oboPopularSearch
            .Load sPopSearchName
            .SearchString = oboPopSearch.GetSavedSearchString(lSavedSearchID)
            If m_oRequest(k_sUpdate) = 1 Then
                .Save True
            Else
                .Save
            End If
            .IndexInVerity
         End With
    End If
    
    If Len(Trim(sPopSearchName)) > 0 Then
        sHTML = ofwHTML.FontWithStyle("Popular search detail", ofwHTML.GetFontStyle(FONTPageHeadingStyle)) & "<P>"
        oboPopSearch.LoadData sPopSearchName
    End If
        
    sHTML = sHTML & ofwHTML.FormStart(gk_sSavePopularSearch, "", "PopularSearchform", , True)
    
    'set up the table properties
    With typTableStyle
        .eALIGN = HALNone
        .eBGCOLOR = COLOURWhite
        .lCELLSPACING = 2
        .lCELLPADDING = 1
        .sWidth = "100%"
        .lCELLSPACING = 1
    End With
    ofwTable.TableStyle = typTableStyle
    
    'now the row properties
    typRowStyle = ofwHTML.GetRowStyle(ROWCustomStyle)
    
    'set cell style properties for the two columns 'Left & Right'
    typCellStyleL = ofwHTML.GetCellStyle(CELLReadOnlyLabelStyle)
    With typCellStyleL
        .sWidth = "30%"
    End With
    typCellStyleR = ofwHTML.GetCellStyle(CELLReadOnlyDataStyle)
    With typCellStyleR
        .sWidth = "70%"
        .lCOLSPAN = 4
    End With
    
    'start a row
    ofwTable.StartRow typRowStyle
    'add CellLeft
    ofwTable.AddCellAsHTML "Name", typCellStyleL
    'add cellright
    ofwTable.AddCellAsHTML oboPopSearch.Name, typCellStyleR
    'end row
    ofwTable.EndRow
    
    'start row
    ofwTable.StartRow typRowStyle
    'cell Left
    ofwTable.AddCellAsHTML "Description", typCellStyleL
    'cell Right
    ofwTable.AddCellAsHTML oboPopSearch.Description, typCellStyleR
    'end row
    ofwTable.EndRow
    
    'start row
    ofwTable.StartRow typRowStyle
    'cell left
    ofwTable.AddCellAsHTML "", typCellStyleL
    'cell right
    ofwTable.AddCellAsHTML ofwHTML.Hyperlink(gk_sChangeSearchString & _
                            "?" & gk_sFORMPSSearchResultsName & "=" & _
                            URLEncode(oboPopSearch.Name), _
                            "change saved search...", _
                            gk_sMainFrame), typCellStyleR
    'end row
    ofwTable.EndRow
    
    'start row
    ofwTable.StartRow typRowStyle
    'cell left
    ofwTable.AddCellAsHTML "Is Reader Search", typCellStyleL
    'cell right
    ofwTable.AddCellAsHTML IIf(oboPopSearch.IsReaderSearch, "Yes", "No"), typCellStyleR
    'end row
    ofwTable.EndRow
    
    ofwTable.AddBlankRow
    
    'change the cell color
    typCellStyleR = ofwHTML.GetCellStyle(CELLNormalStyle)
    typCellStyleR.sWidth = "10%"
    typCellStyleR.lCOLSPAN = 1
      
    ofwTable.StartRow
     
    ofwTable.AddCellAsHTML "", typCellStyleL
    
    'd6
    'cells containing three buttons...Edit, Remove & Done
    
    ofwTable.AddCellAsHTML ofwHTML.Hyperlink(gk_sDisplayCreatePopularSearch & "?" _
                            & gk_sFORMPSSearchResultsName & "=" & _
                            URLEncode(sPopSearchName), ofwHTML.Image("", gk_sImagesPath & _
                            gk_sImageEdit, m_ofwRefData.Search(gToolTip, ETOOLTIP.ttEditPopularSearch)), gk_sMainFrame) _
                            , typCellStyleR
                            
    ofwTable.AddCellAsHTML ofwHTML.Hyperlink(gk_sDeletePopularSearch & "?" & _
                            gk_sFORMPSSearchResultsName & "=" & _
                            URLEncode(sPopSearchName), ofwHTML.Image("", gk_sImagesPath & _
                            gk_sImageRemove, m_ofwRefData.Search(gToolTip, ETOOLTIP.ttDeletePopularSearch)), gk_sMainFrame) _
                            , typCellStyleR

    ofwTable.AddCellAsHTML ofwHTML.Hyperlink(gk_sDisplayPopularSearches & "?" & gk_sFORMPSListUsage & _
                            "=" & EPSBrowserUsage.EPSMaintenance, ofwHTML.Image("", gk_sImagesPath & _
                            gk_sImageDone, m_ofwRefData.Search(gToolTip, ETOOLTIP.ttBackToPopSearchBrowser)), gk_sMainFrame) _
                            , typCellStyleR
                            
    typCellStyleR.sWidth = "40%"
    ofwTable.AddCellAsHTML gk_sHTMLNonBreakSpace, typCellStyleR
    ofwTable.EndRow
    
    sHTML = sHTML & ofwTable.GetHTML
    
    sHTML = sHTML & ofwHTML.HiddenField(gk_sDBCOLPSSearchString, oboPopSearch.SearchString)
    
    sHTML = sHTML & "</FORM>"
    
    sHTML = sHTML & HTMLToAddHelpButton(gk_sHelpLinkToDisplayPopularSearchDetail)
    
    m_oResponse.Write sHTML
    
    'destroy object references
    Set ofwHTML = Nothing
    Set ofwTable = Nothing
    Set oboPopSearch = Nothing
    Set oboPopularSearch = Nothing
    
    m_oObjectContext.SetComplete
    
    m_oSession(k_sChangeSearchString) = ""
    
Exit Sub
ErrorHandler:
    m_oSession(k_sChangeSearchString) = ""
    ShowError sSource
    m_oObjectContext.SetAbort
    Set ofwHTML = Nothing
    Set ofwTable = Nothing
    Set oboPopSearch = Nothing
    Set oboPopularSearch = Nothing
End Sub

Public Sub DeletePopularSearch()
'Description        :       Screen to delete a Popular search
'
'Amendment History
'   d1      crm     10Dec1999       Created
'   d2      crm     13Dec1999       Added delete confirmation page
'   d12     wdp     29Mar2000      Added verity deletion on delete
On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim oboPopSearch As New BOPopSearch
    Dim sPopSearchName As String
    Dim bDeleteConfirm As Boolean
    Dim sMessage As String
    Dim ofwHTML As New FWHTML
    
    sSource = "VWPopularSearch.SavePopularSearch()"
    
    'check is user has permission
    CheckForRightOfUser

    sPopSearchName = m_oRequest.QueryString(gk_sFORMPSSearchResultsName)
    'check if confirm status is part of QueryString
    If m_oRequest.QueryString(k_sDeleteConfirm) = "" Then
        'if not ask user if to proceed with delete
        oboPopSearch.Load sPopSearchName
        sMessage = ofwHTML.Bold("Do you want to delete the popular search " & ofwHTML.Emphasize(oboPopSearch.Description) & " ? ")
        sMessage = sMessage & "<P>"
        sMessage = sMessage & ofwHTML.Hyperlink(gk_sDeletePopularSearch & "?" & _
                            gk_sFORMPSSearchResultsName & "=" & _
                            URLEncode(sPopSearchName) & "&" & k_sDeleteConfirm & "=" & _
                            "Yes", ofwHTML.Image("", gk_sImagesPath & _
                            gk_sImageOK, m_ofwRefData.Search(gToolTip, ETOOLTIP.ttDeletePopularSearch)), gk_sMainFrame)
        sMessage = sMessage & ofwHTML.Hyperlink(gk_sDisplayPopularSearchDetail & "?" & _
                            gk_sFORMPSSearchResultsName & "=" & URLEncode(sPopSearchName), ofwHTML.Image("", gk_sImagesPath & _
                            gk_sImageCancel), gk_sMainFrame, , m_ofwRefData.Search(gToolTip, ETOOLTIP.ttCancelDelete))
        m_oResponse.Write sMessage
    Else
        'if confirm of delete is received then delete
        With oboPopSearch
            .Name = sPopSearchName
            .Delete sPopSearchName
            .IndexInVerity IDXDelete
        End With
        
         m_oResponse.Redirect gk_sDisplayPopularSearches & "?" & gk_sFORMPSListUsage & _
                         "=" & EPSBrowserUsage.EPSMaintenance
    End If
    
    'release object references
    Set oboPopSearch = Nothing
    Set ofwHTML = Nothing
    
    m_oObjectContext.SetComplete
    
Exit Sub
ErrorHandler:
    m_oSession(k_sChangeSearchString) = ""
    ShowError sSource
    m_oObjectContext.SetAbort
    Set oboPopSearch = Nothing
    Set ofwHTML = Nothing
End Sub

Public Sub ChangeSearchString()
'Description        :       Intermediate screen to change a saved search for a Popular search
'
'Amendment History
'   d1      crm     10Dec1999       Created
'
On Error GoTo ErrorHandler
    
    Dim sSource As String
       
    sSource = "VWPopularSearch.ChangeSearchString()"
    
    'check user permission
    CheckForRightOfUser
    
    'set the session variable to notify that we are changing a saved search...
    m_oSession(k_sChangeSearchString) = "TRUE"
    
    'redirect to saved searches page....
    m_oResponse.Redirect gk_sDisplaySavedSearches & "?" & gk_sFORMPSSearchResultsName & "=" & _
                         URLEncode(m_oRequest.QueryString(gk_sFORMPSSearchResultsName))
                             
    m_oObjectContext.SetComplete
    
Exit Sub
ErrorHandler:
     m_oSession(k_sChangeSearchString) = ""
    ShowError sSource
    m_oObjectContext.SetAbort
End Sub




Private Sub Initialise()
'Description    :   Method stores the ObjectContext for the class
'                   and the references to the ASP Response and Request objects
On Error GoTo ErrorHandler

    Dim sSource As String
    sSource = "VWPopularSearch.Initialise()"
    
    
    Set m_oObjectContext = GetObjectContext
    If m_oObjectContext Is Nothing Then
        RaiseError sSource, peOCUnableToGetObjectContext, "Unable to get ObjectContext"
    Else
        Set m_oResponse = m_oObjectContext.Item(gk_sResponseObject)
        Set m_oRequest = m_oObjectContext.Item(gk_sRequestObject)
        Set m_oSession = m_oObjectContext.Item(gk_sSessionObject)
        Set m_ofwRefData = New FWReferenceData
'        m_oSession(k_sChangeSearchString) = ""
    End If

    Set m_oboPermission = New BOEditorialPermissions
    
    'Get the users details
    m_lLoggedUserID = GetDataFromSession(dsUserID, m_oObjectContext.Item(gk_sSessionObject))
    If m_lLoggedUserID = 0 Then     'Not a valid user
        RaiseError sSource, EPROCATINTERNALERROR.peDBInvalidUser, "Invalid User ID : " & m_lLoggedUserID
    End If
    
Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub



Private Sub CheckForRightOfUser()
On Error GoTo ErrorHandler
'Checks if the current user has right to preform the operation

'   d1      crm     10Dec1999      Created

    Dim ePermissionCheckResult As ESecurityCheckResult
    Dim sSource As String
    
    sSource = "VWPopularSearch.CheckForRightOfUser"
    
    m_oboPermission.UserID = m_lLoggedUserID
    ePermissionCheckResult = m_oboPermission.Check(ESecurityCheck.EUserCanMaintainPopularSearches)
    If ePermissionCheckResult <> _
            ESecurityCheckResult.ESRSecurityCheckSuccessful Then
    
        RaiseError sSource, _
            GetErrorEnumForSecurityCheckResult(ePermissionCheckResult), _
            GetDescriptionForSecurityCheckError(GetErrorEnumForSecurityCheckResult(ePermissionCheckResult))
    End If
    
    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub



Private Sub Terminate()
'Description    :       Terminate object references
    Set m_oResponse = Nothing
    Set m_oRequest = Nothing
    Set m_oObjectContext = Nothing
    Set m_oSession = Nothing
    Set m_oboPermission = Nothing
    Set m_ofwRefData = Nothing
End Sub

Private Sub ObjectControl_Activate()
     Initialise
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = True
End Function

Private Sub ObjectControl_Deactivate()
    Terminate
End Sub



