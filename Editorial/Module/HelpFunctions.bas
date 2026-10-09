Attribute VB_Name = "HelpFunctions"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    HelpFunctions.bas
' System:       PRO, PROCAT
' Copyright:    (C) Quidnunc Limited
'
' Description:  Holds the function to display the help button and the list of links to
'               be displayed for every page
'
' Amendment history:
'   d1  ksk    01 Oct 1999 Created
'   d1  ksk    01 Oct 1999 Added code.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Const m_sHelpWindowFunction As String = "javascript:openSeperateWindow"
Private Const m_ksHelp As String = "HLP"


Public Function HTMLToAddHelpButton(t_sHTMLLinkFileName As String) As String
'Function that returns a string containing html to display a help button
'   which opens the passed URL in a seperate help window
'
On Error GoTo ErrHandler
    Dim ofwHelpTable As FWTable
    Dim ofwHTMLRef As FWHTML
    Dim ofwRefData As FWReferenceData
    Dim typHelpCell As CellStyle
    Dim typHelpTable As TableStyle
    Dim sSource As String, sButtonLink As String
    
    Set ofwHelpTable = New FWTable
    Set ofwHTMLRef = New FWHTML
    Set ofwRefData = New FWReferenceData
    
    sSource = "HelpLinksAdnFunctions.HTMLToAddHelpButton"

    typHelpTable = ofwHTMLRef.GetTableStyle(ETableStyle.TABPlainStyle)
    typHelpTable.sWidth = "100%"
    ofwHelpTable.TableStyle = typHelpTable
    
    
    ofwHelpTable.StartRow
    
    sButtonLink = ofwHTMLRef.Hyperlink(OpenHelpWindow(Quotes(t_sHTMLLinkFileName), Quotes(m_ksHelp)), _
                        ofwHTMLRef.Image(m_ksHelp, gk_sImagesPath & _
                                gk_sImageHelp, ofwRefData.Search _
                                (gToolTip, ETOOLTIP.ttButtonHelpHelp)), _
                                 gk_sMainFrame, ofwRefData.Search _
                                (gToolTip, ETOOLTIP.ttButtonHelpHelp))
    
    typHelpCell = ofwHTMLRef.GetCellStyle(ECellStyle.CELLPlainStyle)
    typHelpCell.sWidth = "100%"
    typHelpCell.eALIGN = HALRight
    
    ofwHelpTable.AddCellAsHTML sButtonLink, typHelpCell
    ofwHelpTable.EndRow
    
    HTMLToAddHelpButton = ofwHelpTable.GetHTML & vbNewLine

    HTMLToAddHelpButton = HTMLToAddHelpButton & _
                        ScriptToOpenFormatedWebBrowser(400, 400) _
                        & vbNewLine
    
    Exit Function
    
ErrHandler:
    ConvertError sSource
End Function


Private Function OpenHelpWindow(t_sURL As String, _
                        t_sWindowName As String)
    
    OpenHelpWindow = m_sHelpWindowFunction & "(" & t_sURL & "," & _
                        t_sWindowName & ");"
                        
End Function

Public Function ScriptToOpenFormatedWebBrowser(Optional t_lHeightOfWindow As Long = 200, _
                        Optional t_lWidthOfWindow As Long = 200, _
                        Optional t_lTopWindowPosition As Long = 0, _
                        Optional t_lLeftWindowPosition As Long = 0, _
                        Optional t_bToolbarRequired As Boolean = False, _
                        Optional t_bURLEntryBoxRequired As Boolean = False, _
                        Optional t_bStatusBarRequired As Boolean = False, _
                        Optional t_bMenubarRequired As Boolean = False, _
                        Optional t_bScrollBarRequired As Boolean = True, _
                        Optional t_bResizableRequired As Boolean = True) As String
'A function that returns a text for a javascript to open a new browser window in a given format
'   The controls (like toolbar, status bar) on the browser and placement of the browser etc is
'   controlled through the optional parameters
'

    ScriptToOpenFormatedWebBrowser = vbNewLine & vbNewLine & "<script language='JavaScript'>" & _
                                             vbNewLine & " <!-- begin script  " & vbNewLine
    
    ScriptToOpenFormatedWebBrowser = ScriptToOpenFormatedWebBrowser & " function " & _
                                        "openSeperateWindow(URL, WindowName){ " & vbNewLine
    
    ScriptToOpenFormatedWebBrowser = ScriptToOpenFormatedWebBrowser & _
                                        " window.open( URL, WindowName, '"
    
    ScriptToOpenFormatedWebBrowser = ScriptToOpenFormatedWebBrowser & _
                                        "height=" & CStr(t_lHeightOfWindow) & ","
    ScriptToOpenFormatedWebBrowser = ScriptToOpenFormatedWebBrowser & _
                                        "width=" & CStr(t_lWidthOfWindow) & ","
    ScriptToOpenFormatedWebBrowser = ScriptToOpenFormatedWebBrowser & _
                                        "top=" & CStr(t_lTopWindowPosition) & ","
    ScriptToOpenFormatedWebBrowser = ScriptToOpenFormatedWebBrowser & _
                                        "left=" & CStr(t_lLeftWindowPosition) & ","
    ScriptToOpenFormatedWebBrowser = ScriptToOpenFormatedWebBrowser & _
                            "toolbar=" & IIf(t_bToolbarRequired, "yes", "no") & ","
    ScriptToOpenFormatedWebBrowser = ScriptToOpenFormatedWebBrowser & _
                        "location=" & IIf(t_bURLEntryBoxRequired, "yes", "no") & ","
    ScriptToOpenFormatedWebBrowser = ScriptToOpenFormatedWebBrowser & "directories=no,"
    ScriptToOpenFormatedWebBrowser = ScriptToOpenFormatedWebBrowser & _
                        "status=" & IIf(t_bStatusBarRequired, "yes", "no") & ","
    ScriptToOpenFormatedWebBrowser = ScriptToOpenFormatedWebBrowser & "menubar=" & _
                                        IIf(t_bMenubarRequired, "yes", "no") & ","
    ScriptToOpenFormatedWebBrowser = ScriptToOpenFormatedWebBrowser & "scrollbars=" & _
                                        IIf(t_bScrollBarRequired, "yes", "no") & ","
    ScriptToOpenFormatedWebBrowser = ScriptToOpenFormatedWebBrowser & "resizable=" & _
                                        IIf(t_bResizableRequired, "yes", "no") & "'); }" _
                                        & vbNewLine
    
    ScriptToOpenFormatedWebBrowser = ScriptToOpenFormatedWebBrowser & _
                        vbNewLine & "// end hiding contents from old browsers -->" & _
                        vbNewLine & " </script>" & vbNewLine
    ScriptToOpenFormatedWebBrowser = ScriptToOpenFormatedWebBrowser & vbNewLine & vbNewLine
    
End Function


