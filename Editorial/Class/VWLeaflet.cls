VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "VWLeaflet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    VWLeaflet.cls
' System:       PROCAT, PRO
' Copyright:    (C) Quidnunc Limited
'
' Description:  The Leaflet View Class
'
' Amendment history:
'
'   d1      crm     06Aug1999      Created.
'   d2      crm     02Nov1999      (integrated by ndk) Modified DisplayAllLeaflets() to handle NULLs
'                                  Added CancelAddTermInformationLeaflet,
'                                  AddTermInformationLeaflet,SaveLeafletDetail
'                                  DisplayLeafletDetail,DisplayEditLeafletDetail
'   d3      ndk     19Nov1999      Removed all references to note text column
'                                  Set the framename in all hyperlinks
'   d4      crm     22Nov1999      CreatedOn changed to a display only field,
'                                  added link to cancel button,
'                                  Display of Leaflet detail includes Index Terns
'   d5      ksk     22Nov1999      Added security
'   d6      wdp     02Dec1999      Added verity indexing call
'   d7      crm     13Dec1999      Added cancel button for AllLefaletsFile Display
'                                  made minor alignments proper
'   d8      crm     16Dec1999      Added comments and some formatting
'   d9      crm     27Dec1999      Changed tool tips to be fetched from RefData
'   d10     crm     03Feb2000      Changed cancel button to image link in DisplayEditLeafletDetail
'                                  to make it stop doing validation
'   d11     crm     10Feb2000      Changed title textbox to text area in DisplayEditLeafletDetail()
'   d12     WDP     15Feb2000      Added call to BOLeaflet.DeleteSourceFile in SaveLeafletDetail
'   d13     TMJ     21Feb2000      Changed to use new BOLeaflet.AllLeafletFiles method and check
'                                  files
'   d14     WDP     23Feb2000      Added DateForDisplay calls for displaying created on dates
'   d15     TMJ     01Mar2000      Made cosmetic changes to DisplayEditLeafletDetails
'   d16     TMJ     02Mar2000      Modified so unset dates do not throw errors and extended d15
'   d17     DAJ     26Jun2000      Changed layout of AF term details in leaflet detail page to match
'                                   that of the catalouge layout (mostly)
'   d18     DAJ     10Jul2000      Added DeleteLeaflet and changed DisplayEditLeafletDetail to add a Delete button;
'                                   Changed layout of AF term details in leaflet EDIT details page to match d17
'   d19     MAP     12Oct2000      Changed Leaflet name field from Text Area to Text Field and made
'                                  made max size of field a global value
'   d20     DAJ     06Nov2000      Changed order of SaveLeafletDetail code so that it can understand if the user
'                                       pressed Return in the form instead of clicking OK
'   d21     Tomh    22Mar2002      Have altered displayleafletdetail as it
'                                  did not delete the source file of the
'                                  leaflet when the file was changed.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Implements ObjectControl

'To hold ObjectContext reference
Private m_oObjectContext As ObjectContext

'To hold ASP object references
Private m_oResponse As Response
Private m_oRequest As Request
Private m_oSession As Session
Private m_oboPermission As BOEditorialPermissions
Private m_lLoggedUserID As Long

Private Const ksLeafletName As String = "LName"
Private Const k_sParagraph As String = "<P>"

Private Const k_sEdit As String = "Edit"
Private Const k_sPerson As String = "Person"
Private Const k_sCorporateBody As String = "CorporateBody"
Private Const k_sSubject As String = "Subject"
Private Const k_sPlace As String = "Place"
Private Const k_sPickTerm As String = "PickTerm"
Private Const k_sLoadType As String = "LoadType"
Private Const k_TermType As String = "TermType"

Private Sub Initialise()
'Description    :   Method stores the ObjectContext for the class
'                   and the references to the ASP Response and Request objects
On Error GoTo ErrorHandler

    Dim sSource As String
    sSource = "VWLeaflet.Initialise()"
    
    
    Set m_oObjectContext = GetObjectContext
    If m_oObjectContext Is Nothing Then
        RaiseError sSource, peOCUnableToGetObjectContext, "Unable to get ObjectContext"
    Else
        Set m_oResponse = m_oObjectContext.Item(gk_sResponseObject)
        Set m_oRequest = m_oObjectContext.Item(gk_sRequestObject)
        Set m_oSession = m_oObjectContext.Item(gk_sSessionObject)
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

Private Sub Terminate()
'Description    :       Terminate object references

    Set m_oResponse = Nothing
    Set m_oRequest = Nothing
    Set m_oObjectContext = Nothing
    Set m_oSession = Nothing
    Set m_oboPermission = Nothing
End Sub


Private Sub ObjectControl_Activate()
'Initilase the object
    Initialise
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = True
End Function

Private Sub ObjectControl_Deactivate()
'Terminate the references held inside
    Terminate
End Sub

Public Sub DisplayAllLeaflets()
'Description    :   Displays a list of all Information-Leaflets,their status(released/not released),
'                   creation date and the title-text.Picks data from the SFA table.Displays a link
'                   to add a new leaflet.
'
'
'   d2      crm     02Nov1999      Modified DisplayAllLeaflets() to handle NULLs for CreatedDate
'   d5      ksk     22Nov1999      Added security
'   d8      crm     16Dec1999      Removed code which was commented and written prior
'                                  to using the LeafletBrowser of Reader.


On Error GoTo ErrorHandler
    Dim sSource As String
    Dim sHTML As String
    Dim sURL As String
    
    sSource = "VWLeafLet.DisplayAllLeaflets"
    
    'd5 Check if the user has right to perform the operation
    CheckForRightOfUser
    
    'we redirect this request to the InformationLeafletBrowser.asp page
    sURL = gk_sInformationLeafletBrowser & "?" & "eUsage" & "=" & ELeafletBrowserUsage.ELeafletMaintenance
    m_oResponse.Redirect sURL
    
    m_oObjectContext.SetComplete
Exit Sub
ErrorHandler:
    m_oObjectContext.SetAbort
    ShowError sSource
End Sub

Public Sub DisplayAllLeafletFiles()
'Description        :       Displays a list of leaflet-files which can be registered against
'                           a leaflet.
'Amendment History
'   d1      crm     10Aug1999      Created
'   d5      ksk     22Nov1999      Added security
'   d7      crm     13Dec1999      Added cancel button for AllLefaletFile Display
'   d8      crm     16Dec1999      Added a few comments and formatted code.
'   d13     TMJ     21Feb2000      Changed so uses new ObLeaflet.AllLeafletFiles method and
'                                  added javascript validation, to ensure files not
'                                  accidently overwritten
On Error GoTo ErrorHandler
    Dim sSource As String
    Dim oboLeaflet As BOLeaflet
    Dim sHTML As String
    Dim ofwHTML As FWHTML
    Dim ofwTable As FWTable
    Dim vArrLeafletNames As Variant
    Dim lIndex As Long
    Dim lLeafletID As Long
    'Dim sTempString As String
    Dim lLeafletTitle As String
    Dim sOnClick As String 'd13 String to construct onClick statement
    
    Dim typTableStyle  As TableStyle
    Dim typRowStyle As RowStyle
    Dim typCellStyle As CellStyle
    
    Dim ofwAppData As FWApplicationData
            
    sSource = "VWLeaflet.DisplayAllLeafletFiles()"
    
    'd5 Check if the user has right to perform the operation
    CheckForRightOfUser
    
    'create the helper objects
    Set ofwHTML = New FWHTML
    Set ofwTable = New FWTable
            
    'the ID of the Leaflet, this is to redirect the user after
    'he has choosen the Leaflet-file for a Leaflet
    lLeafletID = m_oRequest.QueryString(gk_sDBCOLSFAReferenceID)
    'the title of the Leaflet if we have a request to change the LeafletFile for
    'an existing Leaflet
    lLeafletTitle = m_oRequest.QueryString(gk_sDBCOLSFATitleText)
            
    'create the BOLeaflet object
    Set oboLeaflet = New BOLeaflet
    
    'check if the ID is > 0
    If lLeafletID > 0 Then
    'sHTML = ofwHTML.Bold(ofwHTML.Font("Change the leaflet file...", COLOURBlack, 4))
    sHTML = ofwHTML.FontWithStyle("Change the leaflet file...", ofwHTML.GetFontStyle(FONTPageHeadingStyle))
    Else
    'sHTML = ofwHTML.Bold(ofwHTML.Font("Locate the leaflet file", COLOURBlack, 4)) & k_sParagraph
    sHTML = ofwHTML.FontWithStyle("Locate the leaflet file", ofwHTML.GetFontStyle(FONTPageHeadingStyle))
    End If
    
    sHTML = sHTML & "<BR>" & ofwHTML.FontWithStyle("The following leaflet files are available. Click on the one you want to use.", ofwHTML.GetFontStyle(FONTGeneralContentStyle)) & _
                    k_sParagraph
                    
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
    typCellStyle = ofwHTML.GetCellStyle(CELLReadOnlyDataStyle)
    With typCellStyle
        .sWidth = "100%"
    End With
    
    'fetch the leaflet names from the BOLeaflet object,this is an variant array
    'so convert to string before sending as argument to any function
    'd13 added boolean variable
    vArrLeafletNames = oboLeaflet.AllLeafletFiles(True)
    
    If Not IsNull(vArrLeafletNames) Then
    For lIndex = LBound(vArrLeafletNames) To UBound(vArrLeafletNames) - 1
        'let each file be in its own row
        ofwTable.StartRow typRowStyle
        'if this page is for choosing a different leaflet-file from
        'the current one for a leaflet
        
        Set ofwAppData = New FWApplicationData
        
        'd13 First check to see if a file with the same name already on the system
        '    And set the the onClick value accordingly
        If oboLeaflet.AlreadyExists(CStr(vArrLeafletNames(lIndex))) Then
            sOnClick = "return confirm('" & ofwAppData.GetDisplayInformation( _
                        EEditorialResourceFile.erfConfirmOverwriteLeafletOne) & _
                       CStr(vArrLeafletNames(lIndex)) & _
                       ofwAppData.GetDisplayInformation( _
                        EEditorialResourceFile.erfConfirmOverwriteLeafletTwo) & "')"
        Else
            sOnClick = ""
        End If
        
        If lLeafletID > 0 Then
            'each file in a single row,so only one cell
            
            'send the Leaflet name and the Id of the Leaflet for which this file has been
            'choosen thru the QueryString
            ofwTable.AddCellAsHTML ofwHTML.Hyperlink(gk_sDisplayLeafletDetail & "?" & _
                                    ksLeafletName & "=" & CStr(vArrLeafletNames(lIndex)) & _
                                    "&" & gk_sDBCOLSFAReferenceID & "=" & lLeafletID, _
                                    CStr(vArrLeafletNames(lIndex)), gk_sMainFrame, _
                                    t_sActionOnClick:=sOnClick), typCellStyle
        Else
            'each file in a single row,so only one cell
            'send the Leaflet Nme alone if this is a request to create a new
            'Leaflet.This is how we identify if it is a new or a old
            'Leaflet in the BO
            ofwTable.AddCellAsHTML ofwHTML.Hyperlink(gk_sDisplayEditLeafletDetail & "?" & _
                                    ksLeafletName & "=" & CStr(vArrLeafletNames(lIndex)), _
                                    CStr(vArrLeafletNames(lIndex)), gk_sMainFrame, _
                                    t_sActionOnClick:=sOnClick), typCellStyle
        End If
        ofwTable.EndRow
    Next
        'get the table HTML
        sHTML = sHTML & ofwTable.GetHTML
    End If
        
    'Cancel link, we send the user back to the Info.LeafletBrowser if the
    'operation is cancelled. Change link here if you have to redirect elsewhere
    'for a Cancel operation
    If lLeafletID > 0 Then
        sHTML = sHTML & ofwHTML.Hyperlink(gk_sDisplayLeafletDetail & "?" & _
                                gk_sDBCOLSFAReferenceID & "=" & lLeafletID, ofwHTML.Image("", gk_sImagesPath & _
                                gk_sImageCancel), gk_sMainFrame)
    Else
        sHTML = sHTML & ofwHTML.Hyperlink(gk_sInformationLeafletBrowser & "?" & _
                                "eUsage" & "=" & 3, ofwHTML.Image("", gk_sImagesPath & _
                                gk_sImageCancel), gk_sMainFrame)
    End If
    
    'the help button
    sHTML = sHTML & HTMLToAddHelpButton(gk_sHelpLinkToDisplayAllLeafletFiles)
    
    'send the HTML out
    m_oResponse.Write sHTML
    
    'destroy object references
    Set ofwHTML = Nothing
    Set ofwTable = Nothing
    Set oboLeaflet = Nothing
    
    'release object context
    m_oObjectContext.SetComplete
    
Exit Sub
ErrorHandler:
    m_oObjectContext.SetAbort
    ShowError sSource
    Set ofwHTML = Nothing
    Set ofwTable = Nothing
    Set oboLeaflet = Nothing
End Sub

Public Sub DisplayEditLeafletDetail()
'Description        :       Method to display the screen where a leaflet file would be setup,
'                           user comes here on choosing a leaflet-file.The leaflet-file which
'                           has been choosen is got thru the Querystring of the Request object
'
'NOTE   :       The BOLeaflet class stores only the IDs of each Index Terms registered against
'               a Leaflet.To get the description we use the FWReference data class
'Amendment History
'   d4      crm     22Nov1999       Modified Cancel button link, see note below
'                                   Added Identifiers for each category of IndexTerms
'   d5      ksk     22Nov1999       Added security
'   d8      crm     16Dec1999       Added few comments, formatted code
'   d10     crm     03Feb2000       Changed cancel button to image link from submit button
'   d11     crm     10Feb2000       Changed title textbox to text area in DisplayEditLeafletDetail()
'   d14     WDP     23Feb2000       Added DateForDisplay calls for displaying created on dates
'   d15     TMJ     01Mar2000       Modified so displays the index term buttons with a pink
'                                   background
'   d16     TMJ     02Mar2000       Modifed so when remove button is displayed look of tables isn't
'                                   affected (extension of d15) and handled case where date not set
'   d18     DAJ     10Jul2000      Changed to add a Delete button, and changed layout of AF terms to
'                                   match the normal detail page in d17
'   d19     MAP     12Oct2000       Changed Leaflet name field from Text Area to Text Field and made
'                                   made max size of field a global value

On Error GoTo ErrorHandler
    Dim sSource As String
    Dim sLeafletName As String
    Dim sHTML As String
    Dim ofwHTML As FWHTML
    Dim ofwTable As FWTable
           
    Dim typTableStyle  As TableStyle
    Dim typRowStyle As RowStyle
    Dim typCellStyleL As CellStyle
    Dim typCellStyleR As CellStyle
    Dim typCellStyleFR As CellStyle
    Dim typCellStyleTR As CellStyle
    
    Dim oboLeaflet As New BOLeaflet
        
    Dim lLeafletID As Long
    Dim sLeafletFile As String
    Dim sPickTermImage As String
    Dim vTemp As Variant
    Dim oTempCol As Collection
    Dim sTempPersonName As String
    Dim vTempName As Variant
    Dim bPutTitle As Boolean
    Dim bNewLeaflet As Boolean
    Dim sDate As String
    
    'variable for the load type which is to be performed
    Dim eTempEAFTermLoadType As EAFTermLoadType
    
    Dim ofwRefdata As New FWReferenceData
    
    sSource = "VWLeaflet.DisplayEditLeafletDetail()"
    
    'd5 Check if the user has right to perform the operation
    CheckForRightOfUser
    
    'create instances of the helper objects
    Set ofwHTML = New FWHTML
    Set ofwTable = New FWTable
          
    'initialise to FALSE
    bNewLeaflet = False
    
    'retrieve the name of the leaflet choosen,if there is one
    sLeafletFile = m_oRequest.QueryString(ksLeafletName)
    If Len(Trim(sLeafletFile)) > 0 Then
        'if there is a name of a Lefalet file then this is the setup for
        'a new leaflet
        sHTML = ofwHTML.FontWithStyle("Setup a new leaflet", ofwHTML.GetFontStyle(FONTPageHeadingStyle)) & "<P>"
        oboLeaflet.LeafletFile = sLeafletFile
        ' set the state for the BO since we have
        'assigned the name.
        oboLeaflet.PROCATTransientInterface_SetState oboLeaflet.Name
        ' new leaflet
        bNewLeaflet = True
    Else
        'check if the user has come here for an edit operation
        lLeafletID = m_oRequest.QueryString(gk_sDBCOLSFAReferenceID)
        'ths first time the user comes here the LeafletID would come along
        'for an edit operation.When adding terms just the load type would
        'come along in the querystring.
        If lLeafletID > 0 Then
            oboLeaflet.SFAReferenceID = lLeafletID
            oboLeaflet.Load EAFLoadFromDatabase, lLeafletID
        Else
            'load the BO state from the Transient store
            eTempEAFTermLoadType = m_oRequest.QueryString(k_sLoadType)
            oboLeaflet.Load eTempEAFTermLoadType
        End If
        sHTML = ofwHTML.FontWithStyle("Edit Leaflet", ofwHTML.GetFontStyle(FONTPageHeadingStyle)) & "<P>"
    End If
        
    sHTML = sHTML & ofwHTML.FormStart(gk_sSaveLeafletDetail, "", "newLeafletform", , True)
    
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
        .sWidth = "20%"
    End With
    typCellStyleR = ofwHTML.GetCellStyle(CELLReadWriteDataStyle)
    With typCellStyleR
         .sWidth = "80%"
         .lCOLSPAN = 2
    End With
    
    'start a row
    ofwTable.StartRow typRowStyle
    'add CellLeft
    ofwTable.AddCellAsHTML "Name", typCellStyleL
    'add cellright
    '   d11     crm     10Feb2000
    '   d19     MAP     12Oct2000
    ofwTable.AddCellAsHTML ofwHTML.TextField(gk_sDBCOLLeafletTitleText, gk_iDBCOLLenLeafletTitle, oboLeaflet.Title, gk_iFormLenLeafletTitle, False) & "*", _
                           typCellStyleR
    'end row
    ofwTable.EndRow
    
    'start row
    ofwTable.StartRow typRowStyle
    'cell left
    ofwTable.AddCellAsHTML "Created on", typCellStyleL
    'cell right
    '   d4      crm     22Nov1999
    'd16 added if statement so error not thrown if date not set
    sDate = oboLeaflet.CreatedOn
    If sDate <> "" Then
        sDate = DateForDisplay(sDate)
    End If
    ofwTable.AddCellAsHTML sDate & ofwHTML.HiddenField(gk_sDBCOLLeafletCreatedOn, _
                           sDate), typCellStyleR
    'end row
    ofwTable.EndRow
    
    'start row
    ofwTable.StartRow typRowStyle
    'cell left
    ofwTable.AddCellAsHTML "Leaflet file", typCellStyleL
    'cell right
    ofwTable.AddCellAsHTML oboLeaflet.LeafletFile, typCellStyleR
    'end row
    ofwTable.EndRow
    
    ofwTable.AddBlankRow
    
    'now change the cell properties for the Index terms, left is smaller & right is bigger
    typCellStyleL.sWidth = "25%"
    typCellStyleR.sWidth = "50%"
    
    'now set up the far right cell for the Remove buttons
    typCellStyleFR = ofwHTML.GetCellStyle(CELLReadWriteDataStyle)
    With typCellStyleFR
        .sWidth = "15%"
        .eALIGN = HALCenter
    End With
    'now set up the right cell for the details
    typCellStyleTR = ofwHTML.GetCellStyle(CELLReadWriteDataStyle)
    With typCellStyleTR
        .sWidth = "60%"
    End With
    
    
    'start row for 'Index Terms...'
    ofwTable.StartRow typRowStyle
    ofwTable.AddCellAsHTML ofwHTML.Bold("Index Terms")
    ofwTable.AddCellAsHTML ""
    ofwTable.EndRow
    
    'get the corporate body terms
    Set oTempCol = oboLeaflet.CorporateBodyTerms
    
    'set flag to show Corporate Body title
    bPutTitle = True
    For Each vTemp In oTempCol
        ofwTable.StartRow typRowStyle
        ofwTable.AddCellAsHTML IIf(bPutTitle, "Corporate Body", ""), typCellStyleL
        ofwTable.AddCellAsHTML ofwRefdata.Search(gCorporateBody, vTemp, _
                               gk_sDBCOLCorpBodyReferenceID, gk_sDBCOLCorpBodyNameText), typCellStyleTR
        ofwTable.AddCellAsHTML ofwHTML.SubmitImage(gk_sImagesPath & _
                         gk_sImageRemove, "Remove" & k_sCorporateBody & vTemp, k_sCorporateBody & vTemp), typCellStyleFR
        ofwTable.EndRow
        bPutTitle = False
    Next
    
    'add row for PickCorporateBody button
    ofwTable.StartRow typRowStyle
    ofwTable.AddCellAsHTML ""
    ofwTable.AddCellAsHTML ofwHTML.SubmitImage(gk_sImagesPath & gk_sImagePickCorporateBody, _
                           k_sPickTerm & EGroup.gCorporateBody, EGroup.gCorporateBody), typCellStyleR
    ofwTable.EndRow
        
    'get the subject terms
    Set oTempCol = oboLeaflet.SubjectTerms
    
    'set flag to show Subject title
    bPutTitle = True
    For Each vTemp In oTempCol
        ofwTable.StartRow typRowStyle
        ofwTable.AddCellAsHTML IIf(bPutTitle, "Subject", ""), typCellStyleL
        ofwTable.AddCellAsHTML ofwRefdata.Search(gSubject, vTemp, _
                               gk_sDBCOLSubjectReferenceID, gk_sDBCOLSubjectTermText), typCellStyleTR
        ofwTable.AddCellAsHTML ofwHTML.SubmitImage(gk_sImagesPath & _
                         gk_sImageRemove, "Remove" & k_sSubject & vTemp, k_sSubject & vTemp), typCellStyleFR
        ofwTable.EndRow
        bPutTitle = False
    Next
    'add row for PickSubject button
    ofwTable.StartRow typRowStyle
    ofwTable.AddCellAsHTML ""
    ofwTable.AddCellAsHTML ofwHTML.SubmitImage(gk_sImagesPath & _
                           gk_sImagePickSubject, k_sPickTerm & EGroup.gSubject, EGroup.gSubject), typCellStyleR
    ofwTable.EndRow
        
    'get the person terms
    Set oTempCol = oboLeaflet.PersonTerms
    
    'set flag to show Person title
    bPutTitle = True
    'each person term has to be displayed in a specific format
    'right now it is 'Title''Surname' 'Forename' ...change code here
    'format in specific way...look at sTempPersonName variable below
    ' d18 - Changed to match the details page change made in d17
    For Each vTemp In oTempCol
        ofwTable.StartRow typRowStyle
        ofwTable.AddCellAsHTML IIf(bPutTitle, "Person", ""), typCellStyleL
        
        vTempName = ofwRefdata.Search(gPerson, vTemp, _
                               gk_sDBCOLPersonReferenceID, gk_sDBCOLPretitle_Text)
        If IsNull(vTempName) Then
            sTempPersonName = ""
        Else
            sTempPersonName = CStr(vTempName)
        End If
        
        vTempName = ofwRefdata.Search(gPerson, vTemp, _
                               gk_sDBCOLPersonReferenceID, gk_sDBCOLForenameText)
        If Not IsNull(vTempName) Then
            sTempPersonName = sTempPersonName & " " & CStr(vTempName)
        End If
              
        vTempName = ofwRefdata.Search(gPerson, vTemp, _
                               gk_sDBCOLPersonReferenceID, gk_sDBCOLSurnameText)
                               
        If Not IsNull(vTempName) Then
            sTempPersonName = sTempPersonName & " " & CStr(vTempName)
        End If
        
        vTempName = ofwRefdata.Search(gPerson, vTemp, _
                               gk_sDBCOLPersonReferenceID, gk_sDBCOLPerson_Title)
                               
        If Not IsNull(vTempName) Then
            sTempPersonName = sTempPersonName & ", " & CStr(vTempName)
        End If
        
        ofwTable.AddCellAsHTML sTempPersonName, typCellStyleTR
        ofwTable.AddCellAsHTML ofwHTML.SubmitImage(gk_sImagesPath & _
                         gk_sImageRemove, "Remove" & k_sPerson & vTemp, k_sPerson & vTemp), typCellStyleFR
        ofwTable.EndRow
         bPutTitle = False
    Next
        
    'add row for PickPerson button
    ofwTable.StartRow typRowStyle
    ofwTable.AddCellAsHTML ""
    ofwTable.AddCellAsHTML ofwHTML.SubmitImage(gk_sImagesPath & _
                           gk_sImagePickPerson, k_sPickTerm & EGroup.gPerson, EGroup.gPerson), typCellStyleR
    ofwTable.EndRow
        
    'get the place terms
    Set oTempCol = oboLeaflet.PlaceTerms
    
    'set flag to show Place title
    bPutTitle = True
    
    For Each vTemp In oTempCol
        ofwTable.StartRow typRowStyle
        ofwTable.AddCellAsHTML IIf(bPutTitle, "Place", ""), typCellStyleL
        ofwTable.AddCellAsHTML ofwRefdata.Search(gPlace, vTemp, _
                               gk_sDBCOLPlaceReferenceID, gk_sDBCOLPlaceNameText), typCellStyleTR
        ofwTable.AddCellAsHTML ofwHTML.SubmitImage(gk_sImagesPath & _
                         gk_sImageRemove, "Remove" & k_sPlace & vTemp, CStr(vTemp)), typCellStyleFR
        ofwTable.EndRow
        bPutTitle = False
    Next
        
    'add row for PickPlace button
    ofwTable.StartRow typRowStyle
    ofwTable.AddCellAsHTML ""
    ofwTable.AddCellAsHTML ofwHTML.SubmitImage(gk_sImagesPath & _
                           gk_sImagePickPlace, k_sPickTerm & EGroup.gPlace, EGroup.gPlace), typCellStyleR
    ofwTable.EndRow
           
    ofwTable.AddBlankRow
    
    'footer cell properties
    With typCellStyleR
        .eStyle = CELLCustomStyle
        .eALIGN = HALLeft
        .eBGCOLOR = COLOURWhite
    End With

    ofwTable.StartRow typRowStyle
    ofwTable.AddCellAsHTML "", typCellStyleL
     '   d10     crm     03Feb2000
    If Not bNewLeaflet Then
        ' d18 - Added Delete button
        ofwTable.AddCellAsHTML ofwHTML.SubmitImage(gk_sImagesPath & _
                            gk_sImageOK, "OK", "OK") & _
                            ofwHTML.ImageButton(gk_sImagesPath & gk_sImageCancel, _
                            "Cancel", gk_sDisplayLeafletDetail & "?" & _
                            gk_sDBCOLSFAReferenceID & "=" & oboLeaflet.SFAReferenceID) & _
                            ofwHTML.ImageButton(gk_sImagesPath & gk_sImageDeleteLeaflet, _
                            "Remove", gk_sPAGEDeleteLeaflet & "?" & _
                            gk_sDBCOLSFAReferenceID & "=" & oboLeaflet.SFAReferenceID) & "</P>"
                           
    Else
        '   d4      crm     22Nov1999
        'NOTE : if a new Leaflet then the Cancel button has link to AllLeafletFiles display page
        ofwTable.AddCellAsHTML ofwHTML.SubmitImage(gk_sImagesPath & _
                            gk_sImageOK, "OK", "OK") & _
                            ofwHTML.ImageButton(gk_sImagesPath & gk_sImageCancel, _
                            "Cancel", gk_sDisplayAllLeaflets), typCellStyleR
    
    End If

    ofwTable.EndRow
    
    sHTML = sHTML & ofwHTML.HiddenField(gk_sDBCOLSFAReferenceID, oboLeaflet.SFAReferenceID)
    
    sHTML = sHTML & ofwTable.GetHTML
    
    sHTML = sHTML & "</FORM>"
    
    'set this state to persist
    oboLeaflet.PROCATTransientInterface_SetState oboLeaflet.Name
    
    sHTML = sHTML & HTMLToAddHelpButton(gk_sHelpLinkToDisplayEditLeafletDetail)
    
    m_oResponse.Write sHTML
    
    'destroy object references
    Set ofwHTML = Nothing
    Set ofwTable = Nothing
    Set oboLeaflet = Nothing
    
    m_oObjectContext.SetComplete
    
Exit Sub
ErrorHandler:
    ShowError sSource
    m_oObjectContext.SetAbort
    Set ofwHTML = Nothing
    Set ofwTable = Nothing
    Set oboLeaflet = Nothing
End Sub


Public Sub DisplayLeafletDetail()
'Description    :       Method to display a static page with the details of the leaflet


'Amendment History
'   d2          crm            02Nov1999    Created
'   d4          crm            22Nov1999    Modified to display IndexTerms of Leaflet
'   d5          ksk            22Nov1999    Added security
'   d10         crm            03Feb2000    Added navigation links..just some ids passed to get back
'                                           to where we started from
'   d11         crm            11Feb2000    Corrected alignment for 'change source document cell'
'                                           Added done button and removed hyperlink on top of page
'   d14         WDP            23Feb2000    Added DateForDisplay calls for displaying created on dates
'   d16         TMJ            02Mar2000    MFixed so does not throw an error when date is not set
'   d17     DAJ     26Jun2000      Changed layout of AF term details in leaflet detail page to match
'                                   that of the catalouge layout (mostly)

On Error GoTo ErrorHandler

    Dim sSource As String
    Dim sHTML As String
    Dim bIsEditPage As Boolean
    Dim lLeafletID As Long
    Dim ofwTable As New FWTable
    Dim ofwHTML As New FWHTML
    Dim ofwInnerTable As New FWTable
    
    Dim typTableStyle As TableStyle
    Dim typRowStyle As RowStyle
    Dim typLeftCellStyle As CellStyle
    Dim typRightCellStyle As CellStyle
    Dim lLeafletFile As String
    Dim oTempCol As Collection
    Dim vTemp As Variant
    Dim sTempPersonName As String
    Dim sTempPlaceName As String
    Dim sTempCorpBodyName As String
    Dim ofwRefdata As New FWReferenceData
    Dim vTempName As Variant
    Dim bPutTitle As Boolean
    Dim oboLeaflet As New BOLeaflet
    Dim sDate As String
    
    sSource = "VWLeaflet.DisplayLeafletDetail"
    
    'd5 Check if the user has right to perform the operation
    CheckForRightOfUser
    
    'get the ID of the SFA/leaflet
    '   d10         crm            03Feb2000
    lLeafletID = m_oRequest.QueryString(gk_sDBCOLSFAReferenceID)
    If lLeafletID = 0 Then
      lLeafletID = m_oRequest.QueryString("sLeafletID")
    End If
    
    If lLeafletID > 0 Then
    oboLeaflet.SFAReferenceID = lLeafletID
    oboLeaflet.Load EAFLoadFromDatabase, lLeafletID
    End If
    
    'check if the user has come from
    'after choosing a leaflet file
    lLeafletFile = m_oRequest.QueryString(ksLeafletName)
    
    'if yes then update this leaflet file for the leaflet
    If Len(Trim(lLeafletFile)) > 0 Then
        oboLeaflet.SFAReferenceID = lLeafletID
        oboLeaflet.LeafletFile = lLeafletFile
        oboLeaflet.Save True
        
        'TODO
        oboLeaflet.IndexInVerity False
        
        ' Tomh d21 (have added this as was a defect that leaflet source
        ' file not being deleted)
        ' If all actions have completed OK, we delete the leaflet source from
        ' the submission directory
        oboLeaflet.DeleteSourceFile
       
        'load again for subsequent use below.
        'this is to ensure that we are using updated
        'data from the DB
        oboLeaflet.Load EAFLoadFromDatabase, lLeafletID
    End If
    
    'set the table style
    With typTableStyle
        .eALIGN = HALCenter
        .eBGCOLOR = COLOURWhite
        .sWidth = "100%"
        .eStyle = TABCustomStyle
        .lCELLPADDING = 2
        .lCELLSPACING = 2
    End With
    
    'set the left/right cell styles
    typLeftCellStyle = ofwHTML.GetCellStyle(CELLReadOnlyLabelStyle)
    With typLeftCellStyle
        .sWidth = "25%"
    End With
    With typRightCellStyle
        .eALIGN = HALRight
        .eStyle = CELLCustomStyle
        .sWidth = "75%"
    End With
    
    'start the table
    ofwTable.TableStyle = typTableStyle
    
    ofwTable.StartRow ofwHTML.GetRowStyle(ROWCustomStyle)
'    ofwTable.AddCellAsHTML ofwHTML.Bold(ofwHTML.Font("Leaflet Details", COLOURBlack, 4)), _
'                           typLeftCellStyle
    ofwTable.AddCellAsHTML ofwHTML.FontWithStyle("Leaflet details", ofwHTML.GetFontStyle(FONTPageHeadingStyle)), _
                           typLeftCellStyle
    ofwTable.AddCellAsHTML "", typRightCellStyle
    ofwTable.EndRow
    
    'add a empty row
    ofwTable.AddBlankRow
    
    'change the cell style for right cell
    typRightCellStyle = ofwHTML.GetCellStyle(CELLReadWriteDataStyle)
    With typRightCellStyle
        .sWidth = "70%"
    End With
    
    'start row for 'Name'
    ofwTable.StartRow ofwHTML.GetRowStyle(ROWCustomStyle)
    ofwTable.AddCellAsHTML "Name", typLeftCellStyle
    ofwTable.AddCellAsHTML oboLeaflet.Title, typRightCellStyle
    ofwTable.EndRow
    
    'start row for 'CreatedOn'
    ofwTable.StartRow ofwHTML.GetRowStyle(ROWCustomStyle)
    ofwTable.AddCellAsHTML "Created on", typLeftCellStyle
    sDate = oboLeaflet.CreatedOn
    'added so to catch case when date is not set (null or the zero length string)
    If sDate <> "" Then
        sDate = DateForDisplay(sDate)
    End If
    ofwTable.AddCellAsHTML sDate, typRightCellStyle
    ofwTable.EndRow
    
    'start row for 'Edit these items...'
    ofwTable.StartRow ofwHTML.GetRowStyle(ROWCustomStyle)
    ofwTable.AddCellAsHTML ofwHTML.Hyperlink(gk_sDisplayEditLeafletDetail & "?" & _
                           gk_sDBCOLSFAReferenceID & "=" & oboLeaflet.SFAReferenceID, _
                           "Edit these details...", gk_sMainFrame, , ofwRefdata.Search(gToolTip, ETOOLTIP.ttEditLeaflet)), typLeftCellStyle
    ofwTable.AddCellAsHTML ""
    ofwTable.EndRow
   
    'start row for LeafletFiles
    ofwTable.StartRow ofwHTML.GetRowStyle(ROWCustomStyle)
    ofwTable.AddCellAsHTML "Leaflet file", typLeftCellStyle
    'TODO
    'prepare the inner table for the second cell
    typTableStyle.lCELLSPACING = 0
    typTableStyle.lCELLPADDING = 0
    ofwInnerTable.TableStyle = typTableStyle

    ofwInnerTable.StartRow ofwHTML.GetRowStyle(ROWCustomStyle)
    typLeftCellStyle.eBGCOLOR = COLOURPink
    ofwInnerTable.AddCellAsHTML oboLeaflet.LeafletFile, typLeftCellStyle
    'set align to right
    typRightCellStyle.eALIGN = HALRight
    typRightCellStyle.eBGCOLOR = COLOURPink
    ofwInnerTable.AddCellAsHTML ofwHTML.Hyperlink(gk_sDisplayAllLeafletFiles & "?" & _
                                gk_sDBCOLSFAReferenceID & "=" & oboLeaflet.SFAReferenceID & "&" & _
                                gk_sDBCOLSFATitleText & "=" & oboLeaflet.LeafletFile, _
                                "Change the source leaflet document...", gk_sMainFrame, , ofwRefdata.Search(gToolTip, ETOOLTIP.ttChangeDocForLeaflet)), typRightCellStyle
    ofwInnerTable.EndRow
    
    'get content of inner-table to outer table's right cell
    typRightCellStyle.eALIGN = HALRight
    ofwTable.AddCellAsHTML ofwInnerTable.GetHTML, typRightCellStyle
    ofwTable.EndRow
    
    '   d4      crm     22Nov1999
    'NOTE   :   Show index terms of Leaflet
    
    'start row for 'Index Terms...'
    ofwTable.StartRow ofwHTML.GetRowStyle(ROWCustomStyle)
    ofwTable.AddCellAsHTML ofwHTML.Bold("Index Terms")
    ofwTable.AddCellAsHTML ""
    ofwTable.EndRow
    
    'set the left/right cell styles
    typLeftCellStyle = ofwHTML.GetCellStyle(CELLReadOnlyLabelStyle)
    With typLeftCellStyle
        .sWidth = "25%"
    End With

    typRightCellStyle = ofwHTML.GetCellStyle(CELLReadOnlyDataStyle)
    With typRightCellStyle
        .sWidth = "75%"
    End With
    
    
    'TODO - Include code to display IndexTerms for Leaflet
    'get the corporate body terms
    Set oTempCol = oboLeaflet.CorporateBodyTerms
    
    'set flag to show Corporate Body title
    bPutTitle = True
    For Each vTemp In oTempCol
        ofwTable.StartRow typRowStyle
        ofwTable.AddCellAsHTML IIf(bPutTitle, "Corporate Body", ""), typLeftCellStyle
        
        vTempName = ofwRefdata.Search(gCorporateBody, vTemp, _
                               gk_sDBCOLCorpBodyReferenceID, gk_sDBCOLCorpBodyNameText)
                               
        If IsNull(vTempName) Then
            sTempCorpBodyName = ""
        Else
            sTempCorpBodyName = CStr(vTempName)
        End If

        ofwTable.AddCellAsHTML sTempCorpBodyName, typRightCellStyle
        ofwTable.EndRow
        bPutTitle = False
    Next
        
    'get the subject terms
    Set oTempCol = oboLeaflet.SubjectTerms
    
    'set flag to show Subject title
    bPutTitle = True
    For Each vTemp In oTempCol
        ofwTable.StartRow typRowStyle
        ofwTable.AddCellAsHTML IIf(bPutTitle, "Subject", ""), typLeftCellStyle
        ofwTable.AddCellAsHTML ofwRefdata.Search(gSubject, vTemp, _
                               gk_sDBCOLSubjectReferenceID, gk_sDBCOLSubjectTermText), _
                               typRightCellStyle
        ofwTable.EndRow
        bPutTitle = False
    Next
        
    'get the person terms
    Set oTempCol = oboLeaflet.PersonTerms
    
    'set flag to show Person title
    bPutTitle = True
    
    'each person term has to be displayed in a specific format
    'right now it is:
    ' 'Pre-title' 'Forename' 'Surname', 'Title' ...change code here
    'format in specific way...look at sTempPersonName variable below
    ' d17 - Added pretitle and moved layout
    For Each vTemp In oTempCol
        ofwTable.StartRow typRowStyle
        ofwTable.AddCellAsHTML IIf(bPutTitle, "Person", ""), typLeftCellStyle
        
        vTempName = ofwRefdata.Search(gPerson, vTemp, _
                               gk_sDBCOLPersonReferenceID, gk_sDBCOLPretitle_Text)
                               
        If IsNull(vTempName) Then
            sTempPersonName = ""
        Else
            sTempPersonName = CStr(vTempName)
        End If
        
        vTempName = ofwRefdata.Search(gPerson, vTemp, _
                               gk_sDBCOLPersonReferenceID, gk_sDBCOLForenameText)
        If Not IsNull(vTempName) Then
            sTempPersonName = sTempPersonName & " " & CStr(vTempName)
        End If
              
        vTempName = ofwRefdata.Search(gPerson, vTemp, _
                               gk_sDBCOLPersonReferenceID, gk_sDBCOLSurnameText)
                               
        If Not IsNull(vTempName) Then
            sTempPersonName = sTempPersonName & " " & CStr(vTempName)
        End If
        
        vTempName = ofwRefdata.Search(gPerson, vTemp, _
                               gk_sDBCOLPersonReferenceID, gk_sDBCOLPerson_Title)
                               
        If Not IsNull(vTempName) Then
            sTempPersonName = sTempPersonName & ", " & CStr(vTempName)
        End If
        
        ofwTable.AddCellAsHTML sTempPersonName, typRightCellStyle
        ofwTable.EndRow
         bPutTitle = False
    Next
        
    'get the place terms
    Set oTempCol = oboLeaflet.PlaceTerms
    
    'set flag to show Place title
    bPutTitle = True
    
    ' d17 - Added parish to Place description
    For Each vTemp In oTempCol
        ofwTable.StartRow typRowStyle
        ofwTable.AddCellAsHTML IIf(bPutTitle, "Place", ""), typLeftCellStyle
        
        vTempName = ofwRefdata.Search(gPlace, vTemp, _
                               gk_sDBCOLPlaceReferenceID, gk_sDBCOLPlaceNameText)
                               
        If IsNull(vTempName) Then
            sTempPlaceName = ""
        Else
            sTempPlaceName = CStr(vTempName)
        End If
        
        vTempName = ofwRefdata.Search(gPlace, vTemp, _
                               gk_sDBCOLPlaceReferenceID, gk_sDBCOLParishName)
        If Not IsNull(vTempName) Then
            sTempPlaceName = sTempPlaceName & ", " & CStr(vTempName)
        End If
        
        ofwTable.AddCellAsHTML sTempPlaceName, typRightCellStyle
        ofwTable.EndRow
        bPutTitle = False
    Next
    
    'get the HTML strings
    sHTML = ofwTable.GetHTML
    
    sHTML = sHTML & "<P ALIGN=RIGHT >" & ofwHTML.Hyperlink( _
            gk_sInformationLeafletBrowser & "?" & "eUsage=" & _
            ELeafletBrowserUsage.ELeafletMaintenance, _
            ofwHTML.Image("", gk_sImagesPath & gk_sImageDone), gk_sMainFrame) & "</P>"
    
    sHTML = sHTML & HTMLToAddHelpButton(gk_sHelpLinkToDisplayLeafletDetail)
    
    'send the HTML outside
    m_oResponse.Write sHTML
    
    'release object references
    Set ofwTable = Nothing
    Set ofwHTML = Nothing
    Set ofwInnerTable = Nothing
    Set oboLeaflet = Nothing
    Set ofwRefdata = Nothing
            
    m_oObjectContext.SetComplete
    
Exit Sub
ErrorHandler:
    ShowError sSource
    m_oObjectContext.SetAbort
    Set ofwTable = Nothing
    Set ofwHTML = Nothing
    Set ofwInnerTable = Nothing
    Set oboLeaflet = Nothing
    Set ofwRefdata = Nothing
End Sub

Public Sub DeleteLeaflet()
' Function to completely remove the leaflet from the system, includin Verity indexes
'
' Amendment History:
'   d18     DAJ     10Jul2000      Added DeleteLeaflet and changed DisplayEditLeafletDetail to add a Delete button
On Error GoTo ErrorHandler

    Dim sSource As String
    sSource = "VWLeaflet.DeleteLeaflet"

    ' Check that user can actually delete a leaflet
    CheckForRightOfUser
    
    ' Get a leaflet business object
    Dim oboLeaflet As New BOLeaflet
    
    ' Get the ID of the SFA/leaflet
    Dim lLeafletID As Long
    lLeafletID = m_oRequest.QueryString(gk_sDBCOLSFAReferenceID)
    If lLeafletID = 0 Then
        lLeafletID = m_oRequest.QueryString("sLeafletID")
    End If
    
    ' Set the ID in the business object
    oboLeaflet.SFAReferenceID = lLeafletID
    
    ' Get the leaflet details
    oboLeaflet.Load EAFLoadFromDatabase, lLeafletID
    
    ' Can the leaflet be deleted?
    If Not oboLeaflet.AbleToDelete Then
        RaiseError sSource, peCanNotDeleteSFAWithReferences, "Leaflet is linked to by one or more catalogue entries; you cannot delete it"
    End If
    
    ' Delete the Verity indexes
    oboLeaflet.DeleteFromVerity

    ' Delete the database entries for the leaflet
    oboLeaflet.Delete

    ' Tidy up the transient store
    oboLeaflet.PROCATTransientInterface_DeleteState oboLeaflet.Name
    
    ' Delete the leaflet file in the live directory
    oboLeaflet.DeleteLiveFile

    ' Redirect to the browser
    m_oResponse.Redirect gk_sInformationLeafletBrowser & "?" & "eUsage=" & _
            ELeafletBrowserUsage.ELeafletMaintenance
    
    m_oObjectContext.SetComplete
    
Exit Sub
ErrorHandler:
    ' If the leaflet file has been deleted... we were going to anyway!
    If Err.Number = 53 Then
        Resume Next
    End If
    
    ShowError sSource
    m_oObjectContext.SetAbort
    Set oboLeaflet = Nothing
End Sub

Public Sub SaveLeafletDetail()
'Description    :       Method to Save Leaflet detail
'
'
'Amendment History
'   d2          crm            09Nov1999   Created
'   d5          ksk            22Nov1999   Added security
'   d6          wdp            02Dec1999   Added Verity indexing
'   d12         WDP            15Feb2000   Added call to BOLeaflet.DeleteSourceFile in SaveLeafletDetail
'   d20     DAJ     06Nov2000      Changed order of SaveLeafletDetail code so that it can understand if the user
'                                       pressed Return in the form instead of clicking OK
On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim sHTML As String
    Dim sTempSring As String
    Dim oboLeaflet As New BOLeaflet
    Dim lIndextermID As Long
    Dim vItem As Variant
    Dim eTempGroup As EGroup
    Dim sRedirectURL As String
    Dim eTempTermType As EAFTermType
    
    sSource = "VWLeaflet.SaveLeafletDetail"
    
    'd5 Check if the user has right to perform the operation
    CheckForRightOfUser
    
    ' The cancel operation was requested, so just display the leaflets
    If m_oRequest.Form("Cancel" & ".x") <> "" Then
        oboLeaflet.PROCATTransientInterface_GetState oboLeaflet.Name
        m_oResponse.Redirect gk_sDisplayAllLeaflets
        oboLeaflet.PROCATTransientInterface_DeleteState oboLeaflet.Name
        Exit Sub
    End If
    
    'NOTE : To get the ID of the id which has to be removed we need the IndexTerm type and the ID
    '       of the term.See the DisplayEditLeafletDetail() where we concatenate the ID and the type
    '       of the IndexTerm to name the Remove button.When we receive the Button after submit a '.x'
    '       is sent along with the button name...trim it and trim the type name which is concatenated
    '       to get the Id and the Type of the item to remove.
    
    'see if we have a remove AF term action
    If InStr(1, m_oRequest.Form, "Remove") > 0 Then
    'now loop thru Form content to fetch the Remove member
    For Each vItem In m_oRequest.Form
    If InStr(1, vItem, "Remove") > 0 Then
          'get the state from Transient storage and assign the new values
          oboLeaflet.PROCATTransientInterface_GetState oboLeaflet.Name
          oboLeaflet.Title = m_oRequest.Form(gk_sDBCOLLeafletTitleText)
          oboLeaflet.CreatedOn = m_oRequest.Form(gk_sDBCOLLeafletCreatedOn)
          
          'if it is a CorporateBody removal
          If InStr(1, vItem, k_sCorporateBody) > 0 Then
                 vItem = Left(CStr(vItem), Len(CStr(vItem)) - 2)
                lIndextermID = CLng(Mid(vItem, Len("Remove" & k_sCorporateBody) + 1))
                oboLeaflet.RemoveIndexTerm lIndextermID, gCorporateBody
                oboLeaflet.PROCATTransientInterface_SetState oboLeaflet.Name
                m_oResponse.Redirect gk_sDisplayEditLeafletDetail & "?" & _
                            k_sLoadType & "=" & EAFTermLoadType.EAFLoadFromStateRemoveTerm
                Exit Sub
          End If
          
          'if it is a Subject removal
          If InStr(1, vItem, k_sSubject) > 0 Then
                vItem = Left(CStr(vItem), Len(CStr(vItem)) - 2)
                lIndextermID = CLng(Mid(vItem, Len("Remove" & k_sSubject) + 1))
                oboLeaflet.RemoveIndexTerm lIndextermID, gSubject
                oboLeaflet.PROCATTransientInterface_SetState oboLeaflet.Name
                m_oResponse.Redirect gk_sDisplayEditLeafletDetail & "?" & _
                            k_sLoadType & "=" & EAFTermLoadType.EAFLoadFromStateRemoveTerm
                Exit Sub
          End If
          
          'if it is a Person removal
          If InStr(1, vItem, k_sPerson) > 0 Then
                 vItem = Left(CStr(vItem), Len(CStr(vItem)) - 2)
                lIndextermID = CLng(Mid(vItem, Len("Remove" & k_sPerson) + 1))
                oboLeaflet.RemoveIndexTerm lIndextermID, gPerson
                oboLeaflet.PROCATTransientInterface_SetState oboLeaflet.Name
                m_oResponse.Redirect gk_sDisplayEditLeafletDetail & "?" & _
                            k_sLoadType & "=" & EAFTermLoadType.EAFLoadFromStateRemoveTerm
                Exit Sub
          End If
          
          'if it is a Place removal
          If InStr(1, vItem, k_sPlace) > 0 Then
                 vItem = Left(CStr(vItem), Len(CStr(vItem)) - 2)
                lIndextermID = CLng(Mid(vItem, Len("Remove" & k_sPlace) + 1))
                oboLeaflet.RemoveIndexTerm lIndextermID, gPlace
                oboLeaflet.PROCATTransientInterface_SetState oboLeaflet.Name
                m_oResponse.Redirect gk_sDisplayEditLeafletDetail & "?" & _
                            k_sLoadType & "=" & EAFTermLoadType.EAFLoadFromStateRemoveTerm
                Exit Sub
          End If
    End If
    Next
    End If
    
    'see if we have received a PickTerm action
    If InStr(1, m_oRequest.Form, k_sPickTerm) > 0 Then
        For Each vItem In m_oRequest.Form
            If InStr(1, vItem, k_sPickTerm) > 0 Then
                oboLeaflet.PROCATTransientInterface_GetState oboLeaflet.Name
                oboLeaflet.Title = m_oRequest.Form(gk_sDBCOLLeafletTitleText)
                oboLeaflet.CreatedOn = m_oRequest.Form(gk_sDBCOLLeafletCreatedOn)
                oboLeaflet.PROCATTransientInterface_SetState oboLeaflet.Name
                
                vItem = Left(CStr(vItem), Len(CStr(vItem)) - 2)
                eTempGroup = Mid(CStr(vItem), Len(CStr(vItem)))
                           
                'set the TermTypes which have to be displayed
                Select Case eTempGroup
                    Case EGroup.gCorporateBody
                        eTempTermType = EAFTermType.AFCorporateBody
                    Case EGroup.gPerson
                        eTempTermType = EAFTermType.AFPerson
                    Case EGroup.gPlace
                        eTempTermType = EAFTermType.AFPlace
                    Case EGroup.gSubject
                        eTempTermType = EAFTermType.AFSubject
                    Case Else
                        eTempTermType = EAFTermType.AFAllTypes
                End Select
                
                'set the redirect URL
                sRedirectURL = gk_sDisplayIndexTermList & "?" & _
                gk_sFORMAFGroup & "=" & eTempTermType & _
                    "&" & gk_sFORMAFListIndex & "=" & "0" & _
                    "&" & gk_sFORMAFTermID & "=" & "1" & _
                    "&" & gk_sFORMAFListType & "=" & "1" & _
                    "&" & gk_sFORMAFListTermType & "=" & "1" & _
                    "&" & gk_sFORMAFListUsage & "=" & EAFBrowserUsage.BROWSEEditorialInformationLeafletPickTerm & _
                    "&" & gk_sFORMAFinSFA & "=" & "1"
                               
                'redirect to the page
                m_oResponse.Redirect sRedirectURL
                Exit Sub
            End If
        Next
    End If
    
    ' d20 - If we passed through all those, the user must have submitted the form
    
    'get the state from Transient storage
    oboLeaflet.PROCATTransientInterface_GetState oboLeaflet.Name
    oboLeaflet.Title = m_oRequest.Form(gk_sDBCOLLeafletTitleText)
    oboLeaflet.CreatedOn = m_oRequest.Form(gk_sDBCOLLeafletCreatedOn)
    
    'save to database
    Dim bIsNew As Boolean
    
    If oboLeaflet.SFAReferenceID > 0 Then
        bIsNew = False
    Else
        bIsNew = True
    End If
    
    oboLeaflet.Save
    
    oboLeaflet.IndexInVerity bIsNew
    
    'redirect to non-editable detail screen...send the ID and the LoadType to be performed
    'on that page
    
    m_oResponse.Redirect gk_sDisplayLeafletDetail & "?" & _
        gk_sDBCOLSFAReferenceID & "=" & oboLeaflet.SFAReferenceID & _
        "&" & k_sLoadType & "=" & EAFTermLoadType.EAFLoadFromDatabase
    oboLeaflet.PROCATTransientInterface_DeleteState oboLeaflet.Name
    
    'If all actions have completed OK, we delete the leaflet source from
    'the submission directory
    oboLeaflet.DeleteSourceFile
    Exit Sub
    
    'release object references
    Set oboLeaflet = Nothing
    
    m_oObjectContext.SetComplete
Exit Sub
ErrorHandler:
    ShowError sSource
    m_oObjectContext.SetAbort
    Set oboLeaflet = Nothing
    
End Sub

Public Sub AddTermInformationLeaflet()
'Description    :       Method to add an index term against a leaflet and redirect page to
'                       edit screen.User comes here after choosing the term from the IndexBrowser


'Amendment History
'   d2          crm            09Nov1999    Created
'   d5          ksk            22Nov1999    Added security
'   d8          crm            16Dec1999    Added comments
On Error GoTo ErrorHandler

    Dim sSource As String
    Dim oboLeaflet As New BOLeaflet
    Dim eChoosenGroup As EGroup
    Dim lChoosenID As Long
    
    sSource = "VWLeaflet.AddTermInformationLeaflet"
    
    'd5 Check if the user has right to perform the operation
    CheckForRightOfUser
    
    'get the previous state
    oboLeaflet.PROCATTransientInterface_GetState oboLeaflet.Name
    
    'get the AFGroup choosen
    eChoosenGroup = m_oRequest.QueryString(gk_sFORMAFGroup)
    'get the ID of the term choosen
    lChoosenID = m_oRequest.QueryString(gk_sFORMAFTermID)
    'add this to the appropriate collection in the BO
    oboLeaflet.AddIndexTerm lChoosenID, eChoosenGroup
    'set the state
    oboLeaflet.PROCATTransientInterface_SetState oboLeaflet.Name
    'now go back to the edit page and set the kind
    'of load which has to be done.
    m_oResponse.Redirect gk_sDisplayEditLeafletDetail & "?" & _
                         k_sLoadType & "=" & EAFTermLoadType.EAFLoadFromStateAddTerm
                             
    'release object references
    Set oboLeaflet = Nothing
    
    m_oObjectContext.SetComplete
Exit Sub
ErrorHandler:
    ShowError sSource
    m_oObjectContext.SetAbort
    Set oboLeaflet = Nothing
End Sub

Public Sub CancelAddTermInformationLeaflet()
'Description    :       Method to redirect after cancelling choosing an Index term from
'                       the Index Browser
'
'
'Amendment History
'   d2          crm            9Nov1999    Created
'   d5          ksk            22Nov1999   Added security
'   d8          crm            16Dec1999   Added comments
On Error GoTo ErrorHandler

    Dim sSource As String
    sSource = "VWLeaflet.CancelAddTermInformationLeaflet"
    
    'd5 Check if the user has right to perform the operation
    CheckForRightOfUser
    'note-----There is nothing to do here.Just set the Load Type in the querystring
    m_oResponse.Redirect gk_sDisplayEditLeafletDetail & "?" & _
                         k_sLoadType & "=" & EAFTermLoadType.EAFLoadFromStateTermNoChange
    m_oObjectContext.SetComplete
Exit Sub
ErrorHandler:
    ShowError sSource
    m_oObjectContext.SetAbort
End Sub



Private Sub CheckForRightOfUser()
On Error GoTo ErrorHandler
'Checks if the current user has right to preform the operation raises an error
' if he does not have one
'   d5      ksk     22Nov1999      Added security. Checks if the user has
'                                  permission to execute the method

    Dim ePermissionCheckResult As ESecurityCheckResult
    Dim sSource As String
    
    sSource = "VWLeafLet.CheckForRightOfUser"
    
    m_oboPermission.UserID = m_lLoggedUserID
    ePermissionCheckResult = m_oboPermission.Check(ESecurityCheck.EUserCanMaintainLeaflets)
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
