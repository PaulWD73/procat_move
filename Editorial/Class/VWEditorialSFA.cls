VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "VWEditorialSFA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    VWEditorialSFA.cls
' System:       PRO, PROCAT
' Copyright:    (C) Quidnunc Limited
'
' Description:  Genertes the HTML for creating , editing and deleting
'               the SFA.
'
' Amendment history:
'   d1  NDK    08Nov1999   Created
'   d2  NDK    17Nov1999   Changed the constant to pass query string
'   d3  NDK    18Nov1999   passing sfa type in query strings
'   d4  NDK    18Nov1999   Changes made to display
'   d5  NDK    19Nov1999   Showing different headings depending on what is created
'   d6  NDK    23Nov1999   Done button displayed to link back to the browser
'   d7  NDK    23Nov1999   'URL' displayed instead of 'FileName'
'   d8  NDK    08Dec1999   Added help buttons, changed the alignment of the details table
'   d9  NDK    21Dec1999   Added security check where required in the methods
'   d10 ndk    27Dec1999   Changed to VI code use customizable styles
'   d11 ndk    28Dec1999   Added tooltips
'   d12 ndk    29Dec1999   Rechecking releasing of variables, especially if an error is deliberately raised
'   d13 ndk    29DEC1999   Using resource file to display captions wherever not present
'   d14 crm    10Feb2000   Using BOSFAs HasReference method to display message if there is
'                          an existing reference
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'Interfaces
Implements ObjectControl

'asp/mts objects
Private m_objObjectContext As ObjectContext
Private m_aspResponse  As Response
Private m_aspRequest As Request

Private m_lLoggedUserID As Long                     'ID of the logged user

'Constants
Private Const ksTableWidth As String = "100%"
Private Const klTablePadding As Long = 1
Private Const klTableSpacing As Long = 2
Private Const ksLabelColumnWidth As String = "25%"

'private types
Private m_typSFADetails As SFADetails

'The SFA loader
Private m_oBOSFALoader As BOSFALoader
Private m_oBOEditorialPermissions As BOEditorialPermissions

'HTML object for reference
Private m_fwTable As FWTable
Private m_oFWHTML As FWHTML
Private m_oFWAppData As FWApplicationData

'constants for query strings


Private Const ksParamSFATitle As String = "Title"

'Names for text fields
Private Const ksNameSFATitle As String = "Title"
Private Const ksNameSFASubTitle As String = "SubTitle"
Private Const ksNameSFAPublisher As String = "Publisher"
Private Const ksNameSFAAuthor As String = "Author"
Private Const ksNameSFAPublicationDate As String = "PublicationDate"
Private Const ksNameSFALocation As String = "Location"
Private Const ksNameSFAExternalReference As String = "ExternalReference"
Private Const ksNameSFAEdition As String = "Edition"
Private Const ksNameSFAPublicationNote As String = "PublicationNote"
Private Const ksNameSFAElectronicFileName As String = "ElectronicFileName"
Private Const ksNameButtonOk As String = "cmdOK"
Private Const ksNameButtonCancel As String = "cmdCancel"

'Widths for text fields
Private Const kiSFATitleWidth As Integer = 40
Private Const kiSFASubTitleWidth As Integer = 40
Private Const kiSFAPublisherWidth As Integer = 40
Private Const kiSFAAuthorWidth As Integer = 40
Private Const kiSFAPublicationDateWidth As Integer = 40
Private Const kiSFAExternalReferenceWidth As Integer = 40
Private Const kiSFAEditionWidth As Integer = 40
Private Const kiSFAPublicationNoteWidth As Integer = 40
Private Const kiSFAElectronicFileNameWidth As Integer = 40

'radio buttonvalues
Private Const kiInProLibraryValue = 1
Private Const kiInProReferenceValue = 2
Private Const kiUnknownLocationValue = 3


Public Property Get Name()
    Name = "VWEditorialSFA"
End Property

Public Sub DisplayCreateEditSFADetailsScreen()
'   Description
'   Used for creating and updating sfas
'   For existing sfas Loads the SFA with BOSFALoader.
'   Call sub functions to add the details to the HTML
'   Writes the HTML back to the ASP.
'
'   d1  NDK    08Nov1999   Created
'   d5  NDK    19Nov1999   Showing different headings depending on what is created
'   d8  NDK    08Dec1999   Added help buttons
'   d9  NDK    21Dec1999   Added security check
'   d13 NDK    29DEC1999   Using resource file to display captions wherever not present

On Error GoTo ErrorHandler

    Dim sSource As String
    Dim lSFAReferenceID As Long, bCreateNew As Boolean, eSFAType As ESFAFilter
    Dim sParamString As String, sHTML As String, sHeading As String
    Dim eSecurityCheckStatus As ESecurityCheckResult
    
    sSource = Me.Name & ".DisplayCreateEditSFADetailsScreen"
    
    'initialise the object
    Set m_oBOSFALoader = New BOSFALoader
    
    'get the values from the query string
    
    bCreateNew = m_aspRequest.QueryString(gk_sParamCreateNewSFA)
    eSFAType = m_aspRequest.QueryString(gk_sParamSFAType)
    lSFAReferenceID = m_aspRequest.QueryString(gk_sParamSFAReferenceID)
    
    'perform security check
    m_oBOEditorialPermissions.UserID = m_lLoggedUserID
    
    If eSFAType = ESFAFilter.ESFAElectronic Then
        eSecurityCheckStatus = m_oBOEditorialPermissions.Check(EUserCanMaintainElectronicSFA)
    Else
        eSecurityCheckStatus = m_oBOEditorialPermissions.Check(EUserCanMaintainPrintedSFA)
    End If
    
    If eSecurityCheckStatus <> ESecurityCheckResult.ESRSecurityCheckSuccessful Then
        RaiseError sSource, _
               GetErrorEnumForSecurityCheckResult(eSecurityCheckStatus), _
               GetDescriptionForSecurityCheckError( _
                        GetErrorEnumForSecurityCheckResult(eSecurityCheckStatus))
    End If
    
    'start the form
    sParamString = _
        gk_sLinkCreateUpdateSFA & _
        "?" & gk_sParamSFAType & "=" & eSFAType & _
        "&" & gk_sParamCreateNewSFA & "=" & bCreateNew & _
        "&" & gk_sParamSFAReferenceID & "=" & lSFAReferenceID
        
    sHTML = m_oFWHTML.FormStart(sParamString, gk_sMainFrame, gk_sHTMLDefaultFormName, MPost, True)
    
    If Not bCreateNew Then      'Editing an existing an SFA
        'load the data from bo to type
        
        m_oBOSFALoader.LoadData (lSFAReferenceID)
        m_typSFADetails = m_oBOSFALoader.Data
        'update sfas
        If eSFAType = ESFAFilter.ESFAElectronic Then
            sHeading = m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayCreateEditSfaCaptionModifyElectronicSFA)
        Else
            sHeading = m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayCreateEditSfaCaptionModifyPrintedSFA)
        End If
    Else
        'create sfa
        If eSFAType = ESFAFilter.ESFAElectronic Then
            sHeading = m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayCreateEditSfaCaptionCreateElectronicSFA)
        Else
            sHeading = m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayCreateEditSfaCaptionCreatePrintedSFA)
        End If
    End If

    m_fwTable.Reset
    'convert data into html
    sHTML = HeaderHTML(sHeading) & sHTML & AddTable(bCreateNew, eSFAType)
        
    'Add the help button
    If bCreateNew Then
        sHTML = sHTML & HTMLToAddHelpButton(gk_sHelpLinkToCreateSFA)
    Else
        sHTML = sHTML & HTMLToAddHelpButton(gk_sHelpLinkToEditSFA)
    End If
    
    'write the html
    m_aspResponse.Write sHTML
    
    Set m_oBOSFALoader = Nothing
    
    'complete the mts transaction
    m_objObjectContext.SetComplete
    
Exit Sub

ErrorHandler:
    Set m_oBOSFALoader = Nothing
    
    m_objObjectContext.SetAbort
    ShowError sSource
End Sub

Private Function AddTable(ByVal t_bCreateNew As Boolean, ByVal t_eSFAType As ESFAFilter) As String
'   Description
'   Builds the html table to display the data for editing
'
'   d1  NDK    08Nov1999   Created
'   d10 ndk    27Dec1999   Changed to VI code use customizable styles
'

On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim sCancelLink As String
    Dim typPinkRow As RowStyle
    Dim typTableStyle As TableStyle, typCellStyle As CellStyle
    
    sSource = Me.Name & ".AddTable"
    
    'Set up row, cell and table styles
    typPinkRow = m_oFWHTML.GetRowStyle(RowBackgroundPink)
    
    typTableStyle = m_oFWHTML.GetTableStyle(TABContentStyle)
    typTableStyle.lCELLPADDING = klTablePadding
    typTableStyle.lCELLSPACING = klTableSpacing
    typTableStyle.sWidth = ksTableWidth
    typTableStyle.eALIGN = HALNone
    
    m_fwTable.DefaultRowStyle = typPinkRow
    m_fwTable.TableStyle = typTableStyle
    
    'Add the data to the table
    
    'Title
    AddTableRow m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfDisplayEditSFALeftHeaderTitle), _
        m_oFWHTML.InputField( _
        ITText, ksNameSFATitle, kiSFATitleWidth, gk_iDBCOLLenSFATitle, _
        VariantStringAsString(m_typSFADetails.sTitle), VCNone, True)

    'Sub Title
    AddTableRow m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfDisplayEditSFALeftHeaderSubTitle), _
        m_oFWHTML.InputField _
        (ITText, ksNameSFASubTitle, kiSFASubTitleWidth, gk_iDBCOLLenSFASubTitle, _
        VariantStringAsString(m_typSFADetails.sSubTitle))

    'Publisher
    AddTableRow m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfDisplayEditSFALeftHeaderPublisher), _
        m_oFWHTML.InputField _
        (ITText, ksNameSFAPublisher, kiSFAPublisherWidth, gk_iDBCOLLenSFAPublisher, VariantStringAsString(m_typSFADetails.sPublisher))

    'Author
    AddTableRow m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfDisplayEditSFALeftHeaderAuthor), _
        m_oFWHTML.InputField _
        (ITText, ksNameSFAAuthor, kiSFAAuthorWidth, gk_iDBCOLLenSFAAuthor, VariantStringAsString(m_typSFADetails.sAuthor))

    'Publication Date
    AddTableRow m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfDisplayEditSFALeftHeaderPublicationDate), _
        m_oFWHTML.InputField _
        (ITText, ksNameSFAPublicationDate, kiSFAPublicationDateWidth, gk_iDBCOLLenSFAPublicationDate, VariantStringAsString(m_typSFADetails.sPublicationDate))

    'Work out where the SFA is held
    AddTableRow m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfDisplayEditSFALeftHeaderLocation), _
        m_oFWHTML.RadioButton( _
            m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfDisplayEditSFALocationNamePROLibrary) _
            , ksNameSFALocation, m_typSFADetails.bInProLibaray, kiInProLibraryValue) & _
        m_oFWHTML.RadioButton( _
            m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfDisplayEditSFALocationNamePROReference) _
            , ksNameSFALocation, m_typSFADetails.bInProReference, kiInProReferenceValue) & _
        m_oFWHTML.RadioButton( _
            m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfDisplayEditSFALocationNameUnknown) _
            , ksNameSFALocation, _
            Not (m_typSFADetails.bInProLibaray Or m_typSFADetails.bInProReference), kiUnknownLocationValue)

    'external reference
    AddTableRow m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfDisplayEditSFALeftHeaderExternalReference), _
        m_oFWHTML.InputField _
        (ITText, ksNameSFAExternalReference, kiSFAExternalReferenceWidth, gk_iDBCOLLenSFAExternalReference, VariantStringAsString(m_typSFADetails.sExternalReference))

    'Edition
    AddTableRow m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfDisplayEditSFALeftHeaderEdition), _
        m_oFWHTML.InputField _
        (ITText, ksNameSFAEdition, kiSFAEditionWidth, gk_iDBCOLLenSFAEdition, VariantStringAsString(m_typSFADetails.sEdition))

    'Publication note
    AddTableRow m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfDisplayEditSFALeftHeaderPublicationNote), _
        m_oFWHTML.InputField _
        (ITText, ksNameSFAPublicationNote, kiSFAPublicationNoteWidth, gk_iDBCOLLenSFAPublicationNote, VariantStringAsString(m_typSFADetails.sPublicationNote))
    
    If t_bCreateNew Then    'when creating
        If t_eSFAType = ESFAFilter.ESFAElectronic Then    'electronic sfa
            AddTableRow m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfDisplayEditSFALeftHeaderFileName), _
                m_oFWHTML.InputField _
                (ITText, ksNameSFAElectronicFileName, kiSFAElectronicFileNameWidth, gk_iDBCOLLenSFAElectronicSFAFilename, "")
        End If
    Else
        'editing
        'If type is electronic then display the filename
        If m_typSFADetails.lTypeID = ESFAFilter.ESFAElectronic Then
            'electronic sfa
            AddTableRow "URL", _
                m_oFWHTML.InputField _
                (ITText, ksNameSFAElectronicFileName, kiSFAElectronicFileNameWidth, gk_iDBCOLLenSFAElectronicSFAFilename, _
                VariantStringAsString(m_typSFADetails.sElectronicSFAFilename))
        End If
    End If
    
    'link for cancel
    'redirect to appropriate page
    If t_bCreateNew Then  'new one was created go to the list
        If t_eSFAType = ESFAFilter.ESFAElectronic Then    'electronic
            sCancelLink = gk_sLinkDisplayElectronicSFAList
        Else
            sCancelLink = gk_sLinkDisplayPrintedSFAList
        End If
    Else
        'editing - go the view mode
        sCancelLink = _
            gk_sPAGESFAEdDetails & _
            "?" & gk_sFORMSFADetailsID & "=" & m_typSFADetails.lID & _
            "&" & gk_sParamSFAType & "=" & t_eSFAType
    End If
      
    'OK, cancel buttons
    typCellStyle = m_oFWHTML.GetCellStyle(CELLNormalStyle)
    typCellStyle.sWidth = ksLabelColumnWidth
    
    m_fwTable.AddCellAsHTML " ", typCellStyle
    
    typCellStyle.sWidth = "75%"
    m_fwTable.AddCellAsHTML _
        m_oFWHTML.SubmitImage(gk_sImagesPath & gk_sImageOK, ksNameButtonOk, _
            ksNameButtonOk) & _
        m_oFWHTML.ImageButton(gk_sImagesPath & gk_sImageCancel, ksNameButtonCancel, _
                                      sCancelLink, gk_sMainFrame), typCellStyle
    
    AddTable = m_fwTable.GetHTML
    
    Exit Function
ErrorHandler:
    ConvertError sSource
End Function

Private Sub AddTableRow(t_sLabel As String, t_sValue As String)
'   This method adds a row to the table
'
'   d1  NDK    08Nov1999   Created
'   d10 ndk    27Dec1999   Changed to VI code use customizable styles
'

On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim typLabelCell As CellStyle
    Dim typDataCell As CellStyle
    sSource = Me.Name & ".AddTableRow"
    
    typLabelCell = m_oFWHTML.GetCellStyle(CELLGeneralLabelStyle)
    typLabelCell.sWidth = ksLabelColumnWidth
    
    typDataCell = m_oFWHTML.GetCellStyle(CELLGeneralContentStyle)
    
    m_fwTable.StartRow
    m_fwTable.AddCellAsHTML m_oFWHTML.Font(t_sLabel, t_iSize:=2), typLabelCell
    m_fwTable.AddCellAsHTML m_oFWHTML.Font(t_sValue, t_iSize:=2), typDataCell
    m_fwTable.EndRow

    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub

Private Function HeaderHTML(ByVal t_sHeading As String) As String
'   This method displays the header for the page
'
'   d1  NDK    08Nov1999   Created


On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim sHTML As String
    Dim typTableStyle As TableStyle, typCellStyle As CellStyle
    
    sSource = Me.Name & ".HeaderHTML"
                
    m_fwTable.Reset
    typTableStyle.eStyle = TABPageHeadingStyle
    m_fwTable.TableStyle = typTableStyle
    
    typCellStyle.eStyle = CELLPageHeadingStyle
    m_fwTable.AddCellAsHTML ( _
                m_oFWHTML.FontWithStyle(t_sHeading, _
                    m_oFWHTML.GetFontStyle(FONTPageHeadingStyle))), typCellStyle
    
    m_fwTable.AddBlankRow
    
    HeaderHTML = sHTML
    
    Exit Function
ErrorHandler:
     ConvertError sSource
End Function


'
' Class initialisation and MTS event handlers
'
Private Sub Initialise()
    Set m_objObjectContext = GetObjectContext()
    
    ' Set asp objects
    Set m_aspResponse = m_objObjectContext.Item("Response")
    Set m_aspRequest = m_objObjectContext.Item("Request")
        
    'user id
    m_lLoggedUserID = GetDataFromSession(dsUserID, m_objObjectContext.Item(gk_sSessionObject))
    
    ' Specific initialisation
    Set m_oFWHTML = New FWHTML
    Set m_fwTable = New FWTable
    Set m_oFWAppData = New FWApplicationData
    Set m_oBOEditorialPermissions = New BOEditorialPermissions
    
End Sub
Private Sub Terminate()
    Set m_objObjectContext = Nothing
    
    ' Set asp objects
    Set m_aspResponse = Nothing
    Set m_aspRequest = Nothing
    
    ' Specific initialisation
    Set m_oFWHTML = Nothing
    Set m_fwTable = Nothing
    Set m_oFWAppData = Nothing
    Set m_oBOEditorialPermissions = Nothing
    Set m_oBOSFALoader = Nothing
End Sub

Private Sub ObjectControl_Activate()
    Initialise
End Sub

Private Sub ObjectControl_Deactivate()
    Set m_objObjectContext = Nothing
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = True
End Function



Public Sub DisplayElectronicSFAList()
'   Description
'
'   Calls the method to display the list of Electronic SFAs
'
'   d1  NDK    08Nov1999   Created
'   d9  NDK    21Dec1999   list not displayed if no permissions to maintain

On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim sParamString As String
    Dim eSecurityCheckStatus As ESecurityCheckResult

    sSource = Me.Name & ".DisplayElectronicSFAList"
    
    'perform security check
    m_oBOEditorialPermissions.UserID = m_lLoggedUserID
    
    eSecurityCheckStatus = m_oBOEditorialPermissions.Check(EUserCanMaintainElectronicSFA)
    
    If eSecurityCheckStatus <> ESecurityCheckResult.ESRSecurityCheckSuccessful Then
        RaiseError sSource, _
               GetErrorEnumForSecurityCheckResult(eSecurityCheckStatus), _
               GetDescriptionForSecurityCheckError( _
                        GetErrorEnumForSecurityCheckResult(eSecurityCheckStatus))
    End If
    
    If m_aspRequest.QueryString.Count = 0 Then      'called from mainmene
        sParamString = _
            gk_sLinkSFABrowserForMaintenance & _
            "?" & gk_sFORMSFAListStartIndex & "=0" & _
            "&" & gk_sFORMSFAListUsage & "=" & ESFABrowserUsage.ESFAEditorialMaintain & _
            "&" & gk_sFORMSFAListFilter & "=" & ESFAFilter.ESFAElectronic & _
            "&" & gk_sParamAddButtonsForEditorialMaintain & "=" & True
    Else
        sParamString = _
            gk_sLinkSFABrowserForMaintenance & "?" & _
            m_aspRequest.QueryString
    End If
    
    m_aspResponse.Redirect sParamString
    
    Exit Sub
ErrorHandler:
    m_objObjectContext.SetAbort
    ShowError sSource
End Sub
Public Sub DisplayPrintedSFAList()
'   Description
'
'   Calls the method to display the list of printed SFAs
'
'   d1  NDK    08Nov1999   Created
'   d9  NDK    21Dec1999   list not displayed if no permissions to maintain
'
On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim sParamString As String
    Dim eSecurityCheckStatus As ESecurityCheckResult
    
    sSource = Me.Name & ".DisplayPrintedSFAList"
    
    'perform security check
    m_oBOEditorialPermissions.UserID = m_lLoggedUserID
    
    eSecurityCheckStatus = m_oBOEditorialPermissions.Check(EUserCanMaintainPrintedSFA)
    
    If eSecurityCheckStatus <> ESecurityCheckResult.ESRSecurityCheckSuccessful Then
        RaiseError sSource, _
               GetErrorEnumForSecurityCheckResult(eSecurityCheckStatus), _
               GetDescriptionForSecurityCheckError( _
                        GetErrorEnumForSecurityCheckResult(eSecurityCheckStatus))
    End If
    
    If m_aspRequest.QueryString.Count = 0 Then      'called from mainmene
        sParamString = _
            gk_sLinkSFABrowserForMaintenance & _
            "?" & gk_sFORMSFAListStartIndex & "=0" & _
            "&" & gk_sFORMSFAListUsage & "=" & ESFABrowserUsage.ESFAEditorialMaintain & _
            "&" & gk_sFORMSFAListFilter & "=" & ESFAFilter.ESFAPrinted & _
            "&" & gk_sParamAddButtonsForEditorialMaintain & "=" & True
    Else
        sParamString = _
            gk_sLinkSFABrowserForMaintenance & "?" & _
            m_aspRequest.QueryString
    End If
    
    m_aspResponse.Redirect sParamString
    
    Exit Sub
ErrorHandler:
    m_objObjectContext.SetAbort
    ShowError sSource
End Sub



Public Sub DisplaySFAForViewWithEditButton()
'   Description
'
'   Calls the method in shared code to display the details of the sfa
'   whose id is passed as query string
'
'   After that it displays one hyperlink to edit the sfa another the delete the sfa
'
'   d1  NDK    08Nov1999   Created
'   d4  NDK    18Nov1999   Changes made to display
'   d6  NDK    23Nov1999   Done button displayed to link back to the browser
'   d8  NDK    08Dec1999   Added help buttons
'   d10 ndk    27Dec1999   Changed to VI code use customizable styles
'   d11 ndk    28Dec1999   Added tooltips
'
On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim sHTML As String, sParamString As String, typTableStyle As TableStyle
    Dim typCellStyle As CellStyle
    Dim lSFAReferenceID As Long, eSFAType  As ESFAFilter, sSFATitle As String
    Dim oVWSharedSFA As VWSharedSFA, eSecurityResult As ESecurityCheckResult
    
    sSource = Me.Name & ".DisplaySFAForViewWithEditButton"
    
    'initialise the object
    Set oVWSharedSFA = New VWSharedSFA
    
    'get the id to be displayed
    lSFAReferenceID = m_aspRequest.QueryString(gk_sFORMSFADetailsID)
    eSFAType = m_aspRequest.QueryString(gk_sParamSFAType)
    sSFATitle = m_aspRequest.QueryString(gk_sParamSFATitle)
    'reset the table
    m_fwTable.Reset
    
    'set the table style
    typTableStyle = m_oFWHTML.GetTableStyle(TABContentStyle)
    typTableStyle.sWidth = "100%"
    typTableStyle.eALIGN = HALNone

    'table style
    m_fwTable.TableStyle = typTableStyle
    
    'Display the details of the sfa
    oVWSharedSFA.DisplayDetails lSFAReferenceID, ESFABrowserUsage.ESFAEditorialMaintain
    
    'check permissions
    m_oBOEditorialPermissions.UserID = m_lLoggedUserID
    If eSFAType = ESFAFilter.ESFAElectronic Then   'electronic
        eSecurityResult = m_oBOEditorialPermissions.Check(EUserCanMaintainElectronicSFA)
    Else
        eSecurityResult = m_oBOEditorialPermissions.Check(EUserCanMaintainPrintedSFA)
    End If
    
    If eSecurityResult = ESecurityCheckResult.ESRSecurityCheckSuccessful Then
        'Link for edit
        sParamString = gk_sLinkDisplayEditSFADetailsScreen & _
            "?" & gk_sParamSFAReferenceID & "=" & lSFAReferenceID & _
            "&" & gk_sParamSFAType & "=" & eSFAType & _
            "&" & gk_sParamSFATitle & "=" & sSFATitle

    Else    'no link
        sParamString = ""
    End If
    
    m_fwTable.AddBlankRow
 
    typCellStyle = m_oFWHTML.GetCellStyle(CELLNormalStyle)
    typCellStyle.sWidth = "25%"
        
    m_fwTable.AddCellAsHTML "", typCellStyle
    
    'edit link
    m_fwTable.AddCellAsHTML _
        m_oFWHTML.Hyperlink(sParamString, _
        m_oFWHTML.Image("", gk_sImagesPath & gk_sImageEdit), gk_sMainFrame, , GetToolTip(ETOOLTIP.ttEditSFADetails)), typCellStyle
        
    If eSecurityResult = ESecurityCheckResult.ESRSecurityCheckSuccessful Then
        'Link for delete
        sParamString = gk_sLinkDisplayConfirmationForSFADeletion & _
            "?" & gk_sParamSFAReferenceID & "=" & lSFAReferenceID & _
            "&" & gk_sParamSFAType & "=" & eSFAType & _
            "&" & gk_sParamSFATitle & "=" & sSFATitle
    Else
        'no link
        sParamString = ""
    End If
    
    m_fwTable.AddCellAsHTML _
        m_oFWHTML.Hyperlink(sParamString, _
        m_oFWHTML.Image("", gk_sImagesPath & gk_sImageRemove), gk_sMainFrame, , GetToolTip(ETOOLTIP.ttDeleteSFA)), typCellStyle
    
    'for the donebutton
    If eSFAType = ESFAElectronic Then
        sParamString = gk_sLinkDisplayElectronicSFAList
    Else
        sParamString = gk_sLinkDisplayPrintedSFAList
    End If
    
    m_fwTable.AddCellAsHTML _
        m_oFWHTML.Hyperlink(sParamString, _
        m_oFWHTML.Image("", gk_sImagesPath & gk_sImageDone), gk_sMainFrame, , GetToolTip(ETOOLTIP.ttBackToSFAList)), typCellStyle
        
    sHTML = sHTML & m_fwTable.GetHTML
        
    'Add the help button
    sHTML = sHTML & HTMLToAddHelpButton(gk_sHelpLinkToViewEditorialSFA)
        
    'display the hyperlinks
    m_aspResponse.Write sHTML

    Set oVWSharedSFA = Nothing
    
    'complete the mts transaction
    m_objObjectContext.SetComplete
    
    Exit Sub
ErrorHandler:
    Set oVWSharedSFA = Nothing
    
    m_objObjectContext.SetAbort
    ShowError sSource
End Sub

Public Sub CreateUpdate()
'   Description
'
'   Calls the method in bo object to create a new sfa
'
'   d1  NDK    08Nov1999   Created
'   d9  NDK    21Dec1999   Added security check
'

    On Error GoTo ErrorHandler
    
    Dim sSource As String, sParamString As String
    Dim lSFAReferenceID As Long, bCreateNew As Boolean, eSFAType As ESFAFilter
    Dim oboSFA As BOSFA
    Dim eSecurityCheckStatus As ESecurityCheckResult
    
    sSource = Me.Name & ".CreateEdit"
    
    'get the values from the query string
    bCreateNew = m_aspRequest.QueryString(gk_sParamCreateNewSFA)
    eSFAType = m_aspRequest.QueryString(gk_sParamSFAType)
    lSFAReferenceID = m_aspRequest.QueryString(gk_sParamSFAReferenceID)
    
    'perform security check
    m_oBOEditorialPermissions.UserID = m_lLoggedUserID
    
    If eSFAType = ESFAFilter.ESFAElectronic Then
        eSecurityCheckStatus = m_oBOEditorialPermissions.Check(EUserCanMaintainElectronicSFA)
    Else
        eSecurityCheckStatus = m_oBOEditorialPermissions.Check(EUserCanMaintainPrintedSFA)
    End If
    
    If eSecurityCheckStatus <> ESecurityCheckResult.ESRSecurityCheckSuccessful Then
        RaiseError sSource, _
               GetErrorEnumForSecurityCheckResult(eSecurityCheckStatus), _
               GetDescriptionForSecurityCheckError( _
                        GetErrorEnumForSecurityCheckResult(eSecurityCheckStatus))
    End If
    
    'initialise the object
    Set oboSFA = New BOSFA
    
    'set the properties
    m_typSFADetails.sTitle = StringToNullOrString(m_aspRequest.Form(ksNameSFATitle))
    m_typSFADetails.sSubTitle = StringToNullOrString(m_aspRequest.Form(ksNameSFASubTitle))
    m_typSFADetails.sPublisher = StringToNullOrString(m_aspRequest.Form(ksNameSFAPublisher))
    m_typSFADetails.sAuthor = StringToNullOrString(m_aspRequest.Form(ksNameSFAAuthor))
    m_typSFADetails.sPublicationDate = StringToNullOrString(m_aspRequest.Form(ksNameSFAPublicationDate))
    m_typSFADetails.sExternalReference = StringToNullOrString(m_aspRequest.Form(ksNameSFAExternalReference))
    m_typSFADetails.sElectronicSFAFilename = StringToNullOrString(m_aspRequest.Form(ksNameSFAElectronicFileName))
    m_typSFADetails.sEdition = StringToNullOrString(m_aspRequest.Form(ksNameSFAEdition))
    m_typSFADetails.sPublicationNote = StringToNullOrString(m_aspRequest.Form(ksNameSFAPublicationNote))
    
    Select Case m_aspRequest.Form(ksNameSFALocation)
        Case kiInProLibraryValue
            m_typSFADetails.bInProLibaray = True
        Case kiInProReferenceValue
            m_typSFADetails.bInProReference = True
        Case Else
            m_typSFADetails.bInProLibaray = False
            m_typSFADetails.bInProReference = False
    End Select
    
    'call the method to save or update the sfa
    If bCreateNew Then  'Creating a new sfa
        m_typSFADetails.lTypeID = eSFAType
        oboSFA.Data = m_typSFADetails
        oboSFA.Create
    Else
        m_typSFADetails.lID = lSFAReferenceID
        oboSFA.Data = m_typSFADetails
        oboSFA.Update
    End If
    
    'redirect to appropriate page
    If bCreateNew Then  'new one was created go to the list
        If eSFAType = ESFAFilter.ESFAElectronic Then    'electronic
            sParamString = gk_sLinkDisplayElectronicSFAList
        Else
            sParamString = gk_sLinkDisplayPrintedSFAList
        End If
    Else
        'editing - go the view mode
        sParamString = _
            gk_sPAGESFAEdDetails & "?" & gk_sFORMSFADetailsID & "=" & lSFAReferenceID & _
            "&" & gk_sParamSFAType & "=" & eSFAType
            
    End If
        
    m_aspResponse.Redirect sParamString
    
    Set oboSFA = Nothing
    
    'complete the mts transaction
    m_objObjectContext.SetComplete
    
    Exit Sub
ErrorHandler:
    
    Set oboSFA = Nothing

    m_objObjectContext.SetAbort
    ShowError sSource
End Sub
Public Sub Delete()
'   Description
'
'   Calls the method in bo to delete the selected SFA
'
'   d1  NDK    08Nov1999   Created
'   d9  NDK    21Dec1999   Added security check
'
On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim sParamString As String
    Dim lSFAReferenceID As Long, eSFAType As ESFAFilter
    Dim oboSFA As BOSFA
    Dim eSecurityCheckStatus As ESecurityCheckResult
    
    sSource = Me.Name & ".Delete"
   
    'get data from query string
    lSFAReferenceID = m_aspRequest.QueryString(gk_sParamSFAReferenceID)
    eSFAType = m_aspRequest.QueryString(gk_sParamSFAType)
        
    'perform security check
    m_oBOEditorialPermissions.UserID = m_lLoggedUserID
    
    If eSFAType = ESFAFilter.ESFAElectronic Then
        eSecurityCheckStatus = m_oBOEditorialPermissions.Check(EUserCanMaintainElectronicSFA)
    Else
        eSecurityCheckStatus = m_oBOEditorialPermissions.Check(EUserCanMaintainPrintedSFA)
    End If
    
    If eSecurityCheckStatus <> ESecurityCheckResult.ESRSecurityCheckSuccessful Then
        RaiseError sSource, _
               GetErrorEnumForSecurityCheckResult(eSecurityCheckStatus), _
               GetDescriptionForSecurityCheckError( _
                        GetErrorEnumForSecurityCheckResult(eSecurityCheckStatus))
    End If
    
    'initialise object
    Set oboSFA = New BOSFA
    
    'call the method to delete the sfa
    oboSFA.Delete lSFAReferenceID
    
    'to redirect to appropriate page
    If eSFAType = ESFAFilter.ESFAElectronic Then    'electronic
        sParamString = gk_sLinkDisplayElectronicSFAList
    Else
        sParamString = gk_sLinkDisplayPrintedSFAList
    End If
    
    'redirect
    m_aspResponse.Redirect sParamString
    
    Set oboSFA = Nothing
    'complete the mts transaction
    m_objObjectContext.SetComplete
    
    Exit Sub
ErrorHandler:

    Set oboSFA = Nothing
    
    m_objObjectContext.SetAbort
    ShowError sSource
End Sub

Public Sub DisplayConfirmationForSFADeletion()
'
'Description
'
'Displays the the confirmation dialog box when the user selects to delete a SFA
'
'Amendment history
'   d1  NDK    08Nov1999   Created
'   d8  NDK    08Dec1999   Added help buttons
'   d9  NDK    21Dec1999   Added security check
'   d10 ndk    27Dec1999   Changed to VI code use customizable styles
'   d14 crm    10Feb2000   Using BOSFAs HasReference method to display message
'                          for existing ref. and show only OK button which takes
'                          user back to details page
    On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim sHTML As String, sOKLink As String, sCancelLink As String
    Dim lSFAReferenceID As Long, eSFAType As ESFAFilter, sSFATitle As String
    Dim eSecurityCheckStatus As ESecurityCheckResult
      
    Dim ofwTable As FWTable
    Dim typCellStyle As CellStyle
    Dim typTableStyle As TableStyle
    Dim typFontStyle As FontStyle
    Dim oboSFA As New BOSFA
    Dim bHasReference As Boolean
    sSource = Me.Name & ".DisplayConfirmationForSFADeletion"
    
    'Read details from the query string
    
    lSFAReferenceID = m_aspRequest.QueryString(gk_sParamSFAReferenceID)
    eSFAType = m_aspRequest.QueryString(gk_sParamSFAType)
    sSFATitle = m_aspRequest.QueryString(gk_sParamSFATitle)
    
    'perform security check
    m_oBOEditorialPermissions.UserID = m_lLoggedUserID
    
    If eSFAType = ESFAFilter.ESFAElectronic Then
        eSecurityCheckStatus = m_oBOEditorialPermissions.Check(EUserCanMaintainElectronicSFA)
    Else
        eSecurityCheckStatus = m_oBOEditorialPermissions.Check(EUserCanMaintainPrintedSFA)
    End If
    
    If eSecurityCheckStatus <> ESecurityCheckResult.ESRSecurityCheckSuccessful Then
        RaiseError sSource, _
               GetErrorEnumForSecurityCheckResult(eSecurityCheckStatus), _
               GetDescriptionForSecurityCheckError( _
                        GetErrorEnumForSecurityCheckResult(eSecurityCheckStatus))
    End If
    
    'Initialise the objects
    Set ofwTable = New FWTable
    
    'Table style for heading
    typTableStyle = m_oFWHTML.GetTableStyle(TABContentStyle)
    typTableStyle.lCELLSPACING = 2
    typTableStyle.lCELLPADDING = 0
    typTableStyle.eALIGN = HALCenter
    
    ofwTable.TableStyle = typTableStyle
    
    'Cell and font style for the heading
    typCellStyle = m_oFWHTML.GetCellStyle(CELLNormalStyle)
    typFontStyle = m_oFWHTML.GetFontStyle(FONTNormalStyle)
    typFontStyle.bIsBold = True
    
    ofwTable.StartRow
    
    'Heading
    typCellStyle.eALIGN = HALCenter
    typCellStyle.sWidth = "50%"
    typCellStyle.eALIGN = HALCenter
    '   d14 crm    10Feb2000
    bHasReference = oboSFA.SFAHasReference(lSFAReferenceID)
    If bHasReference Then
    ofwTable.AddCellAsHTML _
        m_oFWHTML.FontWithStyle(sSFATitle & " SFA has references to one or more catalogue entries, can not delete.", typFontStyle), typCellStyle
    Else
    'Delete catalogue message
    ofwTable.AddCellAsHTML _
        m_oFWHTML.FontWithStyle( _
        m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfDeleteSFAMessageDeleteSFA) & _
        sSFATitle, typFontStyle), typCellStyle
    End If
    
    ofwTable.EndRow
    
    ofwTable.AddBlankRow
    
    ofwTable.StartRow
    'ok link
    sOKLink = _
        gk_sLinkDeleteSFA & _
        "?" & gk_sParamSFAReferenceID & "=" & lSFAReferenceID & _
        "&" & gk_sParamSFAType & "=" & eSFAType
    
    'cancel link
    sCancelLink = gk_sPAGESFAEdDetails & _
        "?" & gk_sFORMSFADetailsID & "=" & lSFAReferenceID & _
        "&" & gk_sParamSFAType & "=" & eSFAType
    
    '   d14 crm    10Feb2000
    'if there is reference then display OK button with the
    'cancel link so that we go back to details page
    If bHasReference Then
     ofwTable.AddCellAsHTML m_oFWHTML.ImageButton(gk_sImagesPath & gk_sImageOK, ksNameButtonOk, _
                        sCancelLink, gk_sMainFrame), typCellStyle
    Else
    'Display Ok and Cancel Button
    ofwTable.AddCellAsHTML _
        m_oFWHTML.ImageButton(gk_sImagesPath & gk_sImageOK, ksNameButtonOk, _
                    sOKLink, gk_sMainFrame) & _
        m_oFWHTML.ImageButton(gk_sImagesPath & gk_sImageCancel, ksNameButtonCancel, _
                    sCancelLink, gk_sMainFrame), typCellStyle
    End If
    ofwTable.EndRow
    
    'Get the table details
    sHTML = sHTML & m_oFWHTML.Paragraph(ofwTable.GetHTML)
    
    'Add the help button
    sHTML = sHTML & HTMLToAddHelpButton(gk_sHelpLinkToDeleteSFA)
    
    'Display the html
    m_aspResponse.Write sHTML
    
    Set ofwTable = Nothing
    
    m_objObjectContext.SetComplete
    
    Exit Sub
    
ErrorHandler:

    Set ofwTable = Nothing
    
    m_objObjectContext.SetAbort
    ShowError sSource
End Sub


