VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "VWAccession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    VWAccession.cls
' System:       PRO, PROCAT
' Copyright:    (C) Quidnunc Limited
'
' Description:  View calss that provides interface to accession of records into
'           PROCAT.
'
' Amendment history:
'   d1  KSK     11Nov1999   Created
'   d2  KSK     12Nov1999   Moved to the Editorial project
'   d3  KSK     24Nov1999   Added a cancel button on the EAD file display list
'   d4  KSK     25Nov1999   Added code the move the display of log to the end of file
'   d5  KSK     10Dec1999   Changes after customer feedback to add DRUID status flags
'   d6  SPB     15Dec1999   Added EAD parse only link and function to display parse log file.
'   d7  SPB     17Dec1999   Added reload button on EAD parse log viewer.
'   d8  CRM     12Jan2000   Added reload button to DisplayLogFileDetails()
'                           Added font style to DisplayBatchListHeading()
'                           Minor correction to heading
'   d9  DAJ     28Feb2000   Added function to use Batch ID for parse log
'   d10 DAJ     28Feb2000   Added parse log column in batch list
'   d11 DAJ     28Feb2000   Added redirection for batch ID log files
'   d12 DAJ     29Feb2000   Added date column to batch list, and merged log columns
'   d13 DAJ     06Mar2000   Added Editset creation failed status to the list of hyperlink
'                           enabled batches
'   d14 DAJ     07Mar2000   Updated DisplayLogFile functions to check the return path
'   d15 JCC     14Mar2000   Spelling mistake - perperation
'   d16 DAJ     21Mar2000   Speed ups to DisplayBatchListContent
'   d17 DAJ     14Jul2000   Added ability to start updates
'   d18 MAP     22Oct2001   Added code to handle EAD and Druid scheduling & automatic laod into Edit Set
'   d19 MAP     22Oct2001   Added extra case to display DRUID load hyperlink for
'   d20 ALM     27Nov2001   Modified scheduling heading
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Implements ObjectControl
Private m_oboAccession As BOAccessionBatch
Private m_objContext As Object
Private m_aspResponse As Response
Private m_aspRequest As Request
Private m_aspServer As Server
Private m_aspSession As Session

Private m_ofwhtmlReference As FWHTML
Private m_oFWTable As FWTable
Private m_oFWAppData As FWApplicationData
Private m_oFWReferenceData As FWReferenceData
Private m_oBOEditorialPermissions As BOEditorialPermissions
Private m_lLoggedUserID As Long

Private Const mk_sEADFileSelectionWildCatString As String = "*.XML"
Private Const mk_sNameButtonOk As String = "btnOK"
Private Const mk_sNameButtonCancel As String = "btnCAN"
Private Const mk_sNameButtonDone As String = "btnDONE"
Private Const mk_sNameButtonReload As String = "btnRELOAD"
Private Const mk_sHyperLink As String = "HyperLink"



''''''''''''''''' Implementation ''''''''''''''''''''''
'
' Class initialisation and MTS event handlers
'

Private Sub Initialise()
    Dim sSource As String
    
    sSource = Me.Name & ".Initialise"
    
    Set m_objContext = GetObjectContext()
    
    ' Set asp objects
    Set m_aspResponse = m_objContext.Item("Response")
    Set m_aspRequest = m_objContext.Item("Request")
    Set m_aspServer = m_objContext.Item("Server")
    Set m_aspSession = m_objContext.Item("Session")
    
    Set m_oboAccession = New BOAccessionBatch
    Set m_ofwhtmlReference = New FWHTML
    Set m_oFWTable = New FWTable
    Set m_oFWAppData = New FWApplicationData
    Set m_oFWReferenceData = New FWReferenceData
    Set m_oBOEditorialPermissions = New BOEditorialPermissions
    
    'Get the users details
    m_lLoggedUserID = GetDataFromSession(dsUserID, m_objContext.Item(gk_sSessionObject))
    If m_lLoggedUserID = 0 Then     'Not a valid user
        RaiseError sSource, EPROCATINTERNALERROR.peDBInvalidUser, "Invalid User ID : " & m_lLoggedUserID
    End If
    
End Sub

Private Sub Terminate()
    
    Set m_objContext = Nothing
    Set m_aspResponse = Nothing
    Set m_aspRequest = Nothing
    Set m_aspServer = Nothing
    Set m_aspSession = Nothing
    
    Set m_oboAccession = Nothing
    Set m_ofwhtmlReference = Nothing
    Set m_oFWTable = Nothing
    Set m_oFWAppData = Nothing
    Set m_oFWReferenceData = Nothing
    Set m_oBOEditorialPermissions = Nothing

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

Public Property Get Name() As String
        Name = "VWAccession"
End Property



Public Sub DisplayPageToListEADFiles()
' Displays a list of XML files in the EAD source directory, so the user
' can click on them to start an accession or update.
'
'   d3  KSK     24Nov1999   Added a cancel button on the EAD file display list
'   d17 DAJ     14Jul2000   Added ability to start updates
'   d18 MAP     11Oct2001   Changed link so that after selecting a EAD file, we check if a
'                           scheduled job exists.
On Error GoTo Err_Handler
    
    Dim sSource As String
    Dim sEADDirectory As String, sFileNames As String, sParam As String
    Dim sHTML As String, sCancelLink As String, sTempHTML  As String
    Dim eResult As ESecurityCheckResult
    Dim bUpdate As Boolean
    
    sSource = Me.Name & ".DisplayPageToListEADFiles"
    
    m_oBOEditorialPermissions.UserID = m_lLoggedUserID
    
    ' Get the update parameter
    bUpdate = False
    If m_aspRequest.QueryString(gk_sParamUpdateEADList) Then
        bUpdate = True
    End If
    
    ' Security check - different for import and update
    If bUpdate Then
        ' Check update permissions
        eResult = m_oBOEditorialPermissions.Check(ESecurityCheck.EUserCanUpdateFromEAD)
        If eResult <> ESecurityCheckResult.ESRSecurityCheckSuccessful Then
            RaiseError sSource, _
                   GetErrorEnumForSecurityCheckResult(eResult), _
                   GetDescriptionForSecurityCheckError( _
                            GetErrorEnumForSecurityCheckResult(eResult))
        End If
    Else
        ' Check import permissions
        eResult = m_oBOEditorialPermissions.Check(ESecurityCheck.EUserCanImportFromEAD)
        If eResult <> ESecurityCheckResult.ESRSecurityCheckSuccessful Then
            RaiseError sSource, _
                   GetErrorEnumForSecurityCheckResult(eResult), _
                   GetDescriptionForSecurityCheckError( _
                            GetErrorEnumForSecurityCheckResult(eResult))
        End If
    End If
    
    ' Get the directory the EAD files are stored in
    sEADDirectory = m_oFWAppData.GetApplicationConfigurationInfomation( _
                            ECONFIGURATIONDATA.pEADAccessionDirectory)
    
    ' d17 - Decide which instructions to show
    If bUpdate Then
        sHTML = m_ofwhtmlReference.Paragraph(m_ofwhtmlReference.Italic( _
                "Select the EAD file you want to use to update the catalogue"))
    Else
        sHTML = m_ofwhtmlReference.Paragraph(m_ofwhtmlReference.Italic( _
                "Select the EAD file you want to accession into the catalogue"))
    End If
    
    sFileNames = Dir(sEADDirectory, vbNormal)
    
    sEADDirectory = IIf(Right(sEADDirectory, 1) <> "\", _
                                sEADDirectory & "\", _
                                sEADDirectory) & _
                    mk_sEADFileSelectionWildCatString
                    
    sFileNames = Dir(sEADDirectory, vbNormal)
        
    'Table with links to load ead file or parse only
    Dim ofwHTML As New FWHTML
    Dim ofwTable As New FWTable
    Dim typTableStyle As TableStyle
    
    typTableStyle = ofwHTML.GetTableStyle(TABPlainStyle)
    typTableStyle.lCELLPADDING = 4
    ofwTable.TableStyle = typTableStyle
    ofwTable.AddBlankRow

    Do While sFileNames <> ""  ' Start the loop.
        ' d18 - Altered sParam
        sParam = "?" & gk_sParamFileName & "=" & EncodeQueryStringValue(sFileNames)
        If bUpdate Then
            sParam = sParam & "&" & gk_sParamJobType & "=" & jtEADUpdate
        Else
            sParam = sParam & "&" & gk_sParamJobType & "=" & jtEADImport
        End If
            
        
        ' d17 - Use different links for update
        ' d18 - Altered page to link to
        If bUpdate Then
            'sParam = sParam & "&" & gk_sParamUpdateEADList
            sFileNames = m_ofwhtmlReference.Hyperlink(gk_sLinkCheckForScheduledJob & sParam, _
                                            sFileNames, gk_sMainFrame, gk_sParamFileName, _
                                            "Click to select file for update")
        Else
            sFileNames = m_ofwhtmlReference.Hyperlink(gk_sLinkCheckForScheduledJob & sParam, _
                                            sFileNames, gk_sMainFrame, gk_sParamFileName, _
                                            "Click to select file for import")
        End If
        
        ' Link to parse only
        Dim sParseOnly As String
        If bUpdate Then
            sParseOnly = m_ofwhtmlReference.Hyperlink(gk_sLinkParseOnly & sParam & "&" & gk_sParamUpdateEADList & "=1", _
                                                "Parse file", gk_sMainFrame, gk_sParamFileName, _
                                                "Click to select file for parsing only")
        Else
            sParseOnly = m_ofwhtmlReference.Hyperlink(gk_sLinkParseOnly & sParam, _
                                                "Parse file", gk_sMainFrame, gk_sParamFileName, _
                                                "Click to select file for parsing only")
        End If
        
        ofwTable.StartRow
        ofwTable.AddCellAsHTML sFileNames
        ofwTable.AddCellAsHTML sParseOnly
        ofwTable.EndRow
        
        sFileNames = Dir   ' Get next entry.
    Loop
        
    ofwTable.AddBlankRow
    
    sCancelLink = gk_sLinkDisplayLoadHome


    sTempHTML = m_ofwhtmlReference.Image(mk_sNameButtonCancel, _
                    gk_sImagesPath & gk_sImageCancel, _
                    m_oFWReferenceData.Search(EGroup.gToolTip, ETOOLTIP.ttButtonCancelHelp))
    
    ofwTable.StartRow
    ofwTable.AddCellAsHTML m_ofwhtmlReference.Hyperlink(gk_sLinkDisplayLoadHome, _
                 sTempHTML, gk_sMainFrame, , m_oFWReferenceData.Search(EGroup.gToolTip, ETOOLTIP.ttButtonCancelHelp))
    ofwTable.EndRow
    
    sHTML = sHTML & ofwTable.GetHTML
    Set ofwTable = Nothing
    m_aspResponse.Write sHTML
    
    
    
    m_objContext.SetComplete
    Exit Sub
    
Err_Handler:
    m_objContext.SetComplete
    ShowError sSource
End Sub



Public Sub DisplayBatchDescriptionEntry()
'method to display conformation screen for the Accession functionality
'
' d18    M Phanides  08-Oct-2001     Added Scheduling functionality
' d20    A Manic     27-Nov-2001     Modified Scheduling heading
On Error GoTo ErrorHandler
    
    Dim sHTML As String, sSource As String, sParam As String
    Dim sTempHTML As String
    Dim sOKLink  As String, sCancelLink  As String
    Dim sFilename As String
    Dim eResult As ESecurityCheckResult
    Dim typCell As CellStyle
    Dim typTable As TableStyle
    
    sSource = Me.Name & ".DisplayBatchDescriptionEntry()"
    
    
    m_oBOEditorialPermissions.UserID = m_lLoggedUserID
    
    eResult = m_oBOEditorialPermissions.Check(ESecurityCheck.EUserCanImportFromEAD)
    If eResult <> ESecurityCheckResult.ESRSecurityCheckSuccessful Then
        RaiseError sSource, _
               GetErrorEnumForSecurityCheckResult(eResult), _
               GetDescriptionForSecurityCheckError( _
                        GetErrorEnumForSecurityCheckResult(eResult))
    End If
    
    
    'Get editset id from the query string
    sFilename = m_aspRequest.QueryString(gk_sParamFileName)
    If Trim(sFilename) = "" Then
        RaiseError sSource, EPROCATINTERNALERROR.PEGeneralInternalError, _
                "File name not found"
    End If
    sParam = "?" & gk_sParamFileName & "=" & sFilename
    
    
    Set m_oFWTable = New FWTable
    
    typTable = m_ofwhtmlReference.GetTableStyle(ETableStyle.TABContentStyle)
    typTable.sWidth = "100%"
    m_oFWTable.TableStyle = typTable
    
    'Display the title to enter the batch description
    m_oFWTable.StartRow
    typCell = m_ofwhtmlReference.GetCellStyle(ECellStyle.CELLHelpTextStyle)
    typCell.lCOLSPAN = 2
    m_oFWTable.AddCellAsHTML "Batch description for importing from EAD file " & _
                    sFilename, typCell
    m_oFWTable.EndRow
    
    'Display the batch description field
    m_oFWTable.StartRow
    
    typCell = m_ofwhtmlReference.GetCellStyle(ECellStyle.CELLReadWriteLabelStyle)
    typCell.sWidth = "25%"
    
    m_oFWTable.AddCellAsHTML "Batch Description", typCell
    
    typCell.sWidth = "75%"
    typCell = m_ofwhtmlReference.GetCellStyle(ECellStyle.CELLReadWriteDataStyle)
    m_oFWTable.AddCellAsHTML m_ofwhtmlReference.TextField(gk_sParamBatchDescription, _
                            gk_iDBCOLLenBatchDescription, , 50), typCell
    m_oFWTable.EndRow
    
    ' d18 Display the scheduling info text description
    
    m_oFWTable.AddBlankRow
    
    m_oFWTable.StartRow
    typCell = m_ofwhtmlReference.GetCellStyle(ECellStyle.CELLHelpTextStyle)
    typCell.lCOLSPAN = 2
    
    'd20 Altered the heading
    m_oFWTable.AddCellAsHTML "Complete below field(s) if scheduling. The time must be between " & _
    m_oFWAppData.GetApplicationConfigurationInfomation(pScheduledJobTimeFrame) & _
    ". The date is optional. Please check the list of currently scheduled jobs and select a time that does not conflict with an existing job", typCell
    m_oFWTable.EndRow
    
    ' d18 Display the time fields
    m_oFWTable.StartRow
    
    typCell = m_ofwhtmlReference.GetCellStyle(ECellStyle.CELLReadWriteLabelStyle)
    
    m_oFWTable.AddCellAsHTML "Start time (hh:mm 24hr clock):", typCell
    
    typCell = m_ofwhtmlReference.GetCellStyle(ECellStyle.CELLReadWriteDataStyle)
    m_oFWTable.AddCellAsHTML m_ofwhtmlReference.TextField(gk_sParamJobStartTime, _
                                gk_iDBCOLLenJobStartTime, , 10), typCell
    m_oFWTable.EndRow
                                
    ' d18 Display the date fields
    m_oFWTable.StartRow
    typCell = m_ofwhtmlReference.GetCellStyle(ECellStyle.CELLReadWriteLabelStyle)
    
    m_oFWTable.AddCellAsHTML "Date (dd/mm/yyyy)", typCell
    
    typCell = m_ofwhtmlReference.GetCellStyle(ECellStyle.CELLReadWriteDataStyle)
    m_oFWTable.AddCellAsHTML m_ofwhtmlReference.TextField(gk_sParamJobStartDate, _
                                gk_iDBCOLLenJobStartDate, , 10), typCell
    
    m_oFWTable.EndRow
    m_oFWTable.AddBlankRow
    
    'Display the buttons
    sCancelLink = _
        gk_sLinkDisplayLoadHome


    sTempHTML = m_ofwhtmlReference.Image(mk_sNameButtonCancel, _
                    gk_sImagesPath & gk_sImageCancel, _
                    m_oFWReferenceData.Search(EGroup.gToolTip, ETOOLTIP.ttButtonCancelHelp))
                    
    sHTML = m_ofwhtmlReference.SubmitImage(gk_sImagesPath & gk_sImageOK, mk_sNameButtonOk, _
                    m_oFWReferenceData.Search(EGroup.gToolTip, ETOOLTIP.ttButtonOKHelp))
    sHTML = sHTML & m_ofwhtmlReference.Hyperlink(gk_sLinkDisplayLoadHome, _
                 sTempHTML, gk_sMainFrame, , m_oFWReferenceData.Search(EGroup.gToolTip, ETOOLTIP.ttButtonCancelHelp))
                
                            
    m_oFWTable.StartRow
    typCell = m_ofwhtmlReference.GetCellStyle(ECellStyle.CELLNoStyle)
    typCell.sWidth = "100%"
    typCell.lCOLSPAN = 2
    m_oFWTable.AddCellAsHTML sHTML
    m_oFWTable.EndRow
    
    sOKLink = _
        gk_sLinkAccessionStarted & sParam
    sHTML = m_ofwhtmlReference.FormStart(sOKLink, gk_sMainFrame)
    sHTML = sHTML & m_oFWTable.GetHTML
    sHTML = sHTML & m_ofwhtmlReference.EndTag(gk_sHTMLForm)
    
    m_aspResponse.Write sHTML
    
    m_objContext.SetComplete
    Exit Sub

ErrorHandler:
    m_objContext.SetComplete
    ShowError sSource
End Sub


Public Sub DisplayAccessionStartedInfoPage()
'method to display conformation screen for the Accession functionality
On Error GoTo ErrorHandler
    
    Dim sHTML As String, sSource As String, sOKLink As String
    
    
    sSource = Me.Name & ".DisplayAccessionStartedPage()"
    
    
    sHTML = m_ofwhtmlReference.Bold(m_ofwhtmlReference.Font( _
                "Accession of EAD file started. This process could take time. " & _
                "Please check back later to create an edit set.", , 15))

    m_oFWTable.StartRow
    m_oFWTable.AddCellAsHTML sHTML
    m_oFWTable.EndRow
                            
    sOKLink = _
        gk_sLinkDisplayLoadHome
                            
    sHTML = m_ofwhtmlReference.ImageButton(gk_sImagesPath & gk_sImageOK, mk_sNameButtonOk, _
                    sOKLink, gk_sMainFrame)
    
    m_oFWTable.StartRow
    m_oFWTable.AddCellAsHTML sHTML
    m_oFWTable.EndRow
    
    m_aspResponse.Write m_oFWTable.GetHTML
    m_objContext.SetComplete
    Exit Sub

ErrorHandler:
    m_objContext.SetComplete
    ShowError sSource
End Sub



Public Sub DisplayBatches()
'This method display the loadable batches available from both DRUID Load
'   and PROCAT EAD Load
On Error GoTo ErrorHandler

    Dim sHTML As String, sSource As String, sOKLink As String
    
    sSource = Me.Name & ".DisplayBatches()"
    
    'Security check for the functionality
    CheckIfCanWorkOnBatchList
    
    m_oFWTable.TableStyle = m_ofwhtmlReference.GetTableStyle(ETableStyle.TABPlainStyle)
    m_oFWTable.TableStyle.sWidth = "100%"
    
    DisplayBatchListHeading
    
    DisplayBatchListContent
    
    sHTML = m_oFWTable.GetHTML
    
    m_aspResponse.Write sHTML
    
    m_objContext.SetComplete
    
    Exit Sub
ErrorHandler:
    m_objContext.SetComplete
    ShowError sSource
End Sub

Private Sub DisplayBatchListHeading()
'   d12 DAJ     29Feb2000   Added date column to batch list, and merged log columns

    Dim typCellStyle As CellStyle
    Dim typTableStyle As TableStyle
    
    Dim sHeading As String
    '   d8  CRM     12Jan2000 minor correction to heading
    sHeading = "List of batches loaded into the system"

    'The page heading
    m_oFWTable.StartRow

    typTableStyle = m_ofwhtmlReference.GetTableStyle(TABContentStyle)
    typTableStyle.lCELLSPACING = 1
    m_oFWTable.TableStyle = typTableStyle

    typCellStyle = m_ofwhtmlReference.GetCellStyle(ECellStyle.CELLPageHeadingStyle)
    typCellStyle.lCOLSPAN = 4
    
    m_oFWTable.AddCellAsHTML sHeading, typCellStyle
    
    m_oFWTable.EndRow

    'Table headers
    m_oFWTable.StartRow
    
    typCellStyle = m_ofwhtmlReference.GetCellStyle(ECellStyle.CELLColumnHeadingStyle)
    '   d8  CRM     12Jan2000  Added FontStyle instead of the default fonts to column headings
    m_oFWTable.AddCellAsHTML m_ofwhtmlReference.FontWithStyle("Batch ID", _
                             m_ofwhtmlReference.GetFontStyle(FONTColumnHeadingStyle)), _
                             typCellStyle
    
    m_oFWTable.AddCellAsHTML m_ofwhtmlReference.FontWithStyle("Batch name", _
                             m_ofwhtmlReference.GetFontStyle(FONTColumnHeadingStyle)), _
                             typCellStyle
    
    m_oFWTable.AddCellAsHTML m_ofwhtmlReference.FontWithStyle("Load from", _
                             m_ofwhtmlReference.GetFontStyle(FONTColumnHeadingStyle)), _
                             typCellStyle
    
    m_oFWTable.AddCellAsHTML m_ofwhtmlReference.FontWithStyle("Status", _
                             m_ofwhtmlReference.GetFontStyle(FONTColumnHeadingStyle)), _
                             typCellStyle
    
    ' d12 - Added date column
    m_oFWTable.AddCellAsHTML m_ofwhtmlReference.FontWithStyle("Date Started", _
                             m_ofwhtmlReference.GetFontStyle(FONTColumnHeadingStyle)), _
                             typCellStyle
        
    ' d10 - Column to display parse log
    ' d12 - Merged all log columns together
    m_oFWTable.AddCellAsHTML m_ofwhtmlReference.FontWithStyle("Parse Log", _
                             m_ofwhtmlReference.GetFontStyle(FONTColumnHeadingStyle)), _
                             typCellStyle
    
    'Column to view the import log
    m_oFWTable.AddCellAsHTML m_ofwhtmlReference.FontWithStyle("Import/Export Log", _
                             m_ofwhtmlReference.GetFontStyle(FONTColumnHeadingStyle)), _
                             typCellStyle
    
    'Column to view the Load log
    m_oFWTable.AddCellAsHTML m_ofwhtmlReference.FontWithStyle("Load Log", _
                             m_ofwhtmlReference.GetFontStyle(FONTColumnHeadingStyle)), _
                             typCellStyle
    
    m_oFWTable.EndRow

End Sub

Private Sub DisplayBatchListContent()
'   d12 DAJ     29Feb2000   Added date column to batch list, and merged log columns
'   d13 DAJ     06Mar2000   Added Editset creation failed status to the list of hyperlink
'                           enabled batches
'   d16 DAJ     21Mar2000   Speed ups to DisplayBatchListContent
'   d18 MAP     22Oct2001   Added code to handle EAD and Druid scheduling & automatic laod into Edit Set
'   d19 MAP     22Oct2001   Added extra case to display DRUID load hyperlink for

    Dim sTempHTML As String, sParam As String, sToolTip As String
    Dim ofwResult As FWRecordList
    Dim typCellStyle As CellStyle
    
    ' d16 - Variables to prevent loads of accesses to the record
    Dim iLoadStatus As EImportStatus, iLoadType As EImportLoadType, lBatchID As Long
    
    Set ofwResult = m_oboAccession.ListBatchesToCreateEditset
    
    While Not ofwResult.IsAtEnd
        
        m_oFWTable.StartRow
        
        ' d16 - Get the type and status now
        iLoadStatus = CInt(ofwResult.GetData(gk_sDBCOLStatus))
        iLoadType = CInt(ofwResult.GetData(gk_sDBCOLLoadType))
        lBatchID = CLng(ofwResult.GetData(gk_sDBCOLBatchID))
                
        typCellStyle = m_ofwhtmlReference.GetCellStyle(ECellStyle.CELLColumnDataStyle)
        
        'Display the batch ID value with the hyper link to start creating an editset
        sParam = "?" & gk_sParamBatchID & "=" & CStr(lBatchID)
        
        'Display the hyper link only if accession was successful
        ' d18 Also display hyperlink if DRUID to edit set failed after successful druid load
        '  Don't display as hyperlink if the DRUID load has been scheduled
        If iLoadStatus = EImportStatus.isEADImportSuccessful Or _
           iLoadStatus = EImportStatus.isDRUIDDataPreparationSuccessful Or _
           iLoadStatus = EImportStatus.isEditsetCreationFailed Then
            
            sToolTip = m_oFWReferenceData.Search(EGroup.gToolTip, ETOOLTIP.ttLoadBatchIntoEditset)
            
            sTempHTML = m_ofwhtmlReference.Hyperlink(gk_sLinkDisplayCreateEmptyEditSetForLoad & sParam, _
                        CStr(lBatchID), gk_sMainFrame, _
                        mk_sHyperLink, _
                        sToolTip)
            
            m_oFWTable.AddCellAsHTML sTempHTML, typCellStyle
        Else
            
            sTempHTML = CStr(lBatchID)
            m_oFWTable.AddCellAsHTML sTempHTML, typCellStyle
        End If
        
        
        'Display the batch description value
        sTempHTML = ofwResult.GetData(gk_sDBCOLDescription)
        
        m_oFWTable.AddCellAsHTML sTempHTML, typCellStyle
        
        'Display the load type of the batch entry. Display a hyper link if DRUID load is done.
        ' This will be used to prepare the DRUID data for PROCAT usage before loading into an edit set
        '
        sTempHTML = GetImportLoadTypeDescription(iLoadType)
        
        ' d18 - Amended links depending on status
        '  Give option to create edit set for druid if DruidHeaders have loaded
        If iLoadType = EImportLoadType.iltDRUID _
            And (iLoadStatus = EImportStatus.isDRUIDHeadersSuccessful Or _
                 iLoadStatus = EImportStatus.isDRUIDAllLoaded) Then
            
            sToolTip = m_oFWReferenceData.Search(EGroup.gToolTip, ETOOLTIP.ttPrepareDRUIDData)
                      
                'Case EImportStatus.isDRUIDHeadersSuccessful
                    sTempHTML = m_ofwhtmlReference.Hyperlink(gk_sLinkDisplayCreateEmptyEditSetScreenForDRUIDLoad & sParam, _
                        sTempHTML, gk_sMainFrame, _
                        mk_sHyperLink, _
                        sToolTip)
                
'            End Select
            
        End If
        
        m_oFWTable.AddCellAsHTML sTempHTML, typCellStyle
        
        'Display the status of the batch entry
        sTempHTML = GetImportLoadStatusDescription(iLoadStatus)
        
        m_oFWTable.AddCellAsHTML sTempHTML, typCellStyle
        
        ' d12 - Display the status of the batch entry
        sTempHTML = DateForDisplay(ofwResult.GetData(gk_sDBCOLUpLoadDate))
        
        m_oFWTable.AddCellAsHTML sTempHTML, typCellStyle
        
        ' d10 - Display the parse log link
        sToolTip = m_oFWReferenceData.Search(EGroup.gToolTip, ETOOLTIP.ttParseLogFileView)
        
        If iLoadStatus = EImportStatus.isDRUIDAllLoaded Or _
           iLoadStatus = EImportStatus.isDRUIDHeadersSuccessful Or _
           iLoadStatus = EImportStatus.isDRUIDLoadStarted Then
            
            sTempHTML = gk_sHTMLNonBreakSpace
            
        Else
        
            If iLoadType = EImportLoadType.iltDRUID Then
            
                sTempHTML = gk_sHTMLNonBreakSpace
            Else
                
                sTempHTML = m_ofwhtmlReference.Hyperlink(gk_sLinkDisplayEADParseLogging & sParam, _
                            "View log", gk_sMainFrame, _
                            mk_sHyperLink, _
                            sToolTip)
            End If
    
        End If
        
        m_oFWTable.AddCellAsHTML sTempHTML, typCellStyle
        
        'Display the view import log link
        sToolTip = m_oFWReferenceData.Search(EGroup.gToolTip, ETOOLTIP.ttImportLogFileView)
        
        If iLoadStatus = EImportStatus.isDRUIDAllLoaded Or _
           iLoadStatus = EImportStatus.isDRUIDHeadersSuccessful Or _
           iLoadStatus = EImportStatus.isDRUIDLoadStarted Then
            
            sTempHTML = gk_sHTMLNonBreakSpace
            
        Else
        
        sTempHTML = m_ofwhtmlReference.Hyperlink(gk_sLinkDisplayAccessionLogFile & sParam, _
                    "View log", gk_sMainFrame, _
                    mk_sHyperLink, _
                    sToolTip)
    
        End If
        
        m_oFWTable.AddCellAsHTML sTempHTML, typCellStyle
        
        
        'Display the view load log link
        sTempHTML = gk_sHTMLNonBreakSpace
        
        If iLoadStatus = EImportStatus.isEditsetCreationStarted Or _
           iLoadStatus = EImportStatus.isEditsetCreationFailed Or _
           iLoadStatus = EImportStatus.isEditsetCreationSuccessful Then
            
            sParam = sParam & "&" & gk_sParamEditSetID & "=" & ofwResult.GetData(gk_sDBCOLEditSetID)
            
            sToolTip = m_oFWReferenceData.Search(EGroup.gToolTip, ETOOLTIP.ttLoadLogFileView)
            
            sTempHTML = m_ofwhtmlReference.Hyperlink(gk_sLinkDisplayAccessionLogFile & sParam, _
                        "View log", gk_sMainFrame, _
                        mk_sHyperLink, _
                        sToolTip)
            
            m_oFWTable.AddCellAsHTML sTempHTML, typCellStyle
        Else
            m_oFWTable.AddCellAsHTML sTempHTML, typCellStyle
        End If
        
        m_oFWTable.EndRow
        
        ofwResult.MoveNext
    Wend
    
End Sub


Private Function GetImportLoadTypeDescription(t_eLoadType As EImportLoadType) As String
    Select Case t_eLoadType
    Case EImportLoadType.iltDRUID
        GetImportLoadTypeDescription = "DRUID"
    Case EImportLoadType.iltEAD
        GetImportLoadTypeDescription = "EAD"
    End Select
End Function

Private Function GetImportLoadStatusDescription(t_eLoadStatus As EImportStatus) As String
'Returns a description of the current status in the database
'   d5  KSK     10Dec1999   Changes after customer feedback to add DRUID status flags
'   d15 JCC     14Mar2000   Spelling mistake - perperation
'   d17 DAJ     14Jul2000   Added ability to recognise update/export batches
'   d18  MAP     17Oct2001  Added Scheduled status
    
    Select Case t_eLoadStatus
    'DRUID status on load by DRUID
    Case EImportStatus.isDRUIDLoadStarted
        GetImportLoadStatusDescription = "DRUID Load started..."
    Case EImportStatus.isDRUIDHeadersSuccessful
        GetImportLoadStatusDescription = "DRUID headers load successful."
    Case EImportStatus.isDRUIDAllLoaded
        GetImportLoadStatusDescription = "DRUID all load successful."
    'PROCAT preperation of data before PROCAT load
    Case EImportStatus.isDRUIDDataPreparationStarted
        GetImportLoadStatusDescription = "DRUID data preparation for load started..."
    Case EImportStatus.isDRUIDDataPreparationFailed
        GetImportLoadStatusDescription = "DRUID data preparation failed!"
    Case EImportStatus.isDRUIDDataPreparationSuccessful
        GetImportLoadStatusDescription = "DRUID data preparation successful."
    'PROCAT impoting data into the system
    Case EImportStatus.isEADImportFailed
        GetImportLoadStatusDescription = "EAD Import failed!"
    Case EImportStatus.isEADImportStarted
        GetImportLoadStatusDescription = "EAD Import started..."
    Case EImportStatus.isEADImportSuccessful
        GetImportLoadStatusDescription = "EAD Import successful."
    'PROCAT creation of edit set from the loaded batches (for both DRUID and EAD)
    Case EImportStatus.isEditsetCreationFailed
        GetImportLoadStatusDescription = "Editset creation failed!"
    Case EImportStatus.isEditsetCreationStarted
        GetImportLoadStatusDescription = "Editset creation started..."
    Case EImportStatus.isEditsetCreationSuccessful
        GetImportLoadStatusDescription = "Editset creation successful."
    
    ' d17 - New descriptions for update/export
    Case EImportStatus.isEADUpdateFailed
        GetImportLoadStatusDescription = "EAD Update failed!"
    Case EImportStatus.isEADUpdateStarted
        GetImportLoadStatusDescription = "EAD Update started..."
    Case EImportStatus.isEADUpdateSuccessful
        GetImportLoadStatusDescription = "EAD Update successful."
    Case EImportStatus.isEADExportFailed
        GetImportLoadStatusDescription = "EAD Export failed!"
    Case EImportStatus.isEADExportStarted
        GetImportLoadStatusDescription = "EAD Export started..."
    Case EImportStatus.isEADExportSuccessful
        GetImportLoadStatusDescription = "EAD Export successful."
        
    ' d18 - Added
    Case EImportStatus.isEADImportScheduled
        GetImportLoadStatusDescription = "EAD Import scheduled."
    Case EImportStatus.isEADUpdateScheduled
        GetImportLoadStatusDescription = "EAD Update scheduled."
    Case EImportStatus.isEADExportScheduled
        GetImportLoadStatusDescription = "EAD Export scheduled."
    Case EImportStatus.isDRUIDLoadScheduled
        GetImportLoadStatusDescription = "DRUID Load scheduled."
        
    End Select
End Function


Public Sub DisplayTheLogFileDetails()
'This method display the loadable batches available from both DRUID Load
'   and PROCAT EAD Load
'   d4  KSK     25Nov1999   Added code the move the display of log to the end of file
'   d8  CRM     12Jan2000   Added reload button to DisplayLogFileDetails(),code copied from
'                           DisplayEADLogFileDetails()
'   d14 DAJ     07Mar2000   Updated DisplayLogFile functions to check the return path

On Error GoTo ErrorHandler
    Dim sSource As String
    Dim lBatchID As Long
    Dim lEditsetID As Long
    Dim sHTML As String
    Dim typCellStyle As CellStyle
    Dim sToolTip As String
    Dim sRedirection As String
    
    sSource = Me.Name & ".DisplayTheLogFileDetails()"
    
    'Security check for the functionality
    CheckIfCanWorkOnBatchList
    
    'Get editset id from the query string
    lBatchID = m_aspRequest.QueryString(gk_sParamBatchID)
    If IsEmpty(m_aspRequest.QueryString(gk_sParamEditSetID)) Then
        lEditsetID = 0
    Else
        lEditsetID = m_aspRequest.QueryString(gk_sParamEditSetID)
    End If
    sRedirection = m_aspRequest(gk_sParamRedirection)
    
    Set m_oFWTable = New FWTable
    m_oFWTable.TableStyle.sWidth = "100%"
    
    If lEditsetID = 0 Then
        sHTML = "Import/Export log content"
    Else
        sHTML = "Load log content"
    End If
    
    m_oFWTable.StartRow
    typCellStyle.sWidth = "100%"
    m_oFWTable.AddCellAsHTML m_ofwhtmlReference.Bold(sHTML), typCellStyle
    m_oFWTable.EndRow
    
    sHTML = "<pre>" & ReadLogFile(, lBatchID, lEditsetID) & "</pre>" & vbNewLine
    
    m_oFWTable.StartRow
    typCellStyle.sWidth = "100%"
    m_oFWTable.AddCellAsHTML sHTML, typCellStyle
    m_oFWTable.EndRow
    
    'Display the buttons
    sToolTip = m_oFWReferenceData.Search(EGroup.gToolTip, ETOOLTIP.ttButtonDone)

    sHTML = m_ofwhtmlReference.Image(mk_sNameButtonDone, _
                    gk_sImagesPath & gk_sImageDone, _
                    sToolTip)
    
    If sRedirection <> "" Then
        sHTML = m_ofwhtmlReference.Hyperlink(sRedirection, sHTML, _
                        gk_sMainFrame, mk_sNameButtonDone, sToolTip)
    Else
        sHTML = m_ofwhtmlReference.Hyperlink(gk_sLinkDisplayBatchList, sHTML, _
                        gk_sMainFrame, mk_sNameButtonDone, sToolTip)
    End If
                        
    '   d8  CRM     12Jan2000   Added reload button
    sHTML = sHTML & m_ofwhtmlReference.Image(mk_sNameButtonReload, _
                    gk_sImagesPath & gk_sImageReload, _
                    "Reload the file", t_sOnClick:="location.reload()")
                    
    m_oFWTable.StartRow
    typCellStyle.sWidth = "100%"
    m_oFWTable.AddCellAsHTML m_ofwhtmlReference.Paragraph(sHTML), typCellStyle
    m_oFWTable.EndRow
    
    sHTML = m_oFWTable.GetHTML & m_ofwhtmlReference.ResetPageTop
    
    m_aspResponse.Write sHTML
    
    m_objContext.SetComplete
    
    Exit Sub
ErrorHandler:
    m_objContext.SetComplete
    ShowError sSource
End Sub


Public Sub DisplayEADLogFileDetails()
'This method display the loadable batches available from both DRUID Load
'   and PROCAT EAD Load
'   d6  SPB     15Dec1999   Created
'   d9  DAJ     28Feb2000   Added function to use Batch ID for parse log
'   d11 DAJ     28Feb2000   Added redirection for batch ID log files
'   d14 DAJ     07Mar2000   Updated DisplayLogFile functions to check the return path
'   d17 DAJ     18Jul2000   Added parameter passing for update list

On Error GoTo ErrorHandler
    Dim sSource As String
    Dim lBatchID As Long
    Dim lEditsetID As Long
    Dim sHTML As String
    Dim typCellStyle As CellStyle
    Dim sToolTip As String
    Dim bBatchLog As Boolean
    Dim sRedirection As String
    Dim sParam As String
    
    sSource = Me.Name & ".DisplayEADLogFileDetails()"
    
    'Security check for the functionality
    CheckIfCanWorkOnBatchList
    
    sRedirection = m_aspRequest(gk_sParamRedirection)

    Set m_oFWTable = New FWTable
    m_oFWTable.TableStyle.sWidth = "100%"
    
    sHTML = "EAD parser log content"
    
    m_oFWTable.StartRow
    typCellStyle.sWidth = "100%"
    m_oFWTable.AddCellAsHTML m_ofwhtmlReference.Bold(sHTML), typCellStyle
    m_oFWTable.EndRow
    
    ' d9 - If filename not found, try Batch ID
    If m_aspRequest(gk_sParamFileName) = "" Then
        ' Must pass an empty string to prevent getting an accession log
        bBatchLog = True
        sHTML = "<pre>" & ReadLogFile("", m_aspRequest(gk_sParamBatchID)) & "</pre>"
    Else
        bBatchLog = False
        sHTML = "<pre>" & ReadLogFile(m_aspRequest(gk_sParamFileName)) & "</pre>"
    End If
    
    m_oFWTable.StartRow
    typCellStyle.sWidth = "100%"
    m_oFWTable.AddCellAsHTML sHTML, typCellStyle
    m_oFWTable.EndRow
    
    'Display the buttons
    sToolTip = m_oFWReferenceData.Search(EGroup.gToolTip, ETOOLTIP.ttButtonDone)

    Dim sDoneButton As String
    Dim sRefreshButton As String
    Dim sButtons As String
    sDoneButton = m_ofwhtmlReference.Image(mk_sNameButtonDone, _
                    gk_sImagesPath & gk_sImageDone, _
                    sToolTip)
    
    ' d17 - Do we need to pass any parameters around?
    If m_aspRequest(gk_sParamUpdateEADList) Then
        sParam = "?" & gk_sParamUpdateEADList & "=1"
    End If
    
    ' d11 - Check if this is a batch ID log file
    If sRedirection <> "" Then
        sDoneButton = m_ofwhtmlReference.Hyperlink(sRedirection, sDoneButton, _
                        gk_sMainFrame, mk_sNameButtonDone, sToolTip)
    ElseIf bBatchLog = True Then
        sDoneButton = m_ofwhtmlReference.Hyperlink(gk_sLinkDisplayBatchList & sParam, sDoneButton, _
                        gk_sMainFrame, mk_sNameButtonDone, sToolTip)
    Else
        sDoneButton = m_ofwhtmlReference.Hyperlink(gk_sLinkDisplayEADFileList & sParam, sDoneButton, _
                        gk_sMainFrame, mk_sNameButtonDone, sToolTip)
    End If
                        
    sRefreshButton = m_ofwhtmlReference.Image(mk_sNameButtonReload, _
                    gk_sImagesPath & gk_sImageReload, _
                    "Reload the file", t_sOnClick:="location.reload()")
                                        
    sButtons = sDoneButton & sRefreshButton & m_ofwhtmlReference.ResetPageTop
    
    m_oFWTable.StartRow
    typCellStyle.sWidth = "100%"
    m_oFWTable.AddCellAsHTML m_ofwhtmlReference.Paragraph(sButtons), typCellStyle
    m_oFWTable.EndRow
    
    sHTML = m_oFWTable.GetHTML
    
    m_aspResponse.Write sHTML
    
    m_objContext.SetComplete
    
    Exit Sub
ErrorHandler:
    m_objContext.SetComplete
    ShowError sSource
End Sub


Private Sub CheckIfCanWorkOnBatchList()
On Error GoTo ErrHandler
Dim eResult1 As ESecurityCheckResult, eResult2 As ESecurityCheckResult
Dim sSource As String
    
    sSource = Me.Name & ".CheckIfCanWorkOnBatchList()"
    
    m_oBOEditorialPermissions.UserID = m_lLoggedUserID
    
    'The batches display is avaialble to all users with accession and contribution right
    ' That is right to load and right to import and prepare data
    eResult1 = m_oBOEditorialPermissions.Check(ESecurityCheck.EUserCanImportFromEAD)
    eResult2 = m_oBOEditorialPermissions.Check(ESecurityCheck.EUserCanLoadIntoEditset)
    
    If eResult1 <> ESecurityCheckResult.ESRSecurityCheckSuccessful And _
        eResult2 <> ESecurityCheckResult.ESRSecurityCheckSuccessful Then
        If eResult1 <> ESecurityCheckResult.ESRSecurityCheckSuccessful Then
            RaiseError sSource, _
               GetErrorEnumForSecurityCheckResult(eResult1), _
               GetDescriptionForSecurityCheckError( _
                        GetErrorEnumForSecurityCheckResult(eResult1))
        Else
            RaiseError sSource, _
               GetErrorEnumForSecurityCheckResult(eResult2), _
               GetDescriptionForSecurityCheckError( _
                        GetErrorEnumForSecurityCheckResult(eResult2))
        End If
    End If
    Exit Sub
ErrHandler:
    ConvertError sSource
End Sub




