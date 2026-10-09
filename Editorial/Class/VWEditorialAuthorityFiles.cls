VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "VWEditorialAuthorityFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    VWEditorialAuthorityFiles.cls
' System:       PRO, PROCAT
' Copyright:    (C) Quidnunc Limited
'
' Description:
'               This view class consists of all display methods used in
'               authority files maintenance.
'
' Amendment history:
'   d1  Geedha Sivanadian     19Jul1999    Created
'   d2  Geedha Sivanadian     4Aug1999      Added UpdateIndexTermDetails method.
'   d3  Geedha Sivanadian     6Aug1999      Added new BOSaver interface.
'   d4  Geedha Sivanadian     10Aug1999     Added code for rest of index term types.
'   d5  AGA                   08Oct1999     Modified to show invisible button
'                                           Added security
'                                           Added method (SetStateBeforeLeave)
'                                             to integrage with index browser and for persistance
'                                             Added methods to add a new AF term
'   d6  AGA                   29Oct1999     Added GetBrowserRedirectPageAfterAddAFTerm function
'   d7  AGA                   02Nov1999     Modified CheckPermissions
'   d8  KSK                   03Nov1999     Changed Checkpermissions method to used the parameter source
'                                           value while raising error
'   d9  AGA                   03Nov1999     Moved ASP form fields & Session parameters from Missing.bas,
'                                               and add term image with pick term images
'   d10 KSK                   04Nov1999     Added redirection for bulkupdate. And method for Cancel
'                                           redirection page name.
'   d11 aga                   04Nov1999     changed from const to function for the cancel link
'                                           Added sessions to DisplayIndexTerms used in case of pressing cancel
'   d12 KSK                   04Nov1999     Changes with AGA.
'   d13 ndk                   04Nov1999     Code put for pickterm for edit catalogue entry
'   d14 aga                   07Dec1999     Modifications for new spec, change in the business login
'                                           Added more categories on the AF home page
'   d15 aga                   14Dec1999     Showing all editsets in all stages and hyperlinks only to the editset in checking stage
'                                           Added reject confirmation for preferred terms
'   d16 aga                   15Dec1999     Adding warning messages for rejection
'   d17 aga                   15Dec1999     Changed the warning messages
'   d18 aga                   21Dec1999     Defect fixes, Changing the cell styles
'   d19 aga                   12Jan2000     Defect fixes
'   d20 crm                   04Feb2000     Added redirection URl after adding a new Index term
'                                           from Leaflets maintenance
'   d20 crm                   07Feb2000     changed redirection links after editing and saving an index term
'   d21 jcc                   07Feb2000     correct spelling mistake
'   d22 crm                   08Feb2000     Functionality to disable and check IsPlaceOfDeposit check box
'                                           when adding term to PlaceOfdeposit, while editing
'                                           PlaceOfDeposit catalogue entries
'   d23 crm                   08Feb2000     Code to redirect cancel operation for adding new term in Leaflets
'   d24 crm                   16Feb2000     Changed field length for Uncertain date code throughout
'   d25 daj                   21Feb2000     Removed all references to maintenance directory
'   d26 wdp                   23Feb2000     Sorted display of see and see also links and buttons
'                                           for various status of terms(preferred, non-preferred etc.)
'   d27 tmj                   29Feb2000     Added call to Request so EditsetID is set when checking permissions
'   d28 tmj                   09Mar2000     Changed URLs for cancel buttons in GetBrowserRedirectForCancelAFTerm
'   d29 tmj                   16Mar2000     Modified so TextArea size is validated.
'   d30 tmj                   16Mar2000     Added code to pass where call to AF list page comes from
'   d31 tmj                   20Mar2000     Added code to improve re-direction
'   d32 tmj                   21Mar2000     Changed parameter that is used to set the displayed Index Term
'                                           when it has just been added. This results in the prefered rather than
'                                           the non prefered term being displayed
'   d33 tmj                   27Mar2000     Added hidden field when is place of deposit checkbox
'                                           is disabled
'   d34 daj                   27Mar2000     Added SFA checking when adding new candidate terms, and also added code
'                                           to pass edit set ID around when adding it to an edit set entry
'   d35 daj                   17Jul2000     Changed the checking in Display<AFType>Details functions so that the links
'                                           currently stored in TransientStore are used (rather than the database) to decide
'                                           which of the See and See Also buttons are shown
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Option Explicit

Implements ObjectControl

Private m_ofwHtmlRef As FWHTML
Private m_oboAuthorityFiles As BOEditorialAuthorityFiles

'd3
Private m_oiAFSaver As IAFSaver

Private m_objContext As ObjectContext
Private m_oFWAppData As FWApplicationData

Private m_aspResponse As Response
Private m_aspRequest As Request
Private m_aspServer As Server
Private m_aspSession As Session



Private m_lLoggedUserID As Long                     'ID of the logged user
Private m_bLoggedUserIsAdministrator As Boolean     'Whether the logged user is an administrator

' Constants
Private Const klEditSetType As Long = 0
Private Const klleafletType As Long = 1
Private Const klInvisibleType As Long = 2
Private Const klIncompleteType As Long = 3
Private Const klNotUsedType As Long = 4
'd14
Private Const klNonCandidateNonPreferredType As Long = 5
Private Const klNotUsedPreferredType As Long = 6
Private Const klNotUsedCandidateType As Long = 7
Private Const klLiveType As Long = 8

'd14    standards
Private Const ksTermCategory As String = "Term_Category"
Public m_bNoPersist As Boolean
Private Const ksFORMCancel As String = "Cancel"

Private Const k_sIndexTermRaisedThrough As String = "IndexTermRaisedType"

Dim m_typPersonData As PersonAFTerm
Dim m_typPlaceData As PlaceAFTerm
Dim m_typSubjectData As SubjectAFTerm
Dim m_typCorporateBodyData As CorporateBodyAFTerm

Private Const k_sApproveCompleteBtn = "ApproveCompleteSubmitButton"
Private Const k_sApproveInCompleteBtn = "ApproveInCompleteSubmitButton"
Private Const k_sRejectBtn = "RejectSubmitButton"
'd5
Private Const k_sInvisibleBtn = "InvisibleSubmitButton"
Private Const k_sOKBtn = "OKSubmitButton"          'aga


'd5  Permissions object
Private m_oboEditorialPermission As BOEditorialPermissions


'd5 Term type and IDs
Private m_lTermID As Long
'd14
Private m_lSpecificTermType As Long

'd16
Public m_lIndexTermRaisedSource As Long

'd5
Private m_colSeeLinks As Collection
Private m_colSeeAlsoLinks As Collection
Private m_oBOAFPerson As BOAFPerson
Private m_oBOAFSubject As BOAFSubject
Private m_oBOAFPlace As BOAFPlace
Private m_oBOAFCorporateBody As BOAFCorporateBody

Private m_iTempRemoveTermID As Integer
Private m_ofwRefData As FWReferenceData

'd9   ASP form field& session parameter names
Private Const mk_sFORMAFSeeLinkType As String = "SeeLinkType"
Private Const mk_sFORMAFRemoveTermID As String = "RemoveTermID"
Private Const mk_sFORMAFTermIsRemoved As String = "IsRemoved"
Private Const mk_sFORMAFTermAddOrRemove As String = "AddOrRemove"
Private Const mk_sAFAuthorityStatus As String = "AuthorityStatus"

'd9 Moved from Editorial constants
'Used to identify which button is clicked on AF Editorial Maintenance
Private Const mk_sAFSee As String = "See"
Private Const mk_sAFSeeAlso As String = "SeeAlso"
Private Const mk_sAFSeeRemove As String = "SeeRemove"
Private Const mk_sAFSeeAlsoRemove As String = "SeeAlsoRemove"

'D14 Constants for return codes of Save methods of the BO classes of AFterms
'IMPORTANT: These values should match the values in the SP_Edit_RejectAFTerm
Private Const mk_lRecordNotChanged As Long = -10
Private Const mk_lRecordDeleted As Long = -20
Private Const mk_lRecordMarkedAsRejected As Long = -30

Private m_sDisablePlaceOfDeposit As String
'TODO
Private Const gk_sDisablePlaceOfDeposit As String = "DisablePlaceOfDeposit"



''''''''''''''''' Implementation ''''''''''''''''''''''

'
' Private functions
'
'
' Class initialisation and MTS event handlers
'

Private Sub Initialise()
    Dim sSource As String
    
    Set m_objContext = GetObjectContext()
    
    sSource = "VWEditorialAuthorityFiles.Initialise"
    
    ' Set asp objects
    Set m_aspResponse = m_objContext.Item("Response")
    Set m_aspRequest = m_objContext.Item("Request")
    Set m_aspServer = m_objContext.Item("Server")
    Set m_aspSession = m_objContext.Item("Session")
    
    ' Specific initialisation
    Set m_ofwHtmlRef = New FWHTML
    Set m_oboAuthorityFiles = New BOEditorialAuthorityFiles
    Set m_oFWAppData = New FWApplicationData
    Set m_oboEditorialPermission = New BOEditorialPermissions
    
    'd5
    Set m_colSeeLinks = New Collection
    Set m_colSeeAlsoLinks = New Collection
    Set m_oBOAFPerson = New BOAFPerson
    Set m_oBOAFSubject = New BOAFSubject
    Set m_oBOAFPlace = New BOAFPlace
    Set m_oBOAFCorporateBody = New BOAFCorporateBody
    
    Set m_ofwRefData = New FWReferenceData
    
    m_lLoggedUserID = GetDataFromSession(dsUserID, m_objContext.Item(gk_sSessionObject))
    If m_lLoggedUserID = 0 Then     'Not a valid user
        RaiseError sSource, EPROCATINTERNALERROR.peDBInvalidUser, "Invalid User"
    End If
End Sub

Private Sub ObjectControl_Activate()
    Initialise
End Sub

Private Sub ObjectControl_Deactivate()
    Set m_objContext = Nothing
    Terminate
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = True
End Function

Private Sub Terminate()
    'release objects
    Set m_objContext = Nothing
        
    Set m_aspResponse = Nothing
    Set m_aspRequest = Nothing
    Set m_aspServer = Nothing
    Set m_aspSession = Nothing
    
    Set m_oiAFSaver = Nothing
    Set m_ofwHtmlRef = Nothing
    Set m_oboAuthorityFiles = Nothing
    Set m_oboEditorialPermission = Nothing
    
    Set m_ofwRefData = Nothing
    'd5
    Set m_colSeeLinks = Nothing
    Set m_colSeeAlsoLinks = Nothing
    Set m_oBOAFPerson = Nothing
    Set m_oBOAFSubject = Nothing
    Set m_oBOAFPlace = Nothing
    Set m_oBOAFCorporateBody = Nothing
    
    Set m_oFWAppData = Nothing
    
End Sub



Public Sub DisplayAFMaintMainPage()

'   This is the main procedure for AF maintenance.
'   This procedure will display all non approved terms raised through leaflets/editsets
'   and corresponding leaflets/editsets details.
'   In addition this should display count of invisible and incomplete terms.
'
'   d1  Geedha Sivanadian     19Jul1999   Created
'   d14 AGA                   07Dec1999   Added new categories of terms
'   d15 aga                   14Dec1999   Showing all editsets in all stages and hyperlinks only to the editset in checking stage

On Error GoTo ErrorHandler
                 
    'Recordssets
    Dim ofwRlResult As FWRecordList
    
    'Table
    Dim ofwTable As FWTable
    
    'Table, row, cell styles
    Dim typTable As TableStyle
    Dim typCustomStyle As RowStyle
    Dim typCustomCell As CellStyle
    
    'Strings to store variables.
    Dim sSource As String
    Dim sHTML As String
    Dim sHeading As String
    Dim sParameter As String
    Dim sListInvisibleTerms As String
    Dim sListIncompleteTerms As String
    Dim sListNotUsedTerms As String
    Dim sNonCandidateNonPreferredTerms As String
    Dim sNotUsedPreferredTerms As String
    Dim sNotUsedCandidateTerms As String
    Dim sLiveTerms As String
    
    sSource = "VWEditorialAuthorityFiles.DisplayAFMaintMainPage"
    
    'Check the permissions
    CheckPermissions ESecurityCheck.EUserCanMaintainAuthorityFiles, sSource, ""
        
    'This is necessary for live terms to come back from Index browser
    m_aspSession(ksTermCategory) = klLiveType
    ' Title
        
    sHeading = m_ofwHtmlRef.Paragraph(m_ofwHtmlRef.Font( _
                m_ofwHtmlRef.Bold(m_oFWAppData.GetDisplayInformation( _
                       EAuthorityFilesMaintResourceFile.afmrAuthorityFilesMaintenance)), COLOURBlack, 4))
    sHeading = sHeading & m_ofwHtmlRef.Bold(m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrTermsRaisedthroughEditSets))
    sHeading = sHeading & m_ofwHtmlRef.Paragraph("")
    sHTML = sHeading
    
    
'*** Edit sets
    'First section for retrieving the index terms raised through edit sets.
    
    
    'Create table
    Set ofwTable = New FWTable
        
    'Table settings
    typTable.eStyle = TABCustomStyle
    typTable.lCELLPADDING = 2
    typTable.lCELLSPACING = 1
    
    
    ofwTable.TableStyle = typTable
    

    'First row to display column captions.
    ofwTable.StartRow (typCustomStyle)
        
            'typCustomCell.eStyle = CELLCustomStyle
            typCustomCell = m_ofwHtmlRef.GetCellStyle(CELLColumnHeadingStyle)
            
            ofwTable.AddHCellAsHTML m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrEditSetName), typCustomCell
            
            ofwTable.AddHCellAsHTML m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrContributor1), typCustomCell
        
            ofwTable.AddHCellAsHTML m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrCreatedon), typCustomCell
            
            ofwTable.AddHCellAsHTML m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrStage), typCustomCell
        
            ofwTable.AddHCellAsHTML m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrDaysinStage), typCustomCell
        
            ofwTable.AddHCellAsHTML m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrNumberOfTermsForApproval), typCustomCell
           
    ofwTable.EndRow
    
'Get the actual data rows from backend.
    
    Set ofwRlResult = m_oboAuthorityFiles.GetEditSetsList(GetDataFromSession(dsUserID))

    'This error condition would have been handled in the business object itself.
    'Just keep it here too for error handling in VWclass.
    If ofwRlResult Is Nothing Then
        RaiseError sSource, PEGeneralInternalError, ""
    End If
    
    ' Display all the editsets one by one.
    If ofwRlResult.RecordCount <> 0 Then
    
       ofwRlResult.MoveFirst
       
       Do While Not ofwRlResult.IsAtEnd
            'Do not show edit sets with number of approavals 0
            If Not ofwRlResult.GetFieldValue(gk_sNumberOfApprovalsRequired) = 0 Then
       
                sParameter = "?" & gk_sNameEditsetID & "=" & ofwRlResult.GetFieldValue(gk_sDBCOLEditSetID) & _
                            "&" & gk_sNameEditSetName & "=" & _
                            EncodeQueryStringValue(CStr(ofwRlResult.GetFieldValue(gk_sDBCOLEditsetName))) & _
                            "&" & ksTermCategory & "=" & klEditSetType
                            
                 ofwTable.StartRow
                    
                 typCustomCell = m_ofwHtmlRef.GetCellStyle(CELLReadOnlyDataStyle)
                 
                    'd15  If checking stage , show hyperlink
                    If ofwRlResult.GetFieldValue(gk_sDBCOLStageID) = gk_lStageCheckingID Then
                        ofwTable.AddCellAsHTML m_ofwHtmlRef.Hyperlink( _
                            gk_sLinkDisplayAuthorityTerms & sParameter, _
                            ofwRlResult.GetFieldValue(gk_sDBCOLEditsetName), gk_sMainFrame), typCustomCell
                    Else
                        ofwTable.AddCellAsHTML ofwRlResult.GetFieldValue(gk_sDBCOLEditsetName), typCustomCell
                    End If
                    
                    ofwTable.AddCellAsHTML ofwRlResult.GetFieldValue( _
                            gk_sDBCOLTempContributor1), typCustomCell
                     
                    ofwTable.AddCellAsHTML VariantDateAsDDMMMYYYY(ofwRlResult.GetFieldValue( _
                            gk_sDBCOLTempEditSetCreationDate)), typCustomCell
                            
        
                    ofwTable.AddCellAsHTML ofwRlResult.GetFieldValue( _
                            gk_sDBCOLStageName), typCustomCell
                    
                    ofwTable.AddCellAsHTML ofwRlResult.GetFieldValue( _
                            gk_sDBCOLTempDaysInCurrentStage), typCustomCell
        
                    ofwTable.AddCellAsHTML ofwRlResult.GetFieldValue( _
                            gk_sNumberOfApprovalsRequired), typCustomCell
                                    
                    ofwTable.EndRow
                End If
                ofwRlResult.MoveNext
        Loop
            
    End If
    
    'Get the table details
    
    sHTML = sHTML & m_ofwHtmlRef.Paragraph(ofwTable.GetHTML)
    
        
    'Write the html
    m_aspResponse.Write sHTML
    
'*** Display leaflets
    sHeading = ""
    sHeading = m_ofwHtmlRef.Paragraph(m_ofwHtmlRef.Bold(m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrTermsRaisedthroughLeaflets)))
    sHeading = sHeading & m_ofwHtmlRef.Paragraph("")
    sHTML = sHeading

'*** Leaflets
    'Second section for retrieving the index terms raised through leaflets.
        
    'Create table
    Set ofwTable = New FWTable
        
    'Table settings
    typTable.eStyle = TABCustomStyle
    typTable.lCELLPADDING = 2
    typTable.lCELLSPACING = 1
    
    ofwTable.TableStyle = typTable
    
    ' To display column-captions.
    ofwTable.DefaultRowStyle = m_ofwHtmlRef.GetRowStyle(RowBackgroundPink)
    typCustomStyle = m_ofwHtmlRef.GetRowStyle(RowBackgroundRed)
    
    'First row to display column captions.
    ofwTable.StartRow (typCustomStyle)
        
            typCustomCell = m_ofwHtmlRef.GetCellStyle(ECellStyle.CELLColumnHeadingStyle)
            
            ofwTable.AddHCellAsHTML m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrLeafletName), typCustomCell
            
            ofwTable.AddHCellAsHTML m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrCreatedon), typCustomCell
        
            ofwTable.AddHCellAsHTML m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrNumberOfTermsForApproval), typCustomCell
           
    ofwTable.EndRow
    
'Get the actual data rows from backend.
    
    Set ofwRlResult = m_oboAuthorityFiles.GetLeafLetsList

    'This error condition would have been handled in the business object itself.
    'Just keep it here too for error handling in VWclass.
    If ofwRlResult Is Nothing Then
        RaiseError sSource, PEGeneralInternalError, ""
    End If
    
    ' Display all the editsets one by one.
    If ofwRlResult.RecordCount <> 0 Then
    
       ofwRlResult.MoveFirst
       
       Do While Not ofwRlResult.IsAtEnd
       
            'Need to change the parameter names after testing
            sParameter = "?" & gk_sSFAReferenceID & "=" & ofwRlResult.GetFieldValue("leaflet_reference_id") & _
                        "&" & gk_sDBCOLSFAName & "=" & ofwRlResult.GetFieldValue("leaflet_name") & _
                        "&" & ksTermCategory & "=" & klleafletType
            
            
             ofwTable.StartRow
                
             typCustomCell = m_ofwHtmlRef.GetCellStyle(ECellStyle.CELLReadOnlyDataStyle)
             
                ofwTable.AddCellAsHTML m_ofwHtmlRef.Hyperlink( _
                        gk_sLinkDisplayAuthorityTerms & sParameter, _
                        ofwRlResult.GetFieldValue("leaflet_name"), gk_sMainFrame), typCustomCell
                
                ofwTable.AddCellAsHTML VariantDateAsDDMMMYYYY(ofwRlResult.GetFieldValue( _
                        gk_sDBCOLTempLeafletCreationDate)), typCustomCell
                        
    
                ofwTable.AddCellAsHTML ofwRlResult.GetFieldValue( _
                        gk_sNumberOfApprovalsRequired), typCustomCell
                                
                ofwTable.EndRow
                
                ofwRlResult.MoveNext
        Loop
            
    End If
    
    'Get the table details
    
    sHTML = sHTML & m_ofwHtmlRef.Paragraph(ofwTable.GetHTML)
    
    'Write the html
    m_aspResponse.Write sHTML
'**********************end of leaflets.


'****Invisible terms
    sHeading = ""
    sHeading = m_ofwHtmlRef.Paragraph(m_ofwHtmlRef.Bold(m_oFWAppData.GetDisplayInformation( _
                            EAuthorityFilesMaintResourceFile.afmrInvisibleTerms) & "..."))
    sHeading = sHeading & m_ofwHtmlRef.Paragraph("")
    sHTML = sHeading

'Get the actual data row from backend.
    
    Set ofwRlResult = m_oboAuthorityFiles.GetInvisibleIndexTermsCount()

    'This error condition would have been handled in the business object itself.
    'Just keep it here too for error handling in VWclass.
    If ofwRlResult Is Nothing Then
        RaiseError sSource, PEGeneralInternalError, ""
    End If

'Display InVisible terms.
     sParameter = "?" & ksTermCategory & "=" & klInvisibleType
     
     
    If Not ofwRlResult.GetFieldValue(gk_sDBCOLTempCounts) = 0 Then
        sListInvisibleTerms = m_ofwHtmlRef.Hyperlink(gk_sLinkDisplayAuthorityTerms & sParameter, _
                     m_ofwHtmlRef.Font(m_ofwHtmlRef.Bold(ofwRlResult.GetFieldValue(gk_sDBCOLTempCounts)) _
                            & " " & m_ofwHtmlRef.Bold(m_oFWAppData.GetDisplayInformation( _
                            EAuthorityFilesMaintResourceFile.afmrInvisibleTerms)) & " "), gk_sMainFrame)
    Else
        sListInvisibleTerms = m_ofwHtmlRef.Font(m_ofwHtmlRef.Bold("No " & m_oFWAppData.GetDisplayInformation( _
                            EAuthorityFilesMaintResourceFile.afmrInvisibleTerms) & ". "))
    End If
                            
    sListInvisibleTerms = m_ofwHtmlRef.List(sListInvisibleTerms & _
                                m_ofwHtmlRef.Font(m_oFWAppData.GetDisplayInformation( _
                            EAuthorityFilesMaintResourceFile.afmrInvisibleComment)))
                            
    sHTML = sHTML & m_ofwHtmlRef.Bullet(sListInvisibleTerms)
    m_aspResponse.Write sHTML
                            
'Display InComplete terms.
'****Incomplete terms
    sHeading = ""
    sHeading = m_ofwHtmlRef.Paragraph(m_ofwHtmlRef.Bold(m_oFWAppData.GetDisplayInformation( _
                            EAuthorityFilesMaintResourceFile.afmrIncompleteTerms) & "..."))
    sHeading = sHeading & m_ofwHtmlRef.Paragraph("")
    sHTML = sHeading

'Get the actual data row from backend.
    
    Set ofwRlResult = m_oboAuthorityFiles.GetIncompleteIndexTermsCount()

    'This error condition would have been handled in the business object itself.
    'Just keep it here too for error handling in VWclass.
    If ofwRlResult Is Nothing Then
        RaiseError sSource, PEGeneralInternalError, ""
    End If
                            
                            
    sParameter = "?" & ksTermCategory & "=" & klIncompleteType
    If Not ofwRlResult.GetFieldValue(gk_sDBCOLTempCounts) = 0 Then
        sListIncompleteTerms = m_ofwHtmlRef.Hyperlink(gk_sLinkDisplayAuthorityTerms & sParameter, _
                     m_ofwHtmlRef.Font(m_ofwHtmlRef.Bold(ofwRlResult.GetFieldValue(gk_sDBCOLTempCounts)) _
                            & " " & m_ofwHtmlRef.Bold(m_oFWAppData.GetDisplayInformation( _
                            EAuthorityFilesMaintResourceFile.afmrIncompleteTerms) & ". ")), gk_sMainFrame)
    Else
        sListIncompleteTerms = m_ofwHtmlRef.Font(m_ofwHtmlRef.Bold("No " & m_oFWAppData.GetDisplayInformation( _
                            EAuthorityFilesMaintResourceFile.afmrIncompleteTerms) & ". "))
    End If
    sListIncompleteTerms = m_ofwHtmlRef.List(sListIncompleteTerms & _
                                m_ofwHtmlRef.Font(m_oFWAppData.GetDisplayInformation( _
                            EAuthorityFilesMaintResourceFile.afmrInCompleteComment)))

    
    sHTML = sHTML & m_ofwHtmlRef.Bullet(sListIncompleteTerms)

    m_aspResponse.Write sHTML
    '****End of incomplete terms
    
    
    'd14
    'Display Non-candidate non-preferred terms.
    sHeading = ""
    sHeading = m_ofwHtmlRef.Paragraph(m_ofwHtmlRef.Bold(m_oFWAppData.GetDisplayInformation( _
                            EAuthorityFilesMaintResourceFile.afmrLiveNonPreferred) & "..."))
    sHeading = sHeading & m_ofwHtmlRef.Paragraph("")
    sHTML = sHeading
    'Get the actual data row from backend.
    Set ofwRlResult = m_oboAuthorityFiles.GetNonCandidateNonPreferredTermsCount()

    'This error condition would have been handled in the business object itself.
    'Just keep it here too for error handling in VWclass.
    If ofwRlResult Is Nothing Then
        RaiseError sSource, PEGeneralInternalError, ""
    End If
                            
                            
    sParameter = "?" & ksTermCategory & "=" & klNonCandidateNonPreferredType
                            
    If Not ofwRlResult.GetFieldValue(gk_sDBCOLTempCounts) = 0 Then
        sNonCandidateNonPreferredTerms = m_ofwHtmlRef.Hyperlink(gk_sLinkDisplayAuthorityTerms & sParameter, _
                     m_ofwHtmlRef.Font(m_ofwHtmlRef.Bold(ofwRlResult.GetFieldValue(gk_sDBCOLTempCounts)) _
                            & " " & m_ofwHtmlRef.Bold(m_oFWAppData.GetDisplayInformation( _
                            EAuthorityFilesMaintResourceFile.afmrLiveNonPreferred) & ".")), gk_sMainFrame)
    Else
        sNonCandidateNonPreferredTerms = m_ofwHtmlRef.Font(m_ofwHtmlRef.Bold("No " & m_oFWAppData.GetDisplayInformation( _
                            EAuthorityFilesMaintResourceFile.afmrLiveNonPreferred) & ". "))
    End If

    sNonCandidateNonPreferredTerms = m_ofwHtmlRef.List(sNonCandidateNonPreferredTerms & _
                                m_ofwHtmlRef.Font(m_oFWAppData.GetDisplayInformation( _
                            EAuthorityFilesMaintResourceFile.afmrLiveNonPreferredComment)))

    sHTML = sHTML & m_ofwHtmlRef.Bullet(sNonCandidateNonPreferredTerms)

    m_aspResponse.Write sHTML
    
    '****End of Non-candidate non-preferred terms.
    
    'd14
    'Display not used preferred terms.
    sHeading = ""
    sHeading = m_ofwHtmlRef.Paragraph(m_ofwHtmlRef.Bold(m_oFWAppData.GetDisplayInformation( _
                            EAuthorityFilesMaintResourceFile.afmrNotUsedPreferred) & "..."))
    sHeading = sHeading & m_ofwHtmlRef.Paragraph("")
    sHTML = sHeading
    'Get the actual data row from backend.
    Set ofwRlResult = m_oboAuthorityFiles.GetNotUsedPreferredTermsCount

    'This error condition would have been handled in the business object itself.
    'Just keep it here too for error handling in VWclass.
    If ofwRlResult Is Nothing Then
        RaiseError sSource, PEGeneralInternalError, ""
    End If
                            
                            
    sParameter = "?" & ksTermCategory & "=" & klNotUsedPreferredType
                            
    If Not ofwRlResult.GetFieldValue(gk_sDBCOLTempCounts) = 0 Then
        sNotUsedPreferredTerms = m_ofwHtmlRef.Hyperlink(gk_sLinkDisplayAuthorityTerms & sParameter, _
                     m_ofwHtmlRef.Font(m_ofwHtmlRef.Bold(ofwRlResult.GetFieldValue(gk_sDBCOLTempCounts)) _
                            & " " & m_ofwHtmlRef.Bold(m_oFWAppData.GetDisplayInformation( _
                            EAuthorityFilesMaintResourceFile.afmrNotUsedPreferred) & ". ")), gk_sMainFrame)
    Else
        sNotUsedPreferredTerms = m_ofwHtmlRef.Font(m_ofwHtmlRef.Bold("No " & m_oFWAppData.GetDisplayInformation( _
                            EAuthorityFilesMaintResourceFile.afmrNotUsedPreferred) & ". "))
    End If

    sNotUsedPreferredTerms = m_ofwHtmlRef.List(sNotUsedPreferredTerms & _
                                m_ofwHtmlRef.Font(m_oFWAppData.GetDisplayInformation( _
                            EAuthorityFilesMaintResourceFile.afmrNotUsedPreferredComment)))

    sHTML = sHTML & m_ofwHtmlRef.Bullet(sNotUsedPreferredTerms)

    m_aspResponse.Write sHTML
    
    '****End of not used preferred terms.
    
    'd14
    'Display not used candidate terms.
    sHeading = ""
    sHeading = m_ofwHtmlRef.Paragraph(m_ofwHtmlRef.Bold(m_oFWAppData.GetDisplayInformation( _
                            EAuthorityFilesMaintResourceFile.afmrNotUsedCandidate) & "..."))
    sHeading = sHeading & m_ofwHtmlRef.Paragraph("")
    sHTML = sHeading
    'Get the actual data row from backend.
    Set ofwRlResult = m_oboAuthorityFiles.GetNotUsedCandidateTermsCount

    'This error condition would have been handled in the business object itself.
    'Just keep it here too for error handling in VWclass.
    If ofwRlResult Is Nothing Then
        RaiseError sSource, PEGeneralInternalError, ""
    End If
                            
                            
    sParameter = "?" & ksTermCategory & "=" & klNotUsedCandidateType
    If Not ofwRlResult.GetFieldValue(gk_sDBCOLTempCounts) = 0 Then
        sNotUsedCandidateTerms = m_ofwHtmlRef.Hyperlink(gk_sLinkDisplayAuthorityTerms & sParameter, _
                     m_ofwHtmlRef.Font(m_ofwHtmlRef.Bold(ofwRlResult.GetFieldValue(gk_sDBCOLTempCounts)) _
                            & " " & m_ofwHtmlRef.Bold(m_oFWAppData.GetDisplayInformation( _
                            EAuthorityFilesMaintResourceFile.afmrNotUsedCandidate) & ". ")), gk_sMainFrame)
    Else
        sNotUsedCandidateTerms = m_ofwHtmlRef.Font(m_ofwHtmlRef.Bold("No " & m_oFWAppData.GetDisplayInformation( _
                            EAuthorityFilesMaintResourceFile.afmrNotUsedCandidate) & ". "))
    End If
                            
    sNotUsedCandidateTerms = m_ofwHtmlRef.List(sNotUsedCandidateTerms & _
                                m_ofwHtmlRef.Font(m_oFWAppData.GetDisplayInformation( _
                            EAuthorityFilesMaintResourceFile.afmrNotUsedCandidateComment)))
                            
    sHTML = sHTML & m_ofwHtmlRef.Bullet(sNotUsedCandidateTerms)

    m_aspResponse.Write sHTML
    
    
    
    '****End of not used candidate terms.
    
    'd14
    'Display live terms.
    'Get the actual data row from backend.
    sHeading = ""
    sHeading = m_ofwHtmlRef.Paragraph(m_ofwHtmlRef.Bold(m_oFWAppData.GetDisplayInformation( _
                            EAuthorityFilesMaintResourceFile.afmrLiveTerms) & "..."))
    sHeading = sHeading & m_ofwHtmlRef.Paragraph("")
    sHTML = sHeading
    'Get the actual data row from backend.
    Set ofwRlResult = m_oboAuthorityFiles.GetTermsCountWithStatus(gk_lApprovedComplete)

    'This error condition would have been handled in the business object itself.
    'Just keep it here too for error handling in VW class.
    If ofwRlResult Is Nothing Then
        RaiseError sSource, PEGeneralInternalError, ""
    End If
                            
                            
    sParameter = "?" & "fldAFGroup" & "=" & EAFTermType.AFAllTypes _
        & "&" & "fldAFListType" & "=" & EAFStatusToDisplay.AFSTATUSApprovedOnly _
        & "&" & gk_sFORMAFListUsage & "=" & BROWSEEditorialEditLiveTerms
                            
    If Not ofwRlResult.GetFieldValue(gk_sDBCOLTempCounts) = 0 Then
        sLiveTerms = m_ofwHtmlRef.Hyperlink(gk_sIndexBrowserAFPickTermPage & sParameter, _
                     m_ofwHtmlRef.Font(m_ofwHtmlRef.Bold(ofwRlResult.GetFieldValue(gk_sDBCOLTempCounts)) _
                            & " " & m_ofwHtmlRef.Bold(m_oFWAppData.GetDisplayInformation( _
                            EAuthorityFilesMaintResourceFile.afmrLiveTerms) & ". ")), gk_sMainFrame)
    Else
        sLiveTerms = m_ofwHtmlRef.Font(m_ofwHtmlRef.Bold("No " & m_oFWAppData.GetDisplayInformation( _
                            EAuthorityFilesMaintResourceFile.afmrLiveTerms) & ". "))
    End If
                            
    sLiveTerms = m_ofwHtmlRef.List(sLiveTerms & _
                                m_ofwHtmlRef.Font(m_oFWAppData.GetDisplayInformation( _
                            EAuthorityFilesMaintResourceFile.afmrLiveTermsComment)))
                            
    sHTML = sHTML & m_ofwHtmlRef.Bullet(sLiveTerms)

    'Add context sensitive help button
    sHTML = sHTML & HTMLToAddHelpButton(gk_sHelpMAFMainPage)

    m_aspResponse.Write sHTML
    
    '****End of live terms.
    
            
    
    'Release the objects
    
    Set ofwRlResult = Nothing
    Set ofwTable = Nothing
        
    m_objContext.SetComplete

Exit Sub

ErrorHandler:
    ' Only use SetAbort if this method did an update,
    ' (and then there should have been no transient storage).
    ' Otherwise use SetComplete
    m_objContext.SetComplete
    
    ShowError sSource
End Sub


Public Sub DisplayIndexTerms()
'
'This procedure will be called from VwDisplayAuthorityTerms ASP page to
'display index terms raised through particular editset/leaflet/live etc
'
'   d1  Geedha Sivanadian     30Jul1999     Created
'   d5  AGA                   08Oct1999     Modifications reg. display of buttons
'   d11 AGA                   04Nov1999     Added sessions to DisplayIndexTerms used in case of pressing cancel
'   d31 TMJ                   20Mar2000     Added code that improved re-direction when editing live terms
'


On Error GoTo ErrorHandler
                 
    'Recordssets
    Dim ofwRlResult As FWRecordList
    
    'Records
    Dim ofwRResult As FWRecord
    'Table
    Dim ofwTable As FWTable
    
    'Table, row, cell styles
    Dim typTable As TableStyle
    Dim typCustomStyle As RowStyle
    Dim typCustomCell As CellStyle
    Dim typCellStyle  As CellStyle
    
    
    'Strings to store variables.
    Dim sSource As String
    Dim sHTML As String
    Dim sHeading As String
    Dim sParameter As String
    Dim sEditSetOrLeafLetName As String
    Dim lEditSetLeafLetId  As Long
    Dim lTermCategoryType As Long
    Dim sText As String
   

    sSource = "VWEditorialAuthorityFiles.DisplayIndexTerms"
    
    'Check the permissions
    CheckPermissions ESecurityCheck.EUserCanMaintainAuthorityFiles, sSource, ""
    
    'Initialise to -1'
    lTermCategoryType = -1
    ' Title
 
 '********** Get PARAMETER values *******************'
 'Get the values from prev page. Check for edit set id or leaflet id based on termtype
  
    
  
   'Set any session variables to be used in AF Maintenance
   If Not m_aspRequest(ksTermCategory) = "" Then
        m_aspSession(ksTermCategory) = m_aspRequest(ksTermCategory)
        lTermCategoryType = m_aspRequest(ksTermCategory)
   Else
        'If term type is blank... redirect to the main page
        m_aspSession(ksTermCategory) = klLiveType
        lTermCategoryType = klLiveType
   End If
  
  Select Case lTermCategoryType
    
    Case klEditSetType
  
              'd11
              If m_aspRequest(gk_sNameEditsetID) = "" Then
                If m_aspSession(gk_sNameEditsetID) = "" Then
                   RaiseError sSource, EPROCATINTERNALERROR.peAuthorityFilesMaintenance, _
                               "Editset ID not found"
                Else
                    lEditSetLeafLetId = CLng(m_aspSession(gk_sNameEditsetID))
                    
                End If
              Else
               
               lEditSetLeafLetId = CLng(m_aspRequest(gk_sNameEditsetID))
               m_aspSession(gk_sNameEditsetID) = lEditSetLeafLetId
               
              End If
               
               If lEditSetLeafLetId <= 0 Then
                   RaiseError sSource, EPROCATINTERNALERROR.peAuthorityFilesMaintenance, _
                                       "Editset ID is not a valid id"
               End If
            
                'Get the value for editset name
                'd11
               If m_aspRequest(gk_sNameEditSetName) = "" Then
                    If m_aspSession(gk_sNameEditSetName) = "" Then
                        RaiseError sSource, EPROCATINTERNALERROR.peEditSet, _
                               "Editset name is blank"
                    Else
                        sEditSetOrLeafLetName = m_aspSession(gk_sNameEditSetName)
                        
                    End If
               Else
                sEditSetOrLeafLetName = m_aspRequest(gk_sNameEditSetName)
                m_aspSession(gk_sNameEditSetName) = sEditSetOrLeafLetName
               End If
               
               
               
       'Get the actual data rows from backend.
        Set ofwRlResult = m_oboAuthorityFiles.GetEditSetIndexTerms(lEditSetLeafLetId)
        
        sHeading = m_ofwHtmlRef.Paragraph(m_ofwHtmlRef.Font(m_oFWAppData.GetDisplayInformation( _
                       EAuthorityFilesMaintResourceFile.afmrIndexTermsDetailHeading1) & _
                       " " & m_ofwHtmlRef.Bold(sEditSetOrLeafLetName), COLOURBlack))
                           
        sHeading = sHeading & ofwRlResult.RecordCount & " " & _
                        m_ofwHtmlRef.Font(m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrIndexTermsDetailHeading2))
                       
    Case klleafletType
            'd11
            If m_aspRequest(gk_sSFAReferenceID) = "" Then
                If m_aspSession(gk_sSFAReferenceID) = "" Then
                   RaiseError sSource, EPROCATINTERNALERROR.peAuthorityFilesMaintenance, _
                               "SFA ID not found"
                Else
                    lEditSetLeafLetId = CLng(m_aspSession(gk_sSFAReferenceID))
                End If
            Else
                lEditSetLeafLetId = CLng(m_aspRequest(gk_sSFAReferenceID))
                m_aspSession(gk_sSFAReferenceID) = lEditSetLeafLetId
            End If
               
               
               
               If lEditSetLeafLetId <= 0 Then
                   RaiseError sSource, EPROCATINTERNALERROR.peAuthorityFilesMaintenance, _
                                       "SFA ID is not a valid id"
               End If
            
                'Get the value for leaflet name
            'd11
            If m_aspRequest(gk_sDBCOLSFAName) = "" Then
                If m_aspSession(gk_sDBCOLSFAName) = "" Then
                   RaiseError sSource, EPROCATINTERNALERROR.peAuthorityFilesMaintenance, _
                               "Leaflet name is blank"
                Else
                   sEditSetOrLeafLetName = m_aspSession(gk_sDBCOLSFAName)
                End If
            Else
               sEditSetOrLeafLetName = m_aspRequest(gk_sDBCOLSFAName)
               m_aspSession(gk_sDBCOLSFAName) = sEditSetOrLeafLetName
            End If
               
            
       
       'Get the actual data rows from backend.
        Set ofwRlResult = m_oboAuthorityFiles.GetLeafLetIndexTerms(lEditSetLeafLetId)
        
        sHeading = m_ofwHtmlRef.Paragraph(m_ofwHtmlRef.Font(m_oFWAppData.GetDisplayInformation( _
                       EAuthorityFilesMaintResourceFile.afmrIndexTermsDetailHeading1) & _
                       " " & m_ofwHtmlRef.Bold(sEditSetOrLeafLetName), COLOURBlack))
                       
        sHeading = sHeading & ofwRlResult.RecordCount & " " & _
                        m_ofwHtmlRef.Font(m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrLeafLetHeading2))
               
    Case klIncompleteType
               'Get the actual data rows from backend.
        Set ofwRlResult = m_oboAuthorityFiles.GetIncompleteIndexTermsDetails
        
        sHeading = sHeading & ofwRlResult.RecordCount & " " & _
                        m_ofwHtmlRef.Font(m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrIncompleteTerms))
        
    Case klInvisibleType
               'Get the actual data rows from backend.
        Set ofwRlResult = m_oboAuthorityFiles.GetInvisibleIndexTermsDetails
        
        sHeading = sHeading & ofwRlResult.RecordCount & " " & _
                        m_ofwHtmlRef.Font(m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrInvisibleTerms))
        
    
    Case klNonCandidateNonPreferredType
               'Get the actual data rows from backend.
        Set ofwRlResult = m_oboAuthorityFiles.GetNonCandidateNonPreferredTermsDetails
        
        sHeading = sHeading & ofwRlResult.RecordCount & " " & _
                        m_ofwHtmlRef.Font(GetTermCategoryDescription(klNonCandidateNonPreferredType))
        
    Case klNotUsedPreferredType
               'Get the actual data rows from backend.
        Set ofwRlResult = m_oboAuthorityFiles.GetNotUsedPreferredTermsDetails
        
        sHeading = sHeading & ofwRlResult.RecordCount & " " & _
                        m_ofwHtmlRef.Font(GetTermCategoryDescription(klNotUsedPreferredType))
    
    Case klNotUsedCandidateType
               'Get the actual data rows from backend.
        Set ofwRlResult = m_oboAuthorityFiles.GetNotUsedCandidateTermsDetails
        
        sHeading = sHeading & ofwRlResult.RecordCount & " " & _
                        m_ofwHtmlRef.Font(GetTermCategoryDescription(klNotUsedCandidateType))
    'd31 Added Case 8
    Case klLiveType
        sParameter = "?" & "fldAFGroup" & "=" & EAFTermType.AFAllTypes _
            & "&" & "fldAFListType" & "=" & EAFStatusToDisplay.AFSTATUSApprovedOnly _
            & "&" & gk_sFORMAFListUsage & "=" & BROWSEEditorialEditLiveTerms
        m_aspResponse.Redirect gk_sIndexBrowserAFPickTermPage & sParameter
    Case Else
        'If someone comes from Index browser (incase of live terms)
        m_aspResponse.Redirect gk_sLinkEditorialAFMaintMain
        m_objContext.SetComplete
        Exit Sub
    End Select
            
    
    'This error condition would have been handled in the business object itself.
    'Just keep it here too for error handling in VWclass.
    If ofwRlResult Is Nothing Then
        RaiseError sSource, PEGeneralInternalError, ""
    End If
    
    sHeading = sHeading & m_ofwHtmlRef.Paragraph("")
    sHTML = sHeading
    
    
   'First section for retrieving the index terms raised through edit sets.
        
    'Create table
    Set ofwTable = New FWTable
        
    'Table settings
    typTable.eStyle = TABCustomStyle
    typTable.lCELLPADDING = 2
    typTable.lCELLSPACING = 1
    typTable.sWidth = "100%"
    
    ofwTable.TableStyle = typTable
    
    ' To display column-captions.
    ofwTable.DefaultRowStyle = m_ofwHtmlRef.GetRowStyle(RowBackgroundPink)
    typCustomStyle = m_ofwHtmlRef.GetRowStyle(RowBackgroundRed)
    
    'Last column alone should have white backgeound.
    
    typCellStyle = m_ofwHtmlRef.GetCellStyle(CELLReadOnlyLabelStyle)
      
    'First row to display column captions.
    ofwTable.StartRow (typCustomStyle)
        
            'typCustomCell.eStyle = CELLCustomStyle
            typCustomCell = m_ofwHtmlRef.GetCellStyle(CELLColumnHeadingStyle)
            
            ofwTable.AddHCellAsHTML m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrType), typCustomCell
            
            ofwTable.AddHCellAsHTML m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrReference), typCustomCell
        
        
            ofwTable.AddHCellAsHTML m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrCreatedBy), typCustomCell
         
            
            ofwTable.AddHCellAsHTML m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrStatus), typCustomCell
                        
          
    ofwTable.EndRow
    
'Display actual rows retrieved from backend.
    typCustomCell = m_ofwHtmlRef.GetCellStyle(CELLReadOnlyDataStyle)
    ' Display all the editsets one by one.
    If ofwRlResult.RecordCount <> 0 Then
    
       ofwRlResult.MoveFirst
       
       Do While Not ofwRlResult.IsAtEnd
    
            sParameter = "?" & gk_sNameEditsetID & "=" & m_aspRequest(gk_sNameEditsetID) & "&" & gk_sFORMAFTermID & "=" & ofwRlResult.GetFieldValue(gk_sTermID) & _
                        "&" & gk_sFORMAFGroup & "=" & ofwRlResult.GetFieldValue(gk_sTermType) & _
                        "&" & ksTermCategory & "=" & lTermCategoryType & _
                        "&" & gk_sFORMAFNoPersist & "=1"
                        
                        
             ofwTable.StartRow
                
                typCustomCell.sWidth = "13%"
                ofwTable.AddCellAsHTML GetTermTypeDescription(ofwRlResult.GetFieldValue(gk_sTermType)), typCustomCell
                typCustomCell.sWidth = "45%"
                ofwTable.AddCellAsHTML VariantStringAsString(ofwRlResult.GetFieldValue( _
                        gk_sDBCOLReference)), typCustomCell
                
                typCustomCell.sWidth = "12%"
                ofwTable.AddCellAsHTML VariantStringAsString(ofwRlResult.GetFieldValue( _
                        gk_sDBCOLUserName)), typCustomCell
                
                typCustomCell.sWidth = "15%"
                ofwTable.AddCellAsHTML VariantStringAsString(ofwRlResult.GetFieldValue( _
                        gk_sDBCOLStatus)), typCustomCell
                        
                'd5   Do not show the edit link for not used terms
                If lTermCategoryType <> klNotUsedType Then
                    typCustomCell.sWidth = "15%"
                    ofwTable.AddCellAsHTML m_ofwHtmlRef.Hyperlink( _
                        gk_sLinkDisplaySpecificIndexTermDetails & sParameter, _
                        m_oFWAppData.GetDisplayInformation( _
                       EAuthorityFilesMaintResourceFile.afmrEditThisTerm), gk_sMainFrame, , m_oFWAppData.GetDisplayInformation( _
                       EAuthorityFilesMaintResourceFile.afmrEditThisTerm)), typCellStyle
                Else
                    ofwTable.AddHCellAsHTML gk_sHTMLNonBreakSpace, typCellStyle
                End If
                ofwTable.EndRow
                
                ofwRlResult.MoveNext
        Loop
        
        'Display done button
        ofwTable.StartRow
        'Display blank columns
        ofwTable.AddCellAsHTML gk_sHTMLNonBreakSpace, typCellStyle
        If lTermCategoryType = klEditSetType Then
            ofwTable.AddCellAsHTML gk_sHTMLNonBreakSpace, typCellStyle
        End If
        ofwTable.AddCellAsHTML gk_sHTMLNonBreakSpace, typCellStyle
        ofwTable.AddCellAsHTML gk_sHTMLNonBreakSpace, typCellStyle
        'Image for the done button
        ofwTable.AddHCellAsHTML m_ofwHtmlRef.ImageButton(gk_sImagesPath & gk_sImageDone, _
                "", gk_sLinkEditorialAFMaintMain, gk_sMainFrame), typCellStyle
                
        ofwTable.EndRow
    End If
    
    'Get the table details
    
    sHTML = sHTML & m_ofwHtmlRef.Paragraph(ofwTable.GetHTML)
        
    'Add context sensitive help button
    sHTML = sHTML & HTMLToAddHelpButton(gk_sHelpMAFIndexTermsPage)
    
    'Write the html
    m_aspResponse.Write sHTML
    
    'Release the objects
    
    Set ofwRlResult = Nothing
    Set ofwTable = Nothing
        
    m_objContext.SetComplete

Exit Sub

ErrorHandler:
    ' Only use SetAbort if this method did an update,
    ' (and then there should have been no transient storage).
    ' Otherwise use SetComplete
    m_objContext.SetComplete
    
    ShowError sSource
End Sub


Public Sub DisplaySpecificIndexTermDetails()
'
'This procedure will be called from VwDisplayAuthorityTerms ASP page to
'display index term details and to update specific index term details.
'
'   d1  Geedha Sivanadian     2Aug1999   Created
'   d5  AGA                        08Oct1999  Modified to include persistance

On Error GoTo ErrorHandler
    
    'Recordssets
    'Dim ofwRlResult As FWRecordList
    
    'Table
    Dim ofwTable As FWTable
    
    'Table, row, cell styles
    Dim typTable As TableStyle
    Dim typCustomCell As CellStyle
    Dim typCellStyle  As CellStyle
       
    Dim sSource As String
    Dim lSpecificTermID As Long
    Dim sParameter As String
    Dim sHeading As String
    Dim sHTML As String
    
    sSource = "VWAuthorityFile.DisplaySpecificIndexTermDetails"
    
    'Check the permissions
    CheckPermissions ESecurityCheck.EUserCanMaintainAuthorityFiles, sSource, ""
     
    If m_aspRequest(gk_sFORMAFGroup) = "" Then
    
            RaiseError sSource, EPROCATINTERNALERROR.peAuthorityFilesMaintenance, _
                           "Term type is blank"
    End If
    m_lSpecificTermType = CLng(m_aspRequest(gk_sFORMAFGroup))
    If m_aspRequest(ksFORMCancel) = "" Then
        'Get the value of term id
        lSpecificTermID = CLng(m_aspRequest(gk_sFORMAFTermID))
        If lSpecificTermID <= 0 Then
    
                 RaiseError sSource, EPROCATINTERNALERROR.peAuthorityFilesMaintenance, _
                                           "Specific Term ID is not a valid id"
        End If
    Else
        lSpecificTermID = CLng(m_aspRequest(gk_sFORMAFTermID))
    End If
    
    'Get the value of the m_lIndexTermRaisedSource which indicates whether this was raised
    'through a leaflet or editset. This is required to putback the user to the same screen.
    'd18
    If m_aspSession(ksTermCategory) = "" Then
        'If the session has no value send the user to the homepage...
        m_lIndexTermRaisedSource = klLiveType
    Else
        m_lIndexTermRaisedSource = m_aspSession(ksTermCategory)
        If m_lIndexTermRaisedSource < 0 Then
            RaiseError sSource, EPROCATINTERNALERROR.peAuthorityFilesMaintenance, _
                           "Index term raised type is not valid"
        End If
    End If
    'Check if the data is coming from persistance object
    If m_aspRequest(gk_sFORMAFNoPersist) = 1 Then
        m_bNoPersist = True
    End If
    
    Select Case m_lSpecificTermType
    'Person details.
        Case EAFTermType.AFPerson
            
           Set ofwTable = DisplayPersonDetails(lSpecificTermID)
           
           sParameter = "?" & gk_sFORMAFGroup & "=" & EAFTermType.AFPerson & "&" & gk_sFORMAFTermID & "=" & _
                        lSpecificTermID & "&" & k_sIndexTermRaisedThrough & "=" & m_lIndexTermRaisedSource
           'SEt the state
           m_oBOAFPerson.PROCATTransientInterface_SetState
            
     'Subject details.
        Case EAFTermType.AFSubject
           
           Set ofwTable = DisplaySubjectDetails(lSpecificTermID)
           
            sParameter = "?" & gk_sFORMAFGroup & "=" & EAFTermType.AFSubject & "&" & gk_sFORMAFTermID & "=" & _
                        lSpecificTermID & "&" & k_sIndexTermRaisedThrough & "=" & m_lIndexTermRaisedSource
            ' SEt the state
            m_oBOAFSubject.PROCATTransientInterface_SetState
            
        Case EAFTermType.AFPlace
           Set ofwTable = DisplayPlaceDetails(lSpecificTermID)
           
            sParameter = "?" & gk_sFORMAFGroup & "=" & EAFTermType.AFPlace & "&" & gk_sFORMAFTermID & "=" & _
                        lSpecificTermID & "&" & k_sIndexTermRaisedThrough & "=" & m_lIndexTermRaisedSource
            'SEt the state
            m_oBOAFPlace.PROCATTransientInterface_SetState
            
        Case EAFTermType.AFCorporateBody
           Set ofwTable = DisplayCorporateBodyDetails(lSpecificTermID)
           
            sParameter = "?" & gk_sFORMAFGroup & "=" & EAFTermType.AFCorporateBody & "&" & gk_sFORMAFTermID & "=" & _
                        lSpecificTermID & "&" & k_sIndexTermRaisedThrough & "=" & m_lIndexTermRaisedSource
            'SEt the state
            m_oBOAFCorporateBody.PROCATTransientInterface_SetState
            
        Case Else
                  RaiseError sSource, EPROCATINTERNALERROR.peAuthorityFilesMaintenance, _
                           "Term type is not valid one"
    End Select
    
    
    'Get the table details
    
'Formstart
    sHTML = sHTML & m_ofwHtmlRef.FormStart(gk_sLinkUpdateIndexTerms & sParameter, gk_sMainFrame, _
                                                gk_sHTMLDefaultFormName, MPost, True)
    'Store the term type etc as hidden fields:
    sHTML = sHTML & m_ofwHtmlRef.HiddenField(gk_sFORMAFGroup, CStr(m_lSpecificTermType))
    
    ' Title
    If Not m_lIndexTermRaisedSource = gk_lNotApproved Then
    
        sHeading = m_ofwHtmlRef.Paragraph(m_ofwHtmlRef.Font( _
                m_ofwHtmlRef.Bold(m_oFWAppData.GetDisplayInformation( _
                       EAuthorityFilesMaintResourceFile.afmrSpecificIndexTermHeading)), COLOURBlack, 4))
    Else
        sHeading = m_ofwHtmlRef.Paragraph(m_ofwHtmlRef.Font( _
                m_ofwHtmlRef.Bold("Index terms"), COLOURBlack, 4))
    End If
    
    sHeading = sHeading & m_ofwHtmlRef.Paragraph("")
    
    sHTML = sHTML & gk_sHTMLLineBreak & sHeading
    
    sHTML = sHTML & m_ofwHtmlRef.Paragraph(ofwTable.GetHTML)
    
    sHTML = sHTML & m_ofwHtmlRef.EndTag(gk_sHTMLForm)
            
    'Add context sensitive help button
    sHTML = sHTML & HTMLToAddHelpButton(gk_sHelpMAFTermDetailsPage)
    
    'Write the html
    m_aspResponse.Write sHTML
    
    'Release the objects
    
    'Set ofwRlResult = Nothing
    Set ofwTable = Nothing
        
    m_objContext.SetComplete

Exit Sub

ErrorHandler:
    ' Only use SetAbort if this method did an update,
    ' (and then there should have been no transient storage).
    ' Otherwise use SetComplete
    m_objContext.SetComplete
    
    ShowError sSource
End Sub

Private Function DisplayPersonDetails(t_lPersonReferenceID As Long) As FWTable
'This procedure will be called from DisplaySpecificTerm procedure.
'This procedure will be called to display person index term page.
'
'   d1  Geedha Sivanadian   2Aug1999    Created
'   d5  AGA                 08Oct1999   Added functionality of See/See also links
'                                       Modifications reg. display of buttons
'   d16 AGA                 16Dec1999   Display message if a see link is added to the term in catelogue
'   d26 wdp                 23Feb2000     Sorted display of see and see also links and buttons
'                                           for various status of terms(preferred, non-preferred etc.)
'   d32 tmj                 21Mar2000   Changed gk_sFORMAFTermValue to gk_sFORMAdvSIndexName
'   d34 tmj                 28Mar2000   Added textarea validation
'   d35 daj                   17Jul2000     Changed the checking in Display<AFType>Details functions so that the links
'                                           currently stored in TransientStore are used (rather than the database) to decide
'                                           which of the See and See Also buttons are shown

On Error GoTo ErrorHandler
    
    'Table
    Dim ofwTable As FWTable
    
    'Table, row, cell styles
    Dim typTable As TableStyle
    Dim typCellLabelStyle As CellStyle
    Dim typCellInputStyle  As CellStyle
    Dim typCellLinks  As CellStyle
    Dim typCellRemove  As CellStyle
    
    Dim sSource As String
    Dim sHTML As String
    Dim bMaleChecked As Boolean
    
    Dim sParameter As String
    Dim SApproveCompleteStr As String
    Dim SApproveInCompleteStr As String
    Dim SRejectStr As String
    
    Dim sTempSeeLinks As String
    
    Dim sMessage As String 'd16
    
    sSource = "VWEditorialAuthorityFiles.DisplayPersonDetails"
    'Check the permissions
    CheckPermissions ESecurityCheck.EUserCanMaintainAuthorityFiles, sSource, ""
    
    Set ofwTable = New FWTable
    
    typTable = m_ofwHtmlRef.GetTableStyle(TABContentStyle)
    'Set the tables format
    typTable.sWidth = "100%"
    ofwTable.TableStyle = typTable
    
    
    'Set the label columns style
    typCellLabelStyle = m_ofwHtmlRef.GetCellStyle(ECellStyle.CELLReadWriteLabelStyle)
    typCellLabelStyle.sWidth = "30%"
    
    'Set the text columns style.
    typCellInputStyle = m_ofwHtmlRef.GetCellStyle(ECellStyle.CELLReadWriteDataStyle)
    typCellInputStyle.sWidth = "70%"
    typCellInputStyle.lCOLSPAN = 4
       
    typCellLinks = m_ofwHtmlRef.GetCellStyle(ECellStyle.CELLReadWriteDataStyle)
    typCellLinks.sWidth = "20%"
    typCellLinks.lCOLSPAN = 3
    
    
    typCellRemove = m_ofwHtmlRef.GetCellStyle(ECellStyle.CELLReadWriteDataStyle)
    typCellRemove.sWidth = "10%"
    typCellRemove.lCOLSPAN = 1
    
    
    If m_aspRequest(mk_sFORMAFTermAddOrRemove) = EAFTermLoadType.EAFLoadFromStateRemoveTerm Then
        'Load from persistant object (After removing a term)
            m_oBOAFPerson.Load EAFTermLoadType.EAFLoadFromStateRemoveTerm, t_lPersonReferenceID, _
                CInt(m_aspSession(mk_sFORMAFRemoveTermID)), "", m_aspSession(mk_sFORMAFSeeLinkType)
            m_oBOAFPerson.PROCATTransientInterface_SetState
    
    Else
        If m_bNoPersist = True Then
            'Lodad from the database
            m_oBOAFPerson.Load EAFLoadFromDatabase, t_lPersonReferenceID
            m_oBOAFPerson.PROCATTransientInterface_SetState

        Else
            If m_aspRequest(ksFORMCancel) = 1 Then
                m_oBOAFPerson.Load EAFTermLoadType.EAFLoadFromStateTermNoChange
            Else
                'd16
                'Check if the term is raised from leaflet or editset: And if someone try to add a see link to it
                If m_aspSession(mk_sFORMAFSeeLinkType) = EAFSeeLinkType.EAFSeeLink And _
                    (m_lIndexTermRaisedSource = klleafletType Or m_lIndexTermRaisedSource = klEditSetType) Then
                   
                    'Display this message in the see links box
                    sMessage = m_ofwHtmlRef.Font("Cannot convert to a non-preferred term while links exist", COLOURRed)
                    
                    m_oBOAFPerson.Load EAFTermLoadType.EAFLoadFromStateTermNoChange
                    m_oBOAFPerson.PROCATTransientInterface_SetState
                Else
                    '   LOAD FROM PERSISTANCE OBJECT (After adding a term)
                    'd32
                    m_oBOAFPerson.Load EAFTermLoadType.EAFLoadFromStateAddTerm, t_lPersonReferenceID, _
                    CLng(m_aspRequest(gk_sFORMAFTermID)), CStr(m_aspRequest(gk_sFORMAdvSIndexName)), m_aspSession(mk_sFORMAFSeeLinkType)
                    
                    If m_oBOAFPerson.IsTermHasLinks = True And _
                        m_aspSession(mk_sFORMAFSeeLinkType) = EAFSeeLinkType.EAFSeeLink Then
                        'Display this message in the see links box
                        sMessage = m_ofwHtmlRef.Font("Cannot convert to a non-preferred term while links exist", COLOURRed)
                    End If
                    
                    m_oBOAFPerson.PROCATTransientInterface_SetState
                End If
                
            End If
        End If
        
    End If
    
    m_typPersonData = m_oBOAFPerson.Data
    Set m_colSeeLinks = m_oBOAFPerson.SeeLinks
    Set m_colSeeAlsoLinks = m_oBOAFPerson.SeeAlsoLinks
        
    
     'First Row
    
     ofwTable.StartRow
     'Label
     ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                         EAuthorityFilesMaintResourceFile.afmrSurname), typCellLabelStyle
     'Text Field
     
     ofwTable.AddCellAsHTML m_ofwHtmlRef.InputField(ITText, gk_sDBCOLSurname_Text, 20, 20, _
                     VariantStringAsString(m_typPersonData.sSurname), VCNone, True), typCellInputStyle
     ofwTable.EndRow
     
    'Third Row
     ofwTable.StartRow
     'Label
     
     ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                         EAuthorityFilesMaintResourceFile.afmrGenderIndicator), typCellLabelStyle
                         
      bMaleChecked = (m_typPersonData.sGender = gk_sMaleGender)
                          
     'Add two radio buttons with the same name to indicate gender.
     'One for Male
     
     ofwTable.AddCellAsHTML m_ofwHtmlRef.RadioButton("Male", gk_sDBCOLGender_Indicator, bMaleChecked, _
                    gk_sMaleGender) & _
                     m_ofwHtmlRef.RadioButton("Female", gk_sDBCOLGender_Indicator, Not bMaleChecked, _
                     gk_sFemaleGender), typCellInputStyle
     ofwTable.EndRow
     
    'Fourth Row
     ofwTable.StartRow
     'Label
     ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                         EAuthorityFilesMaintResourceFile.afmrTitle), typCellLabelStyle
     'Text Field
     ofwTable.AddCellAsHTML m_ofwHtmlRef.InputField(ITText, gk_sDBCOLPerson_Title, 80, 80, _
                  VariantStringAsString(m_typPersonData.sTitle)), typCellInputStyle
     ofwTable.EndRow
     
    'Fifth Row
     ofwTable.StartRow
     'Label
     ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                         EAuthorityFilesMaintResourceFile.afmrAdditionalElementsOfName), typCellLabelStyle
     'Text Field
     ofwTable.AddCellAsHTML m_ofwHtmlRef.InputField(ITText, gk_sDBCOLadditional_elements_of_name, 50, 50, _
                     VariantStringAsString(m_typPersonData.sAdditionalNameElements)), typCellInputStyle
     ofwTable.EndRow
     
    'Sixth Row
     ofwTable.StartRow
     'Label
     ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                         EAuthorityFilesMaintResourceFile.afmrPretitle), typCellLabelStyle
     'Text Field
     ofwTable.AddCellAsHTML m_ofwHtmlRef.InputField(ITText, gk_sDBCOLPretitle_Text, 5, 5, _
                 VariantStringAsString(m_typPersonData.sPretitle)), typCellInputStyle
     ofwTable.EndRow
     
    'Seventh Row
     ofwTable.StartRow
     'Label
     ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                         EAuthorityFilesMaintResourceFile.afmrForename), typCellLabelStyle
     'Text Field
     ofwTable.AddCellAsHTML m_ofwHtmlRef.InputField(ITText, gk_sDBCOLForename_Text, 50, 50, _
                    VariantStringAsString(m_typPersonData.sForenames)), typCellInputStyle
     ofwTable.EndRow
         
    'Uncertain birth date code
     ofwTable.StartRow
     'Label
     ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                         EAuthorityFilesMaintResourceFile.afmrUncertainBirthdateCode), typCellLabelStyle
     'Text Field
     ofwTable.AddCellAsHTML m_ofwHtmlRef.InputField(ITText, gk_sDBCOLUncertain_Birth_Date_Code, 3, 3, _
                     VariantStringAsString(m_typPersonData.sUncertainBirthDateCode)), typCellInputStyle
     ofwTable.EndRow
     
    'Birthdate
     ofwTable.StartRow
     'Label
     ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                         EAuthorityFilesMaintResourceFile.afmrDateOfBirth), typCellLabelStyle
     'Text Field
     ofwTable.AddCellAsHTML m_ofwHtmlRef.InputField(ITText, gk_sDBCOLBirth_Date, gk_iDBCOLLenDateInAF, gk_iDBCOLLenDateInAF, _
                     VariantStringAsString(m_typPersonData.sBirthDate), VCNumber, , , 0, 9999), typCellInputStyle
     ofwTable.EndRow
     
    
     
    'Twelth Row
     ofwTable.StartRow
     'Label
     ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                         EAuthorityFilesMaintResourceFile.afmrBiogHistory), typCellLabelStyle
     'Text Field
     ofwTable.AddCellAsHTML m_ofwHtmlRef.TextArea(gk_sDBCOLBiog_History_Text, _
                     VariantStringAsString(m_typPersonData.sBiogHistory), 40, 4, gk_iDBCOLLenBiogHistory), _
                         typCellInputStyle
     ofwTable.EndRow
     
    'Thirthteen Row
     ofwTable.StartRow
     'Label
     ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                         EAuthorityFilesMaintResourceFile.afmrValidation), typCellLabelStyle
     'Text Field
     ofwTable.AddCellAsHTML m_ofwHtmlRef.InputField(ITText, gk_sDBCOLValidation_text, 50, 50, _
                     VariantStringAsString(m_typPersonData.sValidation)), typCellInputStyle
     ofwTable.EndRow
     
    'Uncertain death date code
     ofwTable.StartRow
     'Label
     ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                         EAuthorityFilesMaintResourceFile.afmrUncertainDeathDateCode), typCellLabelStyle
     'Text Field
     ofwTable.AddCellAsHTML m_ofwHtmlRef.InputField(ITText, gk_sDBCOLUncertain_Death_Date_Code, 3, 3, _
                    VariantStringAsString(m_typPersonData.sUncertainDeathDateCode)), typCellInputStyle
     ofwTable.EndRow
     
    'deathdate
     ofwTable.StartRow
     'Label
     ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                         EAuthorityFilesMaintResourceFile.afmrDeathDate), typCellLabelStyle
     'Text Field
     ofwTable.AddCellAsHTML m_ofwHtmlRef.InputField(ITText, gk_sDBCOLDeath_Date, gk_iDBCOLLenDateInAF, gk_iDBCOLLenDateInAF, _
                     VariantStringAsString(m_typPersonData.sDeathDate), VCNumber, , , 0, 9999), typCellInputStyle
     ofwTable.EndRow
     
     
     'd5
     'See links
      
     'Flags to decide what combination of see and see also
     'buttons we display, there are 3 scenarios:
     '1.    Term is not linked to or from anything - display see and see also buttons
     '2.    Term is a non preferred term - display see link but no button
     '3.    Term is a preferred term or has see also links to/from it - display see also link(s) and button
     
     Dim bIsNonPreferred As Boolean
     Dim bIsPreferredOrLinkedTo As Boolean
     
     If m_colSeeLinks.Count = 0 Then
        bIsNonPreferred = False
     Else
        bIsNonPreferred = True
     End If
     
     ' d35 - Use cache rather than database
     If m_colSeeAlsoLinks.Count = 0 Then
        bIsPreferredOrLinkedTo = False
     Else
        bIsPreferredOrLinkedTo = True
     End If
              
     If Not bIsPreferredOrLinkedTo Or bIsNonPreferred Then
         If m_colSeeLinks.Count = 0 Then
            ofwTable.StartRow
            'Label
            ofwTable.AddCellAsHTML "See ", typCellLabelStyle
            ofwTable.AddCellAsHTML m_ofwHtmlRef.SubmitImage(gk_sImagesPath & gk_sImagePickPerson, _
                mk_sAFSee, mk_sAFSee) & sMessage, typCellInputStyle
            ofwTable.EndRow
         Else
         
            Dim iTemp As Integer
            For iTemp = 1 To m_colSeeLinks.Count
                If iTemp = 1 Then
                    ofwTable.StartRow
                    ofwTable.AddCellAsHTML "See ", typCellLabelStyle
                Else
                    ofwTable.StartRow
                    ofwTable.AddCellAsHTML gk_sHTMLNonBreakSpace, typCellLabelStyle
                End If
                ofwTable.AddCellAsHTML m_colSeeLinks.Item(iTemp).Value, typCellLinks
                ofwTable.AddCellAsHTML m_ofwHtmlRef.SubmitImage(gk_sImagesPath & gk_sImageRemove, _
                mk_sAFSeeRemove & m_colSeeLinks.Item(iTemp).ID, mk_sAFSeeRemove), typCellRemove
                ofwTable.EndRow
            Next
            ofwTable.StartRow
            'Label
            ofwTable.AddCellAsHTML gk_sHTMLNonBreakSpace, typCellLabelStyle
            'Only add the see button if the term is not already a nonpreferred term
            If Not bIsNonPreferred Then
                ofwTable.AddCellAsHTML m_ofwHtmlRef.SubmitImage(gk_sImagesPath & gk_sImagePickPerson, _
                    mk_sAFSee, mk_sAFSee) & sMessage, typCellInputStyle
            End If
            ofwTable.EndRow
        End If
    End If
    
    
    'See also links
    If Not bIsNonPreferred Then
        If m_colSeeAlsoLinks.Count = 0 Then
            ofwTable.StartRow
            'Label
            ofwTable.AddCellAsHTML "See also", typCellLabelStyle
            ofwTable.AddCellAsHTML m_ofwHtmlRef.SubmitImage(gk_sImagesPath & gk_sImagePickPerson, _
                mk_sAFSeeAlso, mk_sAFSeeAlso), typCellInputStyle
            ofwTable.EndRow
        Else
        
            For iTemp = 1 To m_colSeeAlsoLinks.Count
                ofwTable.StartRow
                If iTemp = 1 Then
                    ofwTable.AddCellAsHTML "See also", typCellLabelStyle
                Else
                    ofwTable.AddCellAsHTML gk_sHTMLNonBreakSpace, typCellLabelStyle
                End If
                ofwTable.AddCellAsHTML m_colSeeAlsoLinks.Item(iTemp).Value, typCellLinks
                ofwTable.AddCellAsHTML m_ofwHtmlRef.SubmitImage(gk_sImagesPath & gk_sImageRemove, _
                mk_sAFSeeAlsoRemove & m_colSeeAlsoLinks.Item(iTemp).ID, mk_sAFSeeAlsoRemove), typCellRemove
                 ofwTable.EndRow
                
            Next
            ofwTable.StartRow
            'Label
            ofwTable.AddCellAsHTML gk_sHTMLNonBreakSpace, typCellLabelStyle
            ofwTable.AddCellAsHTML m_ofwHtmlRef.SubmitImage(gk_sImagesPath & gk_sImagePickPerson, _
                mk_sAFSeeAlso, mk_sAFSeeAlso), typCellInputStyle
            ofwTable.EndRow
         End If
     End If
     ofwTable.AddBlankRow
     ofwTable.AddBlankRow
    
     ofwTable.StartRow
     
     ofwTable.AddCellAsHTML gk_sHTMLNonBreakSpace, typCellLabelStyle
     
     'd5 display of buttons
     m_aspSession(mk_sAFAuthorityStatus) = m_typPersonData.lAuthorityStatus
     Dim sImages As String
     If Not m_typPersonData.lAuthorityStatus = gk_lApprovedComplete Then
         sImages = sImages & m_ofwHtmlRef.SubmitImage(gk_sImagesPath & gk_sImageApproveComplete, _
                  k_sApproveCompleteBtn, gk_lApprovedComplete)
     End If
     If Not m_typPersonData.lAuthorityStatus = gk_lApprovedInComplete Then
        If Not m_typPersonData.lAuthorityStatus = gk_lApprovedComplete Then
            sImages = sImages & m_ofwHtmlRef.SubmitImage(gk_sImagesPath & gk_sImageApproveIncomplete, _
                 k_sApproveInCompleteBtn, gk_lApprovedInComplete)
        End If
     End If
     
     If Not m_typPersonData.lAuthorityStatus = gk_lInvisible Then
        If Not m_typPersonData.lAuthorityStatus = gk_lApprovedComplete Then
            sImages = sImages & m_ofwHtmlRef.SubmitImage(gk_sImagesPath & gk_sImageInvisible, _
                 k_sInvisibleBtn, gk_lInvisible)
        End If
     End If
     If Not m_typPersonData.lAuthorityStatus = gk_lRejected Then
         sImages = sImages & m_ofwHtmlRef.SubmitImage(gk_sImagesPath & gk_sImageReject, _
                 k_sRejectBtn, gk_lRejected)
         
     End If
     typCellLabelStyle.lCOLSPAN = 2
     typCellLabelStyle.sWidth = "23%"
     ofwTable.AddCellAsHTML sImages, typCellLabelStyle
     'Show ok/cancel button
     typCellLabelStyle.lCOLSPAN = 2

    typCellLabelStyle.lCOLSPAN = 1
    typCellLabelStyle.sWidth = "12%"
    ofwTable.AddCellAsHTML m_ofwHtmlRef.SubmitImage(gk_sImagesPath & gk_sImageOK, _
        k_sOKBtn, CStr(m_typPersonData.lAuthorityStatus)), typCellLabelStyle
    ofwTable.AddCellAsHTML m_ofwHtmlRef.ImageButton(gk_sImagesPath & gk_sImageCancel, "", _
        gk_sLinkDisplayAuthorityTerms & "?" & ksTermCategory & "=" & m_aspSession(ksTermCategory), , m_ofwRefData.Search(gToolTip, ETOOLTIP.ttButtonCancelHelp)), typCellLabelStyle
     ofwTable.EndRow

    Set DisplayPersonDetails = ofwTable
    
    Set ofwTable = Nothing
    
Exit Function

ErrorHandler:
    ConvertError sSource
End Function

Public Sub UpdateIndexTermDetails()
'This procedure will be called from UpdateIndexTerms.asp page.
'This is used to update index term details.
'
'   d1  Geedha Sivanadian     4Aug1999   Created
'   d2  Aravinda A            08Oct1999  Added code for invisible
'                                        Added security
'                                        Added see/see also links and persistance
'   D14 Aravinda A            08Dec1999  Added error handling for rejection,
'   d15 aga                   14Dec1999  Not deleting state if error occurs while rejecting
'   d20 crm                   07Feb2000     changed redirection links after editing and saving an index term
On Error GoTo ErrorHandler
                 
    'Strings to store variables.
    Dim sSource As String
    Dim sParameter As String
    Dim sSpecificTermType As String
    Dim lIndexID    As Long
    Dim lIndexTermRaisedSource As Long
    Dim lAuthorityStatus As Long
    'd14
    Dim lSaveReturnCode As Long

    sSource = "VWEditorialAuthorityFiles.UpdateIndexTermDetails"
    
    'Check the permissions
    CheckPermissions ESecurityCheck.EUserCanMaintainAuthorityFiles, sSource, ""
    

    If m_aspRequest.Form(mk_sAFSee & ".x") <> "" Then
        SetStateBeforeLeave
    
        Exit Sub
    End If
    If m_aspRequest.Form(mk_sAFSeeAlso & ".x") <> "" Then
        SetStateBeforeLeave
    
        Exit Sub
    End If
    If InStr(m_aspRequest.Form, mk_sAFSeeRemove) > 0 Then
        SetStateBeforeLeave
        m_objContext.SetComplete
        Exit Sub
    End If

    If InStr(m_aspRequest.Form, mk_sAFSeeAlsoRemove) > 0 Then
        SetStateBeforeLeave
    
        Exit Sub
    End If
    
    '''''''''''''''''''''''''''
         '********** Get PARAMETER values *******************'
    'Get the values from prev page. Check for term types.
  
    'Get the value of term type
'Check if the user had gone to Index browser or not before saving

    If m_aspRequest(gk_sFORMAFGroup) = "" Then
          RaiseError sSource, EPROCATINTERNALERROR.peAuthorityFilesMaintenance, _
                           "Term type is blank"
    End If
    m_lSpecificTermType = m_aspRequest(gk_sFORMAFGroup)
    
    If m_aspRequest(gk_sFORMAFTermID) <= 0 Then
          RaiseError sSource, EPROCATINTERNALERROR.peAuthorityFilesMaintenance, _
                           "Term id is not a valid id"
    End If
    lIndexID = m_aspRequest(gk_sFORMAFTermID)
    
    
    If lIndexTermRaisedSource < 0 Then
          RaiseError sSource, EPROCATINTERNALERROR.peAuthorityFilesMaintenance, _
                           "Index term raised type is not valid"
    End If
    lIndexTermRaisedSource = m_aspRequest(k_sIndexTermRaisedThrough)

    'Get the status value based on the button clicked.
    
    If m_aspRequest(k_sApproveCompleteBtn & gk_sXCoordinate) <> 0 Or _
        m_aspRequest(k_sApproveCompleteBtn & gk_sYCoordinate) <> 0 Then
    
        lAuthorityStatus = gk_lApprovedComplete
        
    ElseIf m_aspRequest(k_sApproveInCompleteBtn & gk_sXCoordinate) <> 0 Or _
            m_aspRequest(k_sApproveInCompleteBtn & gk_sYCoordinate) <> 0 Then
    
        lAuthorityStatus = gk_lApprovedInComplete
        
    ElseIf m_aspRequest(k_sInvisibleBtn & gk_sXCoordinate) <> 0 Or _
            m_aspRequest(k_sInvisibleBtn & gk_sYCoordinate) <> 0 Then
            
        lAuthorityStatus = gk_lInvisible
        
    ElseIf m_aspRequest(k_sRejectBtn & gk_sXCoordinate) <> 0 Or _
            m_aspRequest(k_sRejectBtn & gk_sYCoordinate) <> 0 Then
            
        lAuthorityStatus = gk_lRejected
        
    End If
   
    Select Case m_lSpecificTermType
        Case EAFTermType.AFPerson
            
            Set m_oiAFSaver = New BOAFPerson
            
            m_oBOAFPerson.PROCATTransientInterface_GetState
            
            'd5 Save links
            m_oiAFSaver.SeeLinksData = m_oBOAFPerson.SeeLinks
            m_oiAFSaver.SeeAlsoLinksData = m_oBOAFPerson.SeeAlsoLinks
            m_oiAFSaver.SaveLinks m_oBOAFPerson.Data.lID
            
            m_typPersonData.sSurname = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLSurname_Text))
            m_typPersonData.lAuthorityStatus = lAuthorityStatus
            m_typPersonData.sGender = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLGender_Indicator))
            m_typPersonData.sTitle = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLPerson_Title))
            m_typPersonData.sAdditionalNameElements = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLadditional_elements_of_name))
            m_typPersonData.sPretitle = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLPretitle_Text))
            m_typPersonData.sForenames = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLForename_Text))
            m_typPersonData.sAlternativeStatus = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLAlternative_Status_Text))
            m_typPersonData.sBirthDate = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLBirth_Date))
            m_typPersonData.sUncertainBirthDateCode = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLUncertain_Birth_Date_Code))
            m_typPersonData.sBiogHistory = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLBiog_History_Text))
            m_typPersonData.sValidation = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLValidation_text))
            m_typPersonData.sDeathDate = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLDeath_Date))
            m_typPersonData.sUncertainDeathDateCode = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLUncertain_Death_Date_Code))
            'Load the data back to the class.
            m_oiAFSaver.Data = m_typPersonData
            
            If lAuthorityStatus = gk_lRejected And m_oBOAFPerson.SeeLinks.Count = 0 Then
                m_aspResponse.Redirect gk_sLinkAFPreferredTermRejectConfirm & "?" & gk_sFORMAFTermID & "=" & m_oBOAFPerson.Data.lID _
                    & "&" & gk_sFORMAFGroup & "=" & EAFTermType.AFPerson
                Exit Sub
            End If
            'Call the save method.
            lSaveReturnCode = m_oiAFSaver.Save(m_oBOAFPerson.Data.lID)
            If lSaveReturnCode = mk_lRecordNotChanged Then
                'd14
                PROCATEditorialErrorHandling peCanNotDeleteRejectedAFTerm, sSource, _
                           "Cannot delete the AF term. Remove the links and reject it to delete."
                Exit Sub  'd15
            Else
                ''   d20 crm                   07Feb2000
                'm_aspResponse.Redirect gk_sLinkEditorialAFMaintMain
                m_aspResponse.Redirect gk_sLinkDisplayAuthorityTerms & "?" & ksTermCategory & "=" & lIndexTermRaisedSource
            End If
            m_oBOAFPerson.PROCATTransientInterface_DeleteState
            
        Case EAFTermType.AFSubject
        
            Set m_oiAFSaver = New BOAFSubject
            
            m_oBOAFSubject.PROCATTransientInterface_GetState
            
            'd5 Save links
            m_oiAFSaver.SeeLinksData = m_oBOAFSubject.SeeLinks
            m_oiAFSaver.SeeAlsoLinksData = m_oBOAFSubject.SeeAlsoLinks
            m_oiAFSaver.SaveLinks m_oBOAFSubject.Data.lID
            
            m_typSubjectData.sTermText = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLSubjectTermText))
            m_typSubjectData.lAuthorityStatus = lAuthorityStatus
            m_typSubjectData.sDefinition = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLSubjectDefinition))
            m_typSubjectData.sValidation = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLSubjectValidation))
            
            'Load the data back to the class.
            m_oiAFSaver.Data = m_typSubjectData
            If lAuthorityStatus = gk_lRejected And m_oBOAFSubject.SeeLinks.Count = 0 Then
                m_aspResponse.Redirect gk_sLinkAFPreferredTermRejectConfirm & "?" & gk_sFORMAFTermID & "=" & m_oBOAFSubject.Data.lID _
                    & "&" & gk_sFORMAFGroup & "=" & EAFTermType.AFSubject
                Exit Sub
            End If
            'Call the save method.
            lSaveReturnCode = m_oiAFSaver.Save(m_oBOAFSubject.Data.lID)
            If lSaveReturnCode = mk_lRecordNotChanged Then
                'd14
                PROCATEditorialErrorHandling peCanNotDeleteRejectedAFTerm, sSource, _
                            "Cannot delete the AF term. Remove the links and reject it to delete."
                Exit Sub 'd15
            Else
                ''   d20 crm                   07Feb2000
                'm_aspResponse.Redirect gk_sLinkEditorialAFMaintMain
                m_aspResponse.Redirect gk_sLinkDisplayAuthorityTerms & "?" & ksTermCategory & "=" & lIndexTermRaisedSource
            End If
            m_oBOAFSubject.PROCATTransientInterface_DeleteState
            
        Case EAFTermType.AFPlace
        
            Set m_oiAFSaver = New BOAFPlace
            
            m_oBOAFPlace.PROCATTransientInterface_GetState
            
            'd5 Save links
            m_oiAFSaver.SeeLinksData = m_oBOAFPlace.SeeLinks
            m_oiAFSaver.SeeAlsoLinksData = m_oBOAFPlace.SeeAlsoLinks
            m_oiAFSaver.SaveLinks m_oBOAFPlace.Data.lID
            
            m_typPlaceData.sTermText = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLPlaceName))
            m_typPlaceData.lAuthorityStatus = lAuthorityStatus
            m_typPlaceData.sParish = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLParishName))
            m_typPlaceData.sTown = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLTownName))
            m_typPlaceData.sCounty = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLCountyName))
            m_typPlaceData.sStartDate = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLPlaceDateStart))
            m_typPlaceData.sEndDate = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLPlaceDateEnd))
            m_typPlaceData.sGridReference = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLPlaceGridRef))
            m_typPlaceData.sUncertainStartDateCode = StringToNullOrString(m_aspRequest.Form(gk_sDBColUncertainDateCodeStart))
            m_typPlaceData.sUncertainEndDateCode = StringToNullOrString(m_aspRequest.Form(gk_sDBColUncertainDateCodeEnd))
            m_typPlaceData.sValidation = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLValidation_text))
            m_typPlaceData.sCountry = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLCountryName))
            m_typPlaceData.sHistoryText = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLPlaceHistory))
            'Load the data back to the class.
            m_oiAFSaver.Data = m_typPlaceData
            
            If lAuthorityStatus = gk_lRejected And m_oBOAFPlace.SeeLinks.Count = 0 Then
                m_aspResponse.Redirect gk_sLinkAFPreferredTermRejectConfirm & "?" & gk_sFORMAFTermID & "=" & m_oBOAFPlace.Data.lID _
                    & "&" & gk_sFORMAFGroup & "=" & EAFTermType.AFPlace
                Exit Sub
            End If
            'Call the save method.
            lSaveReturnCode = m_oiAFSaver.Save(m_oBOAFPlace.Data.lID)
            If lSaveReturnCode = mk_lRecordNotChanged Then
                'd14
                PROCATEditorialErrorHandling peCanNotDeleteRejectedAFTerm, sSource, _
                            "Cannot delete the AF term. Remove the links and reject it to delete."
                            
                Exit Sub 'd15
            Else
                ''   d20 crm                   07Feb2000
                'm_aspResponse.Redirect gk_sLinkEditorialAFMaintMain
                m_aspResponse.Redirect gk_sLinkDisplayAuthorityTerms & "?" & ksTermCategory & "=" & lIndexTermRaisedSource
            End If
            m_oBOAFPlace.PROCATTransientInterface_DeleteState
        Case EAFTermType.AFCorporateBody
            
            Set m_oiAFSaver = New BOAFCorporateBody
            
            m_oBOAFCorporateBody.PROCATTransientInterface_GetState
            
            'd5 Save links
            m_oiAFSaver.SeeLinksData = m_oBOAFCorporateBody.SeeLinks
            m_oiAFSaver.SeeAlsoLinksData = m_oBOAFCorporateBody.SeeAlsoLinks
            m_oiAFSaver.SaveLinks m_oBOAFCorporateBody.Data.lID
            
            m_typCorporateBodyData.sTermText = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLCorporateBodyName))
            m_typCorporateBodyData.lAuthorityStatus = lAuthorityStatus
            m_typCorporateBodyData.sStartDate = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLCorporateBodyDateStart))
            m_typCorporateBodyData.sEndDate = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLCorporateBodyDateEnd))
            m_typCorporateBodyData.sUncertainStartDateCode = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLCorporateBodyUncertainStartDate))
            m_typCorporateBodyData.sUncertainEndDateCode = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLCorporateBodyUncertainEndDate))
            m_typCorporateBodyData.sRemitFunction = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLRemitFunction))
            m_typCorporateBodyData.sVariantText = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLVariantText))
            m_typCorporateBodyData.sValidation = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLValidation_text))
            m_typCorporateBodyData.sJurisdiction = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLJurisdiction))
            m_typCorporateBodyData.bIsPlaceOfDeposit = m_aspRequest.Form(gk_sDBCOLCorporateBodyIsPlaceOfDep)
            m_typCorporateBodyData.sPlaceOfDepositCode = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLNationalPlaceOfDepositCode))
            
            'Load the data back to the class.
            m_oiAFSaver.Data = m_typCorporateBodyData
            
            If lAuthorityStatus = gk_lRejected And m_oBOAFCorporateBody.SeeLinks.Count = 0 Then
                m_aspResponse.Redirect gk_sLinkAFPreferredTermRejectConfirm & "?" & gk_sFORMAFTermID & "=" & m_oBOAFCorporateBody.Data.lID _
                    & "&" & gk_sFORMAFGroup & "=" & EAFTermType.AFCorporateBody
                Exit Sub
            End If
            
            'Call the save method.
            lSaveReturnCode = m_oiAFSaver.Save(m_oBOAFCorporateBody.Data.lID)
            If lSaveReturnCode = mk_lRecordNotChanged Then
                'd14
                PROCATEditorialErrorHandling peCanNotDeleteRejectedAFTerm, sSource, _
                            "Cannot delete the AF term. Remove the links and reject it to delete."
                            
                Exit Sub 'd15
            Else
                ''   d20 crm                   07Feb2000
                'm_aspResponse.Redirect gk_sLinkEditorialAFMaintMain
                m_aspResponse.Redirect gk_sLinkDisplayAuthorityTerms & "?" & ksTermCategory & "=" & lIndexTermRaisedSource
            End If
            m_oBOAFCorporateBody.PROCATTransientInterface_DeleteState
    End Select
    
    
    m_objContext.SetComplete
    Exit Sub
ErrorHandler:

    m_objContext.SetAbort
    
    ShowError sSource
              
End Sub

Private Function DisplaySubjectDetails(t_lSubjectReferenceId As Long) As FWTable
'This procedure will be called from DisplaySpecificTerm procedure.
'This procedure will be called to display Subject index term page.
'
'   d1  Geedha Sivanadian   2Aug1999    Created
'   d5  AGA                 21Oct1999   Modified to add see & See also links
'   d16 AGA                 16Dec1999   Display message if a see link is added to the term in catelogue
'   d26 wdp                 23Feb2000     Sorted display of see and see also links and buttons
'                                           for various status of terms(preferred, non-preferred etc.)
'   d32 tmj                 21Mar2000   Changed gk_sFORMAFTermValue to gk_sFORMAdvSIndexName
'   d35 daj                   17Jul2000     Changed the checking in Display<AFType>Details functions so that the links
'                                           currently stored in TransientStore are used (rather than the database) to decide
'                                           which of the See and See Also buttons are shown


On Error GoTo ErrorHandler
    
    'Table
    Dim ofwTable As FWTable
    
    'Table, row, cell styles
    Dim typTable As TableStyle
    Dim typCellLabelStyle As CellStyle
    Dim typCellInputStyle  As CellStyle
    'd5
    Dim typCellLinks  As CellStyle
    Dim typCellRemove  As CellStyle
    
    Dim sSource As String
    Dim sHTML As String
    Dim bMaleChecked As Boolean
    
    Dim sParameter As String
    Dim SApproveCompleteStr As String
    Dim SApproveInCompleteStr As String
    Dim SRejectStr As String
    
    Dim sMessage As String 'd16
    
    sSource = "VWEditorialAuthorityFiles.DisplaySubjectDetails"
    
    'Check the permissions
    CheckPermissions ESecurityCheck.EUserCanMaintainAuthorityFiles, sSource, ""
    
    
    If m_aspRequest(mk_sFORMAFTermAddOrRemove) = EAFTermLoadType.EAFLoadFromStateRemoveTerm Then
        'Load from persistant object (After removing a term)
            m_oBOAFSubject.Load EAFTermLoadType.EAFLoadFromStateRemoveTerm, t_lSubjectReferenceId, _
                CInt(m_aspSession(mk_sFORMAFRemoveTermID)), "", m_aspSession(mk_sFORMAFSeeLinkType)
            m_oBOAFSubject.PROCATTransientInterface_SetState
    Else
        If m_bNoPersist = True Then
            'Lodad from the database
            m_oBOAFSubject.Load EAFLoadFromDatabase, t_lSubjectReferenceId
            m_oBOAFSubject.PROCATTransientInterface_SetState
        Else
            If m_aspRequest(ksFORMCancel) = 1 Then
                m_oBOAFSubject.Load EAFTermLoadType.EAFLoadFromStateTermNoChange
            Else
                   'd16
                'Check if the term is raised from leaflet or editset: And if someone try to add a see link to it
                If m_aspSession(mk_sFORMAFSeeLinkType) = EAFSeeLinkType.EAFSeeLink And _
                    (m_lIndexTermRaisedSource = klleafletType Or m_lIndexTermRaisedSource = klEditSetType) Then
                    'Display this message in see links cell
                    sMessage = m_ofwHtmlRef.Font("Cannot convert to a non-preferred term while links exist", COLOURRed)
                
                    m_oBOAFSubject.Load EAFTermLoadType.EAFLoadFromStateTermNoChange
                    m_oBOAFSubject.PROCATTransientInterface_SetState
                Else
                    '   LOAD FROM PERSISTANCE OBJECT (After adding a term)
                    'd32
                    m_oBOAFSubject.Load EAFTermLoadType.EAFLoadFromStateAddTerm, t_lSubjectReferenceId, _
                        CLng(m_aspRequest(gk_sFORMAFTermID)), CStr(m_aspRequest(gk_sFORMAdvSIndexName)), m_aspSession(mk_sFORMAFSeeLinkType)
                    If m_oBOAFSubject.IsTermHasLinks = True And _
                        m_aspSession(mk_sFORMAFSeeLinkType) = EAFSeeLinkType.EAFSeeLink Then
                        'Display this message in the see links box
                        sMessage = m_ofwHtmlRef.Font("Cannot convert to a non-preferred term while links exist", COLOURRed)
                    End If
                    m_oBOAFSubject.PROCATTransientInterface_SetState
                End If
                
            End If
        End If
        
    End If
    
    m_typSubjectData = m_oBOAFSubject.Data
    Set m_colSeeLinks = m_oBOAFSubject.SeeLinks
    Set m_colSeeAlsoLinks = m_oBOAFSubject.SeeAlsoLinks
    
    'Display the data from the collection returned.
    Set ofwTable = New FWTable
    
    
    'Set the tables format
    typTable.eStyle = TABCustomStyle
    typTable.sWidth = "100%"
    ofwTable.TableStyle = typTable
    
    
    'Set the label columns style
    typCellLabelStyle = m_ofwHtmlRef.GetCellStyle(ECellStyle.CELLReadWriteLabelStyle)
    typCellLabelStyle.sWidth = "30%"
    
    'Set the text columns style.
    typCellInputStyle = m_ofwHtmlRef.GetCellStyle(ECellStyle.CELLReadWriteDataStyle)
    typCellInputStyle.sWidth = "70%"
    typCellInputStyle.lCOLSPAN = 4
    
    typCellLinks = m_ofwHtmlRef.GetCellStyle(ECellStyle.CELLReadWriteDataStyle)
    typCellLinks.sWidth = "20%"
    typCellLinks.lCOLSPAN = 3
    
    typCellRemove = m_ofwHtmlRef.GetCellStyle(ECellStyle.CELLReadWriteDataStyle)
    typCellRemove.sWidth = "10%"
    typCellRemove.lCOLSPAN = 1
    
    
    'First Row
    ofwTable.StartRow
    'Label
    ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrName), typCellLabelStyle
    'Text Field
    
    ofwTable.AddCellAsHTML m_ofwHtmlRef.InputField(ITText, gk_sDBCOLSubjectTermText, 100, 100, _
                    VariantStringAsString(m_typSubjectData.sTermText), VCNone, True), typCellInputStyle
    ofwTable.EndRow
    
    'Second Row
    ofwTable.StartRow
    'Label
    ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrDefinition), typCellLabelStyle
    'Text Field
    
    ofwTable.AddCellAsHTML m_ofwHtmlRef.InputField(ITText, gk_sDBCOLSubjectDefinition, 100, 100, _
                    VariantStringAsString(m_typSubjectData.sDefinition)), typCellInputStyle
    ofwTable.EndRow
    
    'Third Row
    ofwTable.StartRow
    'Label
    ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrValidation), typCellLabelStyle
    'Text Field
    
    ofwTable.AddCellAsHTML m_ofwHtmlRef.InputField(ITText, gk_sDBCOLSubjectValidation, 100, 100, _
                    VariantStringAsString(m_typSubjectData.sValidation)), typCellInputStyle
    ofwTable.EndRow
    'Display links here.
  
    'See links
    'Flags to decide what combination of see and see also
     'buttons we display, there are 3 scenarios:
     '1.    Term is not linked to or from anything - display see and see also buttons
     '2.    Term is a non preferred term - display see link but no button
     '3.    Term is a preferred term or has see also links to/from it - display see also link(s) and button
     
     Dim bIsNonPreferred As Boolean
     Dim bIsPreferredOrLinkedTo As Boolean
     
     If m_colSeeLinks.Count = 0 Then
        bIsNonPreferred = False
     Else
        bIsNonPreferred = True
     End If
     
     ' d35 - Use cache rather than database
     If m_colSeeAlsoLinks.Count = 0 Then
        bIsPreferredOrLinkedTo = False
     Else
        bIsPreferredOrLinkedTo = True
     End If
     
     If Not bIsPreferredOrLinkedTo Or bIsNonPreferred Then
         If m_colSeeLinks.Count = 0 Then
            ofwTable.StartRow
            'Label
            ofwTable.AddCellAsHTML "See ", typCellLabelStyle
            ofwTable.AddCellAsHTML m_ofwHtmlRef.SubmitImage(gk_sImagesPath & gk_sImagePickSubject, _
                mk_sAFSee, mk_sAFSee) & sMessage, typCellInputStyle
            ofwTable.EndRow
        Else
            Dim iTemp As Integer
            For iTemp = 1 To m_colSeeLinks.Count
                ofwTable.StartRow
                If iTemp = 1 Then
                    ofwTable.AddCellAsHTML "See ", typCellLabelStyle
                Else
                    ofwTable.AddCellAsHTML gk_sHTMLNonBreakSpace, typCellLabelStyle
                End If
                ofwTable.AddCellAsHTML m_colSeeLinks.Item(iTemp).Value, typCellLinks
                ofwTable.AddCellAsHTML m_ofwHtmlRef.SubmitImage(gk_sImagesPath & gk_sImageRemove, _
                    mk_sAFSeeRemove & m_colSeeLinks.Item(iTemp).ID, mk_sAFSeeRemove), typCellRemove
                ofwTable.EndRow
            Next
            ofwTable.StartRow
            'Label
            ofwTable.AddCellAsHTML gk_sHTMLNonBreakSpace, typCellLabelStyle
            If Not bIsNonPreferred Then
                ofwTable.AddCellAsHTML m_ofwHtmlRef.SubmitImage(gk_sImagesPath & gk_sImagePickSubject, _
                    mk_sAFSee, mk_sAFSee) & sMessage, typCellInputStyle
            End If
            ofwTable.EndRow
         End If
     End If
     
     If Not bIsNonPreferred Then
        If m_colSeeAlsoLinks.Count = 0 Then
         'See also links
            ofwTable.StartRow
            'Label
            ofwTable.AddCellAsHTML "See also", typCellLabelStyle
            ofwTable.AddCellAsHTML m_ofwHtmlRef.SubmitImage(gk_sImagesPath & gk_sImagePickSubject, _
               mk_sAFSeeAlso, mk_sAFSeeAlso), typCellInputStyle
            
            ofwTable.EndRow
         Else
            For iTemp = 1 To m_colSeeAlsoLinks.Count
                ofwTable.StartRow
                If iTemp = 1 Then
                    ofwTable.AddCellAsHTML "See also", typCellLabelStyle
                Else
                    ofwTable.AddCellAsHTML gk_sHTMLNonBreakSpace, typCellLabelStyle
                End If
                ofwTable.AddCellAsHTML m_colSeeAlsoLinks.Item(iTemp).Value, typCellLinks
                ofwTable.AddCellAsHTML m_ofwHtmlRef.SubmitImage(gk_sImagesPath & gk_sImageRemove, _
                    mk_sAFSeeAlsoRemove & m_colSeeAlsoLinks.Item(iTemp).ID, mk_sAFSeeAlsoRemove), typCellRemove
                ofwTable.EndRow
            Next
            ofwTable.StartRow
            'Label
            ofwTable.AddCellAsHTML gk_sHTMLNonBreakSpace, typCellLabelStyle
            ofwTable.AddCellAsHTML m_ofwHtmlRef.SubmitImage(gk_sImagesPath & gk_sImagePickSubject, _
               mk_sAFSeeAlso, mk_sAFSeeAlso), typCellInputStyle
            ofwTable.EndRow
        End If
    End If
    ofwTable.AddBlankRow
    
    ofwTable.StartRow
    
    ofwTable.AddCellAsHTML gk_sHTMLNonBreakSpace, typCellLabelStyle
    
     'd5    displaying images based on the authority status
     m_aspSession(mk_sAFAuthorityStatus) = m_typSubjectData.lAuthorityStatus
     Dim sImages As String
     If Not m_typSubjectData.lAuthorityStatus = gk_lApprovedComplete Then
        
         sImages = sImages & m_ofwHtmlRef.SubmitImage(gk_sImagesPath & gk_sImageApproveComplete, _
                  k_sApproveCompleteBtn, gk_lApprovedComplete)
     End If
     If Not m_typSubjectData.lAuthorityStatus = gk_lApprovedInComplete Then
        If Not m_typSubjectData.lAuthorityStatus = gk_lApprovedComplete Then
            sImages = sImages & m_ofwHtmlRef.SubmitImage(gk_sImagesPath & gk_sImageApproveIncomplete, _
                 k_sApproveInCompleteBtn, gk_lApprovedInComplete)
        End If
     End If
     
     If Not m_typSubjectData.lAuthorityStatus = gk_lInvisible Then
        If Not m_typSubjectData.lAuthorityStatus = gk_lApprovedComplete Then
            sImages = sImages & m_ofwHtmlRef.SubmitImage(gk_sImagesPath & gk_sImageInvisible, _
                 k_sInvisibleBtn, gk_lInvisible)
        End If
     End If
     If Not m_typSubjectData.lAuthorityStatus = gk_lRejected Then
         
         sImages = sImages & m_ofwHtmlRef.SubmitImage(gk_sImagesPath & gk_sImageReject, _
                 k_sRejectBtn, gk_lRejected)
         
     End If
     typCellLabelStyle.lCOLSPAN = 2
     typCellLabelStyle.sWidth = "23%"
     ofwTable.AddCellAsHTML sImages, typCellLabelStyle
     'Show ok/cancel button
    typCellLabelStyle.lCOLSPAN = 1
    typCellLabelStyle.sWidth = "12%"
    ofwTable.AddCellAsHTML m_ofwHtmlRef.SubmitImage(gk_sImagesPath & gk_sImageOK, _
                 k_sOKBtn, CStr(m_typSubjectData.lAuthorityStatus)), typCellLabelStyle
    ofwTable.AddCellAsHTML m_ofwHtmlRef.ImageButton(gk_sImagesPath & gk_sImageCancel, "", _
         gk_sLinkDisplayAuthorityTerms & "?" & ksTermCategory & "=" & m_aspSession(ksTermCategory), , m_ofwRefData.Search(gToolTip, ETOOLTIP.ttButtonCancelHelp)), typCellLabelStyle
         
     ofwTable.EndRow

    Set DisplaySubjectDetails = ofwTable
    Set ofwTable = Nothing
Exit Function

ErrorHandler:
  ConvertError sSource
End Function


Private Function DisplayPlaceDetails(t_lPlaceReferenceID As Long) As FWTable
'This procedure will be called from DisplaySpecificTerm procedure.
'This procedure will be called to display Place index term page.
'
'   d1  Geedha Sivanadian   10Aug1999   Created
'   d5  AGA                 21Oct1999   Modified to add see & See also links
'   d16 AGA                 16Dec1999   Display message if a see link is added to the term in catelogue
'   d24 crm                 16Feb2000   Changed field length
'   d26 wdp                 23Feb2000     Sorted display of see and see also links and buttons
'                                           for various status of terms(preferred, non-preferred etc.)
'   d32 tmj                 21Mar2000   Changed gk_sFORMAFTermValue to gk_sFORMAdvSIndexName
'   d34 tmj                 28Mar2000   Added textarea validation
'   d35 daj                   17Jul2000     Changed the checking in Display<AFType>Details functions so that the links
'                                           currently stored in TransientStore are used (rather than the database) to decide
'                                           which of the See and See Also buttons are shown


On Error GoTo ErrorHandler
    
    'Table
    Dim ofwTable As FWTable
    
    'Table, row, cell styles
    Dim typTable As TableStyle
    Dim typCellLabelStyle As CellStyle
    Dim typCellInputStyle  As CellStyle
    Dim typCellLinks  As CellStyle
    Dim typCellRemove  As CellStyle
            
    Dim sSource As String
    Dim sHTML As String
    Dim bMaleChecked As Boolean
    
    Dim sParameter As String
    Dim SApproveCompleteStr As String
    Dim SApproveInCompleteStr As String
    Dim SRejectStr As String
    Dim sMessage As String 'd16
    
    sSource = "VWEditorialAuthorityFiles.DisplayPlaceDetails"
    
    'Check the permissions
    CheckPermissions ESecurityCheck.EUserCanMaintainAuthorityFiles, sSource, ""
    
    If m_aspRequest(mk_sFORMAFTermAddOrRemove) = EAFTermLoadType.EAFLoadFromStateRemoveTerm Then
        'Load from persistant object (After removing a term)
            m_oBOAFPlace.Load EAFTermLoadType.EAFLoadFromStateRemoveTerm, t_lPlaceReferenceID, _
                CInt(m_aspSession(mk_sFORMAFRemoveTermID)), "", m_aspSession(mk_sFORMAFSeeLinkType)
            m_oBOAFPlace.PROCATTransientInterface_SetState
    Else
        If m_bNoPersist = True Then
            'Lodad from the database
            m_oBOAFPlace.Load EAFLoadFromDatabase, t_lPlaceReferenceID
            m_oBOAFPlace.PROCATTransientInterface_SetState

        Else
            If m_aspRequest(ksFORMCancel) = 1 Then
                m_oBOAFPlace.Load EAFTermLoadType.EAFLoadFromStateTermNoChange
            Else
                'd16
                'Check if the term is raised from leaflet or editset: And if someone try to add a see link to it
                If m_aspSession(mk_sFORMAFSeeLinkType) = EAFSeeLinkType.EAFSeeLink And _
                    (m_lIndexTermRaisedSource = klleafletType Or m_lIndexTermRaisedSource = klEditSetType) Then
                    'Display this message at See links cell
                    sMessage = m_ofwHtmlRef.Font("Cannot convert to a non-preferred term while links exist", COLOURRed)
                    
                    m_oBOAFPlace.Load EAFTermLoadType.EAFLoadFromStateTermNoChange
                    m_oBOAFPlace.PROCATTransientInterface_SetState
                Else
                    '   LOAD FROM PERSISTANCE OBJECT (After adding a term)
                    'd32
                    m_oBOAFPlace.Load EAFTermLoadType.EAFLoadFromStateAddTerm, t_lPlaceReferenceID, _
                        CLng(m_aspRequest(gk_sFORMAFTermID)), CStr(m_aspRequest(gk_sFORMAdvSIndexName)), m_aspSession(mk_sFORMAFSeeLinkType)
                    If m_oBOAFPlace.IsTermHasLinks = True And _
                        m_aspSession(mk_sFORMAFSeeLinkType) = EAFSeeLinkType.EAFSeeLink Then
                        'Display this message in the see links box
                        sMessage = m_ofwHtmlRef.Font("Cannot convert to a non-preferred term while links exist", COLOURRed)
                    End If
                    m_oBOAFPlace.PROCATTransientInterface_SetState
                End If
                
                
            End If
        End If
        
    End If
    
    m_typPlaceData = m_oBOAFPlace.Data
    Set m_colSeeLinks = m_oBOAFPlace.SeeLinks
    Set m_colSeeAlsoLinks = m_oBOAFPlace.SeeAlsoLinks
    
    'Display the data from the collection returned.
    Set ofwTable = New FWTable
    
    'Set the tables format
    typTable.eStyle = TABCustomStyle
    typTable.sWidth = "100%"
    ofwTable.TableStyle = typTable
    
    'Set the label columns style
    typCellLabelStyle = m_ofwHtmlRef.GetCellStyle(ECellStyle.CELLReadWriteLabelStyle)
    typCellLabelStyle.sWidth = "30%"
    
    'Set the text columns style.
    typCellInputStyle = m_ofwHtmlRef.GetCellStyle(ECellStyle.CELLReadWriteDataStyle)
    typCellInputStyle.sWidth = "70%"
    typCellInputStyle.lCOLSPAN = 4
    
    typCellLinks = m_ofwHtmlRef.GetCellStyle(ECellStyle.CELLReadWriteDataStyle)
    typCellLinks.sWidth = "20%"
    typCellLinks.lCOLSPAN = 3
    
    typCellRemove = m_ofwHtmlRef.GetCellStyle(ECellStyle.CELLReadWriteDataStyle)
    typCellRemove.sWidth = "10%"
    typCellRemove.lCOLSPAN = 1
    
    'Display place details.
    
    'Index Name
    ofwTable.StartRow
    'Label
    ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrName), typCellLabelStyle
    'Text Field
    
    ofwTable.AddCellAsHTML m_ofwHtmlRef.InputField(ITText, gk_sDBCOLPlaceName, 30, 30, _
                    VariantStringAsString(m_typPlaceData.sTermText), VCNone, True), typCellInputStyle
    ofwTable.EndRow
    
    'Parish text
    ofwTable.StartRow
    'Label
    ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrParish), typCellLabelStyle
    'Text Field
    
    ofwTable.AddCellAsHTML m_ofwHtmlRef.InputField(ITText, gk_sDBCOLParishName, 30, 30, _
                    VariantStringAsString(m_typPlaceData.sParish)), typCellInputStyle
    ofwTable.EndRow
    
    'Town text
    ofwTable.StartRow
    'Label
    ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrTown), typCellLabelStyle
    'Text Field
    
    ofwTable.AddCellAsHTML m_ofwHtmlRef.InputField(ITText, gk_sDBCOLTownName, 30, 30, _
                    VariantStringAsString(m_typPlaceData.sTown)), typCellInputStyle
    ofwTable.EndRow
    
    'Country ield
    ofwTable.StartRow
    'Label
    ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrCountry), typCellLabelStyle
    'Text Field
    
    ofwTable.AddCellAsHTML m_ofwHtmlRef.InputField(ITText, gk_sDBCOLCountryName, 40, 40, _
                    VariantStringAsString(m_typPlaceData.sCountry)), typCellInputStyle
    ofwTable.EndRow

    'County text
    ofwTable.StartRow
    'Label
    ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrCounty), typCellLabelStyle
    'Text Field
    
    ofwTable.AddCellAsHTML m_ofwHtmlRef.InputField(ITText, gk_sDBCOLCountyName, 50, 50, _
                    VariantStringAsString(m_typPlaceData.sCounty)), typCellInputStyle
    ofwTable.EndRow
    
    'uncertain Date start code
    ofwTable.StartRow
    'Label
    ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrUncertainStartDateCode), typCellLabelStyle
    'Text Field
    'd24    crm     16Feb2000
    ofwTable.AddCellAsHTML m_ofwHtmlRef.InputField(ITText, gk_sDBCOLPlaceUncertainStartDate, 3, 3, _
                    VariantStringAsString(m_typPlaceData.sUncertainStartDateCode), VCNone), typCellInputStyle
    ofwTable.EndRow
    
    'Place Date start
    ofwTable.StartRow
    'Label
    ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrStartDate), typCellLabelStyle
    'Text Field
    
    ofwTable.AddCellAsHTML m_ofwHtmlRef.InputField(ITText, gk_sDBCOLPlaceDateStart, gk_iDBCOLLenDateInAF, gk_iDBCOLLenDateInAF, _
                    VariantStringAsString(m_typPlaceData.sStartDate), VCNumber, , , 0, 9999), typCellInputStyle
    ofwTable.EndRow
    
    'uncertain Date end code
    ofwTable.StartRow
    'Label
    ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrUncertainEndDateCode), typCellLabelStyle
    'Text Field
    
    ofwTable.AddCellAsHTML m_ofwHtmlRef.InputField(ITText, gk_sDBCOLPlaceUncertainEndDate, 3, 3, _
                    VariantStringAsString(m_typPlaceData.sUncertainEndDateCode), VCNone), typCellInputStyle
    ofwTable.EndRow
    
    'Place Date end
    ofwTable.StartRow
    'Label
    ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrEndDate), typCellLabelStyle
    'Text Field
    
    ofwTable.AddCellAsHTML m_ofwHtmlRef.InputField(ITText, gk_sDBCOLPlaceDateEnd, gk_iDBCOLLenDateInAF, gk_iDBCOLLenDateInAF, _
                    VariantStringAsString(m_typPlaceData.sEndDate), VCNumber, , , 0, 9999), typCellInputStyle
    ofwTable.EndRow
 
    'Grid references text
    ofwTable.StartRow
    'Label
    ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrGridReferences), typCellLabelStyle
    'Text Field
    
    ofwTable.AddCellAsHTML m_ofwHtmlRef.InputField(ITText, gk_sDBCOLPlaceGridRef, 50, 50, _
                    VariantStringAsString(m_typPlaceData.sGridReference)), typCellInputStyle
    ofwTable.EndRow
        
    
    
    'Validation text
    ofwTable.StartRow
    'Label
    ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrValidation), typCellLabelStyle
    'Text Field
    
    ofwTable.AddCellAsHTML m_ofwHtmlRef.InputField(ITText, gk_sDBCOLValidation_text, 50, 50, _
                    VariantStringAsString(m_typPlaceData.sValidation), VCNone), typCellInputStyle
    ofwTable.EndRow
        
    'History
    ofwTable.StartRow
    'Label
    ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrHistory), typCellLabelStyle
    'Text Field
    ofwTable.AddCellAsHTML m_ofwHtmlRef.TextArea(gk_sDBCOLPlaceHistory, _
                    VariantStringAsString(m_typPlaceData.sHistoryText), 40, 4, gk_iDBCOLLenPlaceHistory), _
                        typCellInputStyle
    ofwTable.EndRow
    
    ' Display links
     'See links
    'Flags to decide what combination of see and see also
     'buttons we display, there are 3 scenarios:
     '1.    Term is not linked to or from anything - display see and see also buttons
     '2.    Term is a non preferred term - display see link but no button
     '3.    Term is a preferred term or has see also links to/from it - display see also link(s) and button
     
     Dim bIsNonPreferred As Boolean
     Dim bIsPreferredOrLinkedTo As Boolean
     
     If m_colSeeLinks.Count = 0 Then
        bIsNonPreferred = False
     Else
        bIsNonPreferred = True
     End If
     
     ' d35 - Use cache rather than database
     If m_colSeeAlsoLinks.Count = 0 Then
        bIsPreferredOrLinkedTo = False
     Else
        bIsPreferredOrLinkedTo = True
     End If
     
     If Not bIsPreferredOrLinkedTo Or bIsNonPreferred Then
         If m_colSeeLinks.Count = 0 Then
            ofwTable.StartRow
             'Label
            ofwTable.AddCellAsHTML "See ", typCellLabelStyle
            ofwTable.AddCellAsHTML m_ofwHtmlRef.SubmitImage(gk_sImagesPath & gk_sImagePickPlace, _
                mk_sAFSee, mk_sAFSee) & sMessage, typCellInputStyle
            ofwTable.EndRow
        Else
            Dim iTemp As Integer
            For iTemp = 1 To m_colSeeLinks.Count
                ofwTable.StartRow
                If iTemp = 1 Then
                    ofwTable.AddCellAsHTML "See ", typCellLabelStyle
                Else
                    ofwTable.AddCellAsHTML gk_sHTMLNonBreakSpace, typCellLabelStyle
                End If
                ofwTable.AddCellAsHTML m_colSeeLinks.Item(iTemp).Value, typCellLinks
                ofwTable.AddCellAsHTML m_ofwHtmlRef.SubmitImage(gk_sImagesPath & gk_sImageRemove, _
                    mk_sAFSeeRemove & m_colSeeLinks.Item(iTemp).ID, mk_sAFSeeRemove), typCellRemove
                ofwTable.EndRow
            Next
            ofwTable.StartRow
             'Label
            ofwTable.AddCellAsHTML gk_sHTMLNonBreakSpace, typCellLabelStyle
            If Not bIsNonPreferred Then
                ofwTable.AddCellAsHTML m_ofwHtmlRef.SubmitImage(gk_sImagesPath & gk_sImagePickPlace, _
                    mk_sAFSee, mk_sAFSee) & sMessage, typCellInputStyle
            End If
            ofwTable.EndRow
        End If
    End If
     'See also links
    If Not bIsNonPreferred Then
        If m_colSeeAlsoLinks.Count = 0 Then
            ofwTable.StartRow
            'Label
            ofwTable.AddCellAsHTML "See also", typCellLabelStyle
            ofwTable.AddCellAsHTML m_ofwHtmlRef.SubmitImage(gk_sImagesPath & gk_sImagePickPlace, _
                mk_sAFSeeAlso, mk_sAFSeeAlso), typCellInputStyle
            ofwTable.EndRow
        Else
            For iTemp = 1 To m_colSeeAlsoLinks.Count
                ofwTable.StartRow
                If iTemp = 1 Then
                    ofwTable.AddCellAsHTML "See also", typCellLabelStyle
                Else
                    ofwTable.AddCellAsHTML gk_sHTMLNonBreakSpace, typCellLabelStyle
                End If
                ofwTable.AddCellAsHTML m_colSeeAlsoLinks.Item(iTemp).Value, typCellLinks
                ofwTable.AddCellAsHTML m_ofwHtmlRef.SubmitImage(gk_sImagesPath & gk_sImageRemove, _
                   mk_sAFSeeAlsoRemove & m_colSeeAlsoLinks.Item(iTemp).ID, mk_sAFSeeAlsoRemove), typCellRemove
                ofwTable.EndRow
                
            Next
            ofwTable.StartRow
            'Label
            ofwTable.AddCellAsHTML gk_sHTMLNonBreakSpace, typCellLabelStyle
            ofwTable.AddCellAsHTML m_ofwHtmlRef.SubmitImage(gk_sImagesPath & gk_sImagePickPlace, _
                mk_sAFSeeAlso, mk_sAFSeeAlso), typCellInputStyle
            ofwTable.EndRow
        End If
     End If
     ofwTable.AddBlankRow
     ofwTable.AddBlankRow

    ofwTable.StartRow
    
    ofwTable.AddCellAsHTML gk_sHTMLNonBreakSpace, typCellLabelStyle
    
    
     'd5        Display buttons based on the status
     m_aspSession(mk_sAFAuthorityStatus) = m_typPlaceData.lAuthorityStatus
     Dim sImages As String
     If Not m_typPlaceData.lAuthorityStatus = gk_lApprovedComplete Then
         
         sImages = sImages & m_ofwHtmlRef.SubmitImage(gk_sImagesPath & gk_sImageApproveComplete, _
                  k_sApproveCompleteBtn, gk_lApprovedComplete)
     End If
     If Not m_typPlaceData.lAuthorityStatus = gk_lApprovedInComplete Then
        If Not m_typPlaceData.lAuthorityStatus = gk_lApprovedComplete Then
            sImages = sImages & m_ofwHtmlRef.SubmitImage(gk_sImagesPath & gk_sImageApproveIncomplete, _
                 k_sApproveInCompleteBtn, gk_lApprovedInComplete)
        End If
     End If
     
     If Not m_typPlaceData.lAuthorityStatus = gk_lInvisible Then
        If Not m_typPlaceData.lAuthorityStatus = gk_lApprovedComplete Then
            sImages = sImages & m_ofwHtmlRef.SubmitImage(gk_sImagesPath & gk_sImageInvisible, _
                 k_sInvisibleBtn, gk_lInvisible)
        End If
     End If
     If Not m_typPlaceData.lAuthorityStatus = gk_lRejected Then
         
         sImages = sImages & m_ofwHtmlRef.SubmitImage(gk_sImagesPath & gk_sImageReject, _
                 k_sRejectBtn, gk_lRejected)
         
     End If
     typCellLabelStyle.lCOLSPAN = 2
     typCellLabelStyle.sWidth = "23%"
     ofwTable.AddCellAsHTML sImages, typCellLabelStyle
     
     'Show ok/cancel buttons
    typCellLabelStyle.lCOLSPAN = 1
    typCellLabelStyle.sWidth = "12%"
    ofwTable.AddCellAsHTML m_ofwHtmlRef.SubmitImage(gk_sImagesPath & gk_sImageOK, _
                 k_sOKBtn, CStr(m_typPlaceData.lAuthorityStatus)), typCellLabelStyle
    ofwTable.AddCellAsHTML m_ofwHtmlRef.ImageButton(gk_sImagesPath & gk_sImageCancel, "", _
         gk_sLinkDisplayAuthorityTerms & "?" & ksTermCategory & "=" & m_aspSession(ksTermCategory), _
         , m_ofwRefData.Search(gToolTip, ETOOLTIP.ttButtonCancelHelp)), typCellLabelStyle
    
     ofwTable.EndRow
    
    Set DisplayPlaceDetails = ofwTable
    Set ofwTable = Nothing
    
Exit Function

ErrorHandler:
    ConvertError sSource
End Function

Private Function DisplayCorporateBodyDetails(t_lCorporateBodyReferenceID As Long) As FWTable
'This procedure will be called from DisplaySpecificTerm procedure.
'This procedure will be called to display CorporateBody index term page.
'
'   d1  Geedha Sivanadian   2Aug1999    Created
'   d5  AGA                 21Oct1999   Modified to add see & See also links
'   d16 AGA                 16Dec1999   Display message if a see link is added to the term in catelogue
'   d24 crm                 16Feb2000     Changed field length
'   d26 wdp                 23Feb2000     Sorted display of see and see also links and buttons
'                                           for various status of terms(preferred, non-preferred etc.)
'   d29 tmj                 16Mar2000   Added validation for TextArea size
'   d32 tmj                 21Mar2000   Changed gk_sFORMAFTermValue to gk_sFORMAdvSIndexName
'   d34 tmj                 28Mar2000   Added textarea validation
'   d35 daj                   17Jul2000     Changed the checking in Display<AFType>Details functions so that the links
'                                           currently stored in TransientStore are used (rather than the database) to decide
'                                           which of the See and See Also buttons are shown


On Error GoTo ErrorHandler
    
    'Table
    Dim ofwTable As FWTable
    
    'Table, row, cell styles
    Dim typTable As TableStyle
    Dim typCellLabelStyle As CellStyle
    Dim typCellInputStyle  As CellStyle
    Dim typCellLinks  As CellStyle
    Dim typCellRemove  As CellStyle
      
    Dim sSource As String
    Dim sHTML As String
    Dim bChecked As Boolean
    
    Dim sParameter As String
    
    Dim sMessage As String 'd16
    
    sSource = "VWEditorialAuthorityFiles.DisplayCorporateBodyDetails"
    
    'Check the permissions
    CheckPermissions ESecurityCheck.EUserCanMaintainAuthorityFiles, sSource, ""
            
    If m_aspRequest(mk_sFORMAFTermAddOrRemove) = EAFTermLoadType.EAFLoadFromStateRemoveTerm Then
        'Load from persistant object (After removing a term)
            m_oBOAFCorporateBody.Load EAFTermLoadType.EAFLoadFromStateRemoveTerm, t_lCorporateBodyReferenceID, _
                CInt(m_aspSession(mk_sFORMAFRemoveTermID)), "", m_aspSession(mk_sFORMAFSeeLinkType)
            m_oBOAFCorporateBody.PROCATTransientInterface_SetState
    
    Else
        If m_bNoPersist = True Then
            'Lodad from the database
            m_oBOAFCorporateBody.Load EAFLoadFromDatabase, t_lCorporateBodyReferenceID
            m_oBOAFCorporateBody.PROCATTransientInterface_SetState
            
        Else
            If m_aspRequest(ksFORMCancel) = 1 Then
                m_oBOAFCorporateBody.Load EAFTermLoadType.EAFLoadFromStateTermNoChange
            Else
                
                'd16
                'Check if the term is raised from leaflet or editset: And if someone try to add a see link to it
                If m_aspSession(mk_sFORMAFSeeLinkType) = EAFSeeLinkType.EAFSeeLink And _
                    (m_lIndexTermRaisedSource = klleafletType Or m_lIndexTermRaisedSource = klEditSetType) Then
                    'Display this message in the see links box
                    sMessage = m_ofwHtmlRef.Font("Cannot convert to a non-preferred term while links exist", COLOURRed)
                    
                    m_oBOAFCorporateBody.Load EAFTermLoadType.EAFLoadFromStateTermNoChange
                    m_oBOAFCorporateBody.PROCATTransientInterface_SetState
                Else
                    '   LOAD FROM PERSISTANCE OBJECT (After adding a term)
                    'd32
                    m_oBOAFCorporateBody.Load EAFTermLoadType.EAFLoadFromStateAddTerm, t_lCorporateBodyReferenceID, _
                        CLng(m_aspRequest(gk_sFORMAFTermID)), CStr(m_aspRequest(gk_sFORMAdvSIndexName)), m_aspSession(mk_sFORMAFSeeLinkType)
                    If m_oBOAFCorporateBody.IsTermHasLinks = True And _
                        m_aspSession(mk_sFORMAFSeeLinkType) = EAFSeeLinkType.EAFSeeLink Then
                        'Display this message in the see links box
                    sMessage = m_ofwHtmlRef.Font("Cannot convert to a non-preferred term while links exist", COLOURRed)
                    End If
                    m_oBOAFCorporateBody.PROCATTransientInterface_SetState
                End If
                
            End If
        End If
        
    End If
    m_typCorporateBodyData = m_oBOAFCorporateBody.Data
    Set m_colSeeLinks = m_oBOAFCorporateBody.SeeLinks
    Set m_colSeeAlsoLinks = m_oBOAFCorporateBody.SeeAlsoLinks
    
    'Display the data from the collection returned.
    Set ofwTable = New FWTable
    
    'Set the tables format
    typTable.eStyle = TABCustomStyle
    typTable.sWidth = "100%"
    ofwTable.TableStyle = typTable
    
    'Set the label columns style
    typCellLabelStyle = m_ofwHtmlRef.GetCellStyle(ECellStyle.CELLReadWriteLabelStyle)
    typCellLabelStyle.sWidth = "30%"
    
    'Set the text columns style.
    typCellInputStyle = m_ofwHtmlRef.GetCellStyle(ECellStyle.CELLReadWriteDataStyle)
    typCellInputStyle.sWidth = "70%"
    typCellInputStyle.lCOLSPAN = 4
    
    typCellLinks = m_ofwHtmlRef.GetCellStyle(ECellStyle.CELLReadWriteDataStyle)
    typCellLinks.sWidth = "20%"
    typCellLinks.lCOLSPAN = 3
    
    typCellRemove = m_ofwHtmlRef.GetCellStyle(ECellStyle.CELLReadWriteDataStyle)
    typCellRemove.sWidth = "10%"
    typCellRemove.lCOLSPAN = 1
    
    ' Name
    ofwTable.StartRow
    'Label
    ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrName), typCellLabelStyle
    'Text Field
    
    ofwTable.AddCellAsHTML m_ofwHtmlRef.InputField(ITText, gk_sDBCOLCorporateBodyName, 90, 90, _
                    VariantStringAsString(m_typCorporateBodyData.sTermText), VCNone, True), typCellInputStyle
    ofwTable.EndRow

    'Uncertain date code start
    ofwTable.StartRow
    'Label
    ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrUncertainStartDateCode), typCellLabelStyle
    'Text Field
    
    ofwTable.AddCellAsHTML m_ofwHtmlRef.InputField(ITText, gk_sDBColUncertainDateCodeStart, 3, 3, _
                    VariantStringAsString(m_typCorporateBodyData.sUncertainStartDateCode)), typCellInputStyle
    ofwTable.EndRow
    
    
    ' Corporate Body Date Start
    ofwTable.StartRow
    'Label
    ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrStartDate), typCellLabelStyle
    'Text Field
    
    ofwTable.AddCellAsHTML m_ofwHtmlRef.InputField(ITText, gk_sDBCOLCorporateBodyDateStart, gk_iDBCOLLenDateInAF, gk_iDBCOLLenDateInAF, _
                    VariantStringAsString(m_typCorporateBodyData.sStartDate), VCNumber, False, False, 0, 9999), typCellInputStyle
    ofwTable.EndRow

    'Uncertain date code end
    ofwTable.StartRow
    'Label
    ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrUncertainEndDateCode), typCellLabelStyle
    'Text Field
    
    ofwTable.AddCellAsHTML m_ofwHtmlRef.InputField(ITText, gk_sDBColUncertainDateCodeEnd, 3, 3, _
                    VariantStringAsString(m_typCorporateBodyData.sUncertainEndDateCode)), typCellInputStyle
    ofwTable.EndRow
    
    
    ' Corporate Body Date End
    ofwTable.StartRow
    'Label
    ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrEndDate), typCellLabelStyle
    'Text Field
    
    ofwTable.AddCellAsHTML m_ofwHtmlRef.InputField(ITText, gk_sDBCOLCorporateBodyDateEnd, gk_iDBCOLLenDateInAF, gk_iDBCOLLenDateInAF, _
                    VariantStringAsString(m_typCorporateBodyData.sEndDate), VCNumber, False, False, 0, 9999), typCellInputStyle
    ofwTable.EndRow

    'Remit and function
    ofwTable.StartRow
    'Label
    ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrRemitAndFunction), typCellLabelStyle
    'Text Area
    ofwTable.AddCellAsHTML m_ofwHtmlRef.TextArea(gk_sDBCOLRemitFunction, _
                    VariantStringAsString(m_typCorporateBodyData.sRemitFunction), 40, 4, gk_iDBCOLLenRemitFunction), _
                        typCellInputStyle
    ofwTable.EndRow

    
    'Variant text
    ofwTable.StartRow
    'Label
    ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrVariant), typCellLabelStyle
    'Text Field
    
    ofwTable.AddCellAsHTML m_ofwHtmlRef.InputField(ITText, gk_sDBCOLVariantText, 50, 50, _
                    VariantStringAsString(m_typCorporateBodyData.sVariantText)), typCellInputStyle
    ofwTable.EndRow
    
    'Validation text
    ofwTable.StartRow
    'Label
    ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrValidation), typCellLabelStyle
    'Text Field
    
    ofwTable.AddCellAsHTML m_ofwHtmlRef.InputField(ITText, gk_sDBCOLValidation_text, 50, 50, _
                    VariantStringAsString(m_typCorporateBodyData.sValidation)), typCellInputStyle
    ofwTable.EndRow
    
    'Jurisdiction text
    ofwTable.StartRow
    'Label
    ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrJurisdiction), typCellLabelStyle
    'Text Field
    '   d24 crm  16Feb2000
    'Changed the field length to 20 characters
    ofwTable.AddCellAsHTML m_ofwHtmlRef.InputField(ITText, gk_sDBCOLJurisdiction, 20, 20, _
                    VariantStringAsString(m_typCorporateBodyData.sJurisdiction)), typCellInputStyle
    ofwTable.EndRow
    
    'Is Place of Deposit
    ofwTable.StartRow
    'Label
    ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrPlaceOfDeposit), typCellLabelStyle
    'Check box.
        
    ofwTable.AddCellAsHTML m_ofwHtmlRef.Checkbox("", gk_sDBCOLCorporateBodyIsPlaceOfDep, _
          m_typCorporateBodyData.bIsPlaceOfDeposit, 1), typCellInputStyle
    ofwTable.EndRow
    
    'National place of deposit code
    ofwTable.StartRow
    'Label
    ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrNationalPlaceOfDepositCode), typCellLabelStyle
    'Text Field
    
    ofwTable.AddCellAsHTML m_ofwHtmlRef.InputField(ITText, gk_sDBCOLNationalPlaceOfDepositCode, 40, 40, _
                    VariantStringAsString(m_typCorporateBodyData.sPlaceOfDepositCode)), typCellInputStyle
    ofwTable.EndRow
    
    
    'd5     Display links here
     'See links
     'Flags to decide what combination of see and see also
     'buttons we display, there are 3 scenarios:
     '1.    Term is not linked to or from anything - display see and see also buttons
     '2.    Term is a non preferred term - display see link but no button
     '3.    Term is a preferred term or has see also links to/from it - display see also link(s) and button
     
     Dim bIsNonPreferred As Boolean
     Dim bIsPreferredOrLinkedTo As Boolean
     
     If m_colSeeLinks.Count = 0 Then
        bIsNonPreferred = False
     Else
        bIsNonPreferred = True
     End If
     
     ' d35 - Use cache rather than database
     If m_colSeeAlsoLinks.Count = 0 Then
        bIsPreferredOrLinkedTo = False
     Else
        bIsPreferredOrLinkedTo = True
     End If
         
     If Not bIsPreferredOrLinkedTo Or bIsNonPreferred Then
         If m_colSeeLinks.Count = 0 Then
            ofwTable.StartRow
            'Label
            ofwTable.AddCellAsHTML "See ", typCellLabelStyle
            ofwTable.AddCellAsHTML m_ofwHtmlRef.SubmitImage(gk_sImagesPath & gk_sImagePickCorporateBody, _
               mk_sAFSee, mk_sAFSee) & sMessage, typCellInputStyle
            ofwTable.EndRow
        Else
         
            Dim iTemp As Integer
            For iTemp = 1 To m_colSeeLinks.Count
                ofwTable.StartRow
                If iTemp = 1 Then
                    ofwTable.AddCellAsHTML "See ", typCellLabelStyle
                Else
                    ofwTable.AddCellAsHTML gk_sHTMLNonBreakSpace, typCellLabelStyle
                End If
                ofwTable.AddCellAsHTML m_colSeeLinks.Item(iTemp).Value, typCellLinks
                ofwTable.AddCellAsHTML m_ofwHtmlRef.SubmitImage(gk_sImagesPath & gk_sImageRemove, _
                    mk_sAFSeeRemove & m_colSeeLinks.Item(iTemp).ID, mk_sAFSeeRemove), typCellRemove
                ofwTable.EndRow
             Next
             'Label
            ofwTable.AddCellAsHTML gk_sHTMLNonBreakSpace, typCellLabelStyle
            If Not bIsNonPreferred Then
                ofwTable.AddCellAsHTML m_ofwHtmlRef.SubmitImage(gk_sImagesPath & gk_sImagePickCorporateBody, _
                    mk_sAFSee, mk_sAFSee) & sMessage, typCellInputStyle
            End If
            ofwTable.EndRow
         End If
     End If
     'See also links
     
     If Not bIsNonPreferred Then
        If m_colSeeAlsoLinks.Count = 0 Then
            ofwTable.StartRow
            'Label
            ofwTable.AddCellAsHTML "See also", typCellLabelStyle
            ofwTable.AddCellAsHTML m_ofwHtmlRef.SubmitImage(gk_sImagesPath & gk_sImagePickCorporateBody, _
                mk_sAFSeeAlso, mk_sAFSeeAlso), typCellInputStyle
            ofwTable.EndRow
         Else
    
            For iTemp = 1 To m_colSeeAlsoLinks.Count
                ofwTable.StartRow
                If iTemp = 1 Then
                    ofwTable.AddCellAsHTML "See also", typCellLabelStyle
                Else
                    ofwTable.AddCellAsHTML gk_sHTMLNonBreakSpace, typCellLabelStyle
                End If
                
                ofwTable.AddCellAsHTML m_colSeeAlsoLinks.Item(iTemp).Value, typCellLinks
                ofwTable.AddCellAsHTML m_ofwHtmlRef.SubmitImage(gk_sImagesPath & gk_sImageRemove, _
                    mk_sAFSeeAlsoRemove & m_colSeeAlsoLinks.Item(iTemp).ID, mk_sAFSeeAlsoRemove), typCellRemove
                ofwTable.EndRow
            Next
            ofwTable.AddCellAsHTML gk_sHTMLNonBreakSpace, typCellLabelStyle
            ofwTable.AddCellAsHTML m_ofwHtmlRef.SubmitImage(gk_sImagesPath & gk_sImagePickCorporateBody, _
                mk_sAFSeeAlso, mk_sAFSeeAlso), typCellInputStyle
            ofwTable.EndRow
         End If
     End If
     ofwTable.AddBlankRow
     ofwTable.AddBlankRow
     
     ofwTable.StartRow
     
     ofwTable.AddCellAsHTML gk_sHTMLNonBreakSpace, typCellLabelStyle
     
     'd5        display the buttons based on the status
     m_aspSession(mk_sAFAuthorityStatus) = m_typCorporateBodyData.lAuthorityStatus
     Dim sImages As String
     If Not m_typCorporateBodyData.lAuthorityStatus = gk_lApprovedComplete Then
     
         sImages = sImages & m_ofwHtmlRef.SubmitImage(gk_sImagesPath & gk_sImageApproveComplete, _
                  k_sApproveCompleteBtn, gk_lApprovedComplete)
     End If
     If Not m_typCorporateBodyData.lAuthorityStatus = gk_lApprovedInComplete Then
        If Not m_typCorporateBodyData.lAuthorityStatus = gk_lApprovedComplete Then
            sImages = sImages & m_ofwHtmlRef.SubmitImage(gk_sImagesPath & gk_sImageApproveIncomplete, _
                 k_sApproveInCompleteBtn, gk_lApprovedInComplete)
        End If
     End If
     
     If Not m_typCorporateBodyData.lAuthorityStatus = gk_lInvisible Then
        If Not m_typCorporateBodyData.lAuthorityStatus = gk_lApprovedComplete Then
            sImages = sImages & m_ofwHtmlRef.SubmitImage(gk_sImagesPath & gk_sImageInvisible, _
                 k_sInvisibleBtn, gk_lInvisible)
        End If
     End If
     If Not m_typCorporateBodyData.lAuthorityStatus = gk_lRejected Then
        
         sImages = sImages & m_ofwHtmlRef.SubmitImage(gk_sImagesPath & gk_sImageReject, _
                 k_sRejectBtn, gk_lRejected)
         
     End If
     typCellLabelStyle.lCOLSPAN = 2
     typCellLabelStyle.sWidth = "23%"
     ofwTable.AddCellAsHTML sImages, typCellLabelStyle
     'Show ok/cancel buttons

    typCellLabelStyle.lCOLSPAN = 1
    typCellLabelStyle.sWidth = "12%"
    ofwTable.AddCellAsHTML m_ofwHtmlRef.SubmitImage(gk_sImagesPath & gk_sImageOK, _
        k_sOKBtn, CStr(m_typCorporateBodyData.lAuthorityStatus)), typCellLabelStyle
    ofwTable.AddCellAsHTML m_ofwHtmlRef.ImageButton(gk_sImagesPath & gk_sImageCancel, "", _
        gk_sLinkDisplayAuthorityTerms & "?" & ksTermCategory & "=" _
        & m_aspSession(ksTermCategory), , m_ofwRefData.Search(gToolTip, ETOOLTIP.ttButtonCancelHelp)), typCellLabelStyle

     ofwTable.EndRow

    Set DisplayCorporateBodyDetails = ofwTable
    Set ofwTable = Nothing
Exit Function

ErrorHandler:
    ConvertError sSource
End Function

Public Sub SetStateBeforeLeave()
'Set the state before going to the index browser
'
'd5     AGA     21-Oct-1999     Created

    Dim sSource As String
    Dim iSeeType As Integer
    Dim iAddOrRemove As Integer
    Dim sSpecificTermType As String
    Dim stemp As String ' temp string
    Dim iAFGroup As Integer
    
    sSource = "VWEditorialAuthorityFilles.SetStateBeforeLeave()"
    On Error GoTo ErrorHandler:
    
    m_lSpecificTermType = m_aspRequest(gk_sFORMAFGroup)
    Select Case m_lSpecificTermType
        'Person AF
        Case EAFTermType.AFPerson
            'Get the state to modify the fields of the term
            m_oBOAFPerson.PROCATTransientInterface_GetState
            m_typPersonData = m_oBOAFPerson.Data
            
            m_lTermID = m_typPersonData.lID
            
            m_typPersonData.sSurname = m_aspRequest.Form(gk_sDBCOLSurname_Text)
            m_typPersonData.lAuthorityStatus = m_aspSession(mk_sAFAuthorityStatus)
            m_typPersonData.sGender = m_aspRequest.Form(gk_sDBCOLGender_Indicator)
            m_typPersonData.sTitle = m_aspRequest.Form(gk_sDBCOLPerson_Title)
            m_typPersonData.sAdditionalNameElements = m_aspRequest.Form(gk_sDBCOLadditional_elements_of_name)
            m_typPersonData.sPretitle = m_aspRequest.Form(gk_sDBCOLPretitle_Text)
            m_typPersonData.sForenames = m_aspRequest.Form(gk_sDBCOLForename_Text)
            m_typPersonData.sAlternativeStatus = m_aspRequest.Form(gk_sDBCOLAlternative_Status_Text)
            m_typPersonData.sBirthDate = m_aspRequest.Form(gk_sDBCOLBirth_Date)
            m_typPersonData.sUncertainBirthDateCode = m_aspRequest.Form(gk_sDBCOLUncertain_Birth_Date_Code)
            m_typPersonData.sBiogHistory = m_aspRequest.Form(gk_sDBCOLBiog_History_Text)
            m_typPersonData.sValidation = m_aspRequest.Form(gk_sDBCOLValidation_text)
            m_typPersonData.sDeathDate = m_aspRequest.Form(gk_sDBCOLDeath_Date)
            m_typPersonData.sUncertainDeathDateCode = m_aspRequest.Form(gk_sDBCOLUncertain_Death_Date_Code)
            'Update the modified values
            m_oBOAFPerson.Data = m_typPersonData
            'Set the state with modified values
           
            m_oBOAFPerson.PROCATTransientInterface_SetState
            
            iAFGroup = EAFTermType.AFPerson
            
        Case EAFTermType.AFSubject
            'Get the state to modify the fields of the term
            m_oBOAFSubject.PROCATTransientInterface_GetState
            m_typSubjectData = m_oBOAFSubject.Data
            
            m_lTermID = m_typSubjectData.lID
            
            'Update the modified values
            m_typSubjectData.sTermText = m_aspRequest.Form(gk_sDBCOLSubjectTermText)
            m_typSubjectData.lAuthorityStatus = m_aspSession(mk_sAFAuthorityStatus)
            m_typSubjectData.sDefinition = m_aspRequest.Form(gk_sDBCOLSubjectDefinition)
            m_typSubjectData.sValidation = m_aspRequest.Form(gk_sDBCOLSubjectValidation)
            'Update the modified values
            m_oBOAFSubject.Data = m_typSubjectData
            'Set the state with modified values
            m_oBOAFSubject.PROCATTransientInterface_SetState
            iAFGroup = EAFTermType.AFSubject
            
        Case EAFTermType.AFPlace
            'Get the state to modify the fields of the term
            m_oBOAFPlace.PROCATTransientInterface_GetState
            m_typPlaceData = m_oBOAFPlace.Data
            
            m_lTermID = m_typPlaceData.lID
            
            'Update the modified values
            m_typPlaceData.sTermText = m_aspRequest.Form(gk_sDBCOLPlaceName)
            m_typPlaceData.lAuthorityStatus = m_aspSession(mk_sAFAuthorityStatus)
            m_typPlaceData.sParish = m_aspRequest.Form(gk_sDBCOLParishName)
            m_typPlaceData.sTown = m_aspRequest.Form(gk_sDBCOLTownName)
            m_typPlaceData.sCounty = m_aspRequest.Form(gk_sDBCOLCountyName)
            m_typPlaceData.sStartDate = m_aspRequest.Form(gk_sDBCOLPlaceDateStart)
            m_typPlaceData.sEndDate = m_aspRequest.Form(gk_sDBCOLPlaceDateEnd)
            m_typPlaceData.sGridReference = m_aspRequest.Form(gk_sDBCOLPlaceGridRef)
            m_typPlaceData.sUncertainStartDateCode = m_aspRequest.Form(gk_sDBColUncertainDateCodeStart)
            m_typPlaceData.sUncertainEndDateCode = m_aspRequest.Form(gk_sDBColUncertainDateCodeEnd)
            m_typPlaceData.sValidation = m_aspRequest.Form(gk_sDBCOLValidation_text)
            m_typPlaceData.sCountry = m_aspRequest.Form(gk_sDBCOLCountryName)
            m_typPlaceData.sHistoryText = m_aspRequest.Form(gk_sDBCOLPlaceHistory)
            'Update the modified values
            m_oBOAFPlace.Data = m_typPlaceData
            'Set the state with modified values
            m_oBOAFPlace.PROCATTransientInterface_SetState
            
            iAFGroup = EAFTermType.AFPlace
            
        Case EAFTermType.AFCorporateBody
            'Get the state to modify the fields of the term
            m_oBOAFCorporateBody.PROCATTransientInterface_GetState
            m_typCorporateBodyData = m_oBOAFCorporateBody.Data
            
            m_lTermID = m_typCorporateBodyData.lID
            
            
            'Update the modified values
            m_typCorporateBodyData.sTermText = m_aspRequest.Form(gk_sDBCOLCorporateBodyName)
            m_typCorporateBodyData.lAuthorityStatus = m_aspSession(mk_sAFAuthorityStatus)
            m_typCorporateBodyData.sStartDate = m_aspRequest.Form(gk_sDBCOLCorporateBodyDateStart)
            m_typCorporateBodyData.sEndDate = m_aspRequest.Form(gk_sDBCOLCorporateBodyDateEnd)
            m_typCorporateBodyData.sUncertainStartDateCode = m_aspRequest.Form(gk_sDBCOLCorporateBodyUncertainStartDate)
            m_typCorporateBodyData.sUncertainEndDateCode = m_aspRequest.Form(gk_sDBCOLCorporateBodyUncertainEndDate)
            m_typCorporateBodyData.sRemitFunction = m_aspRequest.Form(gk_sDBCOLRemitFunction)
            m_typCorporateBodyData.sVariantText = m_aspRequest.Form(gk_sDBCOLVariantText)
            m_typCorporateBodyData.sValidation = m_aspRequest.Form(gk_sDBCOLValidation_text)
            m_typCorporateBodyData.sJurisdiction = m_aspRequest.Form(gk_sDBCOLJurisdiction)
            m_typCorporateBodyData.bIsPlaceOfDeposit = m_aspRequest.Form(gk_sDBCOLCorporateBodyIsPlaceOfDep)
            m_typCorporateBodyData.sPlaceOfDepositCode = m_aspRequest.Form(gk_sDBCOLNationalPlaceOfDepositCode)
            'Update the modified values
            m_oBOAFCorporateBody.Data = m_typCorporateBodyData
            'Set the state with modified values
            m_oBOAFCorporateBody.PROCATTransientInterface_SetState
            iAFGroup = EAFTermType.AFCorporateBody
            
        Case Else
            RaiseError sSource, EPROCATINTERNALERROR.peAuthorityFilesMaintenance, _
                           "Invalid Term type"
    End Select
    'Now redirect to the indexbrowser
    If m_aspRequest.Form(mk_sAFSee & ".x") <> "" Then
        iSeeType = 1
        iAddOrRemove = EAFTermLoadType.EAFLoadFromStateAddTerm
    End If
    If m_aspRequest.Form(mk_sAFSeeAlso & ".x") <> "" Then
        iSeeType = 2
        iAddOrRemove = EAFTermLoadType.EAFLoadFromStateAddTerm
    End If
    If InStr(m_aspRequest.Form, mk_sAFSeeRemove) Then
        iAddOrRemove = EAFTermLoadType.EAFLoadFromStateRemoveTerm
        'Extract the remove term ID from Querystring. It is appended to the image name
        stemp = m_aspRequest.Form
        stemp = Mid(stemp, InStr(stemp, mk_sAFSeeRemove))
        stemp = Mid(stemp, Len(mk_sAFSeeRemove) + 1, (InStr(stemp, ".x") - (Len(mk_sAFSeeRemove) + 1)))
        m_iTempRemoveTermID = CInt(stemp)
    
        
        stemp = gk_sFORMAFTermID & "=" & m_iTempRemoveTermID & "&" & mk_sFORMAFTermAddOrRemove & "=" & iAddOrRemove _
            & "&" & gk_sFORMAFGroup & "=" & m_lSpecificTermType
            
        'Store the values in the session
        m_aspSession(mk_sFORMAFTermIsRemoved) = 1   'If the term is removed
        m_aspSession(mk_sFORMAFRemoveTermID) = m_iTempRemoveTermID
        m_aspSession(mk_sFORMAFSeeLinkType) = EAFSeeLinkType.EAFSeeLink
        
        m_aspResponse.Redirect gk_sLinkDisplaySpecificIndexTermDetails & "?" & stemp
        
        Exit Sub
    End If
    
    If InStr(m_aspRequest.Form, mk_sAFSeeAlsoRemove) Then
        iAddOrRemove = EAFTermLoadType.EAFLoadFromStateRemoveTerm
        'Extract the remove term ID from Querystring. It is appended to the image name
        stemp = m_aspRequest.Form
        stemp = Mid(stemp, InStr(stemp, mk_sAFSeeAlsoRemove))
        stemp = Mid(stemp, Len(mk_sAFSeeAlsoRemove) + 1, (InStr(stemp, ".x") - (Len(mk_sAFSeeAlsoRemove) + 1)))
        m_iTempRemoveTermID = CInt(stemp)
        
    
        stemp = gk_sFORMAFTermID & "=" & m_iTempRemoveTermID & "&" & mk_sFORMAFTermAddOrRemove & "=" & iAddOrRemove _
            & "&" & gk_sFORMAFGroup & "=" & m_lSpecificTermType
        
        'Store the values in session
        m_aspSession(mk_sFORMAFTermIsRemoved) = 1       'term is removed
        m_aspSession(mk_sFORMAFRemoveTermID) = m_iTempRemoveTermID
        m_aspSession(mk_sFORMAFSeeLinkType) = EAFSeeLinkType.EAFSeealsoLink
        
        m_aspResponse.Redirect gk_sLinkDisplaySpecificIndexTermDetails & "?" & stemp
        m_objContext.SetComplete
        Exit Sub
    End If

    'Store the term type and term ID in session

    m_aspSession(mk_sFORMAFSeeLinkType) = iSeeType

    m_aspResponse.Redirect gk_sIndexBrowserAFPickTermPage & "?" _
         & gk_sFORMAFGroup & "=" & iAFGroup & "&" & gk_sFORMAFListIndex & "=-1" & _
         "&" & gk_sPARAMMaintainedTermID & "=" & m_lTermID
         
    m_objContext.SetComplete
    Exit Sub
ErrorHandler:
    m_objContext.SetComplete
    ShowError sSource
End Sub

Private Sub CheckPermissions(t_eFunctionality As ESecurityCheck, _
                         t_sSource As String, _
                        t_sDescription As String)
'Subroutine that checks and raises an error if the user has no rights to
' the passed function
'
'd5     aga     11Oct1999      Copied from VWCatalogueBulkUpdate and modified
'd7     AGA     02Nov1999      Modified due to change in Check method
'd8     KSK     03Nov1999      Changed Checkpermissions method to used the parameter source
'                              value while raising error
'd27    TMJ     29Feb2000      Added call to set the EditsetID in Editorial permissions
'
    
    Dim eCheckRights As ESecurityCheckResult
    
    m_oboEditorialPermission.UserID = m_lLoggedUserID
    
    If m_aspRequest(gk_sParamEditSetID) <> "" Then
        m_oboEditorialPermission.EditsetID = CLng(m_aspRequest(gk_sParamEditSetID))
    End If
    
    eCheckRights = m_oboEditorialPermission.Check(t_eFunctionality)
    
    If eCheckRights <> _
            ESecurityCheckResult.ESRSecurityCheckSuccessful Then
    
        RaiseError t_sSource, _
            GetErrorEnumForSecurityCheckResult(eCheckRights), _
            GetDescriptionForSecurityCheckError(GetErrorEnumForSecurityCheckResult(eCheckRights))
    End If
    
End Sub

Public Sub DisplayAddNewAFTerm()
'Shows the screen for adding a new AF term
'
'Amendment history:
' d5    aga     22Oct1999       Created
' d22   crm     08Feb2000       Functionality to disable and check IsPlaceOfDeposit check box
' d30   tmj     16Mar2000       Added code to determine origin of the page
'   d34 daj                   27Mar2000     Added SFA checking when adding new candidate terms, and also added code
'                                           to pass edit set ID around when adding it to an edit set entry
    On Error GoTo ErrorHandler
        'Table
    Dim ofwTable As FWTable
       
    Dim sSource As String
    Dim iSpecificTermType As Integer
    Dim eListUsage As EAFBrowserUsage
    Dim sHeading As String
    Dim sHTML As String
    Dim sParameter As String
    Dim bFromSFA
    Dim lEditSetID As Long
    
    sSource = "VWAuthorityFile.DisplayAddNewAFTerm"
    
    'Determine Whether this term originated from the SFA
    If m_aspRequest(gk_sFORMAFinSFA) = "1" Or _
       m_aspRequest(gk_sFORMAFinSFA) = "-1" Then
        bFromSFA = True
    Else
        bFromSFA = False
    End If
    ' d34 - Get the Edit Set ID to pass on if needed
    lEditSetID = 0
    If Not m_aspRequest(gk_sNameEditsetID) = "" Then
        lEditSetID = CLng(m_aspRequest(gk_sNameEditsetID))
    End If
    
    ' d34 - Check the permissions
    If bFromSFA Then
        CheckPermissions ESecurityCheck.EUserCanMaintainLeaflets, sSource, ""
    Else
        CheckPermissions ESecurityCheck.EUserCanEditCatalogueEntryOnEditSet, sSource, ""
    End If

     'Get the value of term type
    If m_aspRequest(gk_sFORMAFGroup) = "" Then
            RaiseError sSource, EPROCATINTERNALERROR.peAuthorityFilesMaintenance, _
                           "Term type is blank"
        
    End If
    iSpecificTermType = m_aspRequest(gk_sFORMAFGroup)
    
    If m_aspRequest(gk_sFORMAFListUsage) = "" Then
            RaiseError sSource, EPROCATINTERNALERROR.peAuthorityFilesMaintenance, _
                           "List usage is blank"
        
    End If
    eListUsage = m_aspRequest(gk_sFORMAFListUsage)
    
    'End of parameters section.
    
    Select Case iSpecificTermType
    'Person details.
        Case EAFTermType.AFPerson
            
           Set ofwTable = DisplayAddNewPersonAFTerm(eListUsage)
     'Subject details.
        Case EAFTermType.AFSubject
           
           Set ofwTable = DisplayAddNewSubjectAFTerm(eListUsage)
     
        Case EAFTermType.AFPlace
           Set ofwTable = DisplayAddNewPlaceAFTerm(eListUsage)
     
        Case EAFTermType.AFCorporateBody
           '   d22 crm    08Feb2000
           m_sDisablePlaceOfDeposit = m_aspRequest.QueryString(gk_sDisablePlaceOfDeposit)
           Set ofwTable = DisplayAddNewCorporateBodyAFTerm(eListUsage)
           
        Case Else
                  RaiseError sSource, EPROCATINTERNALERROR.peAuthorityFilesMaintenance, _
                           "Term type is not valid one"
    End Select
    
    
    'Get the table details
    
    'Formstart
    sHTML = sHTML & m_ofwHtmlRef.FormStart(gk_sAddAFTermResults, gk_sMainFrame, _
                                                gk_sHTMLDefaultFormName, MPost, True)
    'Add hidden fields here to send them to results page
    sHTML = sHTML & m_ofwHtmlRef.HiddenField(gk_sFORMAFGroup, CStr(iSpecificTermType))
    sHTML = sHTML & m_ofwHtmlRef.HiddenField(gk_sFORMAFListUsage, CStr(eListUsage))
    
    ' d34 - Changed inSFA to integer to match Leaflet functionality, and also added editset ID
    sHTML = sHTML & m_ofwHtmlRef.HiddenField(gk_sFORMAFinSFA, CInt(bFromSFA))
    sHTML = sHTML & m_ofwHtmlRef.HiddenField(gk_sNameEditsetID, CInt(lEditSetID))
    ' Title
    sHeading = m_ofwHtmlRef.Paragraph(m_ofwHtmlRef.Font( _
                m_ofwHtmlRef.Bold("Add new index terms"), COLOURBlack, 4))
    
    sHeading = sHeading & m_ofwHtmlRef.Paragraph("")
    
    sHTML = sHTML & gk_sHTMLLineBreak & sHeading
    
    sHTML = sHTML & m_ofwHtmlRef.Paragraph(ofwTable.GetHTML)
    
    sHTML = sHTML & m_ofwHtmlRef.EndTag(gk_sHTMLForm)
            
    'Add context sensitive help button
    sHTML = sHTML & HTMLToAddHelpButton(gk_sHelpMAFAddTermPage)
    
    'Write the html
    m_aspResponse.Write sHTML
    
    'Release the objects
    Set ofwTable = Nothing
    
    m_objContext.SetComplete
    Exit Sub
ErrorHandler:
    m_objContext.SetAbort
    ShowError sSource
End Sub

Private Function DisplayAddNewPersonAFTerm(t_eListUsage As EAFBrowserUsage) As FWTable
'returns the table contains the add new screen
'
'd5     aga     22Oct1999       Created
'd11    aga     04Nov1999       changed from const to function for the cancel link
'd34    tmj     28Mar2000       Added textarea validation

   
    On Error GoTo ErrorHandler
    'Table
    Dim ofwTable As FWTable
    
    'Table, row, cell styles
    Dim typTable As TableStyle
    Dim typCellLabelStyle As CellStyle
    Dim typCellInputStyle  As CellStyle
    
    Dim sSource As String
    Dim sHTML As String
    
    sSource = "VWEditorialAuthorityFiles.DisplayAddNewPersonAFTerm"
        
    Set ofwTable = New FWTable
    
    'Set the tables format
    typTable.eStyle = TABCustomStyle
    typTable.sWidth = "100%"
    ofwTable.TableStyle = typTable
    
    'Set the label columns style
    typCellLabelStyle = m_ofwHtmlRef.GetCellStyle(ECellStyle.CELLReadWriteLabelStyle)
    typCellLabelStyle.sWidth = "30%"
    
    'Set the text columns style.
    typCellInputStyle = m_ofwHtmlRef.GetCellStyle(ECellStyle.CELLReadWriteDataStyle)
    typCellInputStyle.sWidth = "70%"
    typCellInputStyle.lCOLSPAN = 2
    
    'First row
     ofwTable.StartRow
     'Label
     ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                         EAuthorityFilesMaintResourceFile.afmrSurname), typCellLabelStyle
     'Text Field
     
     ofwTable.AddCellAsHTML m_ofwHtmlRef.InputField(ITText, gk_sDBCOLSurname_Text, 20, 20, _
                      "", VCNone, True), typCellInputStyle
     ofwTable.EndRow
     
    'Third Row
     ofwTable.StartRow
     'Label
     
     
     ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                         EAuthorityFilesMaintResourceFile.afmrGenderIndicator), typCellLabelStyle
                          
     'Add two radio buttons with the same name to indicate gender.
     'One for Male
     
     ofwTable.AddCellAsHTML m_ofwHtmlRef.RadioButton("Male", gk_sDBCOLGender_Indicator, , _
                    gk_sMaleGender) & _
                     m_ofwHtmlRef.RadioButton("Female", gk_sDBCOLGender_Indicator, , _
                     gk_sFemaleGender), typCellInputStyle
     ofwTable.EndRow
     
    'Fourth Row
     ofwTable.StartRow
     'Label
     ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                         EAuthorityFilesMaintResourceFile.afmrTitle), typCellLabelStyle
     'Text Field
     ofwTable.AddCellAsHTML m_ofwHtmlRef.InputField(ITText, gk_sDBCOLPerson_Title, 80, 80, _
                  ""), typCellInputStyle
     ofwTable.EndRow
     
    'Fifth Row
     ofwTable.StartRow
     'Label
     ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                         EAuthorityFilesMaintResourceFile.afmrAdditionalElementsOfName), typCellLabelStyle
     'Text Field
     ofwTable.AddCellAsHTML m_ofwHtmlRef.InputField(ITText, gk_sDBCOLadditional_elements_of_name, 50, 50, _
                     ""), typCellInputStyle
     ofwTable.EndRow
     
    'Sixth Row
     ofwTable.StartRow
     'Label
     ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                         EAuthorityFilesMaintResourceFile.afmrPretitle), typCellLabelStyle
     'Text Field
     ofwTable.AddCellAsHTML m_ofwHtmlRef.InputField(ITText, gk_sDBCOLPretitle_Text, 5, 5, _
                 ""), typCellInputStyle
     ofwTable.EndRow
     
    'Seventh Row
     ofwTable.StartRow
     'Label
     ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                         EAuthorityFilesMaintResourceFile.afmrForename), typCellLabelStyle
     'Text Field
     ofwTable.AddCellAsHTML m_ofwHtmlRef.InputField(ITText, gk_sDBCOLForename_Text, 50, 50, _
                    ""), typCellInputStyle
     ofwTable.EndRow
     
    'Uncertain birth date code
     ofwTable.StartRow
     'Label
     ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                         EAuthorityFilesMaintResourceFile.afmrUncertainBirthdateCode), typCellLabelStyle
     'Text Field
     ofwTable.AddCellAsHTML m_ofwHtmlRef.InputField(ITText, gk_sDBCOLUncertain_Birth_Date_Code, 3, 3, _
                     ""), typCellInputStyle
     ofwTable.EndRow
     
     'Birthdate
     ofwTable.StartRow
     'Label
     ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                         EAuthorityFilesMaintResourceFile.afmrDateOfBirth), typCellLabelStyle
     'Text Field
     ofwTable.AddCellAsHTML m_ofwHtmlRef.InputField(ITText, gk_sDBCOLBirth_Date, gk_iDBCOLLenDateInAF, gk_iDBCOLLenDateInAF, _
                     "", VCNumber, , , 0, 9999), typCellInputStyle
     ofwTable.EndRow
     
    'Twelth Row
     ofwTable.StartRow
     'Label
     ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                         EAuthorityFilesMaintResourceFile.afmrBiogHistory), typCellLabelStyle
     'Text Field
     ofwTable.AddCellAsHTML m_ofwHtmlRef.TextArea(gk_sDBCOLBiog_History_Text, _
                     "", 40, 4, gk_iDBCOLLenBiogHistory), _
                         typCellInputStyle
     ofwTable.EndRow
     
    'Thirthteen Row
     ofwTable.StartRow
     'Label
     ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                         EAuthorityFilesMaintResourceFile.afmrValidation), typCellLabelStyle
     'Text Field
     ofwTable.AddCellAsHTML m_ofwHtmlRef.InputField(ITText, gk_sDBCOLValidation_text, 50, 50, _
                     ""), typCellInputStyle
     ofwTable.EndRow
     
     'Uncertain death date code
     ofwTable.StartRow
     'Label
     ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                         EAuthorityFilesMaintResourceFile.afmrUncertainDeathDateCode), typCellLabelStyle
     'Text Field
     ofwTable.AddCellAsHTML m_ofwHtmlRef.InputField(ITText, gk_sDBCOLUncertain_Death_Date_Code, 3, 3, _
                    ""), typCellInputStyle
     ofwTable.EndRow
     
    'death date code
     ofwTable.StartRow
     'Label
     ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                         EAuthorityFilesMaintResourceFile.afmrDeathDate), typCellLabelStyle
     'Text Field
     ofwTable.AddCellAsHTML m_ofwHtmlRef.InputField(ITText, gk_sDBCOLDeath_Date, gk_iDBCOLLenDateInAF, gk_iDBCOLLenDateInAF, _
                     "", VCNumber, , , 0, 9999), typCellInputStyle
     ofwTable.EndRow
     
    
     
     ofwTable.AddBlankRow
     ofwTable.AddBlankRow
     
     ofwTable.StartRow
     ofwTable.AddCellAsHTML gk_sHTMLNonBreakSpace
     
     'd10 ksk changed from const to function for the cancel link
     'Show ok/cancel buttons
     ofwTable.AddCellAsHTML m_ofwHtmlRef.SubmitImage(gk_sImagesPath & gk_sImageOK, _
                 k_sOKBtn, "") & m_ofwHtmlRef.ImageButton(gk_sImagesPath & gk_sImageCancel, "", _
         GetBrowserRedirectPageForCancelAFTerm(t_eListUsage), , m_ofwRefData.Search(gToolTip, ETOOLTIP.ttButtonCancelHelp))
         
     ofwTable.EndRow

    Set DisplayAddNewPersonAFTerm = ofwTable
    Set ofwTable = Nothing
    Exit Function
ErrorHandler:
    ConvertError sSource
End Function
Private Function DisplayAddNewSubjectAFTerm(t_eListUsage As EAFBrowserUsage) As FWTable
'returns the table contains the add new screen
'
'd5     aga     22Oct1999       Created
'd11    aga     04Nov1999       changed from const to function for the cancel link
    On Error GoTo ErrorHandler
    'Table
    Dim ofwTable As FWTable
    
    'Table, row, cell styles
    Dim typTable As TableStyle
    Dim typCellLabelStyle As CellStyle
    Dim typCellInputStyle  As CellStyle
 
    Dim sSource As String
    Dim sHTML As String
    
    Dim sParameter As String
    Dim SApproveCompleteStr As String
    Dim SApproveInCompleteStr As String
    Dim SRejectStr As String
    
    sSource = "VWEditorialAuthorityFiles.DisplayAddNewSubjectAFTerm"
    
    'Display the data from the collection returned.
    Set ofwTable = New FWTable
    
    'Set the tables format
    typTable.eStyle = TABCustomStyle
    typTable.sWidth = "100%"
    ofwTable.TableStyle = typTable
    
    'Set the label columns style
    typCellLabelStyle = m_ofwHtmlRef.GetCellStyle(ECellStyle.CELLReadWriteLabelStyle)
    typCellLabelStyle.sWidth = "30%"
    
    'Set the text columns style.
    typCellInputStyle = m_ofwHtmlRef.GetCellStyle(ECellStyle.CELLReadWriteDataStyle)
    typCellInputStyle.sWidth = "70%"
    typCellInputStyle.lCOLSPAN = 2
        
    'First Row
    ofwTable.StartRow
    'Label
    ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrName), typCellLabelStyle
    'Text Field
    
    ofwTable.AddCellAsHTML m_ofwHtmlRef.InputField(ITText, gk_sDBCOLSubjectTermText, 100, 100, _
                    "", VCNone, True), typCellInputStyle
    ofwTable.EndRow
    
    'Second Row
    ofwTable.StartRow
    'Label
    ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrDefinition), typCellLabelStyle
    'Text Field
    
    ofwTable.AddCellAsHTML m_ofwHtmlRef.InputField(ITText, gk_sDBCOLSubjectDefinition, 100, 100, _
                    VariantStringAsString(m_typSubjectData.sDefinition)), typCellInputStyle
    ofwTable.EndRow
    
    'Third Row
    ofwTable.StartRow
    'Label
    ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrValidation), typCellLabelStyle
    'Text Field
    
    ofwTable.AddCellAsHTML m_ofwHtmlRef.InputField(ITText, gk_sDBCOLSubjectValidation, 100, 100, _
                    ""), typCellInputStyle
    ofwTable.EndRow
    
    ofwTable.AddBlankRow
    ofwTable.AddBlankRow
    'd11   aga      changed from const to function for the cancel link
     'Show ok/cancel button
     ofwTable.AddCellAsHTML m_ofwHtmlRef.SubmitImage(gk_sImagesPath & gk_sImageOK, _
                 k_sOKBtn, "") & m_ofwHtmlRef.ImageButton(gk_sImagesPath & gk_sImageCancel, "", _
         GetBrowserRedirectPageForCancelAFTerm(t_eListUsage), , m_ofwRefData.Search(gToolTip, ETOOLTIP.ttButtonCancelHelp))
         
     ofwTable.EndRow

    
    Set DisplayAddNewSubjectAFTerm = ofwTable
    Set ofwTable = Nothing
    Exit Function
ErrorHandler:
    ConvertError sSource
End Function
Private Function DisplayAddNewPlaceAFTerm(t_eListUsage As EAFBrowserUsage) As FWTable
'returns the table contains the add new screen
'
'd5     aga     22Oct1999       Created
'd11    aga     04Nov1999       changed from const to function for the cancel link
'd24    crm     16Feb2000       Changed field length
'd34    tmj     28Mar2000       Added textarea validation

    Dim sSource As String
    On Error GoTo ErrorHandler
    'Table
    Dim ofwTable As FWTable
    
    'Table, row, cell styles
    Dim typTable As TableStyle
    Dim typCellLabelStyle As CellStyle
    Dim typCellInputStyle  As CellStyle
        
    Dim sHTML As String
    
    sSource = "VWEditorialAuthorityFiles.DisplayAddNewPlaceAFTerm"
    
    'Display the data from the collection returned.
    Set ofwTable = New FWTable
    
    'Set the tables format
    typTable.eStyle = TABCustomStyle
    typTable.sWidth = "100%"
    ofwTable.TableStyle = typTable
    
    'Set the label columns style
    typCellLabelStyle = m_ofwHtmlRef.GetCellStyle(ECellStyle.CELLReadWriteLabelStyle)
    typCellLabelStyle.sWidth = "30%"
    
    'Set the text columns style.
    typCellInputStyle = m_ofwHtmlRef.GetCellStyle(ECellStyle.CELLReadWriteDataStyle)
    typCellInputStyle.sWidth = "70%"
    typCellInputStyle.lCOLSPAN = 2
    
    'Index Name
    ofwTable.StartRow
    'Label
    ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrName), typCellLabelStyle
    'Text Field
    
    ofwTable.AddCellAsHTML m_ofwHtmlRef.InputField(ITText, gk_sDBCOLPlaceName, 30, 30, _
                    "", VCNone, True), typCellInputStyle
    ofwTable.EndRow
    
    'Parish text
    ofwTable.StartRow
    'Label
    ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrParish), typCellLabelStyle
    'Text Field
    
    ofwTable.AddCellAsHTML m_ofwHtmlRef.InputField(ITText, gk_sDBCOLParishName, 30, 30, _
                    ""), typCellInputStyle
    ofwTable.EndRow
    
    'Town text
    ofwTable.StartRow
    'Label
    ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrTown), typCellLabelStyle
    'Text Field
    
    ofwTable.AddCellAsHTML m_ofwHtmlRef.InputField(ITText, gk_sDBCOLTownName, 30, 30, _
                    ""), typCellInputStyle
    ofwTable.EndRow
    
    'Country ield
    ofwTable.StartRow
    'Label
    ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrCountry), typCellLabelStyle
    'Text Field
    
    ofwTable.AddCellAsHTML m_ofwHtmlRef.InputField(ITText, gk_sDBCOLCountryName, 40, 40, _
                    ""), typCellInputStyle
    ofwTable.EndRow

    'County text
    ofwTable.StartRow
    'Label
    ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrCounty), typCellLabelStyle
    'Text Field
    
    ofwTable.AddCellAsHTML m_ofwHtmlRef.InputField(ITText, gk_sDBCOLCountyName, 50, 50, _
                    ""), typCellInputStyle
    ofwTable.EndRow
    
    'uncertain Date start code
    ofwTable.StartRow
    'Label
    ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrUncertainStartDateCode), typCellLabelStyle
    'Text Field
    'd24    crm     16Feb2000       Changed field length
    ofwTable.AddCellAsHTML m_ofwHtmlRef.InputField(ITText, gk_sDBCOLPlaceUncertainStartDate, 3, 3, _
                    "", VCNone), typCellInputStyle
    ofwTable.EndRow
    
    'Place Date start
    ofwTable.StartRow
    'Label
    ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrStartDate), typCellLabelStyle
    'Text Field
    
    ofwTable.AddCellAsHTML m_ofwHtmlRef.InputField(ITText, gk_sDBCOLPlaceDateStart, gk_iDBCOLLenDateInAF, gk_iDBCOLLenDateInAF, _
                    "", VCNumber, , , 0, 9999), typCellInputStyle
    ofwTable.EndRow
    
    'uncertain Date end code
    ofwTable.StartRow
    'Label
    ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrUncertainEndDateCode), typCellLabelStyle
    'Text Field
    
    ofwTable.AddCellAsHTML m_ofwHtmlRef.InputField(ITText, gk_sDBCOLPlaceUncertainEndDate, 3, 3, _
                    "", VCNone), typCellInputStyle
    ofwTable.EndRow
    
    'Place Date end
    ofwTable.StartRow
    'Label
    ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrEndDate), typCellLabelStyle
    'Text Field
    
    ofwTable.AddCellAsHTML m_ofwHtmlRef.InputField(ITText, gk_sDBCOLPlaceDateEnd, gk_iDBCOLLenDateInAF, gk_iDBCOLLenDateInAF, _
                    "", VCNumber, , , 0, 9999), typCellInputStyle
    ofwTable.EndRow

    'Grid references text
    ofwTable.StartRow
    'Label
    ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrGridReferences), typCellLabelStyle
    'Text Field
    
    ofwTable.AddCellAsHTML m_ofwHtmlRef.InputField(ITText, gk_sDBCOLPlaceGridRef, 50, 50, _
                    ""), typCellInputStyle
    ofwTable.EndRow
       
    
    'Validation text
    ofwTable.StartRow
    'Label
    ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrValidation), typCellLabelStyle
    'Text Field
    
    ofwTable.AddCellAsHTML m_ofwHtmlRef.InputField(ITText, gk_sDBCOLValidation_text, 50, 50, _
                    "", VCNone), typCellInputStyle
    ofwTable.EndRow
        
    'History
    ofwTable.StartRow
    'Label
    ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrHistory), typCellLabelStyle
    'Text Field
    ofwTable.AddCellAsHTML m_ofwHtmlRef.TextArea(gk_sDBCOLPlaceHistory, _
                    "", 40, 4, gk_iDBCOLLenPlaceHistory), _
                        typCellInputStyle
    ofwTable.EndRow
    
    ofwTable.AddBlankRow
    ofwTable.AddBlankRow
    
    'd11   aga      04Nov1999       changed from const to function for the cancel link
     'Show ok/cancel button
     ofwTable.AddCellAsHTML m_ofwHtmlRef.SubmitImage(gk_sImagesPath & gk_sImageOK, _
                 k_sOKBtn, "") & m_ofwHtmlRef.ImageButton(gk_sImagesPath & gk_sImageCancel, "", _
         GetBrowserRedirectPageForCancelAFTerm(t_eListUsage), , m_ofwRefData.Search(gToolTip, ETOOLTIP.ttButtonCancelHelp))
         
     ofwTable.EndRow
    
    Set DisplayAddNewPlaceAFTerm = ofwTable
    Set ofwTable = Nothing
    Exit Function
ErrorHandler:
    ConvertError sSource
End Function
Private Function DisplayAddNewCorporateBodyAFTerm(t_eListUsage As EAFBrowserUsage) As FWTable
'returns the table contains the add new screen
'
'd5     aga     22Oct1999       Created
'd11    aga     04Nov1999       changed from const to function for the cancel link
'd24    crm     16Feb2000       Changed field length
'd33    tmj     27Mar2000       Added hidden field when is place of deposit checkbox
'                               is disabled

    On Error GoTo ErrorHandler
    'Table
    Dim ofwTable As FWTable
    
    'Table, row, cell styles
    Dim typTable As TableStyle
    Dim typCellLabelStyle As CellStyle
    Dim typCellInputStyle  As CellStyle
      
    Dim sSource As String
    Dim sHTML As String
    
    sSource = "VWEditorialAuthorityFiles.DisplayAddNewCorporateBodyAFTerm"
    
    'Display the data from the collection returned.
    Set ofwTable = New FWTable
    
    'Set the tables format
    typTable.eStyle = TABCustomStyle
    typTable.sWidth = "100%"
    ofwTable.TableStyle = typTable

    'Set the label columns style
    typCellLabelStyle = m_ofwHtmlRef.GetCellStyle(ECellStyle.CELLReadWriteLabelStyle)
    typCellLabelStyle.sWidth = "30%"
    
    'Set the text columns style.
    typCellInputStyle = m_ofwHtmlRef.GetCellStyle(ECellStyle.CELLReadWriteDataStyle)
    typCellInputStyle.sWidth = "70%"
    typCellInputStyle.lCOLSPAN = 2
    
    ' Name
    ofwTable.StartRow
    'Label
    ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrName), typCellLabelStyle
    'Text Field
    
    ofwTable.AddCellAsHTML m_ofwHtmlRef.InputField(ITText, gk_sDBCOLCorporateBodyName, 90, 90, _
                    "", VCNone, True), typCellInputStyle
    ofwTable.EndRow

    'Uncertain date code start
    ofwTable.StartRow
    'Label
    ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrUncertainStartDateCode), typCellLabelStyle
    'Text Field
    
    ofwTable.AddCellAsHTML m_ofwHtmlRef.InputField(ITText, gk_sDBColUncertainDateCodeStart, 3, 3, _
                    ""), typCellInputStyle
    ofwTable.EndRow
    
    
    ' Corporate Body Date Start
    ofwTable.StartRow
    'Label
    ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrStartDate), typCellLabelStyle
    'Text Field
    
    ofwTable.AddCellAsHTML m_ofwHtmlRef.InputField(ITText, gk_sDBCOLCorporateBodyDateStart, gk_iDBCOLLenDateInAF, gk_iDBCOLLenDateInAF, _
                    "", VCNumber, False, False, 0, 9999), typCellInputStyle
    ofwTable.EndRow

    'Uncertain date code end
    ofwTable.StartRow
    'Label
    ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrUncertainEndDateCode), typCellLabelStyle
    'Text Field
    
    ofwTable.AddCellAsHTML m_ofwHtmlRef.InputField(ITText, gk_sDBColUncertainDateCodeEnd, 3, 3, _
                    ""), typCellInputStyle
    ofwTable.EndRow
    
    ' Corporate Body Date End
    ofwTable.StartRow
    'Label
    ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrEndDate), typCellLabelStyle
    'Text Field
    
    ofwTable.AddCellAsHTML m_ofwHtmlRef.InputField(ITText, gk_sDBCOLCorporateBodyDateEnd, gk_iDBCOLLenDateInAF, gk_iDBCOLLenDateInAF, _
                    "", VCNumber, False, False, 0, 9999), typCellInputStyle
    ofwTable.EndRow

    'Remit and function
    ofwTable.StartRow
    'Label
    ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrRemitAndFunction), typCellLabelStyle
    'Text Area
    ofwTable.AddCellAsHTML m_ofwHtmlRef.TextArea(gk_sDBCOLRemitFunction, _
                    "", 40, 4, gk_iDBCOLLenRemitFunction), _
                        typCellInputStyle
    ofwTable.EndRow
    
    
    'Variant text
    ofwTable.StartRow
    'Label
    ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrVariant), typCellLabelStyle
    'Text Field
    
    ofwTable.AddCellAsHTML m_ofwHtmlRef.InputField(ITText, gk_sDBCOLVariantText, 50, 50, _
                    ""), typCellInputStyle
    ofwTable.EndRow
    
    'Validation text
    ofwTable.StartRow
    'Label
    ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrValidation), typCellLabelStyle
    'Text Field
    
    ofwTable.AddCellAsHTML m_ofwHtmlRef.InputField(ITText, gk_sDBCOLValidation_text, 50, 50, _
                    ""), typCellInputStyle
    ofwTable.EndRow
    
    'Jurisdiction text
    ofwTable.StartRow
    'Label
    ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrJurisdiction), typCellLabelStyle
    'Text Field
    'd24    crm     16Feb2000
    ofwTable.AddCellAsHTML m_ofwHtmlRef.InputField(ITText, gk_sDBCOLJurisdiction, 20, 20, _
                    ""), typCellInputStyle
    ofwTable.EndRow
    
    'Is Place of Deposit
    ofwTable.StartRow
    'Label
    ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrPlaceOfDeposit), typCellLabelStyle
    'Check box.
    'TODO
    'd33 Added hidden field here
    If Trim(m_sDisablePlaceOfDeposit) <> "TRUE" Then
        ofwTable.AddCellAsHTML m_ofwHtmlRef.Checkbox("", gk_sDBCOLCorporateBodyIsPlaceOfDep, _
              False, 1), typCellInputStyle
    Else
        ofwTable.AddCellAsHTML m_ofwHtmlRef.Checkbox("", gk_sDBCOLCorporateBodyIsPlaceOfDep, _
              True, 1, , , False) & " " & _
                               m_ofwHtmlRef.HiddenField(gk_sDBCOLCorporateBodyIsPlaceOfDep, "1") _
                               , typCellInputStyle
        
    End If
    ofwTable.EndRow
    
    'National place of deposit code
    ofwTable.StartRow
    'Label
    ofwTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                        EAuthorityFilesMaintResourceFile.afmrNationalPlaceOfDepositCode), typCellLabelStyle
    'Text Field
    
    ofwTable.AddCellAsHTML m_ofwHtmlRef.InputField(ITText, gk_sDBCOLNationalPlaceOfDepositCode, 40, 40, _
                    ""), typCellInputStyle
    ofwTable.EndRow
    ofwTable.AddBlankRow
    ofwTable.AddBlankRow
    'd11   aga      changed from const to function for the cancel link
     'Show ok/cancel button
     ofwTable.AddCellAsHTML m_ofwHtmlRef.SubmitImage(gk_sImagesPath & gk_sImageOK, _
                 k_sOKBtn, "") & m_ofwHtmlRef.ImageButton(gk_sImagesPath & gk_sImageCancel, "", _
         GetBrowserRedirectPageForCancelAFTerm(t_eListUsage), , m_ofwRefData.Search(gToolTip, ETOOLTIP.ttButtonCancelHelp))
         
     ofwTable.EndRow

    Set DisplayAddNewCorporateBodyAFTerm = ofwTable
    Set ofwTable = Nothing
    Exit Function
ErrorHandler:
    ConvertError sSource
End Function

Public Sub AddNewAFTermResults()
'   Adds the new AF term and redirecs to the relevent page
'
'   d5       aga    22Oct1999        Created
'   d6       aga    29Oct1999        Minor modification: Calling GetBrowserRedirectPageAfterAddAFTerm function
'                                               to return the redirection page name
'   d34 daj                   27Mar2000     Added SFA checking when adding new candidate terms, and also added code
'                                           to pass edit set ID around when adding it to an edit set entry
'
On Error GoTo ErrorHandler
                 
    'Strings to store variables.
    Dim sSource As String
    Dim sParameter As String
    Dim iSpecificTermType As Integer
    Dim eListUsage As EAFBrowserUsage
    Dim lIndexID    As Long
    Dim lIndexTermRaisedSource As Long
    Dim lAuthorityStatus As Long
    Dim sTermValue As String
    Dim bFromSFA As Boolean
    
    Dim sRedirectPage As String
    
    sSource = "VWEditorialAuthorityFiles.AddNewAFTermResults"
    
    ' d34 - Check to see if this came from the leaflet editor
    bFromSFA = CBool(m_aspRequest(gk_sFORMAFinSFA))

    lIndexID = -1  ' For new term make the index < 0
    lAuthorityStatus = gk_lNotApproved ' all new terms will become candidate terms
    
    ' d34 - Check the permissions
    If bFromSFA Then
        CheckPermissions ESecurityCheck.EUserCanMaintainLeaflets, sSource, ""
    Else
        CheckPermissions ESecurityCheck.EUserCanEditCatalogueEntryOnEditSet, sSource, ""
    End If
    

         '********** Get PARAMETER values *******************'
    'Get the values from prev page. Check for term types.
  
    'Get the value of term type
    If m_aspRequest(gk_sFORMAFGroup) = "" Then
          RaiseError sSource, EPROCATINTERNALERROR.peAuthorityFilesMaintenance, _
                           "Term type is blank"
    End If
    iSpecificTermType = m_aspRequest(gk_sFORMAFGroup)
    
    If m_aspRequest(gk_sFORMAFListUsage) = "" Then
          RaiseError sSource, EPROCATINTERNALERROR.peAuthorityFilesMaintenance, _
                           "List usage is blank"
    End If
    eListUsage = CInt(m_aspRequest(gk_sFORMAFListUsage))
    
    Select Case iSpecificTermType
        Case EAFTermType.AFPerson
            
            Set m_oiAFSaver = New BOAFPerson
            
            m_typPersonData.sSurname = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLSurname_Text))
            m_typPersonData.lAuthorityStatus = lAuthorityStatus
            m_typPersonData.sGender = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLGender_Indicator))
            m_typPersonData.sTitle = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLPerson_Title))
            m_typPersonData.sAdditionalNameElements = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLadditional_elements_of_name))
            m_typPersonData.sPretitle = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLPretitle_Text))
            m_typPersonData.sForenames = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLForename_Text))
            m_typPersonData.sAlternativeStatus = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLAlternative_Status_Text))
            m_typPersonData.sBirthDate = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLBirth_Date))
            m_typPersonData.sUncertainBirthDateCode = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLUncertain_Birth_Date_Code))
            m_typPersonData.sBiogHistory = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLBiog_History_Text))
            m_typPersonData.sValidation = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLValidation_text))
            m_typPersonData.sDeathDate = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLDeath_Date))
            m_typPersonData.sUncertainDeathDateCode = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLUncertain_Death_Date_Code))
            'Load the data back to the class.
            m_oiAFSaver.Data = m_typPersonData
            'Call the save method.
            lIndexID = m_oiAFSaver.Save(lIndexID)
            'Now refresh the reference data to show the new records in the recordlist
            m_ofwRefData.Refresh (EGroup.gPersonOnly)
            
            sTermValue = m_typPersonData.sSurname
            
        Case EAFTermType.AFSubject
        
            Set m_oiAFSaver = New BOAFSubject
            
            m_typSubjectData.sTermText = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLSubjectTermText))
            m_typSubjectData.lAuthorityStatus = lAuthorityStatus
            m_typSubjectData.sDefinition = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLSubjectDefinition))
            m_typSubjectData.sValidation = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLSubjectValidation))
            
            'Load the data back to the class.
            m_oiAFSaver.Data = m_typSubjectData

            'Call the save method.
            lIndexID = m_oiAFSaver.Save(lIndexID)
            'Now refresh the reference data to show the new records in the recordlist
            m_ofwRefData.Refresh (EGroup.gSubjectOnly)
            sTermValue = m_typSubjectData.sTermText
            
        Case EAFTermType.AFPlace
        
            Set m_oiAFSaver = New BOAFPlace
            
            m_typPlaceData.sTermText = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLPlaceName))
            m_typPlaceData.lAuthorityStatus = lAuthorityStatus
            m_typPlaceData.sParish = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLParishName))
            m_typPlaceData.sTown = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLTownName))
            m_typPlaceData.sCounty = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLCountyName))
            m_typPlaceData.sStartDate = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLPlaceDateStart))
            m_typPlaceData.sEndDate = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLPlaceDateEnd))
            m_typPlaceData.sGridReference = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLPlaceGridRef))
            m_typPlaceData.sUncertainStartDateCode = StringToNullOrString(m_aspRequest.Form(gk_sDBColUncertainDateCodeStart))
            m_typPlaceData.sUncertainEndDateCode = StringToNullOrString(m_aspRequest.Form(gk_sDBColUncertainDateCodeEnd))
            m_typPlaceData.sValidation = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLValidation_text))
            m_typPlaceData.sCountry = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLCountryName))
            m_typPlaceData.sHistoryText = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLPlaceHistory))
            'Load the data back to the class.
            m_oiAFSaver.Data = m_typPlaceData
            'Call the save method.
            lIndexID = m_oiAFSaver.Save(lIndexID)
            'Now refresh the reference data to show the new records in the recordlist
            m_ofwRefData.Refresh (EGroup.gPlaceOnly)
        
            sTermValue = m_typPlaceData.sTermText
            
        Case EAFTermType.AFCorporateBody
            
            Set m_oiAFSaver = New BOAFCorporateBody
            
            m_typCorporateBodyData.sTermText = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLCorporateBodyName))
            m_typCorporateBodyData.lAuthorityStatus = lAuthorityStatus
            m_typCorporateBodyData.sStartDate = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLCorporateBodyDateStart))
            m_typCorporateBodyData.sEndDate = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLCorporateBodyDateEnd))
            m_typCorporateBodyData.sUncertainStartDateCode = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLCorporateBodyUncertainStartDate))
            m_typCorporateBodyData.sUncertainEndDateCode = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLCorporateBodyUncertainEndDate))
            m_typCorporateBodyData.sRemitFunction = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLRemitFunction))
            m_typCorporateBodyData.sVariantText = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLVariantText))
            m_typCorporateBodyData.sValidation = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLValidation_text))
            m_typCorporateBodyData.sJurisdiction = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLJurisdiction))
            m_typCorporateBodyData.bIsPlaceOfDeposit = m_aspRequest.Form(gk_sDBCOLCorporateBodyIsPlaceOfDep)
            m_typCorporateBodyData.sPlaceOfDepositCode = StringToNullOrString(m_aspRequest.Form(gk_sDBCOLNationalPlaceOfDepositCode))
            
            'Load the data back to the class.
            m_oiAFSaver.Data = m_typCorporateBodyData
            'Call the save method.
            lIndexID = m_oiAFSaver.Save(lIndexID)
            'Now refresh the reference data to show the new records in the recordlist
            m_ofwRefData.Refresh (EGroup.gCorporateBodyOnly)
            
            sTermValue = m_typCorporateBodyData.sTermText
                    
    End Select
    
    'Get the redirection page
    sRedirectPage = GetBrowserRedirectPageAfterAddAFTerm(eListUsage)
    
    'DEPENDS: Dependent modules: KSK, NDK, CRM
    'sParameter = "?" & _
                 gk_sFORMAFTermID & "=" & lIndexID & "&" & _
                gk_sFORMAFListTermType & "=" & iSpecificTermType & "&" & _
                gk_sFORMAFTermValue & "=" & m_aspServer.URLEncode(sTermValue)
    'Changed the parameter name from gk_sFORMAFListTermType to gk_sFORMAFGroup
    sParameter = "?" & _
                 gk_sFORMAFTermID & "=" & lIndexID & "&" & _
                gk_sFORMAFGroup & "=" & iSpecificTermType & "&" & _
                gk_sFORMAFTermValue & "=" & m_aspServer.URLEncode(sTermValue)
                
    If sRedirectPage = "" Then
        RaiseError sSource, PEGeneralInternalError, "Invalid redirection page"
        
    End If
    
     'Redirects to the proper page
    m_aspResponse.Redirect sRedirectPage & sParameter
    
    m_objContext.SetComplete
    Exit Sub
ErrorHandler:
    ' Only use SetAbort if this method did an update,
    ' (and then there should have been no transient storage).
    ' Otherwise use SetComplete
    m_objContext.SetComplete
    
    ShowError sSource
              
End Sub

Private Function GetBrowserRedirectPageAfterAddAFTerm(t_eListUsage As EAFBrowserUsage) As String
'Returns the Page name to be redirected after adding a new AF Term. It depends on BrowserUsage
'
'PARAMETER: t_eListUsage -> Enum to find which module is using (eg,. Editorial Maintenance etc.)
'  d6  AGA      29Oct1999       Created
'  d10 KSK      04Nov1999       Added redirection for bulkupdate. And method for Cancel
'                               redirection page name.
'  d13 ndk      04Nov1999       Code put for pickterm for edit catalogue entry
'  d20 crm      04Feb2000       Added redirection URl for Leaflets maintenance

    Select Case t_eListUsage
        Case EAFBrowserUsage.BROWSEEditorialMaintain
             
             GetBrowserRedirectPageAfterAddAFTerm = gk_sLinkDisplaySpecificIndexTermDetails
        
        Case EAFBrowserUsage.BROWSEEditorialBulkupdatePickTerm
            
            'The link is present in the root and the current directory is maintenance
             GetBrowserRedirectPageAfterAddAFTerm = gk_sLinkAddLinkTermToBulkUpdate
             
        Case EAFBrowserUsage.BROWSEEditorialPickTerm
            
            GetBrowserRedirectPageAfterAddAFTerm = gk_sLinkAddLinkTermToCatalogue
        '  d20 crm      04Feb2000
        Case EAFBrowserUsage.BROWSEEditorialInformationLeafletPickTerm
            GetBrowserRedirectPageAfterAddAFTerm = "/editorial/" & gk_sAddTermInformationLeaflet
            
        Case Else
            
            'Send a blank string
            GetBrowserRedirectPageAfterAddAFTerm = ""
    
    End Select
    
End Function

Private Function GetBrowserRedirectPageForCancelAFTerm(t_eListUsage As EAFBrowserUsage) As String
'Returns the Page name to be redirected after th user hits on cancel button
'
'PARAMETER: t_eListUsage -> Enum to find which module is using (eg,. Editorial Maintenance etc.)
'  d10 KSK      04Nov1999       Added redirection for bulkupdate. And method for Cancel
'                               redirection page name.
'  d13 ndk      04Nov1999       Code put for pickterm for edit catalogue entry
'  d23 crm      08Feb2000       Code to redirect cancel operation for adding new term in Leaflets
'  d28 tmj      09Mar2000       Removed leading ../ from URLs so point to current directory.
    Select Case t_eListUsage
        Case EAFBrowserUsage.BROWSEEditorialMaintain
             
             GetBrowserRedirectPageForCancelAFTerm = gk_sLinkDisplaySpecificIndexTermDetails
        
        Case EAFBrowserUsage.BROWSEEditorialBulkupdatePickTerm
             
            'The link is present in the root and the current directory is maintenance
             GetBrowserRedirectPageForCancelAFTerm = gk_sLinkCancelIndextermSelectionForBulkUpdate
        
        Case EAFBrowserUsage.BROWSEEditorialPickTerm
            
            GetBrowserRedirectPageForCancelAFTerm = gk_sPAGEAFCancelPickTerm
        '  d23 crm      08Feb2000
        Case EAFBrowserUsage.BROWSEEditorialInformationLeafletPickTerm
            
            GetBrowserRedirectPageForCancelAFTerm = gk_sLinkCancelAddTermInformationLeaflet
            
        Case Else
            
            'Send a blank string
            GetBrowserRedirectPageForCancelAFTerm = ""
    
    End Select
    
End Function

Private Function GetTermTypeDescription(t_eTermType As EAFTermType) As String
'   Returns the term type desription.

'   PARAMETER: t_eTermType - > Term type enum
'
'   d14     Aravinda A      09-Dec-1999     Created

    Select Case t_eTermType
        Case EAFTermType.AFPerson
            GetTermTypeDescription = "Person"
        Case EAFTermType.AFPlace
            GetTermTypeDescription = "Place"
        Case EAFTermType.AFSubject
            GetTermTypeDescription = "Subject"
        Case EAFTermType.AFCorporateBody
            GetTermTypeDescription = "Corporate body"
        Case Else
            GetTermTypeDescription = "Unknown"
        End Select
End Function

Private Function GetTermCategoryDescription(t_lTermCategory As Long) As String
'   Returns the term category type desription.

'   PARAMETER: t_lTermType - > Term category type constant
'
'   d14     Aravinda A      09-Dec-1999     Created
'   d21     jcc             07Feb2000     correct spelling mistake

    Select Case t_lTermCategory
        Case klEditSetType
            GetTermCategoryDescription = "Editset"
        Case klleafletType
            GetTermCategoryDescription = "Leaflet"
        Case klInvisibleType
            GetTermCategoryDescription = "Invisible terms"
        Case klIncompleteType
            GetTermCategoryDescription = "Incomplete terms"
        Case klNonCandidateNonPreferredType
            GetTermCategoryDescription = "Non-candidate non-preferred terms"
        Case klNotUsedCandidateType
            GetTermCategoryDescription = "Not used candidate terms"
        Case klNotUsedPreferredType
            GetTermCategoryDescription = "Not used preferred terms"
        Case klLiveType
            GetTermCategoryDescription = "Live terms"
        Case Else
            GetTermCategoryDescription = ""
        End Select
End Function

Public Sub DisplayRejectConfirmationPage()
'Shows confirmation page to reject a preferred term.
'   d15     aga     14Dec1999     Created
    Dim sSource As String
    Dim lTermID As Long
    Dim lTermType As Long
    sSource = "VWEditorialAuthorityFiles.DisplayRejectConfirmationPage"
    On Error GoTo ErrorHandler
    Dim sHTML As String
    
    'Check the permissions
    CheckPermissions ESecurityCheck.EUserCanMaintainAuthorityFiles, sSource, ""
    
    'Get the Term type and ID
    If m_aspRequest(gk_sFORMAFGroup) = "" Then
    
            RaiseError sSource, EPROCATINTERNALERROR.peAuthorityFilesMaintenance, _
                           "Term type is blank"
    End If
    lTermType = m_aspRequest(gk_sFORMAFGroup)
    lTermID = CLng(m_aspRequest(gk_sFORMAFTermID))
    If lTermID <= 0 Then
    
       RaiseError sSource, EPROCATINTERNALERROR.peAuthorityFilesMaintenance, _
                                           "Specific Term ID is not a valid id"
    End If
    'Give the message
    sHTML = sHTML & m_ofwHtmlRef.Font(m_ofwHtmlRef.Bold("Deletion of preferred term deletes all the non-preferred terms linked to it. Do you want to continue?"), EColour.COLOURRed)
    
    sHTML = sHTML & gk_sHTMLLineBreak & gk_sHTMLLineBreak
    
    'Show OK, Cancel buttons
    'Show cancel button
    sHTML = sHTML & m_ofwHtmlRef.ImageButton(gk_sImagesPath & gk_sImageOK, _
            k_sOKBtn, gk_sLinkAFRejectPreferredTerm & "?" & gk_sFORMAFTermID & "=" & lTermID _
            & "&" & gk_sFORMAFGroup & "=" & lTermType) _
            & m_ofwHtmlRef.ImageButton(gk_sImagesPath & gk_sImageCancel, "", _
            gk_sLinkDisplaySpecificIndexTermDetails & "?" & gk_sFORMAFTermID & "=" & lTermID _
            & "&" & gk_sFORMAFGroup & "=" & lTermType & "&Cancel=1", , m_ofwRefData.Search(gToolTip, ETOOLTIP.ttButtonCancelHelp))
    m_aspResponse.Write sHTML
    m_objContext.SetComplete
    Exit Sub
ErrorHandler:
    m_objContext.SetAbort
    ShowError sSource
End Sub

Public Sub DisplayRejectResults()
'Rejects a preferred term and shows the main page or error page
'   d15     aga     14Dec1999     Created
    Dim sSource As String
    Dim lTermID As Long
    Dim lTermType As Long
    Dim lSaveReturnCode As Long
    sSource = "VWEditorialAuthorityFiles.DisplayRejectResults"
    On Error GoTo ErrorHandler
    Dim sHTML As String
    
    'Check the permissions
    CheckPermissions ESecurityCheck.EUserCanMaintainAuthorityFiles, sSource, ""
    
    'Get the Term type and ID
    If m_aspRequest(gk_sFORMAFGroup) = "" Then
    
            RaiseError sSource, EPROCATINTERNALERROR.peAuthorityFilesMaintenance, _
                           "Term type is blank"
    End If
    lTermType = m_aspRequest(gk_sFORMAFGroup)
    lTermID = CLng(m_aspRequest(gk_sFORMAFTermID))
    If lTermID <= 0 Then
    
       RaiseError sSource, EPROCATINTERNALERROR.peAuthorityFilesMaintenance, _
                                           "Specific Term ID is not a valid id"
    End If
    'Save the term with the given ID and term type
    Select Case lTermType
        Case EAFTermType.AFPerson
            Set m_oiAFSaver = New BOAFPerson
            
            m_oBOAFPerson.PROCATTransientInterface_GetState
            m_typPersonData = m_oBOAFPerson.Data
            m_typPersonData.lAuthorityStatus = gk_lRejected
            m_oiAFSaver.Data = m_typPersonData
            lSaveReturnCode = m_oiAFSaver.Save(m_oBOAFPerson.Data.lID)
            If lSaveReturnCode = mk_lRecordNotChanged Then
                
                PROCATEditorialErrorHandling peCanNotDeleteRejectedAFTerm, sSource, _
                           "Cannot delete the AF term. Remove the links and reject it to delete."
                Exit Sub
            Else
                m_aspResponse.Redirect gk_sLinkEditorialAFMaintMain
                
            End If
            m_oBOAFPerson.PROCATTransientInterface_DeleteState
        Case EAFTermType.AFPlace
            Set m_oiAFSaver = New BOAFPlace
            
            m_oBOAFPlace.PROCATTransientInterface_GetState
            m_typPlaceData = m_oBOAFPlace.Data
            m_typPlaceData.lAuthorityStatus = gk_lRejected
            m_oiAFSaver.Data = m_typPlaceData
            lSaveReturnCode = m_oiAFSaver.Save(m_oBOAFPlace.Data.lID)
            If lSaveReturnCode = mk_lRecordNotChanged Then
                
                PROCATEditorialErrorHandling peCanNotDeleteRejectedAFTerm, sSource, _
                           "Cannot delete the AF term. Remove the links and reject it to delete."
                Exit Sub
            Else
                m_aspResponse.Redirect gk_sLinkEditorialAFMaintMain
                
            End If
            m_oBOAFPlace.PROCATTransientInterface_DeleteState
        Case EAFTermType.AFCorporateBody
            Set m_oiAFSaver = New BOAFCorporateBody
            
            m_oBOAFCorporateBody.PROCATTransientInterface_GetState
            m_typCorporateBodyData = m_oBOAFCorporateBody.Data
            m_typCorporateBodyData.lAuthorityStatus = gk_lRejected
            m_oiAFSaver.Data = m_typCorporateBodyData
            lSaveReturnCode = m_oiAFSaver.Save(m_oBOAFCorporateBody.Data.lID)
            If lSaveReturnCode = mk_lRecordNotChanged Then
                
                PROCATEditorialErrorHandling peCanNotDeleteRejectedAFTerm, sSource, _
                           "Cannot delete the AF term. Remove the links and reject it to delete."
                Exit Sub
            Else
                m_aspResponse.Redirect gk_sLinkEditorialAFMaintMain
                
            End If
            m_oBOAFCorporateBody.PROCATTransientInterface_DeleteState
        Case EAFTermType.AFSubject
            Set m_oiAFSaver = New BOAFSubject
            
            m_oBOAFSubject.PROCATTransientInterface_GetState
            m_typSubjectData = m_oBOAFSubject.Data
            m_typSubjectData.lAuthorityStatus = gk_lRejected
            m_oiAFSaver.Data = m_typSubjectData
            lSaveReturnCode = m_oiAFSaver.Save(m_oBOAFSubject.Data.lID)
            If lSaveReturnCode = mk_lRecordNotChanged Then
                
                PROCATEditorialErrorHandling peCanNotDeleteRejectedAFTerm, sSource, _
                           "Cannot delete the AF term. Remove the links and reject it to delete."
                Exit Sub
            Else
                m_aspResponse.Redirect gk_sLinkEditorialAFMaintMain
                
            End If
            m_oBOAFSubject.PROCATTransientInterface_DeleteState
        Case Else
            RaiseError sSource, peAuthorityFilesMaintenance, "Invalid term type"
    End Select
    
    
    m_objContext.SetComplete
    Exit Sub
ErrorHandler:
    m_objContext.SetAbort
    ShowError sSource
End Sub

