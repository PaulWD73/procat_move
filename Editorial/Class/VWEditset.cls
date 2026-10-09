VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "VWEditSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

' File name:    VWEditSet.cls
' System:       PROCAT, PRO
' Copyright:    (C) Quidnunc Limited
'
' Description:  View Edit set class with all display items
'
' Amendment history:
'   d1  ndk     10May1999       Created by Copying from Functional Prototype
'   d2  ndk     14Jun1999       Rewrore existing functions,Added DisplayUserNames, DisplayEditSetsForAllUsers
'   d3  ndk     16Jun1999       Added UpdateEditSetStatus
'   d4  ndk     23Jun1999       Added DisplayUsersAndRolesForEditSet,RemoveRoleOnEditSet
'   d5  ndk     24Jun1999       Added DisplayValidUsersForTheRole, AddUpdateRoleOnEditSet
'   d6  ndk     28Jun1999       Added error handling, reorganised UpdateEditSetStatus
'   d7  ndk     30Jun1999       Added CreateEmptyEditSet, DisplayCreateEmptyEditSetScreen
'   d8  ndk     01Jul1999       Added DisplayEditSetHistory, DisplayEditSetDetails, Using resource strings now
'   d9  ndk     06Jul1999       Changes induced by change due to renaming of business objects
'                               change in name of private variables
'   d10 ndk     07Jul1999       Added DisplayNextStagesForEditSet,DisplayPreviousStagesForEditSet
'                               MoveEditSetToNextStage,MoveEditSetToPreviousStage
'                               replaced page names with constants
'   d11 ksk     08Jul1999       Changed the displayeditsetdetails page to show the
'                               Header editset details and button and the Footer
'                               the catalogue browser
'   d12 ndk     09Jul1999       Added private variables m_lLoggedUserID, m_bLoggedUserIsAdministrator
'                               Removed local variables referring to these from the methods
'   d13 ndk     12Jul1999       Added   DisplayEditsetCatalogueEntries (coded by krish)
'   d14 ndk     12Jul1999       Added DisplayCreateNewCatalogueMenu, CreateNewLetterCode
'                               DisplayCreateNewLetterCodeScreen, CreateNewCatalogueByCopying
'                               DisplayCreateNewCatalogueByCopyingScreen, CreateNewCatalogueFromScratch
'                               DisplayCreateNewCatalogueFromScratchScreen
'   d15     ndk     15Jul1999   Added resource strings, error handling in new functions
'   d16     ndk     15Jul1999   removed references to gk_sHTMLFormEndTag constant
'   d17     ndk     15Jul1999   Added validation for lengths of input fields
'   d18     ndk     16Jul1999   Minor VI Changes after review before demo
'   d19     ndk     16Jul1999   EditSetDetailsHeader now display actual counts of the entries
'   d20     ksk     27Jul1999   Fixing a chnge in the PROCATInternalError Constant
'   d21     ndk     28Jul1999   Changes warranted by
'                               moving of certain methods from boeditset to bouserroles and BOBasicCatalogueEntry and security object
'                               moving of certain methods from vweditset to vwcatalogueentry
'   d22     ksk     02Aug1999   Integrating code in VWEditsetExt
'   d23     ksk     04Aug1999   Changed the global parameter name for the editset ID variable
'   d24     ndk     10Aug1999   Provided hyperlink to move to editsetentry details
'   d25     ndk     11Aug1999   Abandon Button now calls the appropirate asp page,removed updatstatus button
'   d26     ndk     16Aug1999   Checkboxes showing edit set completion status are hyperlinks now
'                               Managing Editor functionality implemented
'                               Moved AddCataloguesToEditSet to VWCatalogueEntry
'   d27     ndk     19Aug1999   Using style constants instead of hardcoded styles for colour, font etc
'                               All the procedures and functions where html was displayed were changed
'   d28     ndk     31Aug1999   Used HelpText style for displaying help texts
'   d29     ndk     08Oct1999   Minor changes (added dummy link for catalogue without reference)
'   d30     ksk     14Oct1999   Integrated code from the new DisplayEditsetDetails method and
'                               the filter form display method
'   d31     ndk     14Oct1999   Changed the link for viewing catalogue details
'   d32     ksk     14Oct1999   Made changes to display filter form
'   d33     ksk     21Oct1999   Add the links to the new images for readonly and imported status
'                               And ignoring the remove button for readonly entries
'   d34     ndk     28Oct1999   Button (on editset details screen) enabling and disabling made dependent on calls
'                               to security/permission object
'   d35     ndk     04Nov1999   Minor change in method DisplayEditsetDetailsHeader
'   d36     ksk     10Nov1999   Changed the functionality for Abandon
'   d37     ndk     11Nov1999   Slight change in interface with security object
'   d37     ksk     12Nov1999   Changed the display of title to use TruncateTitle
'   d39     ndk     17Nov1999   Delete catalogue entry button enabled after doing security check
'   d40     ksk     18Nov1999   Added optional parameters for creating edit set entries
'                               Needed so that it can be integrated with the accession
'                               module
'   d41     ndk     22Nov1999   The filter was displayed twice, now shown once
'   d42     ndk     22Nov1999   Using a different method to get the list of users assignable to a role
'   d43     ndk     23Nov1999   corrected errors in  MoveEditSetToPreviousStage
'   d45     ndk     23Nov1999   Displaying appropirate message if no next/previous stages are available
'                               Disabling 'next stage' button if current stage is pre-release
'   d46     ksk     07Dec1999   Changed the description for the status for the filter
'   d47     ndk     08Dec1999   Added help buttons
'   d48     ndk     10Dec1999   Hyperlink added to Readonly entries to change them to editable entries
'   d49     ndk     16Dec1999   Displaying created on date when displaying editset history
'   d50     ndk     21Dec1999   Added security check where required in the methods
'   d51     ndk     21Dec1999   Showing roles relevant for the user and stage when editing roles
'   d52     ndk     22Dec1999   Slight change in display of text fields
'   d53     ndk     27Dec1999   Changed to VI code use customizable styles
'   d54     ndk     28Dec1999   Added tooltips
'   d55     ndk     28Dec1999   Changed the error raised when length of es name /comments exceed limit
'   d56     ndk     29Dec1999   Rechecking releasing of variables, especially if an error is deliberately raised
'   d57     crm     13Jan2000   Added done button to DisplayEditSetHistory()
'                               modified DisplayCreateEmptyEditSetScreen()
'   d58     aga     14Jan2000   Fixing defects (Added cancel button to cancel adding a role to editset)
'   d59     dtm     02feb2000   corrected typo of 'Catalogye' to 'Catalogue'
'   d60     spb     08Feb2000   Move tooltip and security access out of loop in
'                               DisplayEditsetEntryDetailsContent, for performance.
'                               Change readonly entries to have hyperlink on
'                               add icon at the right, with new tooltip, but
'                               sharing same image as added entry status.
'   d61     dtm     09Feb2000   modified DisplayUsersAndRolesForEditSet so that 'Remove' button only displayed when permissions allow
'                               modified DisplayEditsetEntryDetailsContent so that 'Add read-only' button only hyperlinked if user has permissions
'   d62     jcc     09Feb2000   changed tooltips for arrows on DisplayEditSet
'   d63     jcc     10Feb2000   removed bold on create empty edit set page
'   d64     dtm     10feb2000   Modified DisplayEditsetCatalogueEntries to display filter text for status as well as level
'   d65     tmj     22Feb2000   Changed where a link in DisplayEditsetDetailsHeader points to
'   d66     daj     03Mar2000   Changed DisplayEditsetCatalogueEntries to not show the add button at all...
'   d67     daj     07Mar2000   Changed DisplayEditSetHistory to show parser/accession log links
'   d68     daj     08Mar2000   Fixed display of abandon edit set confirmation to be centred
'   d69     tmj     09Mar2000   Changed so the Editset History displays the first contributor one as the creator
'   d70     daj     21Mar2000   Changed the edit set creation screen to show a context-sensitive title
'   d71     tmj     28Mar2000   Added javascript validation to TextAreas
'   d72     tmj     30Mar2000   Set so titles are now rendered as XML
'   d73     daj     09May2000   Added some encoding/decoding of special characters to query strings
'   d74     wdp     25May2000   Get # entries to display per page from registry
'   d75     daj     17Jul2000   Modified CreateEmptyEditSet and DisplayCreateEmptyEditSetScreen to allow updates;
'                               fixed ordering of recordset position calculation
'   d76     daj     31Oct2000   Fixed problem in DisplayValidUsersForRole where the edit set name could cause problems
'                               if a special character (eg, &, #, ~, ? etc.) was added into the name.
'   d77     daj     12Jan2001   Added 'go to' ability in edit set details screen
'   d78     pnp     17Jan2001   Added 'View' drop down box code to Display filter Form
'   d79     pnp     21Jan2001   Updated DisplayEditsetEntryDetailsHeader to display custom columns
'   d80     pnp     22Jan2001   Updated DisplayEditsetEntryDetailsContent to display custom content
'   d81     gdb     07JFeb2001   Created DisplayEditsetsWithoutEditors method
'   d82     gdb     08Feb2001   Created 'go to' ability - in DisplayEditSetsForOneUser -for managing editor to go to page that
'                                displays all edits sets in editorial stage without editor assigned
'   d83     pnp     08Feb2001   Added functionality for order filter
'   d84     gdb     08Feb2001   Added - in DisplayEditSetsForOneUser - perform security check for managing editor
'   d85     daj     05Mar2001   Added colouring on the edit set details screen
'   d86     map     31May2001   Added function call in DisplayEditsetEntryDetailsContent to format db results correctly
'   d87     alm     19Oct2001   Altered DisplayCreateEmtpyEditSetScreen to support the new
'                               type of editset 'delete'. Altered DisplayEditsetDetailsHeader
'                               to show a new link called 'Move Pieces' and modified function
'                               to display an appropriate header for the new move pieces edit
'                               set usages. Added functions DisplayChoosePiecesToMove,
'                               DisplayChooseParentOfPieces, DisplayChooseParentForPieces and
'                               DisplayEditsetForMultiplePieceMove. Added the parameter t_eUsage
'                               to functions DisplayEditsetEntryDetailsHeading and
'                               DisplayEditsetEntryDetailsContent. Modified
'                               DisplayEditsetEntryDetailsHeading and DisplayEditsetEntryDetailsContent
'                               to support the new move multiple pieces editset states.
'   d86     map     22Oct2001   Added code to deal with new scheduling functionality and automatic edit set load for druid data
'   d87     map     23Oct2001   Added to CreateEditSet to check for a valid date before creating the edit set
'   d88     map     23Oct2001   Changed format of DisplayCreateEditSet screen
'   d89     alm     23Oct2001   Removed references to the sort order in the multiple piece move edit set lists
'   d90     alm     25Oct2001   Major changes were to do with supporting the new editset type - delete
'                               Also some minor changes
'   d91     alm     12Nov2001   Fixed bug where list of pieces to move was curtailed when a piece with an item was hit
'   d92     alm     12Nov2001   Fixed bug where user told that there was no pieces in editset when the current page had no pieces
'   d93     alm     15Nov2001   Fixed bug where series not always remembered. Not elegant. Should be rewritten!!!
'   d94     alm     27Nov2001   Added the time restraint to scheduled job header
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'Interfaces
Implements ObjectControl

'private enums

       
Private Enum EEDITSETSTATUSTYPE
    eNew = 1
    eAccessioned = 2
    eNewAndModified = 3
    eAccessionedAndModified = 4
End Enum


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
Private Const ksNameButtonOk As String = "BtnOK"
Private Const ksNameButtonCancel As String = "BtnCan"
'Private Variables
Private m_objObjectContext As ObjectContext
Private m_aspResponse  As Response
Private m_aspRequest As Request
Private m_oFWHTML As FWHTML
Private m_oFWTable As FWTable
Private m_oFWAppData As FWApplicationData
Private m_ofwRefData  As FWReferenceData
Private m_oBOEditSet As BOEditSet
Private m_oBOEditorialPermissions As BOEditorialPermissions

Private m_lLoggedUserID As Long                     'ID of the logged user

'd30 Added for the filtered edit set list display
Private m_oboFilteredEditset As BOFilteredEditsetList


'Private Constants


'Variables Starting with ksParam are used to pass and read data passed as querystrings
'Variables Starting with ksName are used to pass and read data passed as form elements

Private Const ksParamUserID As String = "USERID"
Private Const ksParamEditSetName As String = "ESNAME"
Private Const ksParamContributor1Name As String = "CONT1NAME"
Private Const ksParamContributor1ForNewEditSetID As String = "CONT1UID"
Private Const ksParamCurrentStageID As String = "ESCURRENTSTAGEID"
Private Const ksParamCurrentStageName As String = "ESCURRENTSTAGENAME"
Private Const ksParamRoleID As String = "ROLEID"
Private Const ksParamStageEditSetIsToBeMoved As String = "STAGETOBEMOVEDTO"
Private Const ksParamAssignedUser As String = "ASSIGNEDUSER"
Private Const ksParamCheckBoxCurrentlyChecked As String = "CHECKBOXCURRENTLYCHECKED"
Private Const ksParamEditSetCreatedDate As String = "EDITSETCREATEDDATE"


Private Const ksNameStatus As String = "CS"

Private Const ksNameEditSetID As String = "ESID"
Private Const ksNameUserID As String = "USERID"
Private Const ksNameContributor1ForNewEditSetID As String = "CONT1NEWEDITSET"   'variable name for cont1 for new editset
Private Const ksNameEditSetName As String = "ESNAME"   'variable name for edit set name
Private Const ksNameEditSetType As String = "EDITSETTYPE"
Private Const ksNameEditSetDescription As String = "EDITSETDESCRIPTION"
Private Const ksNameRoleID As String = "ROLEID" 'variable name for role id
Private Const ksNameEditSetStage As String = "EDITSETSTAGE"
Private Const ksNameStageTransitionComments As String = "STAGETRANSITIONCOMMENTS"

'Field offsets when editsets details are extracted from the recordlist

Private Const KiStep = 5                    'Every 5 coloumns one users details are present
Private Const KiStartForOneUser = 14        'From the 14 th coloumn the status details are present in the fwrecordlist (for one user)
Private Const KiStartForAllUsers = 10       'From the 10 th coloumn the status details are present in the fwrecordlist (for all users)
Private Const KiOffsetUserID As Integer = 0
Private Const KiOffsetInitials As Integer = 1
Private Const KiOffsetUserName As Integer = 2
Private Const KiOffsetStatus As Integer = 3
Private Const KiOffsetRole As Integer = 4

'd30 Added for the filtered edit set list display
Private Const mk_sNameButtonCancel As String = "CAN"
Private Const mk_sNameFilter As String = "FTL"
Private Const mk_sNamePosition As String = "P"
Private Const mk_sNameStatus As String = "CN"
Private Const mk_sNameLevel As String = "CL"
Private Const mk_sNameView As String = "CV"
'd83
Private Const mk_sNameOrder As String = "CO"
Private Const mk_sNameOpenImage As String = "UP"

' d77 - Added
Private Const mk_sNoFilterMessage As String = "No filter set"
Private Const mk_sJumpToReferenceText As String = "Goto Reference"

Public Property Get Name() As String
    Name = "VWEditSet"
End Property




Private Sub ObjectControl_Activate()
    Initialise
End Sub




Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function




Private Sub ObjectControl_Deactivate()
    Terminate
End Sub




Private Sub Terminate()
    Set m_objObjectContext = Nothing
    Set m_aspResponse = Nothing
    Set m_aspRequest = Nothing
    Set m_oFWHTML = Nothing
    Set m_oFWTable = Nothing
    Set m_oFWAppData = Nothing
    Set m_ofwRefData = Nothing
    Set m_oBOEditSet = Nothing
    Set m_oBOEditorialPermissions = Nothing
    Set m_oboFilteredEditset = Nothing
End Sub




Private Sub Initialise()

    On Error GoTo ErrorHandler

    Dim sSource As String
    sSource = Me.Name & ".Initialise"
    
    Set m_objObjectContext = GetObjectContext()
    If m_objObjectContext Is Nothing Then
        RaiseError sSource, peOCUnableToGetObjectContext, "Objectcontext is nothing"
    End If
    
    Set m_aspResponse = m_objObjectContext.Item(gk_sResponseObject)
    Set m_aspRequest = m_objObjectContext.Item(gk_sRequestObject)
    Set m_oFWHTML = New FWHTML
    Set m_oFWTable = New FWTable
    Set m_oFWAppData = New FWApplicationData
    Set m_ofwRefData = New FWReferenceData
    Set m_oBOEditSet = New BOEditSet
    Set m_oBOEditorialPermissions = New BOEditorialPermissions
    
    'd30 Added for the filtered edit set list display
    Set m_oboFilteredEditset = New BOFilteredEditsetList
    
    'Get the users details
    m_lLoggedUserID = GetDataFromSession(dsUserID, m_objObjectContext.Item(gk_sSessionObject))
    If m_lLoggedUserID = 0 Then     'Not a valid user
        RaiseError sSource, EPROCATINTERNALERROR.peDBInvalidUser, "Invalid User ID : " & m_lLoggedUserID
    End If
    
    Exit Sub
    
ErrorHandler:
    m_objObjectContext.SetAbort
    ShowError sSource
End Sub

Public Sub DisplayEditSetsWithoutEditors()
'
'Description
'
'   Displays editset that are in editorial stage without having an editor assigned
'
'   Amendment history
'   d81   GDB  06Feb2001   Created
'

    On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim bCurrentUser   As Boolean   'Whether the editsets are being displayed for current user
    Dim lUserID As Long, sUserName As String, sUserInitials As String
    Dim sAssignedUserInitials As String, lAssignedUserID As Long, iAssignedStatus As Integer
    Dim sCheckBoxLink As String, sCheckBoxImage As String, bCanChangeStatus As Boolean, bCurrentlyChecked As Boolean
    Dim lColoumnsAdded As Long
    Dim oFWRecordlist As FWRecordList
    Dim oBOUser As BOUser
    Dim iColoumnNumber As Integer
    Dim sHTML As String, sParamString   As String
    Dim lMaxNumberOfUsers As Long
    
    Dim typCellStyle As CellStyle
    Dim typCellStyleForDetailLines As CellStyle
    Dim typTableStyle As TableStyle
    Dim typFontStyleForDetailLines As FontStyle
    
    sSource = Me.Name & ".DisplayEditSetsForOneUser"
    
    lUserID = m_aspRequest.QueryString(ksParamUserID)
    
    Set oBOUser = New BOUser
    
    If lUserID = 0 Then        'No user id passed, for the current user
        'Pick up the current users details from the session object
                
        lUserID = GetDataFromSession(dsUserID, m_objObjectContext.Item(gk_sSessionObject))
        sUserName = GetDataFromSession(dsUserName, m_objObjectContext.Item(gk_sSessionObject))
        sUserInitials = GetDataFromSession(dsUserInitials, m_objObjectContext.Item(gk_sSessionObject))
        bCurrentUser = True
    Else    'User ID has been passed as parameter
            'Get the details by running the bo object
        
        oBOUser.Load (lUserID)
        sUserName = oBOUser.UserName
        sUserInitials = oBOUser.UserInitials
        bCurrentUser = False
    End If
    
    m_oFWTable.Reset
   
    typTableStyle = m_oFWHTML.GetTableStyle(TABContentStyle)
    typTableStyle.lCELLSPACING = 0
    
    m_oFWTable.TableStyle = typTableStyle
    
    m_oFWTable.StartRow
        
    'Name of the user who is seeing the editsets
    If bCurrentUser Then
        ShowNormalStyleText _
            sUserName & " (" & sUserInitials & "), " & _
            m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetsInEditorialStageWithoutEditor), _
            "100%", 4
                                
    Else
        ShowNormalStyleText _
            sUserName & " (" & sUserInitials & "), " & _
            m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetsHeader2), _
            "100%", 4
    End If
  
    m_oFWTable.EndRow
    
    m_oFWTable.AddBlankRow
    
    'Colours Indicate
    ShowNormalStyleText _
        m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetsColoursIndicate), "50%", 2
    
    'Click on editset
    typCellStyle = m_oFWHTML.GetCellStyle(CELLHelpTextStyle)
    typCellStyle.sWidth = "50%"
    typCellStyle.lCOLSPAN = 2
    
    m_oFWTable.AddCellAsHTML _
            m_oFWHTML.FontWithStyle( _
            m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetsClickOnEditSet), _
            m_oFWHTML.GetFontStyle(FONTHelpTextStyle)), _
            typCellStyle
            
    m_oFWTable.EndRow
   
    m_oFWTable.StartRow
    
    'Colour Green
    typCellStyle = m_oFWHTML.GetCellStyle(CELLReadyToApproveStyle)    'Reset the settings for the style
    typCellStyle.sWidth = "5%"
    m_oFWTable.AddCellAsHTML gk_sHTMLNonBreakSpace, typCellStyle
    
    'Ready to Approve
    ShowNormalStyleText m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetsReadyToApprove)
    
    typCellStyle = m_oFWHTML.GetCellStyle(CELLHelpTextStyle)
    
    If Not bCurrentUser Then        'Not for the logged user
            
        'Link to Show my editsets
        m_oFWTable.AddCellAsHTML _
            m_oFWHTML.FontWithStyle( _
            m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetsOrShowMe) & " " & _
            m_oFWHTML.Hyperlink(gk_sLinkDisplayEditSetsForOneUser, _
                    m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetsHyperLinkMyEditSets), _
                    gk_sMainFrame, , _
                    GetToolTip(ETOOLTIP.ttShowMyEditSets)), _
            m_oFWHTML.GetFontStyle(FONTHelpTextStyle)), _
            typCellStyle
    Else
        ' Link to show others editsets
        m_oFWTable.AddCellAsHTML _
            m_oFWHTML.FontWithStyle( _
            m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetsOrShowMe) & " " & _
            m_oFWHTML.Hyperlink(gk_sLinkDisplayUserNames, _
                m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetsHyperLinkEditSetsForOtherPeople), _
                gk_sMainFrame, , _
                GetToolTip(ETOOLTIP.ttShowOthersEditSets)), _
            m_oFWHTML.GetFontStyle(FONTHelpTextStyle)), _
            typCellStyle
    End If
    
    
    
    m_oFWTable.StartRow
    
    'Colour Pink
    typCellStyle = m_oFWHTML.GetCellStyle(CELLWorkInProgressStyle)
    typCellStyle.sWidth = "5%"
    m_oFWTable.AddCellAsHTML gk_sHTMLNonBreakSpace, typCellStyle
    
    'Work in Progress
    ShowNormalStyleText m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetsWorkInProgress)

    'Link to show all editets
    typCellStyle = m_oFWHTML.GetCellStyle(CELLHelpTextStyle)
   
    m_oFWTable.AddCellAsHTML _
        m_oFWHTML.FontWithStyle( _
        m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetsOrShowMe) & " " & _
        m_oFWHTML.Hyperlink(gk_sLinkDisplayEditSetsForAllUsers, _
            m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetsHyperLinkAllEditSets), _
            gk_sMainFrame, , _
            GetToolTip(ETOOLTIP.ttShowAllEditSets)), _
        m_oFWHTML.GetFontStyle(FONTHelpTextStyle)), _
        typCellStyle
        
    m_oFWTable.EndRow
    
    m_oFWTable.StartRow
    
    'Colour Grey
    
    typCellStyle = m_oFWHTML.GetCellStyle(CELLNotAtStageStyle)
    typCellStyle.sWidth = "5%"
    m_oFWTable.AddCellAsHTML gk_sHTMLNonBreakSpace, typCellStyle
    
    'Not in users stage
    typCellStyle = m_oFWHTML.GetCellStyle(CELLNormalStyle)
    If bCurrentUser Then
        ShowNormalStyleText m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetsNotAtYourStage)
    Else
        ShowNormalStyleText m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetsNotAtUsersStage)
    End If
    
    'Show me my editsets
    
    typCellStyle = m_oFWHTML.GetCellStyle(CELLHelpTextStyle)
    m_oFWTable.AddCellAsHTML _
        m_oFWHTML.FontWithStyle( _
        m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetsOrShowMe) & " " & _
        m_oFWHTML.Hyperlink(gk_sLinkDisplayEditSetsForOneUser, _
            m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetsHyperLinkMyEditSets), _
            gk_sMainFrame, , _
            GetToolTip(ETOOLTIP.ttShowMyEditSets)), _
        m_oFWHTML.GetFontStyle(FONTHelpTextStyle)), _
        typCellStyle
    
    'm_oFWTable.EndRow
    
    If Not bCurrentUser Then        'This link is not available for logged user
        typCellStyle = m_oFWHTML.GetCellStyle(CELLHelpTextStyle)
        m_oFWTable.AddCellAsHTML _
        m_oFWHTML.FontWithStyle( _
            m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetsOrShowMe) & " " & _
            m_oFWHTML.Hyperlink(gk_sLinkDisplayUserNames, _
                m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetsHyperLinkEditSetsForOtherPeople), _
                gk_sMainFrame, , _
                GetToolTip(ETOOLTIP.ttShowOthersEditSets)), _
            m_oFWHTML.GetFontStyle(FONTHelpTextStyle)), _
            typCellStyle
    End If
  
    m_oFWTable.EndRow
     


    sHTML = sHTML & m_oFWHTML.Paragraph(m_oFWTable.GetHTML)
    
   
    
    'Edit set details start here
   
    m_oFWTable.Reset
    
    typTableStyle = m_oFWHTML.GetTableStyle(TABContentStyle)
    typTableStyle.lCELLSPACING = 2
    
    m_oFWTable.TableStyle = typTableStyle
    
    m_oFWTable.StartRow
    
    'Get the list of editsets for the user
    Set oFWRecordlist = m_oBOEditSet.EditSetListWithoutEditor()

    
    typCellStyle = m_oFWHTML.GetCellStyle(CELLColumnHeadingStyle)
    
    'Coloumn EditSet Name
    ShowColumnHeading m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetsColName)
    
   
    'Coloumn Contributor1 Name
    ShowColumnHeading m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetsColContributor1)

    
    'Coloumn Created On
    ShowColumnHeading m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetsColCreatedOn)
    
    'Coloumn Days in current stage
    ShowColumnHeading m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetsColDaysInCurrentStage)
    
    m_oFWTable.EndRow
    
    'Get the edit set details
   
    If oFWRecordlist.RecordCount = 0 Then       'No recordsets to display
        m_oFWTable.AddBlankRow
    Else
        'Go to the top of recordlist
        oFWRecordlist.MoveFirst
        Do While Not oFWRecordlist.IsAtEnd      'Loop through all the edisets
      
                    typFontStyleForDetailLines = m_oFWHTML.GetFontStyle(FONTReadyToApproveStyle)
                    typCellStyleForDetailLines = m_oFWHTML.GetCellStyle(CELLReadyToApproveStyle)
     
            m_oFWTable.StartRow
            
            'Edit Set id in querystring
            sParamString = "?" & gk_sParamEditSetID & "=" & oFWRecordlist.GetFieldValue(gk_sDBCOLEditSetID)
                
            'EditSet Name
            ShowDetailsLine _
                m_oFWHTML.Hyperlink(gk_sLinkDisplayEditSetDetails & sParamString, oFWRecordlist.GetFieldValue(gk_sDBCOLEditsetName), gk_sMainFrame, , GetToolTip(ETOOLTIP.ttShowEditSetDetails)), typCellStyleForDetailLines, typFontStyleForDetailLines
            
  
            'Contributor1
            ShowDetailsLine oFWRecordlist.GetFieldValue(gk_sDBCOLTempContributor1), typCellStyleForDetailLines, typFontStyleForDetailLines
               
            'Creation Date
            ShowDetailsLine VariantDateAsDDMMMYYYY(oFWRecordlist.GetFieldValue(gk_sDBCOLTempEditSetCreationDate)), typCellStyleForDetailLines, typFontStyleForDetailLines
               
            'Days in current stage
            ShowDetailsLine oFWRecordlist.GetFieldValue(gk_sDBCOLTempDaysInCurrentStage), typCellStyleForDetailLines, typFontStyleForDetailLines
            
            m_oFWTable.EndRow
            
            'Move to the next editset
            oFWRecordlist.MoveNext
        Loop
    End If
    
    'Get the table details
    sHTML = sHTML & m_oFWHTML.Paragraph(m_oFWTable.GetHTML)
    
    'Add the help button
    sHTML = sHTML & HTMLToAddHelpButton(gk_sHelpLinkToEditSetsForOneUser)
    
    'Display the html
    m_aspResponse.Write sHTML
    
    'Release objects
    
    Set oFWRecordlist = Nothing
    Set oBOUser = Nothing
    
    m_objObjectContext.SetComplete
    
    Exit Sub
    
ErrorHandler:
    m_objObjectContext.SetAbort
    ShowError sSource
End Sub

Public Sub DisplayEditSetsForOneUser()
'
'Description
'
'Displays editset for the one user
'If the query string contains a userid that users edit sets are displayed
'If the query string does not contain a user id , the logged users edit sets are displayed
'Amendment history
'   d1   ndk   16Apr1999   Copied from functional Prototype and modified
'   d2   ndk   14Jun1999   Completely rewritten using new framework classes
'   d8   ndk   01Jul1999   Using resource strings now
'   d18  ndk   16Jul1999   Minor VI Changes after review before demo
'   d26  ndk   16Aug1999   Check boxes displayed using images, hyperlinks provided for checkbox
'   d34  ndk   28Oct1999   changed permission check using security object when checking whether to enable checkboxes
'                          Validation is done directly now
'   d47  ndk   08Dec1999   Added help buttons
'   d49  ndk   16Dec1999   passing created on date for editset history
'   d54  ndk   28Dec1999   Added tooltips
'   d69  tmj   09Mar2000   removed passing of contributor1 name
'   d82  gdb   08Feb2001   Created 'go to' ability for managing editor to go to page that
'                          displays all edits sets in editorial stage without editor assigned
'   d84  gdb   08Feb2001   added perform security check for managing editor

    On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim bCurrentUser   As Boolean   'Whether the editsets are being displayed for current user
    Dim lUserID As Long, sUserName As String, sUserInitials As String
    Dim sAssignedUserInitials As String, lAssignedUserID As Long, iAssignedStatus As Integer
    Dim sCheckBoxLink As String, sCheckBoxImage As String, bCanChangeStatus As Boolean, bCurrentlyChecked As Boolean
    Dim lColoumnsAdded As Long
    Dim oFWRecordlist As FWRecordList
    Dim oBOUser As BOUser
    Dim iColoumnNumber As Integer
    Dim sHTML As String, sParamString   As String
    Dim lMaxNumberOfUsers As Long
    
    Dim typCellStyle As CellStyle
    Dim typCellStyleForDetailLines As CellStyle
    Dim typTableStyle As TableStyle
    Dim typFontStyleForDetailLines As FontStyle
    Dim eSecurityCheckStatus As ESecurityCheckResult
    
    sSource = Me.Name & ".DisplayEditSetsForOneUser"
    
    lUserID = m_aspRequest.QueryString(ksParamUserID)
    
    Set oBOUser = New BOUser
    
    If lUserID = 0 Then        'No user id passed, for the current user
        'Pick up the current users details from the session object
                
        lUserID = GetDataFromSession(dsUserID, m_objObjectContext.Item(gk_sSessionObject))
        sUserName = GetDataFromSession(dsUserName, m_objObjectContext.Item(gk_sSessionObject))
        sUserInitials = GetDataFromSession(dsUserInitials, m_objObjectContext.Item(gk_sSessionObject))
        bCurrentUser = True
    Else    'User ID has been passed as parameter
            'Get the details by running the bo object
        
        oBOUser.Load (lUserID)
        sUserName = oBOUser.UserName
        sUserInitials = oBOUser.UserInitials
        bCurrentUser = False
    End If
    
    'perform security check
    m_oBOEditorialPermissions.UserID = lUserID
    'd84 - check if is managing editor
    eSecurityCheckStatus = m_oBOEditorialPermissions.Check(EUserIsManagingEditor)
        
    m_oFWTable.Reset
   
    typTableStyle = m_oFWHTML.GetTableStyle(TABContentStyle)
    typTableStyle.lCELLSPACING = 0
    
    m_oFWTable.TableStyle = typTableStyle
    
    m_oFWTable.StartRow
    
    'Name of the user who is seeing the editsets
    If bCurrentUser Then
        ShowNormalStyleText _
            sUserName & " (" & sUserInitials & "), " & _
            m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetsHeader1), _
            "100%", 4
                                
    Else
        ShowNormalStyleText _
            sUserName & " (" & sUserInitials & "), " & _
            m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetsHeader2), _
            "100%", 4
    End If
  
    m_oFWTable.EndRow
    
    m_oFWTable.AddBlankRow
    
    'Colours Indicate
    ShowNormalStyleText _
        m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetsColoursIndicate), "50%", 2
    
    'Click on editset
    typCellStyle = m_oFWHTML.GetCellStyle(CELLHelpTextStyle)
    typCellStyle.sWidth = "50%"
    typCellStyle.lCOLSPAN = 2
    
    m_oFWTable.AddCellAsHTML _
            m_oFWHTML.FontWithStyle( _
            m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetsClickOnEditSet), _
            m_oFWHTML.GetFontStyle(FONTHelpTextStyle)), _
            typCellStyle
            
    m_oFWTable.EndRow
   
    m_oFWTable.StartRow
    
    'Colour Green
    typCellStyle = m_oFWHTML.GetCellStyle(CELLReadyToApproveStyle)    'Reset the settings for the style
    typCellStyle.sWidth = "5%"
    m_oFWTable.AddCellAsHTML gk_sHTMLNonBreakSpace, typCellStyle
    
    'Ready to Approve
    ShowNormalStyleText m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetsReadyToApprove)
    
    typCellStyle = m_oFWHTML.GetCellStyle(CELLHelpTextStyle)
    
    If Not bCurrentUser Then        'Not for the logged user
            
        'Link to Show my editsets
        m_oFWTable.AddCellAsHTML _
            m_oFWHTML.FontWithStyle( _
            m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetsOrShowMe) & " " & _
            m_oFWHTML.Hyperlink(gk_sLinkDisplayEditSetsForOneUser, _
                    m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetsHyperLinkMyEditSets), _
                    gk_sMainFrame, , _
                    GetToolTip(ETOOLTIP.ttShowMyEditSets)), _
            m_oFWHTML.GetFontStyle(FONTHelpTextStyle)), _
            typCellStyle
    Else
        ' Link to show others editsets
        m_oFWTable.AddCellAsHTML _
            m_oFWHTML.FontWithStyle( _
            m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetsOrShowMe) & " " & _
            m_oFWHTML.Hyperlink(gk_sLinkDisplayUserNames, _
                m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetsHyperLinkEditSetsForOtherPeople), _
                gk_sMainFrame, , _
                GetToolTip(ETOOLTIP.ttShowOthersEditSets)), _
            m_oFWHTML.GetFontStyle(FONTHelpTextStyle)), _
            typCellStyle
    End If
    

    
    m_oFWTable.StartRow
    
    'Colour Pink
    typCellStyle = m_oFWHTML.GetCellStyle(CELLWorkInProgressStyle)
    typCellStyle.sWidth = "5%"
    m_oFWTable.AddCellAsHTML gk_sHTMLNonBreakSpace, typCellStyle
    
    'Work in Progress
    ShowNormalStyleText m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetsWorkInProgress)

    'Link to show all editets
    typCellStyle = m_oFWHTML.GetCellStyle(CELLHelpTextStyle)
   
    m_oFWTable.AddCellAsHTML _
        m_oFWHTML.FontWithStyle( _
        m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetsOrShowMe) & " " & _
        m_oFWHTML.Hyperlink(gk_sLinkDisplayEditSetsForAllUsers, _
            m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetsHyperLinkAllEditSets), _
            gk_sMainFrame, , _
            GetToolTip(ETOOLTIP.ttShowAllEditSets)), _
        m_oFWHTML.GetFontStyle(FONTHelpTextStyle)), _
        typCellStyle
        
    m_oFWTable.EndRow
    
    m_oFWTable.StartRow
    
    'Colour Grey
    
    typCellStyle = m_oFWHTML.GetCellStyle(CELLNotAtStageStyle)
    typCellStyle.sWidth = "5%"
    m_oFWTable.AddCellAsHTML gk_sHTMLNonBreakSpace, typCellStyle
    
    'Not in users stage
    typCellStyle = m_oFWHTML.GetCellStyle(CELLNormalStyle)
    If bCurrentUser Then
        ShowNormalStyleText m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetsNotAtYourStage)
    Else
        ShowNormalStyleText m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetsNotAtUsersStage)
    End If
    
    If Not bCurrentUser Then        'This link is not available for logged user
        typCellStyle = m_oFWHTML.GetCellStyle(CELLHelpTextStyle)
        m_oFWTable.AddCellAsHTML _
        m_oFWHTML.FontWithStyle( _
            m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetsOrShowMe) & " " & _
            m_oFWHTML.Hyperlink(gk_sLinkDisplayUserNames, _
                m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetsHyperLinkEditSetsForOtherPeople), _
                gk_sMainFrame, , _
                GetToolTip(ETOOLTIP.ttShowOthersEditSets)), _
            m_oFWHTML.GetFontStyle(FONTHelpTextStyle)), _
            typCellStyle
    End If
  ' check if is the current user
  If bCurrentUser Then
      'd82 - if user is a managing editor , display this link to all edit sets without editor
    If eSecurityCheckStatus = ESecurityCheckResult.ESRSecurityCheckSuccessful Then
        m_oFWTable.AddCellAsHTML _
            m_oFWHTML.FontWithStyle( _
            m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetsOrShowMe) & " " & _
            m_oFWHTML.Hyperlink(gk_sLinkDisplayEditorialSetsWithoutEditor, _
                m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetsHyperLinkEditorialSetsWithoutEditors), _
                gk_sMainFrame, , _
                GetToolTip(ETOOLTIP.ttShowOthersEditSets)), _
            m_oFWHTML.GetFontStyle(FONTHelpTextStyle)), _
            typCellStyle
       Else
    End If
  End If
  
  
    m_oFWTable.EndRow
    
    sHTML = sHTML & m_oFWHTML.Paragraph(m_oFWTable.GetHTML)
    
    'Edit set details start here
   
    m_oFWTable.Reset
    
    typTableStyle = m_oFWHTML.GetTableStyle(TABContentStyle)
    typTableStyle.lCELLSPACING = 2
    
    m_oFWTable.TableStyle = typTableStyle
    
    m_oFWTable.StartRow
    
    'Get the list of editsets for the user
    Set oFWRecordlist = m_oBOEditSet.EditSetListForUser(lUserID)

    'Max number of users on any editset for its current stage
    If Not oFWRecordlist.RecordCount = 0 Then
        oFWRecordlist.MoveFirst
        lMaxNumberOfUsers = oFWRecordlist.GetFieldValue(gk_sDBCOLTempMaxUsersForAllEditSets)
    End If
    If lMaxNumberOfUsers = 0 Then
        lMaxNumberOfUsers = 1
    End If
    
    typCellStyle = m_oFWHTML.GetCellStyle(CELLColumnHeadingStyle)
    
    'Coloumn EditSet Name
    ShowColumnHeading m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetsColName)
    
    'Coloumn EditSetType
    ShowColumnHeading m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetsColType)
    
    'Coloumn Contributor1 Name
    ShowColumnHeading m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetsColContributor1)
    
    'Coloumn Role Name
    If bCurrentUser Then
        ShowColumnHeading m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetsColYourRole)
    Else
        ShowColumnHeading m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetsColUsersRole)
    End If
    
    'Coloumn Created On
    ShowColumnHeading m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetsColCreatedOn)
    
    'Coloumn Current Stage
    ShowColumnHeading m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetsColWorkStage)
    
    'Coloumn Days in current stage
    ShowColumnHeading m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetsColDaysInCurrentStage)
    
    'Coloumn Completion Status
    ShowColumnHeading m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetsColWorkStatus), _
        lMaxNumberOfUsers
    
    m_oFWTable.EndRow
    
    'Get the edit set details
   
    If oFWRecordlist.RecordCount = 0 Then       'No recordsets to display
        m_oFWTable.AddBlankRow
    Else
        'Go to the top of recordlist
        oFWRecordlist.MoveFirst
        Do While Not oFWRecordlist.IsAtEnd      'Loop through all the edisets
            If oFWRecordlist.GetFieldValue(gk_sDBCOLTempIsInYourStage) = 0 Then     'Not at users stage
                
                typFontStyleForDetailLines = m_oFWHTML.GetFontStyle(FontNotAtStageStyle)
                typCellStyleForDetailLines = m_oFWHTML.GetCellStyle(CELLNotAtStageStyle)
                
            Else
                If oFWRecordlist.GetFieldValue(gk_sDBCOLComplete) = 1 Then      'Complete  (all users status completed)
                    
                    typFontStyleForDetailLines = m_oFWHTML.GetFontStyle(FONTReadyToApproveStyle)
                    typCellStyleForDetailLines = m_oFWHTML.GetCellStyle(CELLReadyToApproveStyle)
                Else
                    
                    typFontStyleForDetailLines = m_oFWHTML.GetFontStyle(FONTWorkInProgressStyle)
                    typCellStyleForDetailLines = m_oFWHTML.GetCellStyle(CELLWorkInProgressStyle)

                End If
            End If
        
            m_oFWTable.StartRow
            
            'Edit Set id in querystring
            sParamString = "?" & gk_sParamEditSetID & "=" & oFWRecordlist.GetFieldValue(gk_sDBCOLEditSetID)
                
            'EditSet Name
            ShowDetailsLine _
                m_oFWHTML.Hyperlink(gk_sLinkDisplayEditSetDetails & sParamString, oFWRecordlist.GetFieldValue(gk_sDBCOLEditsetName), gk_sMainFrame, , GetToolTip(ETOOLTIP.ttShowEditSetDetails)), typCellStyleForDetailLines, typFontStyleForDetailLines
            
            'Edit set type
            ShowDetailsLine _
                oFWRecordlist.GetFieldValue(gk_sDBCOLEditSetTypeDescription), typCellStyleForDetailLines, typFontStyleForDetailLines
            
            'Contributor1
            ShowDetailsLine oFWRecordlist.GetFieldValue(gk_sDBCOLTempContributor1), typCellStyleForDetailLines, typFontStyleForDetailLines
            
            'Users Role Name
            ShowDetailsLine oFWRecordlist.GetFieldValue(gk_sDBCOLTempYourRoleName), typCellStyleForDetailLines, typFontStyleForDetailLines
            
            'Creation Date
            ShowDetailsLine VariantDateAsDDMMMYYYY(oFWRecordlist.GetFieldValue(gk_sDBCOLTempEditSetCreationDate)), typCellStyleForDetailLines, typFontStyleForDetailLines
            
            'Current Stage
            ShowDetailsLine oFWRecordlist.GetFieldValue(gk_sDBCOLEditsetStageDescription), typCellStyleForDetailLines, typFontStyleForDetailLines
            
            'Days in current stage
            ShowDetailsLine oFWRecordlist.GetFieldValue(gk_sDBCOLTempDaysInCurrentStage), typCellStyleForDetailLines, typFontStyleForDetailLines
            
            lColoumnsAdded = 0
            
            'Do display the completion status
            For iColoumnNumber = KiStartForOneUser To oFWRecordlist.FieldCount Step KiStep
                If IsNull(oFWRecordlist.GetFieldValue(iColoumnNumber)) Then
                    Exit For
                End If
                
                sAssignedUserInitials = oFWRecordlist.GetFieldValue(iColoumnNumber + KiOffsetInitials)
                lAssignedUserID = oFWRecordlist.GetFieldValue(iColoumnNumber + KiOffsetUserID)
                iAssignedStatus = oFWRecordlist.GetFieldValue(iColoumnNumber + KiOffsetStatus)
                
                lColoumnsAdded = lColoumnsAdded + 1
                
                bCurrentlyChecked = IIf(iAssignedStatus = 1, True, False)
                
                bCanChangeStatus = False    'Initial value
                
                If bCurrentUser And _
                    Not oFWRecordlist.GetFieldValue(gk_sDBCOLTempIsInYourStage) = 0 Then    'Only when seeing his own editsets and in his stage
                    
                    If lAssignedUserID = m_lLoggedUserID Then   'His own editset
                        bCanChangeStatus = True
                    End If
                End If
                
                sParamString = _
                    "?" & gk_sParamEditSetID & "=" & oFWRecordlist.GetFieldValue(gk_sDBCOLEditSetID) & _
                    "&" & ksParamAssignedUser & "=" & lAssignedUserID & _
                    "&" & ksParamCheckBoxCurrentlyChecked & "=" & bCurrentlyChecked
                
                If bCanChangeStatus Then        'Show enabled checkboxes
                    If bCurrentlyChecked Then
                        sCheckBoxImage = gk_sImagesPath & gk_sImageCheckBoxCheckedEnabled
                        sCheckBoxLink = _
                            gk_sLinkUpdateEditSetStatus & sParamString
                    Else
                        sCheckBoxImage = gk_sImagesPath & gk_sImageCheckBoxUnCheckedEnabled
                        sCheckBoxLink = _
                            gk_sLinkUpdateEditSetStatus & sParamString
                    End If
                Else
                    If bCurrentlyChecked Then   'Show disabled checkboxes
                        sCheckBoxImage = gk_sImagesPath & gk_sImageCheckBoxCheckedDisabled
                    Else
                        sCheckBoxImage = gk_sImagesPath & gk_sImageCheckBoxUnCheckedDisabled
                    End If
                    
                    sCheckBoxLink = ""       'No link

                End If
                
                'Show check box
                ShowDetailsLine _
                    m_oFWHTML.ImageButton(sCheckBoxImage, "", sCheckBoxLink, gk_sMainFrame) & sAssignedUserInitials, _
                    typCellStyleForDetailLines, typFontStyleForDetailLines
                
            Next
            
            'Add blank coloumns
            Do While lColoumnsAdded < lMaxNumberOfUsers
                lColoumnsAdded = lColoumnsAdded + 1
                
                ShowDetailsLine _
                    gk_sHTMLNonBreakSpace, typCellStyleForDetailLines, typFontStyleForDetailLines
            Loop
                
            typCellStyle = m_oFWHTML.GetCellStyle(CELLNormalStyle)
           
            'To display editset history
            'd69 removed passing of contributor1 name
            ShowDetailsLine _
                m_oFWHTML.Hyperlink(gk_sLinkDisplayEditSetHistory & _
                "?" & gk_sParamEditSetID & "=" & oFWRecordlist.GetFieldValue(gk_sDBCOLEditSetID) & _
                "&" & ksParamEditSetName & "=" & oFWRecordlist.GetFieldValue(gk_sDBCOLEditsetName) & _
                "&" & ksParamEditSetCreatedDate & "=" & VariantDateAsDDMMMYYYY(oFWRecordlist.GetFieldValue(gk_sDBCOLTempEditSetCreationDate)), _
                m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetsHyperLinkHistory), gk_sMainFrame, , GetToolTip(ETOOLTIP.ttShowEditSetHistory)), _
                typCellStyle, m_oFWHTML.GetFontStyle(FONTNormalStyle)
            
            m_oFWTable.EndRow
            
            'Move to the next editset
            oFWRecordlist.MoveNext
        Loop
    End If
        
    'Get the table details
    sHTML = sHTML & m_oFWHTML.Paragraph(m_oFWTable.GetHTML)
    
    'Add the help button
    sHTML = sHTML & HTMLToAddHelpButton(gk_sHelpLinkToEditSetsForOneUser)
    
    'Display the html
    m_aspResponse.Write sHTML
    
    'Release objects
    
    Set oFWRecordlist = Nothing
    Set oBOUser = Nothing
    
    m_objObjectContext.SetComplete
    
    Exit Sub
    
ErrorHandler:
    m_objObjectContext.SetAbort
    ShowError sSource
End Sub



Public Sub DisplayEditSetsForAllUsers()
'
'Displays all the editsets
'
''Amendment history
'   d2      ndk     14Jun1999   Created
'   d8      ndk     01Jul1999   Using resource strings now
'   d26     ndk     16Aug1999   Check boxes displayed using images, hyperlinks provided for checkbox
'   d47     ndk     08Dec1999   Added help buttons
'   d49     ndk     16Dec1999   passing created on date for editset history
'   d54     ndk     28Dec1999   Added tooltips
'   d69     tmj     09Mar2000   removed passing of contributor1 name
'

    On Error GoTo ErrorHandler
    
    Dim lUserID As Long, sUserName As String, sUserInitials As String
    Dim sSource As String
    Dim oFWRecordlist As FWRecordList
    Dim iColoumnNumber As Integer
    Dim sHTML As String, sParamString   As String, sCheckBoxImage As String
    Dim lMaxNumberOfUsers As Long
    Dim lColoumnsAdded As Long
    
    Dim typCellStyle As CellStyle
    Dim typCellStyleForDetailLines As CellStyle
    Dim typTableStyle As TableStyle
    Dim typFontStyleForDetailLines As FontStyle
    
    sSource = Me.Name & ".DisplayEditSetsForAllUsers()"

    m_oFWTable.Reset
    
    typTableStyle = m_oFWHTML.GetTableStyle(TABContentStyle)
    typTableStyle.lCELLSPACING = 2
    
    m_oFWTable.TableStyle = typTableStyle

    m_oFWTable.StartRow
    
    ' Current Edit Sets message
    typCellStyle = m_oFWHTML.GetCellStyle(CELLPageHeadingStyle)
    typCellStyle.sWidth = "50%"
    typCellStyle.lCOLSPAN = 2
    
    m_oFWTable.AddCellAsHTML _
        m_oFWHTML.FontWithStyle( _
        m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetsCurrentEditSets), _
        m_oFWHTML.GetFontStyle(FONTPageHeadingStyle)), _
        typCellStyle
    
    m_oFWTable.EndRow
    
    m_oFWTable.AddBlankRow
    
    'Colours Indicate
    ShowNormalStyleText _
        m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetsColoursIndicate), "50%", 2
        
    'click on editset
    typCellStyle = m_oFWHTML.GetCellStyle(CELLHelpTextStyle)
    typCellStyle.sWidth = "50%"
    typCellStyle.lCOLSPAN = 2
    
    m_oFWTable.AddCellAsHTML _
        m_oFWHTML.FontWithStyle( _
        m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetsClickOnEditSet), _
        m_oFWHTML.GetFontStyle(FONTHelpTextStyle)), _
        typCellStyle
    
    m_oFWTable.EndRow
   
   'Colour Green
    typCellStyle = m_oFWHTML.GetCellStyle(CELLReadyToApproveStyle)    'Reset the settings for the style
    typCellStyle.sWidth = "5%"
    m_oFWTable.AddCellAsHTML gk_sHTMLNonBreakSpace, typCellStyle
    
    'Ready to Approve
    ShowNormalStyleText m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetsReadyToApprove)
    
    'Show me editsets of other people
    typCellStyle = m_oFWHTML.GetCellStyle(CELLHelpTextStyle)
    
    m_oFWTable.AddCellAsHTML _
        m_oFWHTML.FontWithStyle( _
        m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetsOrShowMe) & " " & _
        m_oFWHTML.Hyperlink(gk_sLinkDisplayUserNames, _
                m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetsHyperLinkEditSetsForOtherPeople), _
                gk_sMainFrame, , _
                GetToolTip(ETOOLTIP.ttShowOthersEditSets)), _
        m_oFWHTML.GetFontStyle(FONTHelpTextStyle)), _
        typCellStyle
        
    m_oFWTable.EndRow
    
    m_oFWTable.StartRow
       
    'Colour Pink
    typCellStyle = m_oFWHTML.GetCellStyle(CELLWorkInProgressStyle)
    typCellStyle.sWidth = "5%"
    m_oFWTable.AddCellAsHTML gk_sHTMLNonBreakSpace, typCellStyle
    
    'Work in progress
    ShowNormalStyleText m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetsWorkInProgress)
    
    'Show me my editsets
    
    typCellStyle = m_oFWHTML.GetCellStyle(CELLHelpTextStyle)
    m_oFWTable.AddCellAsHTML _
        m_oFWHTML.FontWithStyle( _
        m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetsOrShowMe) & " " & _
        m_oFWHTML.Hyperlink(gk_sLinkDisplayEditSetsForOneUser, _
            m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetsHyperLinkMyEditSets), _
            gk_sMainFrame, , _
            GetToolTip(ETOOLTIP.ttShowMyEditSets)), _
        m_oFWHTML.GetFontStyle(FONTHelpTextStyle)), _
        typCellStyle
    
    m_oFWTable.EndRow

    sHTML = sHTML & m_oFWHTML.Paragraph(m_oFWTable.GetHTML)
    
    'Edit set details start here
    
    m_oFWTable.Reset
    
    typTableStyle.sWidth = "100%"
    typTableStyle.lCELLSPACING = 2
    m_oFWTable.TableStyle = typTableStyle
    
    m_oFWTable.StartRow
    
    'Get the list of all editsets
    Set oFWRecordlist = m_oBOEditSet.EditSetListForAllUsers

    'Get the max users for any editset for current stage
    If Not oFWRecordlist.RecordCount = 0 Then
        oFWRecordlist.MoveFirst
        lMaxNumberOfUsers = oFWRecordlist.GetFieldValue(gk_sDBCOLTempMaxUsersForAllEditSets)
    End If
    If lMaxNumberOfUsers = 0 Then
        lMaxNumberOfUsers = 1
    End If
    
    'Coloumn EditSet Name
    ShowColumnHeading m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetsColName)
    
    'Coloumn Contriburtor1
    ShowColumnHeading m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetsColContributor1)
        
    'Coloumn Created On
    ShowColumnHeading m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetsColCreatedOn)
    
    'Coloumn Current Stage
    ShowColumnHeading m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetsColWorkStage)
    
    'Coloumn Days in current stage
    ShowColumnHeading m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetsColDaysInCurrentStage)
    
    'Coloumn Workstatus
    ShowColumnHeading _
        m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetsColWorkStatus), _
        lMaxNumberOfUsers
    
    m_oFWTable.EndRow
    
    'Get the edit set details

    If Not oFWRecordlist.RecordCount = 0 Then
        oFWRecordlist.MoveFirst
        Do While Not oFWRecordlist.IsAtEnd
            If oFWRecordlist.GetFieldValue(gk_sDBCOLComplete) = 1 Then  'All editsets status are complete
                typFontStyleForDetailLines = m_oFWHTML.GetFontStyle(FONTReadyToApproveStyle)
                typCellStyleForDetailLines = m_oFWHTML.GetCellStyle(CELLReadyToApproveStyle)
            Else
                typFontStyleForDetailLines = m_oFWHTML.GetFontStyle(FONTWorkInProgressStyle)
                typCellStyleForDetailLines = m_oFWHTML.GetCellStyle(CELLWorkInProgressStyle)
            End If
    
            m_oFWTable.StartRow
    
            'Edit Set id in querystring
            sParamString = "?" & gk_sParamEditSetID & "=" & oFWRecordlist.GetFieldValue(gk_sDBCOLEditSetID)
                
            'EditSet Name
            ShowDetailsLine _
                m_oFWHTML.Hyperlink(gk_sLinkDisplayEditSetDetails & sParamString, _
                    oFWRecordlist.GetFieldValue(gk_sDBCOLEditsetName), _
                    gk_sMainFrame, , _
                    GetToolTip(ETOOLTIP.ttShowEditSetDetails)), typCellStyleForDetailLines, typFontStyleForDetailLines
            
            'Contributor1
            ShowDetailsLine _
                oFWRecordlist.GetFieldValue(gk_sDBCOLTempContributor1), typCellStyleForDetailLines, typFontStyleForDetailLines
            
            'Creation Date
            ShowDetailsLine _
                VariantDateAsDDMMMYYYY(oFWRecordlist.GetFieldValue(gk_sDBCOLTempEditSetCreationDate)), typCellStyleForDetailLines, typFontStyleForDetailLines
            
            'Current Stage Description
            ShowDetailsLine _
                oFWRecordlist.GetFieldValue(gk_sDBCOLEditsetStageDescription), typCellStyleForDetailLines, typFontStyleForDetailLines
            
            'Days in Current Stage
            ShowDetailsLine _
                oFWRecordlist.GetFieldValue(gk_sDBCOLTempDaysInCurrentStage), typCellStyleForDetailLines, typFontStyleForDetailLines
            
            
            'To display the status
            lColoumnsAdded = 0
            
            For iColoumnNumber = KiStartForAllUsers To oFWRecordlist.FieldCount Step KiStep
                If IsNull(oFWRecordlist.GetFieldValue(iColoumnNumber)) Then
                    Exit For
                End If
                
                lColoumnsAdded = lColoumnsAdded + 1
                
                If oFWRecordlist.GetFieldValue(iColoumnNumber + KiOffsetStatus) = 1 Then   'Checked
                    sCheckBoxImage = gk_sImagesPath & gk_sImageCheckBoxCheckedDisabled
                Else
                    sCheckBoxImage = gk_sImagesPath & gk_sImageCheckBoxUnCheckedDisabled
                End If

                'Show check box
                ShowDetailsLine _
                    m_oFWHTML.ImageButton( _
                        sCheckBoxImage, "", "", gk_sMainFrame) & oFWRecordlist.GetFieldValue(iColoumnNumber + KiOffsetInitials), _
                        typCellStyleForDetailLines, typFontStyleForDetailLines
            Next
            
            'Add blank coloumns
            Do While lColoumnsAdded < lMaxNumberOfUsers
                lColoumnsAdded = lColoumnsAdded + 1
                ShowDetailsLine _
                        gk_sHTMLNonBreakSpace, _
                        typCellStyleForDetailLines, typFontStyleForDetailLines
                
            Loop
                
            typCellStyle = m_oFWHTML.GetCellStyle(CELLNormalStyle)
            
            'To display editset history
            'd69 removed passing of contributor1 name
            ShowDetailsLine _
                m_oFWHTML.Hyperlink(gk_sLinkDisplayEditSetHistory & _
                "?" & gk_sParamEditSetID & "=" & oFWRecordlist.GetFieldValue(gk_sDBCOLEditSetID) & _
                "&" & ksParamEditSetName & "=" & oFWRecordlist.GetFieldValue(gk_sDBCOLEditsetName) & _
                "&" & ksParamEditSetCreatedDate & "=" & VariantDateAsDDMMMYYYY(oFWRecordlist.GetFieldValue(gk_sDBCOLTempEditSetCreationDate)), _
                m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetsHyperLinkHistory), gk_sMainFrame, , GetToolTip(ETOOLTIP.ttShowEditSetHistory)), _
                typCellStyle, m_oFWHTML.GetFontStyle(FONTNormalStyle)
            
            m_oFWTable.EndRow
            
            'Move to the next editset
            oFWRecordlist.MoveNext
        Loop
        
    End If
            
    'Get the table details
    
    sHTML = sHTML & m_oFWHTML.Paragraph(m_oFWTable.GetHTML)
    
    'Add the help button
    sHTML = sHTML & HTMLToAddHelpButton(gk_sHelpLinkToEditSetsForAllUsers)
    
    'Write the html
    m_aspResponse.Write sHTML
    
    'Release the objects
    
    Set oFWRecordlist = Nothing
    
    m_objObjectContext.SetComplete
    
    Exit Sub
ErrorHandler:
    m_objObjectContext.SetAbort
    ShowError sSource
End Sub





Public Sub DisplayUserNames()
'Description :
'
'This method is meant to display the names of other users
'A list of users is displayed and after the user picks up a name from the list
'his edit sets are displayed
'
'Amendment history
'
'   d2      ndk     14Jun1999       Created
'   d8      ndk     01Jul1999       Using resource strings now
'   d47     ndk     08Dec1999       Added help buttons

On Error GoTo ErrorHandler


    Dim sSource As String
    Dim sUserName As String   'To store the current users name
    Dim lUserID As Long
    Dim iCounter As Integer
    Dim sIndex As String
    
    Dim sFirstCharOfUserName As String
    Dim sOldFirstChar As String
    Dim colNames As New Collection
    Dim colItem As Boolean
    
    Dim sHTML As String, sParamString   As String
    
    Dim oFWRecordlist As FWRecordList
    Dim oBOUser As BOUser
    
    Dim typCellStyle As CellStyle
    Dim typTableStyle As TableStyle
    
    sSource = Me.Name & ".DisplayUserNames()"
    
    Set oBOUser = New BOUser
        
    'Get the list of all users of the system
    Set oFWRecordlist = oBOUser.AllUserNames
        
    If oFWRecordlist.RecordCount = 0 Then
        Exit Sub
    End If
    
    'Loop through all the names and make a list of characters which have
    'some name starting with them
    
    'The names are sorted in ascending order
    For iCounter = 0 To 25
        colNames.Add False, Chr(iCounter + 65)
    Next
    
    oFWRecordlist.MoveFirst
    sOldFirstChar = ""
    
    Do While Not oFWRecordlist.IsAtEnd
        If Not oFWRecordlist.GetFieldValue(gk_sDBCOLUserID) = m_lLoggedUserID Then       'Dont display current users name
            sFirstCharOfUserName = UCase(Left(oFWRecordlist.GetFieldValue(gk_sDBCOLUserName), 1))    'Extract the first char of the name
            If sFirstCharOfUserName <> sOldFirstChar Then
                colNames.Remove (sFirstCharOfUserName)
                colNames.Add True, sFirstCharOfUserName
                sOldFirstChar = sFirstCharOfUserName
            End If
            iCounter = iCounter + 1
        End If
        oFWRecordlist.MoveNext
    Loop
    
    sIndex = ""
    For iCounter = 65 To 65 + 25
        colItem = colNames.Item(Chr(iCounter))
        If colItem = True Then       'There is a name starting with this character
            sIndex = sIndex & m_oFWHTML.Hyperlink(CStr("#" & Chr(iCounter)), CStr(Chr(iCounter)), gk_sMainFrame)
        Else
            sIndex = sIndex & Chr(iCounter)
        End If
        sIndex = sIndex & " "
    Next
    
    typTableStyle = m_oFWHTML.GetTableStyle(TABPageHeadingStyle)
    typTableStyle.lCELLSPACING = 2
    
    m_oFWTable.Reset
    m_oFWTable.TableStyle = typTableStyle
    
    m_oFWTable.StartRow
    
    ShowNormalStyleText _
        m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetsWhoseEditSetToShow), "50%", 2
    
    m_oFWTable.EndRow
    
    m_oFWTable.AddBlankRow
    
    m_oFWTable.StartRow
    
    ShowNormalStyleText _
        sIndex, "50%", 2
    
    m_oFWTable.EndRow
    
    sHTML = sHTML & m_oFWTable.GetHTML
    
    m_oFWTable.Reset
    
    typTableStyle = m_oFWHTML.GetTableStyle(TABContentStyle)
    typTableStyle.lCELLSPACING = 0
    typTableStyle.lCELLPADDING = 0
    
    m_oFWTable.TableStyle = typTableStyle
    
    oFWRecordlist.MoveFirst
    sOldFirstChar = ""  'So that the first name is displayed as using "href name"
    
    Do While Not oFWRecordlist.IsAtEnd
        If Not oFWRecordlist.GetFieldValue(gk_sDBCOLUserID) = m_lLoggedUserID Then      'Dont display current users name
            sFirstCharOfUserName = UCase(Left(oFWRecordlist.GetFieldValue(gk_sDBCOLUserName), 1))    'Extract the first char of the name
            sUserName = oFWRecordlist.GetFieldValue(gk_sDBCOLUserName)
            lUserID = oFWRecordlist.GetFieldValue(gk_sDBCOLUserID)
            If sFirstCharOfUserName <> sOldFirstChar Then
                m_oFWTable.StartRow
                
                ShowNormalStyleText _
                    m_oFWHTML.Hyperlink(gk_sLinkDisplayEditSetsForOneUser & "?" & ksParamUserID & "=" & lUserID, sUserName, gk_sMainFrame, sFirstCharOfUserName)
                
                m_oFWTable.EndRow
            
                sOldFirstChar = sFirstCharOfUserName
            Else
                m_oFWTable.StartRow
                ShowNormalStyleText _
                    m_oFWHTML.Hyperlink(gk_sLinkDisplayEditSetsForOneUser & "?" & ksParamUserID & "=" & lUserID, sUserName, gk_sMainFrame)
                m_oFWTable.EndRow
            End If
        End If
            oFWRecordlist.MoveNext
    Loop
    
    sHTML = sHTML & m_oFWHTML.Paragraph(m_oFWTable.GetHTML)
    
    'Add the help button
    sHTML = sHTML & HTMLToAddHelpButton(gk_sHelpLinkToUserNames)
    
    m_aspResponse.Write sHTML
    
    Set oFWRecordlist = Nothing
    Set oBOUser = Nothing
    
    m_objObjectContext.SetComplete
    
    Exit Sub
ErrorHandler:
    m_objObjectContext.SetAbort
    ShowError sSource
End Sub
    




Public Sub UpdateEditSetStatus()
'
'   Update the database with the changes made to the completion status of
'   the edit set
'
'   d3  ndk     16Jun1999       Created
'   d6  ndk     28Jun1999       Re-organised and moved some functionality to business object
'   d21 ndk     28Jul1999       Only updates one editset and one user at time

On Error GoTo ErrorHandler

    Dim sSource As String
    Dim lEditSetID As Long, lAssignedUserID As Long, bIsChecked As Boolean
    
    sSource = Me.Name & ".UpdateEditSetStatus()"
    
    'Read the query string
    
    lEditSetID = m_aspRequest.QueryString(gk_sParamEditSetID)
    lAssignedUserID = m_aspRequest.QueryString(ksParamAssignedUser)
    bIsChecked = m_aspRequest.QueryString(ksParamCheckBoxCurrentlyChecked)
    
    'Update the status
    m_oBOEditSet.UpdateEditSetStatus m_lLoggedUserID, lEditSetID, lAssignedUserID, Not bIsChecked
       
   'Release the objects
    
    m_aspResponse.Redirect (gk_sLinkDisplayEditSetsForOneUser)
    
    m_objObjectContext.SetComplete
    
    Exit Sub
ErrorHandler:
    m_objObjectContext.SetAbort
    ShowError sSource

End Sub





Public Sub DisplayUsersAndRolesForEditSet()
' Displays the user and their roles on the edit set
' d4      ndk     23Jun1999       Created
' d47     ndk     08Dec1999       Added help buttons
' d50     ndk     21Dec1999       Added security check
' d51     ndk     21Dec1999       Showing roles relevant for the user and stage when editing roles
' d54     ndk     28Dec1999       Added tooltips
' d61     dtm     09Feb2000       modified so that 'Remove' button only displayed when permissions allow
'   d73     daj     09May2000   Added some encoding of special characters to edit set name

On Error GoTo ErrorHandler

    Dim sSource As String, sHTML As String, sParamString
    Dim lEditSetID As Long, sEditSetName As String
    Dim oFWRecordlist  As FWRecordList, oFWRolesList As FWRecordList
    Dim sHyperlink As String
    Dim sText As String, eSecurityCheckStatus As ESecurityCheckResult
            
    Dim typTableStyle As TableStyle, typCellStyle As CellStyle, typFontStyle As FontStyle
    
    'sSource used in error handling
    sSource = Me.Name & ".DisplayUsersAndRolesForEditSet"
    
    'Get the edit set id from the request object
    lEditSetID = CLng(m_aspRequest.QueryString(gk_sParamEditSetID))
    sEditSetName = m_aspRequest.QueryString(ksParamEditSetName)
    
    'perform security check
    m_oBOEditorialPermissions.EditsetID = lEditSetID
    m_oBOEditorialPermissions.UserID = m_lLoggedUserID
    
    eSecurityCheckStatus = m_oBOEditorialPermissions.Check(EUserCanEditRolesOnEditSet)
    
    If eSecurityCheckStatus <> ESecurityCheckResult.ESRSecurityCheckSuccessful Then
        RaiseError sSource, _
               GetErrorEnumForSecurityCheckResult(eSecurityCheckStatus), _
               GetDescriptionForSecurityCheckError( _
                        GetErrorEnumForSecurityCheckResult(eSecurityCheckStatus))
    End If
    
    typTableStyle = m_oFWHTML.GetTableStyle(TABPageHeadingStyle)
    typTableStyle.lCELLSPACING = 2
    
    m_oFWTable.Reset
    m_oFWTable.TableStyle = typTableStyle
    
    m_oFWTable.AddBlankRow
    
    m_oFWTable.StartRow
    
    ShowNormalStyleText _
        m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayUsersAndRolesHeader1) & _
        m_oFWHTML.Bold(sEditSetName)
    
    m_oFWTable.EndRow
    
    m_oFWTable.AddBlankRow
    
    sHTML = sHTML & m_oFWHTML.Paragraph(m_oFWTable.GetHTML)
    
    sHTML = sHTML & m_oFWHTML.FormStart(gk_sLinkDisplayValidUsersForTheRole & _
                    "?" & gk_sParamEditSetID & "=" & lEditSetID & _
                    "&" & ksParamEditSetName & "=" & EncodeQueryStringValue(sEditSetName), gk_sMainFrame, gk_sHTMLDefaultFormName, MPost)
    
    m_oFWTable.Reset
    
    typTableStyle = m_oFWHTML.GetTableStyle(TABContentStyle)
    typTableStyle.lCELLSPACING = 2
    m_oFWTable.TableStyle = typTableStyle
  
    m_oFWTable.StartRow
    
    ShowColumnHeading _
        m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayUsersAndRolesColPerson), 1, "15%"
    
    ShowColumnHeading _
        m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayUsersAndRolesColRole), 1, "25%"
    
    ShowColumnHeading _
        m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayUsersAndRolesColRoleActiveInStages), 1, "50%"
    
    ShowNormalStyleText gk_sHTMLNonBreakSpace, "10%"
    
    'Get the users and roles for the editset
    Set oFWRecordlist = m_oBOEditSet.UsersAndRolesForEditSet(lEditSetID)
    
    If oFWRecordlist.RecordCount = 0 Then
        m_oFWTable.StartRow
        
        m_oFWTable.AddCellAsHTML "", typCellStyle     'Display Nothing
        
        m_oFWTable.EndRow
    Else
        
        oFWRecordlist.MoveFirst
        
        Do While Not oFWRecordlist.IsAtEnd
            m_oFWTable.StartRow
            
            'User Name
            ShowColumnData _
                oFWRecordlist.GetFieldValue(gk_sDBCOLUserName)
            
            'Role Description
            ShowColumnData _
                oFWRecordlist.GetFieldValue(gk_sDBCOLRoleDescription)
            
            'Stages where role is active
            ShowColumnData _
                oFWRecordlist.GetFieldValue(gk_sDBCOLTempStagesWhereRoleIsActive)

            m_oBOEditorialPermissions.EditsetID = lEditSetID
            m_oBOEditorialPermissions.UserID = m_lLoggedUserID
            m_oBOEditorialPermissions.RoleToBeRemovedID = oFWRecordlist.GetFieldValue(gk_sDBCOLRoleID)
            
            If (m_oBOEditorialPermissions.Check(ESecurityCheck.EUserCanRemoveRoleFromEditSet) = ESecurityCheckResult.ESRSecurityCheckSuccessful) Then

                If Not m_oBOEditorialPermissions.RoleCanOnlyBeReplaced( _
                        oFWRecordlist.GetFieldValue(gk_sDBCOLRoleID)) And _
                    Not m_oBOEditSet.UsersRoleOnEditSet(m_lLoggedUserID, lEditSetID) = _
                        oFWRecordlist.GetFieldValue(gk_sDBCOLRoleID) Then

                    'if the role is not removable (only replacable) and
                    'if it is not his own role

                    'Hyperlink for the remove button
                    sHyperlink = gk_sLinkRemoveRoleFromEditSet & "?" & _
                        gk_sParamEditSetID & "=" & lEditSetID & "&" & _
                        ksParamEditSetName & "=" & EncodeQueryStringValue(sEditSetName) & "&" & _
                        ksParamUserID & "=" & oFWRecordlist.GetFieldValue(gk_sDBCOLUserID) & "&" & _
                        ksParamRoleID & "=" & oFWRecordlist.GetFieldValue(gk_sDBCOLRoleID)

                    'Image for the remove button
                    sText = m_oFWHTML.ImageButton(gk_sImagesPath & gk_sImageRemove, "", sHyperlink, gk_sMainFrame, GetToolTip(ETOOLTIP.ttRemoveRole))

                    'Display the remove button
                    ShowNormalStyleText sText

                End If
            End If
            m_oFWTable.EndRow
            
            oFWRecordlist.MoveNext
        Loop
        
    End If
    
    typCellStyle = m_oFWHTML.GetCellStyle(CELLNormalStyle)       'To reinit
    typCellStyle.eVALIGN = VALNone
    
    m_oFWTable.AddBlankRow
    
    m_oFWTable.StartRow
    
    ShowNormalStyleText _
        m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayUsersAndRolesFooter1), "", 3
    
    m_oFWTable.EndRow
    
    'Get the list of roles
    Set oFWRolesList = m_oBOEditSet.RolesAssignableByUserInEditSet( _
        m_lLoggedUserID, lEditSetID)
    
    'Display the list of roles as a list box
    ShowNormalStyleText _
        m_oFWHTML.Listbox(ksNameRoleID, oFWRolesList, False, -1, 0, gk_sDBCOLRoleID, gk_sDBCOLRoleDescription), "", 3
    
    'Add button
    ShowNormalStyleText _
        m_oFWHTML.SubmitImage(gk_sImagesPath & gk_sImageAdd, "", "", ""), "", 3
         
    m_oFWTable.EndRow
        
    ShowNormalStyleText _
            gk_sHTMLNonBreakSpace, "", 3
    
    'Hyperlink for the done button
    sHyperlink = gk_sLinkDisplayEditSetDetails & _
                        "?" & gk_sParamEditSetID & "=" & lEditSetID
    
    'Image for the done button
    sText = m_oFWHTML.ImageButton(gk_sImagesPath & gk_sImageDone, "", sHyperlink, gk_sMainFrame, GetToolTip(ETOOLTIP.ttFinishedEditingRoles))
            
    'Display the done button
    ShowNormalStyleText sText

    m_oFWTable.EndRow
    
    sHTML = sHTML & m_oFWTable.GetHTML
    
    'Add the help button
    sHTML = sHTML & HTMLToAddHelpButton(gk_sHelpLinkToUsersAndRolesOnEditSet)
    
    m_aspResponse.Write sHTML
    
    'Release the objects
    Set oFWRecordlist = Nothing
    Set oFWRolesList = Nothing
    
    m_objObjectContext.SetComplete
    
    Exit Sub
    
ErrorHandler:

    Set oFWRecordlist = Nothing
    Set oFWRolesList = Nothing

    m_objObjectContext.SetAbort
    ShowError sSource
End Sub





Public Sub RemoveRoleFromEditSet()
'  Removes the role from an edit set
'
' d4      ndk     23Jun1999     Created
' d50     ndk     21Dec1999     Added security check
'   d73     daj     09May2000   Added some encoding of special characters to edit set name
'

On Error GoTo ErrorHandler

    Dim sParamString As String
    Dim sSource As String
    Dim lEditSetID As Long, lRoleIDToBeRemoved As Long, sEditSetName As String
    Dim eSecurityCheckStatus As ESecurityCheckResult
 
    'sSource used in error handling
    sSource = Me.Name & ".RemoveRoleFromEditSet()"

    'Get the details from the request object
    lEditSetID = CLng(m_aspRequest.QueryString(gk_sParamEditSetID))
    sEditSetName = m_aspRequest.QueryString(ksParamEditSetName)
    lRoleIDToBeRemoved = CLng(m_aspRequest.QueryString(ksParamRoleID))
    
    'perform security check
    m_oBOEditorialPermissions.EditsetID = lEditSetID
    m_oBOEditorialPermissions.UserID = m_lLoggedUserID

    eSecurityCheckStatus = m_oBOEditorialPermissions.Check(EUserCanEditRolesOnEditSet)
    
    If eSecurityCheckStatus <> ESecurityCheckResult.ESRSecurityCheckSuccessful Then
        RaiseError sSource, _
               GetErrorEnumForSecurityCheckResult(eSecurityCheckStatus), _
               GetDescriptionForSecurityCheckError( _
                        GetErrorEnumForSecurityCheckResult(eSecurityCheckStatus))
    End If
    
    'Call the method in the business object to remove the role from the editset
    m_oBOEditSet.RemoveRoleFromEditSet lEditSetID, lRoleIDToBeRemoved, m_lLoggedUserID
        
    sParamString = "?" & gk_sParamEditSetID & "=" & lEditSetID
    sParamString = sParamString & "&" & ksParamEditSetName & "=" & EncodeQueryStringValue(sEditSetName)
    
    'Send the control back to the previous page
    m_aspResponse.Redirect gk_sLinkDisplayUsersAndRolesForEditSet & sParamString
    
    m_objObjectContext.SetComplete
        
    Exit Sub

ErrorHandler:
    m_objObjectContext.SetAbort
    ShowError sSource
End Sub





Public Sub AddUpdateRoleOnEditSet()
'  Adds a new role, or updates an existing role with a new user
'
'   d5  ndk     24Jun1999   Created
'   d8  ndk     01Jul1999   Now using resource strings
'   d50 ndk     21Dec1999   Added security check
'   d72 tmj     28Mar2000   Changed so final redirects only redirects to
'                           role list if user has permissions to view it
'                           else redirects to the homepage
'   d73     daj     09May2000   Added some encoding of special characters to edit set name
'

On Error GoTo ErrorHandler
    
    Dim sSource As String, sParamString As String
    Dim lEditSetID As Long, lRoleIDToBeAdded As Long, lUserIDToBeAdded  As Long, sEditSetName As String
    Dim eSecurityCheckStatus As ESecurityCheckResult
    
    sSource = Me.Name & ".AddUpdateRoleOnEditSet()"
     
     'Get the details from the request object
    lEditSetID = CLng(m_aspRequest.QueryString(gk_sParamEditSetID))
    sEditSetName = m_aspRequest.QueryString(ksParamEditSetName)
    lRoleIDToBeAdded = CLng(m_aspRequest.QueryString(ksParamRoleID))
    lUserIDToBeAdded = CLng(m_aspRequest.QueryString(ksParamUserID))
    
    'perform security check
    m_oBOEditorialPermissions.EditsetID = lEditSetID
    m_oBOEditorialPermissions.UserID = m_lLoggedUserID
    
    eSecurityCheckStatus = m_oBOEditorialPermissions.Check(EUserCanEditRolesOnEditSet)
    
    If eSecurityCheckStatus <> ESecurityCheckResult.ESRSecurityCheckSuccessful Then
        RaiseError sSource, _
               GetErrorEnumForSecurityCheckResult(eSecurityCheckStatus), _
               GetDescriptionForSecurityCheckError( _
                        GetErrorEnumForSecurityCheckResult(eSecurityCheckStatus))
    End If
    
    'Call the method in the business object to add/replace the role on the editset
    m_oBOEditSet.AddUpdateRoleOnEditSet lEditSetID, lRoleIDToBeAdded, lUserIDToBeAdded, m_lLoggedUserID
    
    'Redirect to the roles page
    sParamString = "?" & gk_sParamEditSetID & "=" & lEditSetID
    sParamString = sParamString & "&" & ksParamEditSetName & "=" & EncodeQueryStringValue(sEditSetName)
    
    'Checking to see if user has permissions to view the edit page
    eSecurityCheckStatus = m_oBOEditorialPermissions.Check(EUserCanEditRolesOnEditSet)
    If eSecurityCheckStatus = ESecurityCheckResult.ESRSecurityCheckSuccessful Then
        'Send the control back to the previous page
        m_aspResponse.Redirect gk_sLinkDisplayUsersAndRolesForEditSet & sParamString
    Else
        'Send the control back to the home page if user does not have permissions
        m_aspResponse.Redirect gk_sLinkDisplayEditSetsForOneUser
    End If
    
    m_objObjectContext.SetComplete
    
    Exit Sub
ErrorHandler:
    m_objObjectContext.SetAbort
    ShowError sSource

End Sub





Public Sub DisplayValidUsersForTheRole()
'   Displays the list of users who can take on the role selected
'   in add role page
'
'   d5  ndk     24Jun1999   Created
'   d8  ndk     01Jul1999   Now using resource strings
'   d42 ndk     22Nov1999   Using a different method to get the list of users assignable to a role
'   d47 ndk     08Dec1999   Added help buttons
'   d50 ndk     21Dec1999   Added security check
'   d54 ndk     28Dec1999   Added tooltips
'   d58 aga     14Jan2000   Added cancel button
'   d76     daj     31Oct2000   Fixed problem in DisplayValidUsersForRole where the edit set name could cause problems
'                               if a special character (eg, &, #, ~, ? etc.) was added into the name.

On Error GoTo ErrorHandler

    Dim sSource As String
    Dim lEditSetID As Long, lRoleIDToBeAdded As Long, lUserIDToBeAdded  As Long, sUserName As String, sEditSetName
    
    Dim sHTML As String, sHyperlink As String
    Dim typTableStyle As TableStyle, typCellStyle As CellStyle
    Dim oFWRecordlist As FWRecordList
    Dim eSecurityCheckStatus As ESecurityCheckResult

    sSource = Me.Name & ".DisplayValidUsersForTheRole()"

     'Get the details from the request object
    lEditSetID = CLng(m_aspRequest.QueryString(gk_sParamEditSetID))
    
    ' d76 - Added encoding to fix dodgy characters problem
    sEditSetName = EncodeQueryStringValue(m_aspRequest.QueryString(ksParamEditSetName))
    lRoleIDToBeAdded = CLng(m_aspRequest.Form(ksNameRoleID))
   
    'perform security check
    m_oBOEditorialPermissions.EditsetID = lEditSetID
    m_oBOEditorialPermissions.UserID = m_lLoggedUserID
   
    eSecurityCheckStatus = m_oBOEditorialPermissions.Check(EUserCanEditRolesOnEditSet)
    
    If eSecurityCheckStatus <> ESecurityCheckResult.ESRSecurityCheckSuccessful Then
        RaiseError sSource, _
               GetErrorEnumForSecurityCheckResult(eSecurityCheckStatus), _
               GetDescriptionForSecurityCheckError( _
                        GetErrorEnumForSecurityCheckResult(eSecurityCheckStatus))
    End If
   
    typTableStyle = m_oFWHTML.GetTableStyle(TABPageHeadingStyle)
    typTableStyle.lCELLSPACING = 2
    
    m_oFWTable.Reset
    m_oFWTable.TableStyle = typTableStyle
    
    m_oFWTable.AddBlankRow
    
    m_oFWTable.StartRow
    
    ShowNormalStyleText _
        m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayUsersAndRolesDisplayValidUsers) & m_aspRequest.QueryString(ksParamEditSetName)
    
    m_oFWTable.EndRow
    
    m_oFWTable.AddBlankRow
    
    sHTML = sHTML & m_oFWHTML.Paragraph(m_oFWTable.GetHTML)
    
    m_oFWTable.Reset
    
    typTableStyle = m_oFWHTML.GetTableStyle(TABContentStyle)
    typTableStyle.sWidth = "70%"
    typTableStyle.lCELLSPACING = 2
    
    m_oFWTable.TableStyle = typTableStyle
    
    'Call the method in the business object to get the list of users who can be assigned
    'the selected role
    
    Set oFWRecordlist = _
        m_oBOEditSet.ValidUsersForTheRoleOnEditSet(lRoleIDToBeAdded, lEditSetID)
    
    If oFWRecordlist.RecordCount <> 0 Then
        oFWRecordlist.MoveFirst
        
        'To display the users names
        Do While Not oFWRecordlist.IsAtEnd
            m_oFWTable.StartRow
            
            lUserIDToBeAdded = oFWRecordlist.GetFieldValue(gk_sDBCOLUserID)
            
            sUserName = oFWRecordlist.GetFieldValue(gk_sDBCOLUserName)
            
            sHyperlink = gk_sLinkAddUpdateRoleOnEditSet & _
                 "?" & gk_sParamEditSetID & "=" & lEditSetID & _
                 "&" & ksParamRoleID & "=" & lRoleIDToBeAdded & _
                 "&" & ksParamEditSetName & "=" & sEditSetName & _
                 "&" & ksParamUserID & "=" & lUserIDToBeAdded
                 
            ShowNormalStyleText _
                m_oFWHTML.Hyperlink(sHyperlink, sUserName, gk_sMainFrame, , GetToolTip(ETOOLTIP.ttSelectUserForRole))
            
            m_oFWTable.EndRow
            
            oFWRecordlist.MoveNext
        Loop
    End If
    
    sHTML = sHTML & m_oFWHTML.Paragraph(m_oFWTable.GetHTML)
    
    'd58
    'Add cancel button
    sHTML = sHTML & gk_sHTMLLineBreak & m_oFWHTML.ImageButton(gk_sImagesPath & gk_sImageCancel, , _
        gk_sLinkDisplayUsersAndRolesForEditSet & "?" _
        & gk_sParamEditSetID & "=" & lEditSetID _
        & "&" & ksParamEditSetName & "=" & sEditSetName, _
        gk_sMainFrame, "Cancel adding a person to editset")
        
    'Add the help button
    sHTML = sHTML & HTMLToAddHelpButton(gk_sHelpLinkToValidUsersForRole)
    
    m_aspResponse.Write sHTML
    
    'release instances
    Set oFWRecordlist = Nothing

    m_objObjectContext.SetComplete
        
    Exit Sub

ErrorHandler:

    Set oFWRecordlist = Nothing
    
    m_objObjectContext.SetAbort
    ShowError sSource
End Sub





Public Sub DisplayCreateEmptyEditSetScreen(Optional t_eUsage As ECreateEditSetUsage = cesNormal)
'
'Description
'
'Displays the screen to create an empty edit list
'
'Amendment history
'   d7  ndk     30Jun1999       Created
'   d8  ndk     01Jul1999       now using resource strings
'   d40 ksk     18Nov1999       Added optional parameters for creating edit set entries
'                               Needed so that it can be integrated with the accession
'                               module
'   d47 ndk     08Dec1999       Added help buttons
'   d57 crm     13Jan2000       Modified to show only new accession in Editset type list
'                               box if creation is by import
'   d70     daj     21Mar2000   Changed the edit set creation screen to show a context-sensitive title
'   d75     daj     17Jul2000   Modified CreateEmptyEditSet and DisplayCreateEmptyEditSetScreen to allow updates
'   d86  map     08Oct2001       Added scheduling functionality
'   d87     alm     19Oct2001   Supports the new edit set type 'delete'
'   d88 map     23Oct2001       Changed format of screen
'   d94 alm     27Nov2001       Added the time restraint to scheduled job header
On Error GoTo ErrorHandler

   
    Dim sSource As String, sErrorMessage As String
    Dim bCanCreateEditSet As Boolean
    Dim bShowListBoxToChooseContributor1 As Boolean
    Dim iColoumnNumber As Integer
    Dim sHTML As String, sParamString   As String
    Dim lMaxNumberOfUsers As Long
    Dim vRowColour As Variant
    
    Dim typRowStyle As RowStyle
    Dim typCellStyle As CellStyle
    Dim typTableStyle As TableStyle
    ' d86
    Dim typCell As CellStyle
    Dim ofwhtmlReference As FWHTML
    Dim oFWEditSetTypesList As FWRecordList
    Dim oFWUsersForContributor1 As FWRecordList
    Dim oboUserRoles As BoUserRoles
    
    '   d40 ksk Variable to hold the link depending on the source of call
    Dim sLinkSubmitpage As String, sLinkCancelPage As String
    Dim sParam As String
        
    sSource = Me.Name & ".DisplayCreateEmptyEditScreen"
    
    'Create instances
    Set oboUserRoles = New BoUserRoles
    Set ofwhtmlReference = New FWHTML
    m_oFWTable.Reset
        
    bCanCreateEditSet = True
    bShowListBoxToChooseContributor1 = False
    
    oboUserRoles.UserID = m_lLoggedUserID
    If oboUserRoles.UserHasAdministratorsRole() Then
        'Logged User is Administrator   '
        'Get list of users to become contributor1
        
        oboUserRoles.RoleID = gk_lContributor1RoleID
        Set oFWUsersForContributor1 = oboUserRoles.ValidUsersForTheRole
                
        If oFWUsersForContributor1.RecordCount = 0 Then
            'No users who can take on the contributor1 role
            bCanCreateEditSet = False
        Else
            bCanCreateEditSet = True
            bShowListBoxToChooseContributor1 = True
        End If
        
        'd87 If the user is an administrator and the editset type was normal
        'make it delete
        If t_eUsage = cesNormal Then
            t_eUsage = cesDelete
        End If
        
    Else
        'Some other user
        m_oBOEditorialPermissions.UserID = m_lLoggedUserID
        If m_oBOEditorialPermissions.UserCanBeContributor1() Then
            'The same user will be cont1
            bCanCreateEditSet = True
            bShowListBoxToChooseContributor1 = False
        Else
            'The user cant have cont1 role, so he can not create an editset
            sErrorMessage = m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfCreateEmptyEditSetErrorCantCreateCont1)
            bCanCreateEditSet = False
        End If
    End If
    
    
    ' d75 - Decide on how to handle the call
    Select Case t_eUsage
        Case cesNormal, cesDelete 'd87 - Added cesDelete
            sLinkSubmitpage = gk_sLinkCreateEmptyEditSet
            sLinkCancelPage = gk_sLinkDisplayEditSetsForOneUser
        Case cesImport
            'If the method is called from Loading then the batch ID is passes as a parameter#
            ' for loading which needs to be passed on to the create edit set asp
            sParam = "?" & gk_sParamBatchID & "=" & m_aspRequest(gk_sParamBatchID)
            sLinkSubmitpage = gk_sLinkCreateEmptyEditSetForLoad & sParam
            sLinkCancelPage = gk_sLinkDisplayBatchList
        Case cesUpdate
            ' If the method is called from updating then we need to pass the filename through
            ' The return link has to be set to the file list
            sParam = "?" & gk_sParamFileName & "=" & m_aspRequest(gk_sParamFileName)
            sLinkSubmitpage = gk_sLinkCreateEmptyEditSetForUpdate & sParam
            sLinkCancelPage = gk_sLinkDisplayEADFileList & "?" & gk_sParamUpdateEADList & "=1"
        Case cesDRUIDLoad
            ' If the method is called for creating an edit set for a DRUID Load then need...
            sParam = "?" & gk_sParamBatchID & "=" & m_aspRequest(gk_sParamBatchID)
            sLinkSubmitpage = gk_sLinkCreateEmptyEditSetForDRUIDLoad & sParam
            sLinkCancelPage = gk_sLinkDisplayBatchList
    End Select
    
    If bCanCreateEditSet Then
        sHTML = sHTML & m_oFWHTML.FormStart(sLinkSubmitpage, _
                                gk_sMainFrame, gk_sHTMLDefaultFormName, MPost, True)
    
        typTableStyle = m_oFWHTML.GetTableStyle(TABContentStyle)
        typTableStyle.lCELLSPACING = 2
    
        m_oFWTable.TableStyle = typTableStyle
        
        m_oFWTable.StartRow
    
        ' d75 - Use a context sensitive title
        Select Case t_eUsage
            Case cesImport, cesDRUIDLoad
                ShowDetailsLine _
                    m_oFWHTML.Bold(m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfCreateLoadEditSetHeader)), _
                    typCellStyle, m_oFWHTML.GetFontStyle(FONTNormalStyle)
            Case cesNormal, cesDelete 'd87 - Added cesDelete
                ShowDetailsLine _
                    m_oFWHTML.Bold(m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfCreateEmptyEditSetHeader1)), _
                    typCellStyle, m_oFWHTML.GetFontStyle(FONTNormalStyle)
            Case cesUpdate
                ShowDetailsLine _
                    m_oFWHTML.Bold(m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfCreateUpdateEditSetHeader)), _
                    typCellStyle, m_oFWHTML.GetFontStyle(FONTNormalStyle)

        End Select
    
        m_oFWTable.EndRow
    
        sHTML = sHTML & m_oFWHTML.Paragraph(m_oFWTable.GetHTML)
        
        m_oFWTable.Reset
    
        typTableStyle = m_oFWHTML.GetTableStyle(TABContentStyle)
        typTableStyle.lCELLSPACING = 0
        typTableStyle.lCELLPADDING = 3
      
        m_oFWTable.TableStyle = typTableStyle
        
        'Name of edit set
        
        m_oFWTable.StartRow
           
        ShowGeneralLeftHeading _
            m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfCreateEmptyEditSetColName)
        
        ShowGeneralRightContent _
            m_oFWHTML.InputField(ITText, ksNameEditSetName, 50, gk_iDBCOLLenEditSetName, "", VCNone, True)
        
        m_oFWTable.EndRow
        
        'Name of contributor1
        
        m_oFWTable.StartRow
           
        ShowGeneralLeftHeading _
            "Contributor1 :"
           
        If bShowListBoxToChooseContributor1 Then
            ShowGeneralRightContent _
                m_oFWHTML.Listbox(ksNameContributor1ForNewEditSetID, oFWUsersForContributor1, False, -1, 0, gk_sDBCOLUserID, gk_sDBCOLUserName)
        Else
            ShowGeneralRightContent _
                GetDataFromSession(dsUserName, m_objObjectContext.Item(gk_sSessionObject)) & _
                m_oFWHTML.HiddenField(ksNameContributor1ForNewEditSetID, CStr(m_lLoggedUserID))
        End If
        
        m_oFWTable.EndRow
        
        'Type of editset
        m_oFWTable.StartRow
           
        ShowGeneralLeftHeading _
            m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfCreateEmptyEditSetColType)
            
        'Get the list of editset types
        Set oFWEditSetTypesList = m_oBOEditSet.EditSetTypes(t_eUsage)
               
        ShowGeneralRightContent _
            m_oFWHTML.Listbox(ksNameEditSetType, oFWEditSetTypesList)
            
        m_oFWTable.EndRow
        
        'Comments
        
        m_oFWTable.StartRow
        ShowGeneralLeftHeading _
            m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfCreateEmptyEditSetColComments)
            
        ShowGeneralRightContent _
            m_oFWHTML.TextArea(ksNameEditSetDescription, "", 40, 3, 1000)
            
        m_oFWTable.EndRow
        
        ' d86 - Need option to schedule update or DRUID load
        If t_eUsage = cesUpdate Or t_eUsage = cesDRUIDLoad Then
        
            ' d88 - Amended heading
            ' d94 - Amended heading
            m_oFWTable.StartRow
            typCell = ofwhtmlReference.GetCellStyle(ECellStyle.CELLPageHeadingStyle)
            typCell.lCOLSPAN = 2
            m_oFWTable.AddCellAsHTML "Enter scheduling details below. The time must be between " _
            & m_oFWAppData.GetApplicationConfigurationInfomation(pScheduledJobTimeFrame) & _
            ". The date is optional. Please check the list of currently scheduled jobs and select a time that does not conflict with an existing job", typCell
            m_oFWTable.EndRow
            
            ' d88 - Added format info to start time field
            m_oFWTable.StartRow
            typCell = ofwhtmlReference.GetCellStyle(ECellStyle.CELLGeneralLabelStyle)
            typCell.lCOLSPAN = 1
            m_oFWTable.AddCellAsHTML "Start time (hh:mm 24hr clock):", typCell
            typCell = ofwhtmlReference.GetCellStyle(ECellStyle.CELLReadWriteDataStyle)
            m_oFWTable.AddCellAsHTML ofwhtmlReference.TextField(gk_sParamJobStartTime, _
                                gk_iDBCOLLenJobStartTime, , 10), typCell
            m_oFWTable.EndRow
            
            ' d88 - Added format info to date field
            m_oFWTable.StartRow
            typCell = ofwhtmlReference.GetCellStyle(ECellStyle.CELLGeneralLabelStyle)
            typCell.lCOLSPAN = 1
            m_oFWTable.AddCellAsHTML "Date (dd/mm/yyyy):", typCell
            typCell = ofwhtmlReference.GetCellStyle(ECellStyle.CELLReadWriteDataStyle)
            m_oFWTable.AddCellAsHTML ofwhtmlReference.TextField(gk_sParamJobStartDate, _
                                gk_iDBCOLLenJobStartDate, , 10), typCell
            m_oFWTable.EndRow
            
            
        End If
        
        'Ok and cancel buttons
        
        m_oFWTable.StartRow
        
        ShowGeneralLeftHeading ""
        ShowGeneralRightContent _
            m_oFWHTML.SubmitImage(gk_sImagesPath & gk_sImageOK, "", "") & _
            m_oFWHTML.ImageButton(gk_sImagesPath & gk_sImageCancel, "", _
                                      sLinkCancelPage, gk_sMainFrame)
                
        m_oFWTable.EndRow
  
        sHTML = sHTML & m_oFWHTML.Paragraph(m_oFWTable.GetHTML)
        sHTML = sHTML & m_oFWHTML.EndTag(gk_sHTMLForm)
    Else
        sHTML = sHTML & m_oFWHTML.FormStart(sLinkCancelPage, gk_sMainFrame, gk_sHTMLDefaultFormName, MPost)
        
        typTableStyle = m_oFWHTML.GetTableStyle(TABContentStyle)
        typTableStyle.lCELLSPACING = 2
    
        m_oFWTable.TableStyle = typTableStyle
        
        typCellStyle = m_oFWHTML.GetCellStyle(CELLNormalStyle)       'To reinit
        
        m_oFWTable.StartRow
    
        m_oFWTable.AddCellAsHTML m_oFWHTML.Bold(sErrorMessage), typCellStyle
   
        m_oFWTable.EndRow
        
        m_oFWTable.AddBlankRow
        
        m_oFWTable.StartRow
    
        m_oFWTable.AddCellAsHTML m_oFWHTML.SubmitImage(gk_sImagesPath & gk_sImageOK, "", ""), typCellStyle
   
        m_oFWTable.EndRow
        
        sHTML = sHTML & m_oFWHTML.Paragraph(m_oFWTable.GetHTML)
        sHTML = sHTML & m_oFWHTML.EndTag(gk_sHTMLForm)

    End If
    
    'Add the help button
    sHTML = sHTML & HTMLToAddHelpButton(gk_sHelpLinkToCreateEmptyEditSet)
    
    m_aspResponse.Write sHTML

    Set oFWUsersForContributor1 = Nothing
    Set oFWEditSetTypesList = Nothing
    Set oboUserRoles = Nothing

    m_objObjectContext.SetComplete

    Exit Sub

ErrorHandler:
    m_objObjectContext.SetAbort
    ShowError sSource
End Sub





Public Sub CreateEmptyEditSet(Optional t_eUsage As ECreateEditSetUsage = cesNormal)
'
'Description
'
' Adds a new empty edit set.  Can be used from Editorial normally, accessioning or updating.
'
'Amendment history
'   d7  ndk     30Jun1999   Created
'   d55 ndk     28Dec1999   Changed the error raised when length of es name /comments exceed limit
'   d75     daj     17Jul2000   Modified CreateEmptyEditSet to allow updates
'   d86  map     09Oct2001   Added Scheduling functionality
'   d87 map     23Oct2001   Check for valid time and date if scheduling, BEFORE creating a new edit set
'
On Error GoTo ErrorHandler

    Dim sSource As String
    Dim lEditSetType As Long, sEditSetName  As String, sComments As String
    Dim lUserIDForContributor1 As Long
    Dim sParam  As String
    Dim sLinkRedirectPage As String
    ' d87 - Added
    Dim sTime As String, sDate As String
    
    sSource = Me.Name & ".CreateEmptyEditSet()"
    
    ' d75 - Retrieve the details from the form, or generate them
    lEditSetType = CLng(m_aspRequest.Form(ksNameEditSetType))      'Type for the new editset
    sEditSetName = m_aspRequest.Form(ksNameEditSetName)    'Name for the new editset
    sComments = m_aspRequest.Form(ksNameEditSetDescription)    'Description
    lUserIDForContributor1 = m_aspRequest.Form(ksNameContributor1ForNewEditSetID)   'Logged user is creating editset
        
    If LTrim(RTrim(sEditSetName)) = "" Then   'Edit Set Name not entered
        RaiseError sSource, peIncorrectLength, "EditSet name should not be blank"
    End If
    
    If Len(LTrim(RTrim(sComments))) > gk_iDBCOLLenFirstComment Then
        RaiseError sSource, peIncorrectLength, "Comments should not exceed " & gk_iDBCOLLenComment & " characters"
    End If
            

    'set the properties
    m_oBOEditSet.EditSetTypeID = lEditSetType
    m_oBOEditSet.EditSetName = sEditSetName
    m_oBOEditSet.Contributor1UserID = lUserIDForContributor1
    m_oBOEditSet.Description = sComments

    ' d87 - Check if a valid time and date have been entered
    sTime = m_aspRequest(gk_sParamJobStartTime)
    If sTime <> "" Then
        If Not ValidTime(sTime) Then
            RaiseError sSource, peUSERInvalidTime, "Scheduled start time entered is invalid."
        End If
    End If
    sDate = m_aspRequest(gk_sParamJobStartDate)
    If sDate <> "" Then
        If Not IsDate(sDate) Then
            RaiseError sSource, peUSERInvalidDate, "Scheduled date entered is invalid."
        End If
    End If

    'Create the editset
    m_oBOEditSet.CreateNewEmpty (m_lLoggedUserID)

        
    ' d75 - SELECT the usage
    Select Case t_eUsage
        Case cesNormal
            sLinkRedirectPage = gk_sLinkDisplayEditSetsForOneUser
        
        Case cesImport
            'If the method is called from the batch Loading process then the
            '  batch ID is passes as a parameter and both batch Id and editset ID
            '  need to be passed to the Loading start files for loading which needs
            '  to be passed on to the create edit set asp
            sParam = "?" & gk_sParamBatchID & "=" & m_aspRequest(gk_sParamBatchID) & _
                        "&" & gk_sParamEditSetID & "=" & m_oBOEditSet.EditsetID
            sLinkRedirectPage = gk_sLinkStartLoadingIntoEditset & sParam
        
        Case cesUpdate
            ' Need to pass through the edit set name, filename and edit set
            ' d86 - Added schedule times to param
            sParam = "?" & gk_sParamEditSetID & "=" & m_oBOEditSet.EditsetID & _
                "&" & gk_sParamFileName & "=" & m_aspRequest(gk_sParamFileName) & _
                "&" & gk_sParamBatchDescription & "=" & EncodeQueryStringValue(m_oBOEditSet.EditSetName) & _
                "&" & gk_sParamJobStartTime & "=" & m_aspRequest(gk_sParamJobStartTime) & _
                "&" & gk_sParamJobStartDate & "=" & m_aspRequest(gk_sParamJobStartDate)
            sLinkRedirectPage = gk_sLinkStartUpdatingIntoEditset & sParam
        
        Case cesDRUIDLoad
            ' d86 - Now need to kick off DRUID Load and if this is successful, will need edit set details
            sParam = "?" & gk_sParamEditSetID & "=" & m_oBOEditSet.EditsetID & _
                "&" & gk_sParamBatchID & "=" & m_aspRequest(gk_sParamBatchID) & _
                "&" & gk_sParamFileName & "=" & m_aspRequest(gk_sParamFileName) & _
                "&" & gk_sParamBatchDescription & "=" & EncodeQueryStringValue(m_oBOEditSet.EditSetName) & _
                "&" & gk_sParamJobStartTime & "=" & m_aspRequest(gk_sParamJobStartTime) & _
                "&" & gk_sParamJobStartDate & "=" & m_aspRequest(gk_sParamJobStartDate)
            sLinkRedirectPage = gk_sLinkStartDRUIDLoadIntoEditset & sParam
            
    End Select
    
    'goto the home page
    m_aspResponse.Redirect sLinkRedirectPage
    
    m_objObjectContext.SetComplete
    
    Exit Sub

ErrorHandler:
    m_objObjectContext.SetAbort
    ShowError sSource
End Sub





Public Sub DisplayEditSetHistory()
'
'Description
'
'Displays the history of the edit set
'
'Amendment history
'   d8  ndk     01Jul1999       Created
'   d18 ndk     16Jul1999       Minor VI Changes after review before demo
'   d47 ndk     08Dec1999       Added help buttons
'   d49 ndk     16Dec1999       Displaying created on date
'   d53 ndk     27Dec1999       Changed to VI code use customizable styles
'   d57 crm     12Jan2000       Added done button
'   d67 daj     07Mar2000       Changed DisplayEditSetHistory to show parser/accession log links
'   d69 tmj     09Mar2000       Changed so displays original contributor, and removed contributor1 from history
'                               link

On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim sHTML As String, sParamString   As String
    Dim lEditSetID As Long, sEditSetName As String, sCreationDate As String, sContributor1 As String
    Dim iTotalDays As Integer
    Dim typCellStyle As CellStyle
    Dim typTableStyle As TableStyle
    Dim typFontStyle As FontStyle
    Dim oFWRecordlist As FWRecordList
  
    ' d67 - Filename and log strings
    Dim lBatchID As Long
    Dim sEADFilename As String, sParam As String, sToolTip As String
    Dim ofwReferenceData As New FWReferenceData

    sSource = Me.Name & ".DisplayEditHistory"
    
    'Get the id of the editset to be displayed
    lEditSetID = CLng(m_aspRequest.QueryString(gk_sParamEditSetID))
    sEditSetName = m_aspRequest.QueryString(ksParamEditSetName)
    sCreationDate = m_aspRequest.QueryString(ksParamEditSetCreatedDate)
    
    'd69 Get the creator name
    sContributor1 = m_oBOEditSet.GetCreator(lEditSetID)
    
    ' d67 - Find the EAD filename that created the edit set
    sEADFilename = m_oBOEditSet.GetFilenameForEditSet(lEditSetID)
    lBatchID = m_oBOEditSet.GetBatchIDForEditSet(lEditSetID)
    
    m_oFWTable.Reset
    
    typTableStyle = m_oFWHTML.GetTableStyle(TABContentStyle)
    typTableStyle.sWidth = "100%"
    typTableStyle.lCELLSPACING = 2
    
    m_oFWTable.TableStyle = typTableStyle
    
    m_oFWTable.StartRow
    
    'Headings
    ShowNormalStyleText _
            m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetHistoryHeader1) & _
            m_oFWHTML.Bold(sEditSetName)
    
    m_oFWTable.EndRow
    
    m_oFWTable.AddBlankRow
    
    'Headings
    ShowNormalStyleText _
        m_oFWHTML.Bold(sEditSetName) _
        & " " & m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetHistoryHeader2A) _
        & " " & sCreationDate _
        & " " & m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetHistoryHeader2B) _
        & " " & sContributor1
    
    ' d67 - Add the logs
    If lBatchID > 0 Then
        ' Parameters for the log viewing pages
        'd69 removed passing of contributor1 name
        sParam = "?" & gk_sParamBatchID & "=" & lBatchID & _
            "&" & gk_sParamRedirection & "=" & _
            URLEncode(gk_sLinkDisplayEditSetHistory & _
                "?" & gk_sParamEditSetID & "=" & lEditSetID & _
                "&" & ksParamEditSetName & "=" & sEditSetName & _
                "&" & ksParamEditSetCreatedDate & "=" & sCreationDate)
        
        ' Only show the parsing log if this is an EAD batch
        If sEADFilename <> "" Then
            sToolTip = ofwReferenceData.Search(EGroup.gToolTip, ETOOLTIP.ttParseLogFileView)
            ShowNormalStyleText _
                m_oFWHTML.Hyperlink(gk_sLinkDisplayEADParseLogging & sParam, _
                    "Parse log", gk_sMainFrame, , _
                    sToolTip)
        End If
        
        ' Import log
        sToolTip = ofwReferenceData.Search(EGroup.gToolTip, ETOOLTIP.ttImportLogFileView)
        ShowNormalStyleText _
            m_oFWHTML.Hyperlink(gk_sLinkDisplayAccessionLogFile & sParam, _
                "Import log", gk_sMainFrame, , _
                sToolTip)
        
        ' Load log
        sParam = sParam & "&" & gk_sParamEditSetID & "=" & lEditSetID
        sToolTip = ofwReferenceData.Search(EGroup.gToolTip, ETOOLTIP.ttLoadLogFileView)
        ShowNormalStyleText _
            m_oFWHTML.Hyperlink(gk_sLinkDisplayAccessionLogFile & sParam, _
                "Load log", gk_sMainFrame, , _
                sToolTip)
        
    End If
    
    sHTML = sHTML & m_oFWTable.GetHTML

    'To display the column headers
    
    m_oFWTable.Reset
    
    typTableStyle.sWidth = "100%"
    typTableStyle.lCELLSPACING = 2
    
    m_oFWTable.TableStyle = typTableStyle
    m_oFWTable.StartRow
    
    'Stage name
    ShowColumnHeading _
            m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetHistoryColStage), 1, "10%"
    
    'Date stage entered
    ShowColumnHeading _
        m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetHistoryColStageStart), 1, "15%"
    
    'Date stage left
    ShowColumnHeading _
        m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetHistoryColStageEnd), 1, "15%"
    
    'No of days in the stage
    ShowColumnHeading _
        m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetHistoryColDaysInStage), 1, "15%"
    
    'Comments
    ShowColumnHeading _
        m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetHistoryColComments), 1, "45%"
    
    m_oFWTable.EndRow
    
    'Get the history of the ediset
    Set oFWRecordlist = m_oBOEditSet.EditSetHistory(lEditSetID)

    typCellStyle = m_oFWHTML.GetCellStyle(CELLReadWriteDataStyle)
    typFontStyle = m_oFWHTML.GetFontStyle(FONTGeneralContentStyle)
    
    oFWRecordlist.MoveFirst
    
    iTotalDays = 0
    
    'Loop through the stages of the editset
    Do While Not oFWRecordlist.IsAtEnd
        m_oFWTable.StartRow
        
        'Stage name
        ShowColumnData _
            oFWRecordlist.GetFieldValue(gk_sDBCOLEditsetStageDescription), 1, "", True
        
        'Date stage entered
        ShowColumnData _
            VariantDateAsDDMMMYYYY(oFWRecordlist.GetFieldValue(gk_sDBCOLTempStageStartDate))

        'Date stage left
        If IsNull(oFWRecordlist.GetFieldValue(gk_sDBCOLTempStageEndDate)) Then
            ShowColumnData _
                gk_sHTMLNonBreakSpace
        Else
            ShowColumnData _
                VariantDateAsDDMMMYYYY(oFWRecordlist.GetFieldValue(gk_sDBCOLTempStageEndDate))
        End If

        'No of days spent in the stage
        If IsNull(oFWRecordlist.GetFieldValue(gk_sDBCOLTempDaysInCurrentStage)) Then
            ShowColumnData CStr(0)
        Else
            iTotalDays = iTotalDays + oFWRecordlist.GetFieldValue(gk_sDBCOLTempDaysInCurrentStage)
            ShowColumnData _
                oFWRecordlist.GetFieldValue(gk_sDBCOLTempDaysInCurrentStage)
        End If

        'Comments for the stage transition
        ShowColumnData _
            VariantStringAsString(oFWRecordlist.GetFieldValue(gk_sDBCOLTempComment)), 1, "", True
        
        oFWRecordlist.MoveNext
    Loop
        
    m_oFWTable.AddBlankRow
    
    m_oFWTable.StartRow
    
    m_oFWTable.AddCellAsHTML gk_sHTMLNonBreakSpace
    m_oFWTable.AddCellAsHTML gk_sHTMLNonBreakSpace
    
    typCellStyle = m_oFWHTML.GetCellStyle(CELLNormalStyle)
    typCellStyle.eALIGN = HALRight
    
    'Total days spent in all stages
    ShowDetailsLine _
        m_oFWAppData.GetDisplayInformation(erfDisplayEditSetHistoryFooterDaysTotal) & _
                            gk_sHTMLNonBreakSpace & _
                            gk_sHTMLNonBreakSpace, typCellStyle, m_oFWHTML.GetFontStyle(FONTNormalStyle)
                           
    typCellStyle.eALIGN = HALLeft
    ShowDetailsLine _
        m_oFWHTML.Bold(CStr(iTotalDays)), typCellStyle, m_oFWHTML.GetFontStyle(FONTNormalStyle)
        
    m_oFWTable.EndRow
        
    sHTML = sHTML & m_oFWTable.GetHTML
    
    '   d57 crm 12Jan2000   Add a done button
    sHTML = sHTML & "<P ALIGN=Right>" & m_oFWHTML.ImageButton(gk_sImagesPath & gk_sImageDone, "", _
            gk_sLinkDisplayEditSetsForOneUser, gk_sMainFrame, GetToolTip(ETOOLTIP.ttButtonDone))
     
    'Add the help button
    sHTML = sHTML & HTMLToAddHelpButton(gk_sHelpLinkToEditSetHistory)
    
    m_aspResponse.Write sHTML
    
    Set oFWRecordlist = Nothing
    
    m_objObjectContext.SetComplete
    
    Exit Sub

ErrorHandler:
    m_objObjectContext.SetAbort
    ShowError sSource
End Sub





Private Sub DisplayEditsetDetailsHeader(Optional t_eUsage As EEditSetActionType = esaNormal)
'
'
'Description
'
'Displays the details of the editset
'
'Amendment history
'   d8      ndk     01Jul1999       Created
'   d11     ksk     08Jul1999       Moved code from DisplayEditSetDetails into this method
'   d19     ndk     16Jul1999       now displays actual counts of the entries
'   d25     ndk     11Aug1999       Abandon Button now calls the appropirate asp page
'   d34     ndk     28Oct1999       Button (on editset details screen) enabling and disabling made dependent on calls
'                                   to security/permission object
'   d35     ndk     04Nov1999        Minor change
'   d41     ndk     22Nov1999       The filter was displayed twice, now shown once
'   d45     ndk     23Nov1999       Disabling 'next stage' button if current stage is pre-release
'   d53     ndk     27Dec1999       Changed to VI code use customizable styles
'   d54     ndk     28Dec1999       Added tooltips
'   d62     jcc     09Feb2000       changed tooltips for arrows on DisplayEditSet
'   d65     tmj     22Feb2000       Changed link relating to adding an entry so now goes to an intermediate choose search page
'                                   rather than directly to the basic search
'
'   d73     daj     09May2000       Added some decoding of special characters to query strings
'   d87     alm     19Oct2001       Added a link to move multiple pieces. Changed header display depending
'                                   on new optional param t_eUsage
'   d89     alm     23Oct2001       Removed references to the sort order
'   d90     alm     23Oct2001       Checks that an editsetID is passed and moved the loading of the editset
'                                   details to DisplayEditSetDetails. Also supports the new editset
'                                   type - delete
On Error GoTo ErrorHandler

   
    Dim sSource As String
    Dim sHTML As String, sParamString   As String, sDisplayString As String
    
    Dim sEditSetName As String, lEditSetID As Long, sContributor1 As String
    Dim sEditSetCreationDate As String, iDaysInStage As Integer, lEditSetStageID As Long, sEditSetStageName As String
    Dim iTotalDepartMents As Integer, iTotalDivisions As Integer, iTotalSeries As Integer
    Dim iTotalSubSeries As Integer, iTotalSubSubSeries As Integer, iTotalPices As Integer, iTotalItems As Integer
    Dim sCurrentFilter As String
    
    Dim typCellStyle As CellStyle
    Dim typTableStyle As TableStyle
    
    Dim oFWRecordlist As FWRecordList
    Dim oFWEditSetEntries As FWRecordList
    
    Dim sEditsetParam As String
    
    sSource = Me.Name & ".DisplayEditSetDetails"
    
    ' Make sure that the editset ID was passes to this page
    If m_aspRequest.QueryString(gk_sParamEditSetID) = "" Then
    
        RaiseError sSource, EPROCATINTERNALERROR.peEditSetIDInvalue, _
                        "Editset ID not found"
                        
    End If
    
    lEditSetID = CLng(m_aspRequest(gk_sParamEditSetID))
    
    If lEditSetID <= 0 Then
    
        RaiseError sSource, EPROCATINTERNALERROR.peEditSetIDInvalue, _
                            "Editset ID not found"
                            
    End If
    
    'Create instances
    Set m_oFWTable = New FWTable
        
    typTableStyle = m_oFWHTML.GetTableStyle(TABPageHeadingStyle)
    typTableStyle.sWidth = "100%"
    typTableStyle.lCELLSPACING = 0
    m_oFWTable.TableStyle = typTableStyle
    
    'd87 What header we output depends on the usage
    Select Case t_eUsage
    
        Case esaNormal:
            
            sEditSetName = m_oBOEditSet.EditSetName
            sContributor1 = m_oBOEditSet.Contributor1UserName
            sEditSetCreationDate = m_oBOEditSet.CreationDate
            iDaysInStage = m_oBOEditSet.DaysInCurrentStage
            lEditSetStageID = m_oBOEditSet.CurrentStageID
            sEditSetStageName = m_oBOEditSet.CurrentStageName
            
            'The count of entries
            
            iTotalDepartMents = m_oBOEditSet.TotalDepartments
            iTotalDivisions = m_oBOEditSet.TotalDivisions
            iTotalSeries = m_oBOEditSet.TotalSeries
            iTotalSubSeries = m_oBOEditSet.TotalSubSeries
            iTotalSubSubSeries = m_oBOEditSet.TotalSubSubSeries
            iTotalPices = m_oBOEditSet.TotalPieces
            iTotalItems = m_oBOEditSet.TotalItems
        
            'Line displaying the properties of the editset
            '   d33     ksk     21Oct1999   Added space before the stage for
            sDisplayString = m_oFWAppData.GetDisplayInformation( _
                            EEditorialResourceFile.erfDisplayEditSetDetailsEditSet) & _
                            " " & m_oFWHTML.Bold(sEditSetName) & ", "             'Edit Set
            sDisplayString = sDisplayString & m_oFWAppData.GetDisplayInformation( _
                            EEditorialResourceFile.erfDisplayEditSetDetailsOwnedBy) & _
                            " " & m_oFWHTML.Bold(sContributor1) & ", "      'owned by
            sDisplayString = sDisplayString & m_oFWAppData.GetDisplayInformation( _
                            EEditorialResourceFile.erfDisplayEditSetDetailsCreatedOn) & _
                            " " & m_oFWHTML.Bold(VariantDateAsDDMMMYYYY(sEditSetCreationDate)) & " , "       'created On
            sDisplayString = sDisplayString & m_oFWAppData.GetDisplayInformation( _
                            EEditorialResourceFile.erfDisplayEditSetDetailsHasBeenIn) & _
                            " " & m_oFWHTML.Bold(sEditSetStageName) & " "    ''has been in
            sDisplayString = sDisplayString & m_oFWAppData.GetDisplayInformation( _
                            EEditorialResourceFile.erfDisplayEditSetDetailsStageFor) & " " 'Stage for
            sDisplayString = sDisplayString & m_oFWHTML.Bold(CStr(iDaysInStage)) & _
                            " " & m_oFWAppData.GetDisplayInformation( _
                            EEditorialResourceFile.erfDisplayEditSetDetailsDays) 'days
            
            ShowNormalStyleText sDisplayString, "", 4
        
            m_oFWTable.EndRow
            
            m_oFWTable.AddBlankRow
            
            'Line displaying the count of the entries in the editset
        
            sDisplayString = iTotalDepartMents & " " & m_oFWAppData.GetDisplayInformation( _
                EEditorialResourceFile.erfDisplayEditSetDetailsDepartments) & ", "           'departments
        
            sDisplayString = sDisplayString & iTotalDivisions & " " & m_oFWAppData.GetDisplayInformation( _
                EEditorialResourceFile.erfDisplayEditSetDetailsDivisions) & ", "     'divisions
        
            sDisplayString = sDisplayString & iTotalSeries & " " & m_oFWAppData.GetDisplayInformation( _
                EEditorialResourceFile.erfDisplayEditSetDetailsSeries) & ", "   'series
            
            sDisplayString = sDisplayString & iTotalSubSeries & " " & " " & m_oFWAppData.GetDisplayInformation( _
                EEditorialResourceFile.erfDisplayEditSetDetailsSubSeries) & ", " ' sub series
            
            sDisplayString = sDisplayString & iTotalSubSubSeries & "  " & " " & m_oFWAppData.GetDisplayInformation( _
                EEditorialResourceFile.erfDisplayEditSetDetailsSubSubSeries) & ", "   '  sub-sub-series
            
            sDisplayString = sDisplayString & iTotalPices & "  " & " " & m_oFWAppData.GetDisplayInformation( _
                EEditorialResourceFile.erfDisplayEditSetDetailsPieces) & ", "       'pieces
            
            sDisplayString = sDisplayString & iTotalItems & " " & " " & m_oFWAppData.GetDisplayInformation( _
                EEditorialResourceFile.erfDisplayEditSetDetailsItems)         'items
        
            m_oFWTable.StartRow
            ShowNormalStyleText _
                m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetDetailsLeftHeaderContents) & " " & m_oFWHTML.Bold(sDisplayString), _
                "", 4 'Contents
            m_oFWTable.EndRow
            
            m_oFWTable.AddBlankRow
            
            typCellStyle.lCOLSPAN = 1
            typCellStyle.eStyle = CELLGeneralContentStyle
            
            m_oFWTable.StartRow
          
            'Bulk Update
            typCellStyle.eALIGN = HALLeft
            m_oBOEditorialPermissions.UserID = m_lLoggedUserID
            m_oBOEditorialPermissions.EditsetID = lEditSetID
            If m_oBOEditorialPermissions.Check(EUserCanBulkUpdate) = ESecurityCheckResult.ESRSecurityCheckSuccessful Then
                m_oFWTable.AddCellAsHTML _
                    m_oFWHTML.Hyperlink(gk_sLinkDisplayBulkUpdateFields _
                    & "?" & gk_sParamEditSetID & "=" & lEditSetID, _
                    m_oFWAppData.GetDisplayInformation( _
                        EEditorialResourceFile.erfDisplayEditSetDetailsHyperLinkBulkUpdate), gk_sMainFrame)
            Else
                m_oFWTable.AddCellAsHTML _
                    m_oFWHTML.Underline(m_oFWAppData.GetDisplayInformation( _
                        EEditorialResourceFile.erfDisplayEditSetDetailsHyperLinkBulkUpdate))
            End If
            
            'd87 Move pieces
            If m_oBOEditorialPermissions.Check(EUserCanEditCatalogueEntryOnEditSet) = ESecurityCheckResult.ESRSecurityCheckSuccessful Then
                m_oFWTable.AddCellAsHTML _
                    m_oFWHTML.Hyperlink(gk_sLinkDisplayChooseParentOfPieces _
                    & "?" & gk_sParamEditSetID & "=" & lEditSetID, _
                    m_oFWAppData.GetDisplayInformation( _
                        EEditorialResourceFile.erfDisplayEditSetDetailsHyperLinkMovePieces), gk_sMainFrame)
            Else
                m_oFWTable.AddCellAsHTML _
                    m_oFWHTML.Underline(m_oFWAppData.GetDisplayInformation( _
                        EEditorialResourceFile.erfDisplayEditSetDetailsHyperLinkMovePieces))
            End If
            
            ' Add input box
            typCellStyle.eALIGN = HALRight
            sEditsetParam = "?" & gk_sParamEditSetID & "=" & lEditSetID
            m_oFWTable.AddCellAsHTML m_oFWHTML.FormStart(gk_sLinkDisplayEditSetDetails & sEditsetParam) & _
                mk_sJumpToReferenceText & " " & _
                m_oFWHTML.InputField(ITText, gk_sNameCatalogueReference, 15) & _
                m_oFWHTML.SubmitImage(gk_sIMAGESmallGo, "", "") & _
                m_oFWHTML.EndTag(gk_sHTMLForm), typCellStyle
           
            m_oFWTable.EndRow
            
            typCellStyle.lCOLSPAN = 1
            
            m_oFWTable.StartRow
            
            'Get the contents of the table
            sHTML = sHTML & m_oFWHTML.Paragraph(m_oFWTable.GetHTML)
        
            m_oFWTable.Reset
            
            'To display the buttons
            typTableStyle = m_oFWHTML.GetTableStyle(TABContentStyle)
            typTableStyle.lCELLSPACING = 0
            m_oFWTable.TableStyle = typTableStyle
            
            'Filter
            typCellStyle = m_oFWHTML.GetCellStyle(CELLNormalStyle)
            
            m_oBOEditorialPermissions.UserID = m_lLoggedUserID
            m_oBOEditorialPermissions.EditsetID = lEditSetID
            If m_oBOEditorialPermissions.Check(EUserCanSetFilterOnEditSet) = _
                        ESecurityCheckResult.ESRSecurityCheckSuccessful Then
        
                m_oFWTable.AddCellAsHTML _
                    m_oFWHTML.ImageButton(gk_sImagesPath & gk_sImageFilter, _
                     mk_sNameFilter, _
                        gk_sDisplayFilterForm & _
                        "?" & gk_sParamEditSetID & "=" & lEditSetID, _
                        gk_sMainFrame, GetToolTip(ETOOLTIP.ttFilter)), typCellStyle
            Else
                m_oFWTable.AddCellAsHTML _
                    m_oFWHTML.ImageButton(gk_sImagesPath & gk_sImageFilter, "", "", gk_sMainFrame), typCellStyle
            End If
            
            'Add catalogue to editset
            m_oBOEditorialPermissions.UserID = m_lLoggedUserID
            m_oBOEditorialPermissions.EditsetID = lEditSetID
            If m_oBOEditorialPermissions.Check(EUserCanAddCatalogueEntryToEditSet) = _
                        ESecurityCheckResult.ESRSecurityCheckSuccessful Then
        
                m_oFWTable.AddCellAsHTML _
                    m_oFWHTML.ImageButton(gk_sImagesPath & gk_sImageAddEntry, _
                    "", _
                        gk_sLinkChooseSearchMethod & _
                        "?" & gk_sParamEditSetID & "=" & lEditSetID, _
                        gk_sMainFrame, GetToolTip(ETOOLTIP.ttAddEntry)), typCellStyle
            Else
                m_oFWTable.AddCellAsHTML _
                    m_oFWHTML.ImageButton(gk_sImagesPath & gk_sImageAddEntry), typCellStyle
            End If
            
            'Create new catalogue entry
            m_oBOEditorialPermissions.UserID = m_lLoggedUserID
            m_oBOEditorialPermissions.EditsetID = lEditSetID
            If m_oBOEditorialPermissions.Check(EUserCanCreateCatalogueEntryOnEditSet) = _
                        ESecurityCheckResult.ESRSecurityCheckSuccessful Then
        
                m_oFWTable.AddCellAsHTML _
                    m_oFWHTML.ImageButton(gk_sImagesPath & gk_sImageCreateEntry, _
                    "", _
                        gk_sLinkDisplayCreateNewCatalogueMenu & _
                        "?" & gk_sParamEditSetID & "=" & lEditSetID, _
                        gk_sMainFrame, GetToolTip(ETOOLTIP.ttCreateEntry)), typCellStyle
            Else
                m_oFWTable.AddCellAsHTML m_oFWHTML.ImageButton(gk_sImagesPath & gk_sImageCreateEntry), typCellStyle
            End If
            
            'Move editset to next stage
            If Not lEditSetStageID = gk_lStagePreReleaseID Then
                m_oBOEditorialPermissions.UserID = m_lLoggedUserID
                m_oBOEditorialPermissions.EditsetID = lEditSetID
                If m_oBOEditorialPermissions.Check(EUserCanMoveEditSetFromCurrentStage) = _
                            ESecurityCheckResult.ESRSecurityCheckSuccessful Then
            
                    m_oFWTable.AddCellAsHTML _
                        m_oFWHTML.ImageButton(gk_sImagesPath & gk_sImageNextStage, _
                        "", _
                        gk_sLinkDisplayNextStagesForEditSet & _
                            "?" & gk_sParamEditSetID & "=" & lEditSetID & _
                            "&" & ksParamEditSetName & "=" & EncodeQueryStringValue(sEditSetName) & _
                            "&" & ksParamCurrentStageID & "=" & lEditSetStageID & _
                            "&" & ksParamCurrentStageName & "=" & sEditSetStageName, _
                        gk_sMainFrame, GetToolTip(ETOOLTIP.ttNextStage)), typCellStyle
                Else
                    m_oFWTable.AddCellAsHTML m_oFWHTML.ImageButton(gk_sImagesPath & gk_sImageNextStage), typCellStyle
                End If
            End If
            
            'Move editset to previous stage
            If Not lEditSetStageID = gk_lStageContributionID Then
                
                m_oBOEditorialPermissions.UserID = m_lLoggedUserID
                m_oBOEditorialPermissions.EditsetID = lEditSetID
                If m_oBOEditorialPermissions.Check(EUserCanMoveEditSetFromCurrentStage) = _
                        ESecurityCheckResult.ESRSecurityCheckSuccessful Then
        
                    'When in contribution stage , send back button should not be displayed
                    m_oFWTable.AddCellAsHTML _
                        m_oFWHTML.ImageButton(gk_sImagesPath & gk_sImageSendBack, _
                        "", _
                        gk_sLinkDisplayPreviousStagesForEditSet & _
                            "?" & gk_sParamEditSetID & "=" & lEditSetID & _
                            "&" & ksParamEditSetName & "=" & EncodeQueryStringValue(sEditSetName) & _
                            "&" & ksParamCurrentStageID & "=" & lEditSetStageID & _
                            "&" & ksParamCurrentStageName & "=" & sEditSetStageName, _
                        gk_sMainFrame, GetToolTip(ETOOLTIP.ttPreviousStage)), typCellStyle
                Else
                    m_oFWTable.AddCellAsHTML _
                        m_oFWHTML.ImageButton(gk_sImagesPath & gk_sImageSendBack)
                End If
            End If
        
            'Display and edit users and roles on the editset
            m_oBOEditorialPermissions.UserID = m_lLoggedUserID
            m_oBOEditorialPermissions.EditsetID = lEditSetID
            If m_oBOEditorialPermissions.Check(EUserCanEditRolesOnEditSet) = _
                        ESecurityCheckResult.ESRSecurityCheckSuccessful Then
        
                m_oFWTable.AddCellAsHTML _
                    m_oFWHTML.ImageButton(gk_sImagesPath & gk_sImageEditRoles, _
                    "", _
                    gk_sLinkDisplayUsersAndRolesForEditSet & _
                        "?" & gk_sParamEditSetID & "=" & lEditSetID & _
                        "&" & ksParamEditSetName & "=" & EncodeQueryStringValue(sEditSetName), _
                    gk_sMainFrame, GetToolTip(ETOOLTIP.ttEditRoles)), typCellStyle
            Else
                m_oFWTable.AddCellAsHTML m_oFWHTML.ImageButton(gk_sImagesPath & gk_sImageEditRoles), typCellStyle
            End If
            
            'Abandon the editset
            '   d36     Changed to Redirect To AbandoneConformation page
            m_oBOEditorialPermissions.UserID = m_lLoggedUserID
            m_oBOEditorialPermissions.EditsetID = lEditSetID
            If m_oBOEditorialPermissions.Check(EUserCanAbandonEditSet) = _
                        ESecurityCheckResult.ESRSecurityCheckSuccessful Then
        
                m_oFWTable.AddCellAsHTML _
                        m_oFWHTML.ImageButton(gk_sImagesPath & gk_sImageAbandon, _
                            "", _
                            gk_sLinkDisplayConformationForAbandon & _
                            "?" & gk_sParamEditSetID & "=" & lEditSetID & _
                            "&" & ksParamCurrentStageID & "=" & lEditSetStageID & _
                            "&" & ksParamStageEditSetIsToBeMoved & "=" & gk_lStageAbandonID, _
                            gk_sMainFrame, GetToolTip(ETOOLTIP.ttAbandonEditSet)), typCellStyle
            Else
                m_oFWTable.AddCellAsHTML m_oFWHTML.ImageButton(gk_sImagesPath & gk_sImageAbandon), typCellStyle
            End If
        
            m_oFWTable.EndRow
            
            Set oFWRecordlist = Nothing
            Set oFWEditSetEntries = Nothing
            
        ' d90 - Added
        Case esaDelete:
            
            ' Let the user know this is a delete editset
            sDisplayString = m_oFWAppData.GetDisplayInformation( _
                             EEditorialResourceFile.erfDisplayDeleteEditSetHeader)
            ShowNormalStyleText m_oFWHTML.Bold(sDisplayString), "", 4
            
            m_oFWTable.EndRow
            
            m_oFWTable.AddBlankRow
            
            sEditSetName = m_oBOEditSet.EditSetName
            sContributor1 = m_oBOEditSet.Contributor1UserName
            sEditSetCreationDate = m_oBOEditSet.CreationDate
            iDaysInStage = m_oBOEditSet.DaysInCurrentStage
            lEditSetStageID = m_oBOEditSet.CurrentStageID
            sEditSetStageName = m_oBOEditSet.CurrentStageName
            
            'The count of entries
            
            iTotalDepartMents = m_oBOEditSet.TotalDepartments
            iTotalDivisions = m_oBOEditSet.TotalDivisions
            iTotalSeries = m_oBOEditSet.TotalSeries
            iTotalSubSeries = m_oBOEditSet.TotalSubSeries
            iTotalSubSubSeries = m_oBOEditSet.TotalSubSubSeries
            iTotalPices = m_oBOEditSet.TotalPieces
            iTotalItems = m_oBOEditSet.TotalItems
        
            'Line displaying the properties of the editset
            '   d33     ksk     21Oct1999   Added space before the stage for
            sDisplayString = m_oFWAppData.GetDisplayInformation( _
                            EEditorialResourceFile.erfDisplayEditSetDetailsEditSet) & _
                            " " & m_oFWHTML.Bold(sEditSetName) & ", "             'Edit Set
            sDisplayString = sDisplayString & m_oFWAppData.GetDisplayInformation( _
                            EEditorialResourceFile.erfDisplayEditSetDetailsOwnedBy) & _
                            " " & m_oFWHTML.Bold(sContributor1) & ", "      'owned by
            sDisplayString = sDisplayString & m_oFWAppData.GetDisplayInformation( _
                            EEditorialResourceFile.erfDisplayEditSetDetailsCreatedOn) & _
                            " " & m_oFWHTML.Bold(VariantDateAsDDMMMYYYY(sEditSetCreationDate)) & " , "       'created On
            sDisplayString = sDisplayString & m_oFWAppData.GetDisplayInformation( _
                            EEditorialResourceFile.erfDisplayEditSetDetailsHasBeenIn) & _
                            " " & m_oFWHTML.Bold(sEditSetStageName) & " "    ''has been in
            sDisplayString = sDisplayString & m_oFWAppData.GetDisplayInformation( _
                            EEditorialResourceFile.erfDisplayEditSetDetailsStageFor) & " " 'Stage for
            sDisplayString = sDisplayString & m_oFWHTML.Bold(CStr(iDaysInStage)) & _
                            " " & m_oFWAppData.GetDisplayInformation( _
                            EEditorialResourceFile.erfDisplayEditSetDetailsDays) 'days
            
            ShowNormalStyleText sDisplayString, "", 4
        
            m_oFWTable.EndRow
            
            m_oFWTable.AddBlankRow
            
            'Line displaying the count of the entries in the editset
        
            sDisplayString = iTotalDepartMents & " " & m_oFWAppData.GetDisplayInformation( _
                EEditorialResourceFile.erfDisplayEditSetDetailsDepartments) & ", "           'departments
        
            sDisplayString = sDisplayString & iTotalDivisions & " " & m_oFWAppData.GetDisplayInformation( _
                EEditorialResourceFile.erfDisplayEditSetDetailsDivisions) & ", "     'divisions
        
            sDisplayString = sDisplayString & iTotalSeries & " " & m_oFWAppData.GetDisplayInformation( _
                EEditorialResourceFile.erfDisplayEditSetDetailsSeries) & ", "   'series
            
            sDisplayString = sDisplayString & iTotalSubSeries & " " & " " & m_oFWAppData.GetDisplayInformation( _
                EEditorialResourceFile.erfDisplayEditSetDetailsSubSeries) & ", " ' sub series
            
            sDisplayString = sDisplayString & iTotalSubSubSeries & "  " & " " & m_oFWAppData.GetDisplayInformation( _
                EEditorialResourceFile.erfDisplayEditSetDetailsSubSubSeries) & ", "   '  sub-sub-series
            
            sDisplayString = sDisplayString & iTotalPices & "  " & " " & m_oFWAppData.GetDisplayInformation( _
                EEditorialResourceFile.erfDisplayEditSetDetailsPieces) & ", "       'pieces
            
            sDisplayString = sDisplayString & iTotalItems & " " & " " & m_oFWAppData.GetDisplayInformation( _
                EEditorialResourceFile.erfDisplayEditSetDetailsItems)         'items
        
            m_oFWTable.StartRow
            ShowNormalStyleText _
                m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayEditSetDetailsLeftHeaderContents) & " " & m_oFWHTML.Bold(sDisplayString), _
                "", 4 'Contents
            m_oFWTable.EndRow
            
            m_oFWTable.AddBlankRow
            
            typCellStyle.lCOLSPAN = 1
            typCellStyle.eStyle = CELLGeneralContentStyle
            
            m_oFWTable.StartRow
            
            ' Add input box
            typCellStyle.lCOLSPAN = 5
            typCellStyle.eALIGN = HALRight
            sEditsetParam = "?" & gk_sParamEditSetID & "=" & lEditSetID
            m_oFWTable.AddCellAsHTML m_oFWHTML.FormStart(gk_sLinkDisplayEditSetDetails & sEditsetParam) & _
                mk_sJumpToReferenceText & " " & _
                m_oFWHTML.InputField(ITText, gk_sNameCatalogueReference, 15) & _
                m_oFWHTML.SubmitImage(gk_sIMAGESmallGo, "", "") & _
                m_oFWHTML.EndTag(gk_sHTMLForm), typCellStyle
            m_oFWTable.EndRow
            
            typCellStyle.lCOLSPAN = 1
            
            m_oFWTable.StartRow
            
            'Get the contents of the table
            sHTML = sHTML & m_oFWHTML.Paragraph(m_oFWTable.GetHTML)
        
            m_oFWTable.Reset
            
            'To display the buttons
            typTableStyle = m_oFWHTML.GetTableStyle(TABContentStyle)
            typTableStyle.lCELLSPACING = 0
            m_oFWTable.TableStyle = typTableStyle
            
            'Filter
            typCellStyle = m_oFWHTML.GetCellStyle(CELLNormalStyle)
            
            m_oBOEditorialPermissions.UserID = m_lLoggedUserID
            m_oBOEditorialPermissions.EditsetID = lEditSetID
            If m_oBOEditorialPermissions.Check(EUserCanSetFilterOnEditSet) = _
                        ESecurityCheckResult.ESRSecurityCheckSuccessful Then
        
                m_oFWTable.AddCellAsHTML _
                    m_oFWHTML.ImageButton(gk_sImagesPath & gk_sImageFilter, _
                     mk_sNameFilter, _
                        gk_sDisplayFilterForm & _
                        "?" & gk_sParamEditSetID & "=" & lEditSetID, _
                        gk_sMainFrame, GetToolTip(ETOOLTIP.ttFilter)), typCellStyle
            Else
                m_oFWTable.AddCellAsHTML _
                    m_oFWHTML.ImageButton(gk_sImagesPath & gk_sImageFilter, "", "", gk_sMainFrame), typCellStyle
            End If
            
            'Add catalogue to editset
            m_oBOEditorialPermissions.UserID = m_lLoggedUserID
            m_oBOEditorialPermissions.EditsetID = lEditSetID
            If m_oBOEditorialPermissions.Check(EUserCanAddCatalogueEntryToEditSet) = _
                        ESecurityCheckResult.ESRSecurityCheckSuccessful Then
        
                m_oFWTable.AddCellAsHTML _
                    m_oFWHTML.ImageButton(gk_sImagesPath & gk_sImageAddEntry, _
                    "", _
                        gk_sLinkChooseSearchMethod & _
                        "?" & gk_sParamEditSetID & "=" & lEditSetID, _
                        gk_sMainFrame, GetToolTip(ETOOLTIP.ttAddEntry)), typCellStyle
            Else
                m_oFWTable.AddCellAsHTML _
                    m_oFWHTML.ImageButton(gk_sImagesPath & gk_sImageAddEntry), typCellStyle
            End If
            
            ' Move editset to next stage
            If Not lEditSetStageID = gk_lStagePreReleaseID Then
                m_oBOEditorialPermissions.UserID = m_lLoggedUserID
                m_oBOEditorialPermissions.EditsetID = lEditSetID
                If m_oBOEditorialPermissions.Check(EUserCanMoveEditSetFromCurrentStage) = _
                            ESecurityCheckResult.ESRSecurityCheckSuccessful Then
            
                    m_oFWTable.AddCellAsHTML _
                        m_oFWHTML.ImageButton(gk_sImagesPath & gk_sImageNextStage, _
                        "", _
                        gk_sLinkDisplayNextStagesForEditSet & _
                            "?" & gk_sParamEditSetID & "=" & lEditSetID & _
                            "&" & ksParamEditSetName & "=" & EncodeQueryStringValue(sEditSetName) & _
                            "&" & ksParamCurrentStageID & "=" & lEditSetStageID & _
                            "&" & ksParamCurrentStageName & "=" & sEditSetStageName, _
                        gk_sMainFrame, GetToolTip(ETOOLTIP.ttNextStage)), typCellStyle
                Else
                    m_oFWTable.AddCellAsHTML m_oFWHTML.ImageButton(gk_sImagesPath & gk_sImageNextStage), typCellStyle
                End If
            End If
            
            'Move editset to previous stage
            If Not lEditSetStageID = gk_lStageContributionID Then
                
                m_oBOEditorialPermissions.UserID = m_lLoggedUserID
                m_oBOEditorialPermissions.EditsetID = lEditSetID
                If m_oBOEditorialPermissions.Check(EUserCanMoveEditSetFromCurrentStage) = _
                        ESecurityCheckResult.ESRSecurityCheckSuccessful Then
        
                    'When in contribution stage , send back button should not be displayed
                    m_oFWTable.AddCellAsHTML _
                        m_oFWHTML.ImageButton(gk_sImagesPath & gk_sImageSendBack, _
                        "", _
                        gk_sLinkDisplayPreviousStagesForEditSet & _
                            "?" & gk_sParamEditSetID & "=" & lEditSetID & _
                            "&" & ksParamEditSetName & "=" & EncodeQueryStringValue(sEditSetName) & _
                            "&" & ksParamCurrentStageID & "=" & lEditSetStageID & _
                            "&" & ksParamCurrentStageName & "=" & sEditSetStageName, _
                        gk_sMainFrame, GetToolTip(ETOOLTIP.ttPreviousStage)), typCellStyle
                Else
                    m_oFWTable.AddCellAsHTML _
                        m_oFWHTML.ImageButton(gk_sImagesPath & gk_sImageSendBack)
                End If
            End If
        
            ' Display and edit users and roles on the editset
            m_oBOEditorialPermissions.UserID = m_lLoggedUserID
            m_oBOEditorialPermissions.EditsetID = lEditSetID
            If m_oBOEditorialPermissions.Check(EUserCanEditRolesOnEditSet) = _
                        ESecurityCheckResult.ESRSecurityCheckSuccessful Then
        
                m_oFWTable.AddCellAsHTML _
                    m_oFWHTML.ImageButton(gk_sImagesPath & gk_sImageEditRoles, _
                    "", _
                    gk_sLinkDisplayUsersAndRolesForEditSet & _
                        "?" & gk_sParamEditSetID & "=" & lEditSetID & _
                        "&" & ksParamEditSetName & "=" & EncodeQueryStringValue(sEditSetName), _
                    gk_sMainFrame, GetToolTip(ETOOLTIP.ttEditRoles)), typCellStyle
            Else
                m_oFWTable.AddCellAsHTML m_oFWHTML.ImageButton(gk_sImagesPath & gk_sImageEditRoles), typCellStyle
            End If
            
            ' Abandon the editset
            m_oBOEditorialPermissions.UserID = m_lLoggedUserID
            m_oBOEditorialPermissions.EditsetID = lEditSetID
            If m_oBOEditorialPermissions.Check(EUserCanAbandonEditSet) = _
                        ESecurityCheckResult.ESRSecurityCheckSuccessful Then
        
                m_oFWTable.AddCellAsHTML _
                        m_oFWHTML.ImageButton(gk_sImagesPath & gk_sImageAbandon, _
                            "", _
                            gk_sLinkDisplayConformationForAbandon & _
                            "?" & gk_sParamEditSetID & "=" & lEditSetID & _
                            "&" & ksParamCurrentStageID & "=" & lEditSetStageID & _
                            "&" & ksParamStageEditSetIsToBeMoved & "=" & gk_lStageAbandonID, _
                            gk_sMainFrame, GetToolTip(ETOOLTIP.ttAbandonEditSet)), typCellStyle
            Else
                m_oFWTable.AddCellAsHTML m_oFWHTML.ImageButton(gk_sImagesPath & gk_sImageAbandon), typCellStyle
            End If
        
            m_oFWTable.EndRow
            
            Set oFWRecordlist = Nothing
            Set oFWEditSetEntries = Nothing
                        
        Case esaChooseParentOfPieces:
            
            ' Heading
            m_oFWTable.StartRow
            ShowNormalStyleText "Please select the catalogue entry <b>from</b> which pieces should be moved", "", 4
            m_oFWTable.EndRow
            
            ' Add input box
            typCellStyle.lCOLSPAN = 5
            typCellStyle.eALIGN = HALRight
            typCellStyle.eStyle = CELLWorkInProgressStyle
            m_oFWTable.AddCellAsHTML "&nbsp;", typCellStyle
            sEditsetParam = "?" & gk_sParamEditSetID & "=" & lEditSetID
            m_oFWTable.AddCellAsHTML m_oFWHTML.FormStart(gk_sLinkDisplayChooseParentOfPieces & sEditsetParam) & _
                mk_sJumpToReferenceText & " " & _
                m_oFWHTML.InputField(ITText, gk_sNameCatalogueReference, 15) & _
                m_oFWHTML.SubmitImage(gk_sIMAGESmallGo, "", "") & _
                m_oFWHTML.EndTag(gk_sHTMLForm), typCellStyle
            m_oFWTable.EndRow
            
            ' Cancel button
            m_oFWTable.StartRow
            typCellStyle.eALIGN = HALLeft
            sEditsetParam = "?" & gk_sParamEditSetID & "=" & lEditSetID
            m_oFWTable.AddCellAsHTML m_oFWHTML.ImageButton(gk_sIMAGEURLCancelButton, _
                                       t_sURL:=gk_sLinkDisplayEditSetDetails & sEditsetParam)
            m_oFWTable.EndRow
            
        Case esaChoosePieces:
            
            ' Heading
            m_oFWTable.StartRow
            ShowNormalStyleText "Please select the pieces that should be moved to a new parent catalogue entry", "", 4
            m_oFWTable.EndRow
            
            ' Cancel button
            m_oFWTable.AddBlankRow
            m_oFWTable.StartRow
            typCellStyle.eALIGN = HALLeft
            sEditsetParam = "?" & gk_sParamEditSetID & "=" & lEditSetID
            m_oFWTable.AddCellAsHTML m_oFWHTML.ImageButton(gk_sIMAGEURLCancelButton, _
                                       t_sURL:=gk_sLinkDisplayEditSetDetails & sEditsetParam)
            m_oFWTable.EndRow
    
        Case esaChooseParentForPieces:
        
            ' Heading
            m_oFWTable.StartRow
            ShowNormalStyleText "Please select the catalogue entry <b>to</b> which pieces should be moved", "", 4
            m_oFWTable.EndRow
            
            ' Cancel button
            m_oFWTable.AddBlankRow
            m_oFWTable.StartRow
            typCellStyle.eALIGN = HALLeft
            sEditsetParam = "?" & gk_sParamEditSetID & "=" & lEditSetID
            m_oFWTable.AddCellAsHTML m_oFWHTML.ImageButton(gk_sIMAGEURLCancelButton, _
                                       t_sURL:=gk_sLinkDisplayEditSetDetails & sEditsetParam)
            m_oFWTable.EndRow
    
    End Select
    
    'Get the contents of the table
    sHTML = sHTML & m_oFWHTML.Paragraph(m_oFWTable.GetHTML)
            
    'Display the html
    m_aspResponse.Write sHTML
    
    Exit Sub

ErrorHandler:
    m_objObjectContext.SetAbort
    ShowError sSource

End Sub
    
    



Public Sub DisplayConformationForAbandon()
'method to display conformation screen for the abandon functionality
'   d36     ksk     created
'   d47     ndk     08Dec1999       Added help buttons
'   d68     daj     08Mar2000   Fixed display of abandon edit set confirmation to be centred
On Error GoTo ErrorHandler
    
    Dim sHTML As String, sSource As String, sEditsetParam As String
    Dim sOKLink  As String, sCancelLink  As String
    Dim lEditSetID  As Long
    Dim eResult As ESecurityCheckResult
    Dim typTableStyle As TableStyle
    Dim typCellStyle As CellStyle
    
    sSource = Me.Name & ".DisplayConformationForAbandon()"
    
    Set m_oFWTable = New FWTable
    
    'Table style for heading
    typTableStyle = m_oFWHTML.GetTableStyle(TABContentStyle)
    typTableStyle.lCELLSPACING = 2
    typTableStyle.lCELLPADDING = 0
    typTableStyle.eALIGN = HALCenter
    typCellStyle = m_oFWHTML.GetCellStyle(CELLNormalStyle)
    typCellStyle.eALIGN = HALCenter
    
    m_oFWTable.TableStyle = typTableStyle

    'Get editset id from the query string
    lEditSetID = m_aspRequest.QueryString(gk_sParamEditSetID)
    If lEditSetID <= 0 Then
        RaiseError sSource, EPROCATINTERNALERROR.peEditSetIDInvalue, _
                "Editset ID not found"
    End If
    sEditsetParam = "?" & gk_sParamEditSetID & "=" & lEditSetID
    
    m_oBOEditorialPermissions.EditsetID = lEditSetID
    m_oBOEditorialPermissions.UserID = m_lLoggedUserID
    
    eResult = m_oBOEditorialPermissions.Check(EUserCanAbandonEditSet)
    If eResult <> ESecurityCheckResult.ESRSecurityCheckSuccessful Then
        RaiseError sSource, _
               GetErrorEnumForSecurityCheckResult(eResult), _
               GetDescriptionForSecurityCheckError( _
                        GetErrorEnumForSecurityCheckResult(eResult))
    End If
    
    'Create instance of the editset
    m_oBOEditSet.EditsetID = lEditSetID
    m_oBOEditSet.Load
     
    sOKLink = _
        gk_sLinkAbandonEditset & sEditsetParam
        
    sCancelLink = _
        gk_sLinkDisplayEditSetDetails & sEditsetParam

    ' d68 - Fixed to user resource file and center text
    sHTML = m_oFWHTML.Bold(m_oFWHTML.Font( _
                m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfConfirmAbandonEditSet) & _
                " '" & m_oBOEditSet.EditSetName & "' ?", COLOURRed, 15))

    m_oFWTable.StartRow
    m_oFWTable.AddCellAsHTML sHTML, typCellStyle
    m_oFWTable.EndRow
                            
    sHTML = m_oFWHTML.ImageButton(gk_sImagesPath & gk_sImageOK, ksNameButtonOk, _
                    sOKLink, gk_sMainFrame)
    sHTML = sHTML & m_oFWHTML.ImageButton(gk_sImagesPath & gk_sImageCancel, ksNameButtonCancel, _
                    sCancelLink, gk_sMainFrame)
                            
    m_oFWTable.StartRow
    m_oFWTable.AddCellAsHTML sHTML, typCellStyle
    m_oFWTable.EndRow
    
    m_aspResponse.Write m_oFWTable.GetHTML
    
    'Add the help button
    m_aspResponse.Write HTMLToAddHelpButton(gk_sHelpLinkToAbandnon)

    Exit Sub

ErrorHandler:
    m_objObjectContext.SetAbort
    ShowError sSource

End Sub





Public Sub AbandonEditset()
'method to display conformation screen for the abandon functionality
'   d36     ksk     created
On Error GoTo ErrorHandler
    
    Dim sSource As String, sEditsetParam As String
    Dim lEditSetID  As Long
    Dim eResult As ESecurityCheckResult
    
    sSource = Me.Name & ".AbandonEditset()"
    
    Set m_oFWTable = New FWTable
    
    'Get editset id from the query string
    lEditSetID = m_aspRequest.QueryString(gk_sParamEditSetID)
    If lEditSetID <= 0 Then
        RaiseError sSource, EPROCATINTERNALERROR.peEditSetIDInvalue, _
                "Editset ID not found"
    End If
    sEditsetParam = "?" & gk_sParamEditSetID & "=" & lEditSetID
    
    m_oBOEditorialPermissions.EditsetID = lEditSetID
    m_oBOEditorialPermissions.UserID = m_lLoggedUserID
    
    eResult = m_oBOEditorialPermissions.Check(EUserCanAbandonEditSet)
    If eResult <> ESecurityCheckResult.ESRSecurityCheckSuccessful Then
        RaiseError sSource, _
               GetErrorEnumForSecurityCheckResult(eResult), _
               GetDescriptionForSecurityCheckError( _
                        GetErrorEnumForSecurityCheckResult(eResult))
    End If
    
    m_oBOEditSet.EditsetID = lEditSetID
    m_oBOEditSet.AbandonEditset
    
    m_objObjectContext.SetComplete
    
    m_aspResponse.Redirect gk_sLinkDisplayEditSetsForOneUser
    
    Exit Sub
ErrorHandler:
    m_objObjectContext.SetAbort
    ConvertError sSource
End Sub

Public Sub DisplayEditSetDetails()
'd11 ksk Moved code from DisplayEditSetDetailsHeader
'       Coded with a call to DisplayEditSetDetailsHeader and the VWEditSet
On Error GoTo ErrorHandler
    
    ' Declare locals
    Dim sSource    As String
    Dim lEditSetID As Long
    Dim eUsage     As EEditSetActionType
    
    ' Make sure that the editset ID was passes to this page
    If m_aspRequest.QueryString(gk_sParamEditSetID) = "" Then
    
        RaiseError sSource, EPROCATINTERNALERROR.peEditSetIDInvalue, _
                        "Editset ID not found"
                        
    End If
    
    lEditSetID = CLng(m_aspRequest(gk_sParamEditSetID))
    
    If lEditSetID <= 0 Then
    
        RaiseError sSource, EPROCATINTERNALERROR.peEditSetIDInvalue, _
                            "Editset ID not found"
                            
    End If
    
    ' Create instance of the editset
    m_oBOEditSet.EditsetID = lEditSetID
    m_oBOEditSet.Load
    
    ' d90 Determine if the editset is a delete edit set
    ' (4 is hardcoded as the edit set types do not match the values
    ' in the enumeration ECreateEditSetUsage)
    If m_oBOEditSet.EditSetTypeID = 4 Then
        eUsage = esaDelete
    Else
        eUsage = esaNormal
    End If
        
    Call DisplayEditsetDetailsHeader(eUsage)
    Call DisplayEditsetCatalogueEntries(eUsage)
    
    ' Clean up
    Set m_oBOEditSet = Nothing
    
    ' Success
    m_objObjectContext.SetComplete
    
    Exit Sub
ErrorHandler:
    m_objObjectContext.SetAbort
    ShowError sSource
End Sub

Public Sub DisplayChoosePiecesToMove()
'Description
'
'Displays an edit set where pieces can be selected to be moved
'to a new parent
'
'Amendment history
'   d87     alm     19Oct2001   Created (based on DisplayEditSetDetails)
'
    Call DisplayEditsetDetailsHeader(esaChoosePieces)
    Call DisplayEditsetForMultiplePieceMove(esaChoosePieces)

End Sub

Public Sub DisplayChooseParentOfPieces()
'Description
'
'Displays an edit set where pieces can be selected to be moved
'to a new parent
'
'Amendment history
'   d87     alm     19Oct2001   Created (based on DisplayEditSetDetails)
'
    Call DisplayEditsetDetailsHeader(esaChooseParentOfPieces)
    Call DisplayEditsetForMultiplePieceMove(esaChooseParentOfPieces)

End Sub

Public Sub DisplayChooseParentForPieces()
'Description
'
'Displays an edit set (one lettercode only) levels 1 - 5 so that the
'user can select a new parent for selected pieces
'
'Amendment history
'   d87     alm     19Oct2001   Created (based on DisplayEditSetDetails)
'
    Call DisplayEditsetDetailsHeader(esaChooseParentForPieces)
    Call DisplayEditsetForMultiplePieceMove(esaChooseParentForPieces)

End Sub




Private Sub DisplayStagesForEditSet( _
    ByVal t_lEditSetID As Long, _
    ByVal t_sEditSetName As String, _
    ByVal t_lCurrentStageID As Long, _
    ByVal t_sCurrentStageName As String, _
    ByVal t_bMoveToNextStage As Boolean)
    
'Description
'
'Displays the next/previous stages available for the edit set to be moved
'
'Amendment history
'   d45     ndk     23Nov1999   Displaying appropirate message if no next/previous stages are available
'   d47     ndk     08Dec1999   Added help buttons
'   d50     ndk     21Dec1999   Added security check
'

On Error GoTo ErrorHandler

    Dim sSource As String
    Dim sHTML As String
    Dim sDisplayString As String, sHyperlink As String
    
    Dim typRowStyle As RowStyle
    Dim typCellStyle As CellStyle
    Dim typTableStyle As TableStyle
    Dim eSecurityCheckStatus As ESecurityCheckResult
    
    Dim oFWRecordlist As FWRecordList
    
    Dim sUserName As String

    sSource = Me.Name & ".DisplayStagesForEditSet"
    
    m_oFWTable.Reset
    
    'perform security check
    m_oBOEditorialPermissions.EditsetID = t_lEditSetID
    m_oBOEditorialPermissions.UserID = m_lLoggedUserID
    
    eSecurityCheckStatus = m_oBOEditorialPermissions.Check(EUserCanMoveEditSetFromCurrentStage)
    
    If eSecurityCheckStatus <> ESecurityCheckResult.ESRSecurityCheckSuccessful Then
        RaiseError sSource, _
               GetErrorEnumForSecurityCheckResult(eSecurityCheckStatus), _
               GetDescriptionForSecurityCheckError( _
                        GetErrorEnumForSecurityCheckResult(eSecurityCheckStatus))
    End If
    
    'Get the list of next/previous stages for the editset
    
    Set oFWRecordlist = m_oBOEditSet.NextPreviousStagesForEditSet(t_lCurrentStageID, t_bMoveToNextStage)
    
    typTableStyle = m_oFWHTML.GetTableStyle(TABContentStyle)
    m_oFWTable.TableStyle = typTableStyle

    If Not oFWRecordlist.RecordCount = 0 Then
        
        'Headings
        If t_bMoveToNextStage Then
            sDisplayString = m_oFWAppData.GetDisplayInformation(erfDisplayNextStagesForEditSetMessageApproveAndAdvance) & " " & _
                m_oFWHTML.Bold(t_sEditSetName) & " " & _
                m_oFWAppData.GetDisplayInformation(erfDisplayNextStagesForEditSetMessageToTheNextStage)

        Else
            sDisplayString = m_oFWAppData.GetDisplayInformation(erfDisplayPreviousStagesForEditSetMessageApproveAndSendBack) & " " & _
                m_oFWHTML.Bold(t_sEditSetName) & " " & _
                m_oFWAppData.GetDisplayInformation(erfDisplayPreviousStagesForEditSetMessageToThePreviousStage)
        End If
        
        m_oFWTable.StartRow
        ShowNormalStyleText sDisplayString, "100%", 2
        m_oFWTable.EndRow
        
        'Headings
        If t_bMoveToNextStage Then
            sDisplayString = m_oFWAppData.GetDisplayInformation(erfDisplayNextStagesForEditSetMessageTheEditSetIsCurrentlyIn) & " " & _
                m_oFWHTML.Bold(t_sCurrentStageName) & " " & _
                m_oFWAppData.GetDisplayInformation(erfDisplayNextStagesForEditSetMessageStage) & " " & _
                m_oFWAppData.GetDisplayInformation(erfDisplayNextStagesForEditSetMessageSelectANewStage)
        Else
            sDisplayString = m_oFWAppData.GetDisplayInformation(erfDisplayPreviousStagesForEditSetMessageTheEditSetIsCurrentlyIn) & " " & _
                m_oFWHTML.Bold(t_sCurrentStageName) & " " & _
                m_oFWAppData.GetDisplayInformation(erfDisplayPreviousStagesForEditSetMessageStage) & " " & _
                m_oFWAppData.GetDisplayInformation(erfDisplayPreviousStagesForEditSetMessageSelectANewStage)
        End If
        
        m_oFWTable.StartRow
        ShowNormalStyleText sDisplayString, "100%", 2
        m_oFWTable.EndRow
        
        sHTML = sHTML & m_oFWTable.GetHTML
        
        m_oFWTable.Reset
        
        'Form start
        If t_bMoveToNextStage Then
            sHyperlink = gk_sLinkMoveEditSetToNextStage
        Else
            sHyperlink = gk_sLinkMoveEditSetToPreviousStage
        End If
        sHTML = sHTML & m_oFWHTML.FormStart(sHyperlink & _
                    "?" & gk_sParamEditSetID & "=" & t_lEditSetID & _
                    "&" & ksParamCurrentStageID & "=" & t_lCurrentStageID, _
                    gk_sMainFrame, "", MPost, True)

        typTableStyle = m_oFWHTML.GetTableStyle(TABContentStyle)
        typTableStyle.sWidth = "60%"
        
        m_oFWTable.TableStyle = typTableStyle
        m_oFWTable.StartRow
        
        'Headings
        If t_bMoveToNextStage Then
            sDisplayString = m_oFWAppData.GetDisplayInformation(erfDisplayNextStagesForEditSetLeftHeaderNextStage)
        Else
            sDisplayString = m_oFWAppData.GetDisplayInformation(erfDisplayPreviousStagesForEditSetLeftHeaderPreviousStage)
        End If
        ShowGeneralLeftHeading sDisplayString
        
        'To choose the stage
        ShowGeneralRightContent _
            m_oFWHTML.Listbox(ksNameEditSetStage, oFWRecordlist, False), "", 2
        m_oFWTable.EndRow
        
        'Headings
        m_oFWTable.StartRow
        If t_bMoveToNextStage Then
            sDisplayString = m_oFWAppData.GetDisplayInformation(erfDisplayNextStagesForEditSetLeftHeaderComments)
        Else
            sDisplayString = m_oFWAppData.GetDisplayInformation(erfDisplayPreviousStagesForEditSetLeftHeaderComments)
        End If
        ShowGeneralLeftHeading sDisplayString
        
        'To key in the comments
        ShowGeneralRightContent _
            m_oFWHTML.TextArea(ksNameStageTransitionComments, "", 50, 3, 500), "", 2
            
        m_oFWTable.StartRow
        ShowGeneralLeftHeading gk_sHTMLNonBreakSpace
                
        'OK Button

        typCellStyle.lCOLSPAN = 1
        typCellStyle.sWidth = "50%"
        typCellStyle.eALIGN = HALLeft
        m_oFWTable.AddCellAsHTML _
            m_oFWHTML.SubmitImage(gk_sImagesPath & gk_sImageOK, "", "", ""), typCellStyle
        
        'Cancel Button
        typCellStyle.lCOLSPAN = 1
        typCellStyle.sWidth = "50%"
        typCellStyle.eALIGN = HALRight
        m_oFWTable.AddCellAsHTML _
            m_oFWHTML.ImageButton(gk_sImagesPath & gk_sImageCancel, _
                "", _
                gk_sLinkDisplayEditSetDetails & _
                    "?" & gk_sParamEditSetID & "=" & t_lEditSetID, _
                gk_sMainFrame)
                
                
        m_oFWTable.EndRow
    Else
        'No stages available
        If t_bMoveToNextStage Then
            m_oFWTable.AddCellAsHTML "No stage present after this stage"
        Else
            m_oFWTable.AddCellAsHTML "No stage present before this stage"
        End If
    End If
    
    sHTML = sHTML & m_oFWHTML.Paragraph(m_oFWTable.GetHTML)
    
    'Add the help button
    If t_bMoveToNextStage Then
        sHTML = sHTML & HTMLToAddHelpButton(gk_sHelpLinkToNextStages)
    Else
        sHTML = sHTML & HTMLToAddHelpButton(gk_sHelpLinkToPreviousStages)
    End If
    
    m_aspResponse.Write sHTML
        
    Set oFWRecordlist = Nothing
        
    Exit Sub

ErrorHandler:
    Set oFWRecordlist = Nothing
    
    ConvertError sSource
End Sub





Public Sub DisplayNextStagesForEditSet()
'
'Description
'
'Displays the next stages available for the edit set to be moved
'
'Amendment history
'   d10 ndk     07Jul1999   Created
'   d50 ndk     21Dec1999   Added security check
'

On Error GoTo ErrorHandler

    Dim sSource As String
    Dim lEditSetID As Long, sEditSetName As String, lCurrentStageID As Long, sCurrentStageName As String
    Dim eSecurityCheckStatus As ESecurityCheckResult

    sSource = Me.Name & ".DisplayNextStagesForEditSet"
    
    'Get editset id, current stage from the query string
    lEditSetID = m_aspRequest.QueryString(gk_sParamEditSetID)
    sEditSetName = m_aspRequest.QueryString(ksParamEditSetName)
    lCurrentStageID = m_aspRequest.QueryString(ksParamCurrentStageID)
    sCurrentStageName = m_aspRequest.QueryString(ksParamCurrentStageName)
    
    'perform security check
    m_oBOEditorialPermissions.EditsetID = lEditSetID
    m_oBOEditorialPermissions.UserID = m_lLoggedUserID
    
    eSecurityCheckStatus = m_oBOEditorialPermissions.Check(EUserCanMoveEditSetFromCurrentStage)
    
    If eSecurityCheckStatus <> ESecurityCheckResult.ESRSecurityCheckSuccessful Then
        RaiseError sSource, _
               GetErrorEnumForSecurityCheckResult(eSecurityCheckStatus), _
               GetDescriptionForSecurityCheckError( _
                        GetErrorEnumForSecurityCheckResult(eSecurityCheckStatus))
    End If
    
    DisplayStagesForEditSet lEditSetID, sEditSetName, lCurrentStageID, sCurrentStageName, True
            
    m_objObjectContext.SetComplete
    
    Exit Sub
    
ErrorHandler:
    m_objObjectContext.SetAbort
    ShowError sSource
End Sub

    



Public Sub MoveEditSetToNextStage()
'
'Description
'
'Moves the editset to the next stage
'
'Amendment history
'   d10     ndk     07Jul1999       Created
'   d25     ndk     11Aug1999       Now handles abandoing an editset
'   d50     ndk     21Dec1999       Added security check
'

On Error GoTo ErrorHandler

    Dim sSource As String
    
    Dim lEditSetID As Long, lCurrentStageID As Long, lStageEditSetIsToBeMoved As Long
    Dim sComment As String
    Dim eSecurityCheckStatus As ESecurityCheckResult
        
    sSource = Me.Name & ".MoveEditSetToNextStage"
    
    'Get the details from the request object and the form object, session
    lEditSetID = m_aspRequest.QueryString(gk_sParamEditSetID)
    lCurrentStageID = m_aspRequest.QueryString(ksParamCurrentStageID)
    If IsEmpty(m_aspRequest.Form(ksNameEditSetStage)) Then      'This happens when Abandon/Release is directly done from the menu
        lStageEditSetIsToBeMoved = m_aspRequest.QueryString(ksParamStageEditSetIsToBeMoved)
    Else
        lStageEditSetIsToBeMoved = m_aspRequest.Form(ksNameEditSetStage)
    End If
    sComment = m_aspRequest.Form(ksNameStageTransitionComments)
    
    'perform security check
    m_oBOEditorialPermissions.EditsetID = lEditSetID
    m_oBOEditorialPermissions.UserID = m_lLoggedUserID
    
    eSecurityCheckStatus = m_oBOEditorialPermissions.Check(EUserCanMoveEditSetFromCurrentStage)
    
    If eSecurityCheckStatus <> ESecurityCheckResult.ESRSecurityCheckSuccessful Then
        RaiseError sSource, _
               GetErrorEnumForSecurityCheckResult(eSecurityCheckStatus), _
               GetDescriptionForSecurityCheckError( _
                        GetErrorEnumForSecurityCheckResult(eSecurityCheckStatus))
    End If
        
    'check the length of the comment
    If Len(LTrim(RTrim(sComment))) > gk_iDBCOLLenComment Then
        RaiseError sSource, peIncorrectLength, "Comments should not exceed " & gk_iDBCOLLenComment & " characters"
    End If
    
    'Call the business object's method to move the editset to the next stage
    m_oBOEditSet.MoveEditSetToDifferentStage _
                lEditSetID, _
                lCurrentStageID, _
                lStageEditSetIsToBeMoved, _
                True, _
                m_lLoggedUserID, _
                sComment

    'Redirect to homepage
    m_aspResponse.Redirect gk_sLinkDisplayEditSetsForOneUser
    
    m_objObjectContext.SetComplete
    
    Exit Sub

ErrorHandler:
    m_objObjectContext.SetAbort
    ShowError sSource

End Sub





Public Sub DisplayPreviousStagesForEditSet()
'
'Description
'
'Displays the previous stages available for the edit set to be moved
'
'Amendment history
'   d10 ndk     07Jul1999   Created
'   d50 ndk     21Dec1999   Added security check
'

On Error GoTo ErrorHandler

    Dim sSource As String
    
    Dim lEditSetID As Long, sEditSetName As String, lCurrentStageID As Long, sCurrentStageName As String
    Dim eSecurityCheckStatus As ESecurityCheckResult
    
    sSource = Me.Name & ".DisplayPreviousStagesForEditSet"
    
    'Get editset id, current stage from the query string
    
    lEditSetID = m_aspRequest.QueryString(gk_sParamEditSetID)
    sEditSetName = m_aspRequest.QueryString(ksParamEditSetName)
    lCurrentStageID = m_aspRequest.QueryString(ksParamCurrentStageID)
    sCurrentStageName = m_aspRequest.QueryString(ksParamCurrentStageName)
        
    'perform security check
    m_oBOEditorialPermissions.EditsetID = lEditSetID
    m_oBOEditorialPermissions.UserID = m_lLoggedUserID
    
    eSecurityCheckStatus = m_oBOEditorialPermissions.Check(EUserCanMoveEditSetFromCurrentStage)
    
    If eSecurityCheckStatus <> ESecurityCheckResult.ESRSecurityCheckSuccessful Then
        RaiseError sSource, _
               GetErrorEnumForSecurityCheckResult(eSecurityCheckStatus), _
               GetDescriptionForSecurityCheckError( _
                        GetErrorEnumForSecurityCheckResult(eSecurityCheckStatus))
    End If
        
    DisplayStagesForEditSet lEditSetID, sEditSetName, lCurrentStageID, sCurrentStageName, False
            
    m_objObjectContext.SetComplete
    
    Exit Sub

ErrorHandler:
    m_objObjectContext.SetAbort
    ShowError sSource
End Sub

    



Public Sub MoveEditSetToPreviousStage()
'
'Description
'
'Displays the next stages available for the edit set to be moved
'
'Amendment history
'   d10 ndk     07Jul1999       Created
'   d43 ndk     23Nov1999       corrected errors
'   d50 ndk     21Dec1999       Added security check

On Error GoTo ErrorHandler

    Dim sSource As String
    
    Dim lEditSetID As Long, lCurrentStageID As Long, lStageEditSetIsToBeMoved As Long
    Dim sComment As String
    Dim eSecurityCheckStatus As ESecurityCheckResult
   
    sSource = Me.Name & ".MoveEditSetToPreviousStage"
    
    'Get the details from the request object and the form object, session
    
    lEditSetID = m_aspRequest.QueryString(gk_sParamEditSetID)
    lCurrentStageID = m_aspRequest.QueryString(ksParamCurrentStageID)
    lStageEditSetIsToBeMoved = m_aspRequest.Form(ksNameEditSetStage)
    sComment = m_aspRequest.Form(ksNameStageTransitionComments)
    
    'perform security check
    m_oBOEditorialPermissions.EditsetID = lEditSetID
    m_oBOEditorialPermissions.UserID = m_lLoggedUserID
    
    eSecurityCheckStatus = m_oBOEditorialPermissions.Check(EUserCanMoveEditSetFromCurrentStage)
    
    If eSecurityCheckStatus <> ESecurityCheckResult.ESRSecurityCheckSuccessful Then
        RaiseError sSource, _
               GetErrorEnumForSecurityCheckResult(eSecurityCheckStatus), _
               GetDescriptionForSecurityCheckError( _
                        GetErrorEnumForSecurityCheckResult(eSecurityCheckStatus))
    End If
    
    'check length of comment
    If Len(LTrim(RTrim(sComment))) > gk_iDBCOLLenComment Then
        RaiseError sSource, PEGeneralInternalError, "Comments should not exceed " & gk_iDBCOLLenComment & " characters"
    End If
    
    'Call the business object's method to move the editset to the next stage
    m_oBOEditSet.MoveEditSetToDifferentStage _
                lEditSetID, _
                lCurrentStageID, _
                lStageEditSetIsToBeMoved, _
                False, _
                m_lLoggedUserID, _
                sComment
    
    'Release the objects
    
    m_aspResponse.Redirect gk_sLinkDisplayEditSetsForOneUser
    
    m_objObjectContext.SetComplete
    
    Exit Sub

ErrorHandler:
    m_objObjectContext.SetAbort
    ShowError sSource

End Sub


Public Sub DisplayEditsetForMultiplePieceMove(t_eUsage As EEditSetActionType)
' Description: Displays edit sets for multiple piece movement. What is
'              displayed depends on the usage
'
'   d87     alm     19Oct2001   Created (based on DisplayEditsetCatalogueEntries)
'   d89     alm     23Oct2001   Removed references to the sort order
On Error GoTo ErrorHandler

    Dim typCellStyle As CellStyle, typTableStyle As TableStyle
    Dim oFWRecordlist As FWRecordList
    Dim sEditsetParam As String
    Dim sSource  As String
    Dim lEditSetID As Long
    Dim i As Integer
    Dim bGoToReferenceSpecified As Boolean
    
    sSource = Me.Name & ".DisplayEditsetForMultiplePieceMove"
    
    ' Make sure that the editset ID was passed to this page
    If m_aspRequest.QueryString(gk_sParamEditSetID) = "" Then
        RaiseError sSource, EPROCATINTERNALERROR.peEditSetIDInvalue, _
                        "Editset ID not found"
    End If
    
    lEditSetID = CLng(m_aspRequest(gk_sParamEditSetID))
    If lEditSetID <= 0 Then
        RaiseError sSource, EPROCATINTERNALERROR.peEditSetIDInvalue, _
                            "Editset ID not found"
    End If
    
    ' Create the parameter string - include the new parent id and level if
    ' choosing piece to move
    If t_eUsage = esaChoosePieces Then
        sEditsetParam = "?" & gk_sParamEditSetID & "=" & lEditSetID & _
                        "&" & gk_sParamCatalogueID & "=" & m_aspRequest.QueryString(gk_sParamCatalogueID) & _
                        "&" & gk_sParamCatalogueLevel & "=" & m_aspRequest.QueryString(gk_sParamCatalogueLevel)
    Else
        sEditsetParam = "?" & gk_sParamEditSetID & "=" & lEditSetID
    End If

    ' Check for the 'go to' catalogue reference
    If (Not (IsEmpty(m_aspRequest(gk_sNameCatalogueReference)) Or _
        m_aspRequest(gk_sNameCatalogueReference) = "")) And t_eUsage = esaChooseParentOfPieces Then
        m_oboFilteredEditset.PositionCatalogueReference = m_aspRequest(gk_sNameCatalogueReference)
        bGoToReferenceSpecified = True
    Else
        m_oboFilteredEditset.PositionCatalogueReference = ""
        bGoToReferenceSpecified = False
    End If

    ' Check for the position to get
    If Not (IsEmpty(m_aspRequest(mk_sNamePosition)) Or _
        m_aspRequest(mk_sNamePosition) = "") Then
        m_oboFilteredEditset.EditsetPosition = CLng(m_aspRequest(mk_sNamePosition))
    End If
    
    ' Set up to get the edit set
    m_oboFilteredEditset.CatalogueEntryStatus = cesNONE
    m_oboFilteredEditset.CatalogueLevelFiltered = ctNone
    m_oboFilteredEditset.EditsetID = lEditSetID
    m_oboFilteredEditset.MaxCataloguePerPage = _
        m_oFWAppData.GetApplicationConfigurationInfomation(pEditsetEntriesPerPage)
        
    ' d89 The sort order is always ceoHierarchical
    m_oboFilteredEditset.CatalogueOrderFiltered = ceoHierarchical
    
    ' Get the entries
    Set oFWRecordlist = m_oboFilteredEditset.GetEditsetEntries
    
    ' Check to see if the entry has been found if necessary
    If m_oboFilteredEditset.EditsetPositionValid = False Then
        
        If m_oboFilteredEditset.EditsetPositionValid = False Then
            RaiseError sSource, PEUSERInvalidGoToWildcardSearch, "Entry not in editset"
        End If
    End If
    
    ' If we are showing a list of pieces to move to a new parent start a form
    If t_eUsage = esaChoosePieces Then
                
        m_aspResponse.Write m_oFWHTML.FormStart(gk_sLinkDisplayMoveMultiplePieces & sEditsetParam _
                                                , gk_sMainFrame, "", MPost)
                                                
    End If
    
    ' Output the page
    Set m_oFWTable = New FWTable
    m_oFWTable.Reset
    typTableStyle = m_oFWHTML.GetTableStyle(TABContentStyle)
    typTableStyle.sWidth = "100%"
    typTableStyle.lCELLSPACING = 2
    m_oFWTable.TableStyle = typTableStyle
    
    ' Add the column headers to the table
    DisplayEditsetEntryDetailsHeading t_eUsage
    
    ' Add the actual content
    DisplayEditsetEntryDetailsContent oFWRecordlist, lEditSetID, t_eUsage
    m_aspResponse.Write m_oFWTable.GetHTML
    
    ' Display the Help Button
    Select Case t_eUsage
    
        Case esaChooseParentOfPieces:
        
            m_aspResponse.Write HTMLToAddHelpButton(gk_sHelpLinkParentOfPiecesEditset)
        
        Case esaChooseParentForPieces:
        
            m_aspResponse.Write HTMLToAddHelpButton(gk_sHelpLinkParentForPiecesEditset)
        
        Case esaChoosePieces:
        
            m_aspResponse.Write HTMLToAddHelpButton(gk_sHelpLinkChoosePiecesEditset)
        
    End Select
    
    ' If we are showing a list of pieces to move to a new parent end a form
    If t_eUsage = esaChoosePieces Then
        m_aspResponse.Write "</FORM>"
    End If
    
    m_objObjectContext.SetComplete
    Set oFWRecordlist = Nothing
    
    Exit Sub
    
ErrorHandler:
    
    Set oFWRecordlist = Nothing

    ShowError sSource
    m_objObjectContext.SetComplete
End Sub



Public Sub DisplayEditsetCatalogueEntries(Optional t_eUsage As EEditSetActionType = esaNormal)
'   d13     ndk     12Jul1999   Added   (Coded by Krish)
'   d24     ndk     10Aug1999   Provided hyperlink to move to editsetentry details
'   d30     ksk     14Oct1999   Display the filtered catalogue entry list in an open tree
'   d46     ksk     07Dec1999   Changed the description for the status for the filter
'   d53     ndk     27Dec1999   Changed to VI code use customizable styles
'   d64     dtm     10feb2000   Modified to display filter text for status as well as level
'   d74     wdp     25May2000   Get # entries to display per page from registry
'   d77     daj     12Jan2001   Added 'go to' ability in edit set details screen
'   d90     alm     24Oct2001   Added optional parameter t_eUsage
On Error GoTo ErrorHandler

    Dim typCellStyle As CellStyle, typTableStyle As TableStyle
    Dim oFWRecordlist As FWRecordList
    Dim sEditsetParam As String
    Dim sSource  As String
    Dim sTempHTML As String
    Dim lEditSetID As Long, eStatus As ECATALOGUEENTRYSTATUS, eLevel As ECATALOGUELEVEL
    Dim i As Integer
    Dim bGoToReferenceSpecified As Boolean
    
    sSource = Me.Name & ".DisplayEditsetCatalogueEntries"
    
    'Make sure that the editset ID was passes to this page
    If m_aspRequest.QueryString(gk_sParamEditSetID) = "" Then
        ' d2    changed the enumeration
        RaiseError sSource, EPROCATINTERNALERROR.peEditSetIDInvalue, _
                        "Editset ID not found"
    End If
    
    lEditSetID = CLng(m_aspRequest(gk_sParamEditSetID))
    If lEditSetID <= 0 Then
        RaiseError sSource, EPROCATINTERNALERROR.peEditSetIDInvalue, _
                            "Editset ID not found"
    End If
    sEditsetParam = "?" & gk_sParamEditSetID & "=" & lEditSetID

    m_oboFilteredEditset.PROCATTransientInterface_GetState CStr(lEditSetID)

    ' d77 - Check for the 'go to' catalogue reference
    If Not (IsEmpty(m_aspRequest(gk_sNameCatalogueReference)) Or _
        m_aspRequest(gk_sNameCatalogueReference) = "") Then
        m_oboFilteredEditset.PositionCatalogueReference = m_aspRequest(gk_sNameCatalogueReference)
        bGoToReferenceSpecified = True
    Else
        m_oboFilteredEditset.PositionCatalogueReference = ""
        bGoToReferenceSpecified = False
    End If
    
    ' d77 - Check for the position to get
    If Not (IsEmpty(m_aspRequest(mk_sNamePosition)) Or _
        m_aspRequest(mk_sNamePosition) = "") Then
        m_oboFilteredEditset.EditsetPosition = CLng(m_aspRequest(mk_sNamePosition))
    End If
    
    
    'd80
    'Check for the filtered view value
    If Not IsEmpty(m_aspRequest(mk_sNameView)) Then
        m_oboFilteredEditset.CatalogueViewFiltered = m_aspRequest(mk_sNameView)
    End If
    
    'd83
    'Check for the filtered order value
    If Not IsEmpty(m_aspRequest(mk_sNameOrder)) Then
        m_oboFilteredEditset.CatalogueOrderFiltered = m_aspRequest(mk_sNameOrder)
    End If
    
    'Check for the filtered status value
    If IsEmpty(m_aspRequest(mk_sNameStatus)) Then
        eStatus = m_oboFilteredEditset.CatalogueEntryStatus
    'If none selected then set to enumeration none value
    ElseIf CInt(m_aspRequest(mk_sNameStatus)) = -1 Then
        eStatus = ECATALOGUEENTRYSTATUS.cesNONE
        m_oboFilteredEditset.EditsetPositionValid = False
    Else
        eStatus = CInt(m_aspRequest(mk_sNameStatus))
        m_oboFilteredEditset.EditsetPositionValid = False
    End If

    'Check for the filtered catalogue level value
    If IsEmpty(m_aspRequest(mk_sNameLevel)) Then
        eLevel = m_oboFilteredEditset.CatalogueLevelFiltered
    'If none selected then set to enumeration none value
    ElseIf CInt(m_aspRequest(mk_sNameLevel)) = -1 Then
        eLevel = ECATALOGUELEVEL.ctNone
        m_oboFilteredEditset.EditsetPositionValid = False
    Else
        eLevel = CInt(m_aspRequest(mk_sNameLevel))
        m_oboFilteredEditset.EditsetPositionValid = False
    End If
    
    m_oboFilteredEditset.CatalogueEntryStatus = eStatus
    m_oboFilteredEditset.CatalogueLevelFiltered = eLevel
    m_oboFilteredEditset.EditsetID = lEditSetID
    m_oboFilteredEditset.MaxCataloguePerPage = _
        m_oFWAppData.GetApplicationConfigurationInfomation(pEditsetEntriesPerPage)
    
    Set oFWRecordlist = m_oboFilteredEditset.GetEditsetEntries
    
    ' d10 - Check to see if the entry has been found if necessary
    If bGoToReferenceSpecified And _
        m_oboFilteredEditset.EditsetPositionValid = False Then
        
        ' Remove the filter and try again
        m_oboFilteredEditset.CatalogueEntryStatus = ECATALOGUEENTRYSTATUS.cesNONE
        m_oboFilteredEditset.CatalogueLevelFiltered = ECATALOGUELEVEL.ctNone
        Set oFWRecordlist = m_oboFilteredEditset.GetEditsetEntries
        
        If m_oboFilteredEditset.EditsetPositionValid = False Then
            RaiseError sSource, PEUSERInvalidGoToWildcardSearch, "Entry not in editset"
        End If
    End If
    
    Set m_oFWTable = New FWTable

    m_oFWTable.Reset
    
    typTableStyle = m_oFWHTML.GetTableStyle(TABContentStyle)
    typTableStyle.sWidth = "100%"
    typTableStyle.lCELLSPACING = 2
    
    m_oFWTable.TableStyle = typTableStyle
    
    ' Add the filter info
    DisplayEditsetEntryDetailsFilterInfo GetFilterString()
    
    ' Add the column headers to the table
    DisplayEditsetEntryDetailsHeading t_eUsage
    
    ' Add the actual content
    DisplayEditsetEntryDetailsContent oFWRecordlist, lEditSetID, t_eUsage
    
    m_aspResponse.Write m_oFWTable.GetHTML
    
    ' d90 The help that is displayed is different if this is a delete editset
    If t_eUsage = esaDelete Then
        m_aspResponse.Write HTMLToAddHelpButton(gk_sHelpLinkToDeleteEditsetEntryList)
    Else
        m_aspResponse.Write HTMLToAddHelpButton(gk_sHelpLinkToFilteredEditsetEntryList)
    End If
    
    m_oboFilteredEditset.PROCATTransientInterface_SetState CStr(lEditSetID)
    
    m_objObjectContext.SetComplete
    
    Set oFWRecordlist = Nothing
    
    Exit Sub
    
ErrorHandler:
    
    Set oFWRecordlist = Nothing

    ShowError sSource
    m_oboFilteredEditset.PROCATTransientInterface_DeleteState CStr(lEditSetID)
    m_objObjectContext.SetComplete
End Sub


Private Function GetFilterString() As String
' Generation for the required description to display the current applicable filter in the BOFilteredEditset
'
'   d77     daj     12Jan2001   Created (split out from DisplayEditsetCatalogueEntries)
'   d80     pnp     24Jan2001   Added statements for view filtering

    Dim sFilterText As String
    sFilterText = ""
    'd80
    Dim sViewText As String
    sViewText = ""
    'd83
    Dim sOrderText As String
    sOrderText = ""
    
    'd80
    'get the view filtering
    
    sViewText = m_oboFilteredEditset.GetEditsetView(m_oboFilteredEditset.CatalogueViewFiltered)
    sViewText = sViewText & " view, "
    
    
    
    ' Get the status filtering
    If (m_oboFilteredEditset.CatalogueEntryStatus <> ECATALOGUEENTRYSTATUS.cesNONE) Then
        sFilterText = sFilterText & m_oboFilteredEditset.GetEditsetStatusDescription(m_oboFilteredEditset.CatalogueEntryStatus)
    End If
    
    ' Get the level filtering
    If (m_oboFilteredEditset.CatalogueLevelFiltered <> ECATALOGUELEVEL.ctNone) Then
        sFilterText = IIf(sFilterText = "", "", sFilterText & " which are ")
        sFilterText = sFilterText & GetCatalogueLevelDescription(GetCatalogueLevelCode(m_oboFilteredEditset.CatalogueLevelFiltered)) & " entries "
    End If
    
    ' If we have no filter, set a suitable message
    If Trim(sFilterText) = "" Then
        sFilterText = mk_sNoFilterMessage
    End If
    
    'd83
    'get the view filtering
    
    sOrderText = m_oboFilteredEditset.GetEditsetOrder(m_oboFilteredEditset.CatalogueOrderFiltered)
    sOrderText = " - " & sOrderText

    GetFilterString = sViewText & sFilterText & sOrderText
    
End Function

Private Sub DisplayEditsetEntryDetailsFilterInfo(t_sFilterString As String)
' Displays the current filter information and the 'go to' input box to allow
' navigation around the edit set quickly
'
'   d77     daj     12Jan2001   Created

    Dim typCellStyle As CellStyle

    typCellStyle.eStyle = CELLGeneralContentStyle
    typCellStyle.lCOLSPAN = 11
    
    m_oFWTable.StartRow
    
    ' Add filter info
    m_oFWTable.AddCellAsHTML m_oFWHTML.Bold(m_oFWAppData.GetDisplayInformation( _
                    EEditorialResourceFile.erfFilterDescriptionText) _
                    & " " & t_sFilterString), typCellStyle
    
    ' Finish the row off
    m_oFWTable.EndRow
    
End Sub


Private Sub DisplayEditsetEntryDetailsHeading(t_eUsage As EEditSetActionType)
'Display the table header for the editset entry list
'   d30     ksk     14Oct1999   Display the filtered catalogue entry list in an open tree
'   d79     pnp     21Jan2001   Updated DisplayEditsetEntryDetailsHeader to display custom columns
'   d87     alm     15Oct2001   Modified function to support the new move pieces editset states
'   d90     alm     24Oct2001   Modified function to support the new delete editset state
    
    'Display the Table header
    
    Dim iViewType As Integer
    Dim oFWRecordlist As FWRecordList
    Dim sColName As String
    Dim iColSpan As Integer
    Dim sColWidth As String
    Dim sColText As String
        
    
    iViewType = m_oboFilteredEditset.CatalogueViewFiltered
    Set oFWRecordlist = m_oboFilteredEditset.GetEditsetColumnList(iViewType)
    m_oFWTable.StartRow
    
    ' d87 The headings we show depend on the usage
    Select Case t_eUsage
    
        ' The catalogue levels 1 - 7 are not shown and niether is status
        Case esaChoosePieces:
            ShowColumnHeading _
                m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfCatalogueReference), 1, "20%"
    
        ' Only levels 3, 4, 5 are shown and not status
        Case esaChooseParentForPieces
            
            ShowColumnHeading _
                m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfClassShortName), 1, "1%"
            ShowColumnHeading _
                m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfHeaderShortName), 1, "1%"
            ShowColumnHeading _
                m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfSubheaderShortName), 1, "1%"
            ShowColumnHeading _
                m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfCatalogueReference), 1, "20%"
                
        ' All levels are shown but not status
        Case esaChooseParentOfPieces, esaDelete 'd90 Added esaDelete
        
            ShowColumnHeading _
                m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfLettercodeShortName), 1, "1%"
            ShowColumnHeading _
                m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfDivisionShortName), 1, "1%"
            ShowColumnHeading _
                m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfClassShortName), 1, "1%"
            ShowColumnHeading _
                m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfHeaderShortName), 1, "1%"
            ShowColumnHeading _
                m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfSubheaderShortName), 1, "1%"
            ShowColumnHeading _
                m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfPieceShortName), 1, "1%"
            ShowColumnHeading _
                m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfItemShortName), 1, "1%"
            ShowColumnHeading _
                m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfCatalogueReference), 1, "20%"
    
        Case Else:
            
            ShowColumnHeading _
                m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfLettercodeShortName), 1, "1%"
            ShowColumnHeading _
                m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfDivisionShortName), 1, "1%"
            ShowColumnHeading _
                m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfClassShortName), 1, "1%"
            ShowColumnHeading _
                m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfHeaderShortName), 1, "1%"
            ShowColumnHeading _
                m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfSubheaderShortName), 1, "1%"
            ShowColumnHeading _
                m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfPieceShortName), 1, "1%"
            ShowColumnHeading _
                m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfItemShortName), 1, "1%"
            ShowColumnHeading _
                m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfCatalogueReference), 1, "20%"
            ShowColumnHeading _
                m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfStatus), 1, "6%"
                                
    End Select
        
    'd79
    oFWRecordlist.MoveFirst
    While Not oFWRecordlist.IsAtEnd
        sColName = oFWRecordlist.GetFieldValue(gk_sDBCOLValue)
        sColWidth = oFWRecordlist.GetFieldValue(gk_sDBCOLWidth)
        m_oboFilteredEditset.AddColumnHeading sColName, sColText, iColSpan
        oFWRecordlist.MoveNext
        
        'd87 The last column only spans 2 if we are not showing the 'Move' column
        If oFWRecordlist.IsAtEnd And t_eUsage <> esaChoosePieces Then
            iColSpan = 2
        End If
        
        ShowColumnHeading sColText, iColSpan, sColWidth
    Wend
    
    'd87 If we are selecting multiple pieces to move show a move
    'column heading
    If t_eUsage = esaChoosePieces Then
        ShowColumnHeading _
            m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfMoveHeader), 1, "1%"
    End If
    
    m_oFWTable.EndRow

End Sub





Private Sub DisplayEditsetEntryDetailsContent(t_ofwRecordList As FWRecordList, t_lEditSetID As Long, _
                                              t_eUsage As EEditSetActionType)
'Displays the content part, edit set entry list, in a tree format in the list
'   d30     ksk     14Oct1999   Display the filtered catalogue entry list in an open tree
'   d31     ndk     14Oct1999   Changed the link for viewing catalogue details
'   d33     ksk     21Oct1999   Add the links to the new images for readonly and imported status
'                               And ignoring the remove button for readonly entries
'   d39     ndk     17Nov1999   Delete catalogue entry button enabled after doing security check
'   d48     ndk     10Dec1999   Hyperlink added to Readonly entries to change them to editable entries
'   d54     ndk     28Dec1999   Added tooltips
'   d60     spb     08Feb2000   Move tooltip and security access out of loop in
'                               DisplayEditsetEntryDetailsContent.
'                               Change readonly entries to have hyperlink on
'                               add icon at the right, with new tooltip, but
'                               sharing same image as added entry status.
'   d61     dtm     09Feb2000   modified DisplayEditsetEntryDetailsContent so that 'Add read-only' button only hyperlinked if user has permissions
'   d62     jcc     09Feb2000   changed tooltips for arrows on DisplayEditSet
'   d66     daj     03Mar2000   Changed DisplayEditsetCatalogueEntries to not show the add button at all...
'   d72     tmj     30Mar2000   Changed so title is now rendered as XML
'   d75     daj     17Jul2000   Fixed ordering of recordset position calculation to prevent negative positions
'   d77     daj     12Jan2001   Modified to use absolute positions instead of pages
'   d80     pnp     22Jan2001   Updated DisplayEditsetEntryDetailsContent to display custom content
'   d85     daj     05Mar2001   Added colouring on the edit set details screen
'   d86     map     31May2001   Added call to function to add end tags truncated database results
'   d87     alm     19Oct2001   Modified function to display the edit set content of the new multiple piece editset states
'   d89     alm     23Oct2001   Removed references to the sort order
'   d90     alm     24Oct2001   Now supports the new delete editset
'   d91     alm     12Nov2001   Fixed bug where list of pieces to move was curtailed when a piece with an item was hit
'   d92     alm     12Nov2001   Fixed bug where user told that there was no pieces in editset when the current page had no pieces
'   d93     alm     15Nov2001   Fixed bug where series not always remembered. Not elegant. Should be rewritten!!!
On Error GoTo ErrHandler
    
    Dim sSource As String, sParamString As String, sReference As String, sHyperlink As String
    Dim bPreviousPagePresent As Boolean, bNextPagePresent As Boolean
    Dim sParameter As String, sHTML As String, sTempHTML  As String
    Dim sTitleToDisplay As String
    Dim lPosition As Long, lStartPosition As Long
    Dim i As Integer
    Dim eCatalogueEntryStatusValue As ECATALOGUEENTRYSTATUS
    Dim bUserCanDeleteCatalogueEntry As Boolean, bUserCanAddCatalogueEntry As Boolean
    Dim typCellStyle As CellStyle
    Dim lMaxCatPerPage As Long, lNumRecords As Long
    'd80
    Dim oFWRecordlist As FWRecordList
    Dim iViewType As Integer
    Dim sColName As String
    Dim iColSpan As Integer
    Dim sColWidth As String
    Dim sColText As String
    Dim bAsString As Boolean
    Dim sPreviousNextPage As String
    Dim sSubmit As String
    Dim lSeriesPos As Long
    Dim lParentPos As Long
    Dim bValidChoice As Boolean
    Dim bSkip As Boolean
    
    Dim oFWRecord As FWRecord
    
    sSource = Me.Name & ".DisplayEditsetEntryDetailsContent"
    
    ' d85 - Set the default cell style
    typCellStyle = m_oFWHTML.GetCellStyle(CELLColumnDataStyle)
    
    'Skip the display if no records to display
    lNumRecords = t_ofwRecordList.RecordCount
    If lNumRecords <= 0 Then
        Exit Sub
    End If
    
    ' Set up images & tooltips
    Dim sTooltipDisplayDetails As String
    Dim sImageNewentry As String
    Dim sImageAddedEntry As String
    Dim sImageImportedEntry As String
    Dim sImageModifiedEntry As String
    Dim sImageReadOnly As String
    Dim sImageRemove As String
    Dim sImageAddReadOnly As String
    
    sTooltipDisplayDetails = m_ofwRefData.Search(gToolTip, ETOOLTIP.ttViewCatalogueDetails)
    
    sImageNewentry = m_oFWHTML.Image(mk_sNameStatus, gk_sImagesPath & gk_sImageNew, _
                                    m_ofwRefData.Search(gToolTip, ETOOLTIP.ttNewCatalogue))
    sImageAddedEntry = m_oFWHTML.Image(mk_sNameStatus, gk_sImagesPath & gk_sImageAddedFromLive, _
                                    m_ofwRefData.Search(gToolTip, ETOOLTIP.ttAddedCatalogue))
    sImageImportedEntry = m_oFWHTML.Image(mk_sNameStatus, gk_sImagesPath & gk_sImageImported, _
                                    m_ofwRefData.Search(gToolTip, ETOOLTIP.ttImportedCatalogue))
    sImageModifiedEntry = m_oFWHTML.Image(mk_sNameStatus, gk_sImagesPath & gk_sImageChecked, _
                                    m_ofwRefData.Search(gToolTip, ETOOLTIP.ttModifiedCatalogue))
    sImageReadOnly = m_oFWHTML.Image(mk_sNameStatus, gk_sImagesPath & gk_sImageReadOnly, _
                                    m_ofwRefData.Search(gToolTip, ETOOLTIP.ttReadOnlyCatalogue))
    sImageRemove = m_oFWHTML.Image(mk_sNameOpenImage, gk_sImagesPath & gk_sImageRemoveButton, _
                                    m_ofwRefData.Search(gToolTip, ETOOLTIP.ttButtonRemove))
    sImageAddReadOnly = m_oFWHTML.Image(mk_sNameOpenImage, gk_sImagesPath & gk_sImageAddReadOnlyButton, _
                                    m_ofwRefData.Search(gToolTip, ETOOLTIP.ttButtonAddReadOnly))
                                                
    ' d87 The display of the whole edit set, or so many per page
    ' depends on the editset usage
    Select Case t_eUsage
    
        ' We display so many items per page
        Case esaNormal, esaChooseParentOfPieces, esaDelete: 'd90 Added esaDelete
            
            ' d77 - Get the positions
            lMaxCatPerPage = m_oboFilteredEditset.MaxCataloguePerPage
            lStartPosition = m_oboFilteredEditset.EditsetPosition
            lStartPosition = IIf(lStartPosition >= lNumRecords, (lNumRecords - 1), lStartPosition)
            lStartPosition = IIf(lStartPosition < 0, 0, lStartPosition)
    
            'Set value for previous page
            bPreviousPagePresent = False
            If lStartPosition > 0 Then
                bPreviousPagePresent = True
            End If
    
            'Set value for next page
            bNextPagePresent = False
            If lStartPosition < (lNumRecords - lMaxCatPerPage) Then
                bNextPagePresent = True
            End If
            
        ' We display all pieces from the parent
        Case esaChoosePieces
        
            lMaxCatPerPage = t_ofwRecordList.RecordCount
            lStartPosition = CLng(m_aspRequest.QueryString(gk_sParamParentPos))
            bPreviousPagePresent = False
            bNextPagePresent = False
                    
        ' We display all items from a single lettercode
        Case esaChooseParentForPieces:
        
            lMaxCatPerPage = t_ofwRecordList.RecordCount
            lStartPosition = CLng(m_aspRequest.QueryString(gk_sParamSeriesPos)) - 1
            bPreviousPagePresent = False
            bNextPagePresent = False
            
    End Select
    
    ' d87 The asp page that we call to display the previous or next page
    ' depends on the usage of the editorial
    Select Case t_eUsage
    
        Case esaNormal, esaDelete: 'd90 Added esaDelete
            
            sPreviousNextPage = gk_sLinkDisplayEditSetDetails
            
        Case esaChooseParentOfPieces:
        
            sPreviousNextPage = gk_sLinkDisplayChooseParentOfPieces
                        
        Case esaChoosePieces
            
            ' Although there are no previous / next functions on this page
            ' we define it here for clarity
            sPreviousNextPage = gk_sLinkDisplayChoosePiecesToMove
        
        Case esaChooseParentForPieces
            
            ' Although there are no previous / next functions on this page
            ' we define it here for clarity
            sPreviousNextPage = gk_sLinkDisplayChooseParentForPieces
            
    End Select
    
    If bPreviousPagePresent Then
        m_oFWTable.StartRow
            
            sParameter = "?" & gk_sParamEditSetID & "=" & t_lEditSetID _
                    & "&" & mk_sNamePosition & "=" & IIf((lStartPosition - lMaxCatPerPage) < 0, 0, (lStartPosition - lMaxCatPerPage))
            
            sHTML = m_oFWHTML.Image(mk_sNameOpenImage, gk_sImagesPath & gk_sImageScrollUP, _
                                m_ofwRefData.Search(gToolTip, ETOOLTIP.ttButtonPageUpHelp))
            sHTML = m_oFWHTML.Hyperlink(sPreviousNextPage & sParameter _
                                    , sHTML, gk_sMainFrame, , _
                                    m_ofwRefData.Search(gToolTip, ETOOLTIP.ttButtonPageUpHelp))
        
           ShowNormalStyleText sHTML, "100%", 8, False
    
        m_oFWTable.EndRow
    End If
    
    'Move to the position
    t_ofwRecordList.MoveTo lStartPosition
                
    ' Check user has delete permission in this editset
    m_oBOEditorialPermissions.UserID = m_lLoggedUserID
    m_oBOEditorialPermissions.EditsetID = t_lEditSetID
        
    If m_oBOEditorialPermissions.Check(EUserCanDeleteCatalogueEntryOnEditSet) = _
        ESecurityCheckResult.ESRSecurityCheckSuccessful Then
        
        bUserCanDeleteCatalogueEntry = True
    Else
        bUserCanDeleteCatalogueEntry = False
    End If
        
    If m_oBOEditorialPermissions.Check(EUserCanAddCatalogueEntryToEditSet) = _
        ESecurityCheckResult.ESRSecurityCheckSuccessful Then
        
        bUserCanAddCatalogueEntry = True
    Else
        bUserCanAddCatalogueEntry = False
    End If

        
    Dim sDeleteButton As String
    Dim sAddReadOnlyEntryButton As String
    Dim iCatalogueLevel As Integer
    
    'd87
    Dim iChildCatalogueLevel As Integer
    
    
    ' d87 Set the default for valid choice, which is true when an edit set list
    ' contains something valid for the specific usage
    If t_eUsage = esaNormal Or t_eUsage = esaDelete Then
        bValidChoice = True
    Else
        bValidChoice = False
    End If
    
    ' d93 This is not elegant. When we are displaying a list of parents to move
    ' multiple pieces from we need to cycle backwards to find the first series
    ' position as the latter code would not find it
    If t_eUsage = esaChooseParentOfPieces Then
        
        t_ofwRecordList.MovePrevious
        
        Do While Not t_ofwRecordList.IsAtBeginning
        
            ' If we find a series, remember then exit
            If t_ofwRecordList.GetData(gk_sDBCOLCatalogueLevel) = 3 Then
                lSeriesPos = t_ofwRecordList.GetIndex()
                Exit Do
            End If
            
            ' Cycle back
            t_ofwRecordList.MovePrevious
    
        Loop
            
        ' Restore the position
        t_ofwRecordList.MoveTo lStartPosition
                    
    End If
    
    'Loop through the list and display the number of elements
    lPosition = 1
    Do While Not t_ofwRecordList.IsAtEnd
        
        ' Get the levels
        iCatalogueLevel = t_ofwRecordList.GetData(gk_sDBCOLCatalogueLevel)
        iChildCatalogueLevel = 0
        
        ' d87 If we are only displaying the pieces of the parent and
        ' we have moved to another catalogue level, stop displaying any
        ' more items, unless it is an item in which case skip (d91)
        bSkip = False
        If (t_eUsage = esaChoosePieces) Then
        
            If iCatalogueLevel < 6 Then
                Exit Do
            ElseIf iCatalogueLevel = 7 Then
                bSkip = True
            End If
            
        End If
        
        ' d91 Only continue if we are not skipping this item
        If bSkip = False Then
        
            ' d87 If we are only displaying the items of a series to choose
            ' a new parent from and we have moved to another catalogue level,
            ' stop displaying any more items
            If (t_eUsage = esaChooseParentForPieces) And (lPosition > 1) And _
               (iCatalogueLevel <= 3) Then
                Exit Do
            End If
            
            ' d87 If we are choosing a parent to move multiple pieces from
            ' remember the current series position
            If (t_eUsage = esaChooseParentOfPieces) And (iCatalogueLevel = 3) Then
                lSeriesPos = t_ofwRecordList.GetIndex()
            End If
            
            ' Get the status of the entry
            eCatalogueEntryStatusValue = IIf(IsNull(t_ofwRecordList.GetFieldValue(gk_sDBCOLEntryTypeID)), 0, _
                                         t_ofwRecordList.GetFieldValue(gk_sDBCOLEntryTypeID))
                                         
            ' d87 If we are displaying a list of possible new parents dont show levels 6 and 7 or
            ' if we are displaying a list of pieces to move dont display those which are not live
            If Not ((t_eUsage = esaChoosePieces And eCatalogueEntryStatusValue = ECATALOGUEENTRYSTATUS.cesReadeOnly) Or _
                    (t_eUsage = esaChooseParentForPieces And iCatalogueLevel > 5)) Then
    
                m_oFWTable.StartRow
            
                ' d87 The levels we display depend on the usage of the edit set
                Select Case t_eUsage
                
                    ' Show no levels
                    Case esaChoosePieces:
                    
                    ' Show first 5 levels
                    Case esaChooseParentForPieces:
                    
                        If iCatalogueLevel > 2 And iCatalogueLevel < 6 Then
                        
                            ' SKIP Levels not belonging to the current levels
                            iColSpan = t_ofwRecordList.GetData(gk_sDBCOLCatalogueLevel) - 1
                            If iColSpan > 2 Then
                                typCellStyle.lCOLSPAN = iColSpan - 2
                                m_oFWTable.AddCellAsHTML "&nbsp;", typCellStyle
                            End If
                        
                            ' Calculate the number of levels left to span over
                            iColSpan = ECATALOGUELEVEL.ctClass - (iCatalogueLevel - 3)
                            typCellStyle.lCOLSPAN = iColSpan
                        
                            ' Display the code in the current level
                            m_oFWTable.AddCellAsString GetCatalogueLevelCode(t_ofwRecordList.GetData(gk_sDBCOLCatalogueLevel)), typCellStyle
                        
                            ' Reset the colspan
                            typCellStyle.lCOLSPAN = 1
                                            
                        End If
                
                    ' Show all levels
                    Case Else:
                            
                        ' SKIP Levels not belonging to the current levels
                        iColSpan = iCatalogueLevel - 1
                        If iColSpan > 0 Then
                            typCellStyle.lCOLSPAN = iColSpan
                            m_oFWTable.AddCellAsHTML "&nbsp;", typCellStyle
                        End If
                    
                        ' Calculate the number of levels left to span over
                        iColSpan = ECATALOGUELEVEL.ctItem - (iCatalogueLevel - 1)
                        typCellStyle.lCOLSPAN = iColSpan
                    
                        ' Display the code in the current level
                        m_oFWTable.AddCellAsString GetCatalogueLevelCode(t_ofwRecordList.GetData(gk_sDBCOLCatalogueLevel)), typCellStyle
                    
                        ' Reset the colspan
                        typCellStyle.lCOLSPAN = 1
                    
                End Select
                 
                ' Display Reference
                sReference = CreateReference( _
                    VariantStringAsString(t_ofwRecordList.GetData(gk_sDBCOLAliasLettercodeRef)), _
                    VariantStringAsString(t_ofwRecordList.GetData(gk_sDBCOLAliasDivisionRef)), _
                    VariantStringAsString(t_ofwRecordList.GetData(gk_sDBCOLAliasClassRef)), _
                    VariantStringAsString(t_ofwRecordList.GetData(gk_sDBCOLAliasHeaderRef)), _
                    VariantStringAsString(t_ofwRecordList.GetData(gk_sDBCOLAliasSubHeaderRef)), _
                    VariantStringAsString(t_ofwRecordList.GetData(gk_sDBCOLAliasPieceRef)), _
                    VariantStringAsString(t_ofwRecordList.GetData(gk_sDBCOLAliasItemRef)), _
                    t_ofwRecordList.GetData(gk_sDBCOLCatalogueLevel), _
                    VariantStringAsString(t_ofwRecordList.GetData(gk_sDBCOLAliasSubClassRef)))
                
                ' D87 Determine if we have been passed the record position of the parent
                ' to move pieces from
                If m_aspRequest.QueryString(gk_sParamParentPos) = "" Then
                    lParentPos = t_ofwRecordList.GetIndex()
                Else
                    lParentPos = CLng(m_aspRequest.QueryString(gk_sParamParentPos))
                End If
                
                'Link to view the details of or delete the catalogue entry
                sParamString = "?" & gk_sParamEditSetID & "=" & t_lEditSetID & _
                    "&" & gk_sParamCatalogueID & "=" & t_ofwRecordList.GetFieldValue(gk_sDBCOLCatalogueID) & _
                    "&" & gk_sParamCatalogueLevel & "=" & t_ofwRecordList.GetFieldValue(gk_sDBCOLCatalogueLevel) & _
                    "&" & gk_sParamDetailsTabNumber & "=" & ETabType.etTabSummary & _
                    "&" & gk_sParamSectionType & "=" & ESectionType.SECeditorial & _
                    "&" & gk_sParamParentPos & "=" & lParentPos & _
                    "&" & gk_sParamSeriesPos & "=" & CStr(lSeriesPos)
                     
                ' d87 The hyperlink we generate and display depends on the usage of the editset
                Select Case t_eUsage
                
                    Case esaNormal, esaDelete: 'd90 Added esaDelete
                            
                        'constuct the hyperlink to view details
                        sHyperlink = gk_sLinkDisplayViewEditorialCatalogueDetails & sParamString
                        
                        'display the hyperlink only if the catalogue entry status is not read only
                        If eCatalogueEntryStatusValue = ECATALOGUEENTRYSTATUS.cesReadeOnly Then
                            m_oFWTable.AddCellAsString sReference, typCellStyle
                        Else
                            m_oFWTable.AddCellAsHTML m_oFWHTML.Hyperlink(sHyperlink, sReference, gk_sMainFrame, , sTooltipDisplayDetails), typCellStyle
                        End If
                         
                    Case esaChooseParentOfPieces:
                            
                        ' We only construct a hyperlink if the entry is a parent of one or more pieces
                        
                        ' Get the child entry
                        If (iCatalogueLevel = 3 Or iCatalogueLevel = 4 Or iCatalogueLevel = 5) Then
                            t_ofwRecordList.MoveNext
                            
                            If t_ofwRecordList.IsAtEnd = False Then
                                iChildCatalogueLevel = t_ofwRecordList.GetData(gk_sDBCOLCatalogueLevel)
                            End If
                            
                            t_ofwRecordList.MovePrevious
                        End If
                                                                                
                        'constuct the hyperlink to view details
                        sHyperlink = gk_sLinkDisplayChooseParentForPieces & sParamString
                        
                        'display the hyperlink
                        If iChildCatalogueLevel = 6 Then
                            m_oFWTable.AddCellAsHTML m_oFWHTML.Hyperlink(sHyperlink, sReference, gk_sMainFrame, , sTooltipDisplayDetails), typCellStyle
                        Else
                            m_oFWTable.AddCellAsString sReference, typCellStyle
                        End If
                         
                    Case esaChoosePieces:
                        
                        ' We dont show any hyperlinks
                        m_oFWTable.AddCellAsString sReference, typCellStyle
                        bValidChoice = True
                        
                    Case esaChooseParentForPieces:
                     
                        ' We only construct a hyperlink if the entry can be a valid parent
                                                                                
                        'constuct the hyperlink to view details
                        sHyperlink = gk_sLinkDisplayChoosePiecesToMove & sParamString
                        
                        'display the hyperlink
                        If (iCatalogueLevel = 3 Or iCatalogueLevel = 4 Or iCatalogueLevel = 5) And _
                           t_ofwRecordList.GetIndex() <> lParentPos Then
                            m_oFWTable.AddCellAsHTML m_oFWHTML.Hyperlink(sHyperlink, sReference, gk_sMainFrame, , sTooltipDisplayDetails), typCellStyle
                            bValidChoice = True
                        Else
                            m_oFWTable.AddCellAsString sReference, typCellStyle
                        End If
                 
                End Select
                     
                ' d87 Display the status of the editset entry (only if the usage is normal)
                If t_eUsage = esaNormal Then
                
                    Select Case eCatalogueEntryStatusValue
                    Case ECATALOGUEENTRYSTATUS.cesPROCATCreated
                        sTempHTML = sImageNewentry
                    Case ECATALOGUEENTRYSTATUS.cesAddedFromLive
                        sTempHTML = sImageAddedEntry
                    Case ECATALOGUEENTRYSTATUS.cesImported
                        sTempHTML = sImageImportedEntry
                    Case ECATALOGUEENTRYSTATUS.cesReadeOnly
                        sTempHTML = sImageReadOnly
                    Case ECATALOGUEENTRYSTATUS.cesPROCATCreatedModified
                        sTempHTML = sImageNewentry & sImageModifiedEntry
                    Case ECATALOGUEENTRYSTATUS.cesAddedFromLiveModified
                        sTempHTML = sImageAddedEntry & sImageModifiedEntry
                    Case ECATALOGUEENTRYSTATUS.cesImportedModified
                        sTempHTML = sImageImportedEntry & sImageModifiedEntry
                    Case Else
                        sTempHTML = gk_sHTMLNonBreakSpace
                    End Select
                    
                    m_oFWTable.AddCellAsHTML sTempHTML, typCellStyle
                
                End If
                
                'd80
                iViewType = m_oboFilteredEditset.CatalogueViewFiltered
                Set oFWRecordlist = m_oboFilteredEditset.GetEditsetColumnList(iViewType)
                oFWRecordlist.MoveFirst
                While Not oFWRecordlist.IsAtEnd
                    sColName = oFWRecordlist.GetFieldValue(gk_sDBCOLValue)
                    If sColName = gk_sDBCOLTitle Then
                        
                      'Display Title or scope and content field upto the set length
                      'd72 The field will now contain XML, so has to be rendered
                      'd86 Need to add end tags first to strings from database that have been truncated
                        sTitleToDisplay = VariantStringAsString(t_ofwRecordList.GetFieldValue(gk_sDBCOLTitle))
                        If Len(sTitleToDisplay) = gk_iDBCOLLenShortTitle Then
                            sTitleToDisplay = AddEndTagsToString(sTitleToDisplay)
                        End If
                            
                        sTitleToDisplay = RenderXML(sTitleToDisplay)
                        sTitleToDisplay = TruncateTitle(sTitleToDisplay)
                        sTitleToDisplay = AlterStringForFullSentenceTrim(sTitleToDisplay, _
                                        gk_iMaxTitleDisplayChars)
                                        
                        sTitleToDisplay = m_oFWHTML.FontWithStyle(sTitleToDisplay, m_oFWHTML.GetFontStyle(FONTNormalStyle))
                        m_oFWTable.AddCellAsHTML sTitleToDisplay, typCellStyle
                    
                    ElseIf sColName = gk_sDBColPiecekeyOrder Then
                    
                        If iCatalogueLevel = 6 Then
                        
                            m_oFWTable.AddCellAsString VariantStringAsString(oFWRecord.GetValueForField(gk_sDBColPiecekeyOrd)), typCellStyle
                        
                        ElseIf iCatalogueLevel = 7 Then
                       
                            m_oFWTable.AddCellAsString VariantStringAsString(oFWRecord.GetValueForField(gk_sDBColItemkeyOrd)), typCellStyle
                        
                        Else
                        
                            m_oFWTable.AddCellAsString "", typCellStyle
                        
                        End If
                 
                    Else
                    
                        Set oFWRecord = t_ofwRecordList.GetRecord()
                        m_oboFilteredEditset.AddContentCell oFWRecord, sColName, sColText, sColWidth, iColSpan, bAsString
                        With typCellStyle
                            .sWidth = sColWidth
                            .lCOLSPAN = iColSpan
                        End With
                        
                        If bAsString Then
                            m_oFWTable.AddCellAsString sColText, typCellStyle
                        Else
                            m_oFWTable.AddCellAsHTML sColText, typCellStyle
                        End If
                    
                    End If
                    oFWRecordlist.MoveNext
                     
                Wend
            
                ' d87 What we display depends on the usage of the editset
                Select Case t_eUsage
                
                    ' Display a trash can or add button
                    Case esaNormal:
                                 
                        ' For readonly entries show add image, otherwise the remove image
                        If eCatalogueEntryStatusValue = ECATALOGUEENTRYSTATUS.cesReadeOnly Then
                            If bUserCanAddCatalogueEntry Then
                                sParamString = _
                                    gk_sLinkAddReadOnlyEntryIntoEditSet & "?" & _
                                    gk_sParamEditSetID & "=" & t_lEditSetID & "&" & _
                                    gk_sParamCatalogueID & "=" & t_ofwRecordList.GetFieldValue(gk_sDBCOLCatalogueID) & "&" & _
                                    gk_sParamCatalogueLevel & "=" & t_ofwRecordList.GetFieldValue(gk_sDBCOLCatalogueLevel)
                                    
                                sTempHTML = m_oFWHTML.Hyperlink(sParamString, sImageAddReadOnly, gk_sMainFrame)
                            End If
                            
                            typCellStyle.sWidth = "1%"
                            m_oFWTable.AddCellAsHTML sTempHTML, typCellStyle
                             
                        Else
                            ' Construct the remove button with a hyperlink with parameters worked out above
                            If bUserCanDeleteCatalogueEntry Then
                                sDeleteButton = m_oFWHTML.Hyperlink(gk_sLinkDisplayConfirmationForCatalogueDeletion & _
                                                sParamString, sImageRemove, gk_sMainFrame, , _
                                                "")
                            Else
                                sDeleteButton = m_oFWHTML.Hyperlink("", sImageRemove, gk_sMainFrame, , _
                                                "")
                            End If
                             
                            typCellStyle.sWidth = "1%"
                            m_oFWTable.AddCellAsHTML sDeleteButton, typCellStyle
                        End If
                 
                        m_oFWTable.EndRow
                        
                    ' d90 Display a trash can / read only button
                    Case esaDelete
                    
                        ' For readonly entries show readonly image, otherwise the remove image
                        If eCatalogueEntryStatusValue = ECATALOGUEENTRYSTATUS.cesReadeOnly Then
    
                            typCellStyle.sWidth = "1%"
                            m_oFWTable.AddCellAsHTML sImageReadOnly, typCellStyle
                             
                        Else
                             
                           ' Construct the remove button with a hyperlink
                            If bUserCanDeleteCatalogueEntry Then
                                sDeleteButton = m_oFWHTML.Hyperlink(gk_sLinkDisplayConfirmationForCatalogueUnDeletion & _
                                                sParamString, sImageRemove, gk_sMainFrame, , _
                                                "")
                            Else
                                sDeleteButton = m_oFWHTML.Hyperlink("", sImageRemove, gk_sMainFrame, , _
                                                "")
                            End If
                             
                            typCellStyle.sWidth = "1%"
                            m_oFWTable.AddCellAsHTML sDeleteButton, typCellStyle
                        End If
                 
                        m_oFWTable.EndRow
                                 
                    ' Dont display a trash can / add button
                    Case esaChooseParentOfPieces:
                         
                        m_oFWTable.EndRow
                         
                    ' Display a check box
                    Case esaChoosePieces:
                     
                        typCellStyle.eALIGN = HALCenter
                        typCellStyle.sWidth = "1%"
                        m_oFWTable.AddCellAsHTML m_oFWHTML.Checkbox("", t_sName:=gk_sFORMCheckboxName + CStr(lPosition), _
                                                                        t_sValue:=t_ofwRecordList.GetFieldValue(gk_sDBCOLCatalogueID)), typCellStyle
                        typCellStyle.eALIGN = HALLeft
                    
                        m_oFWTable.EndRow
                         
                    Case esaChooseParentForPieces
                         
                        m_oFWTable.EndRow
                         
                End Select
            
                lPosition = lPosition + 1
                If lPosition > lMaxCatPerPage Then
                    Exit Do
                End If
                t_ofwRecordList.MoveNext
            
            Else
                t_ofwRecordList.MoveNext
            End If
        Else
            t_ofwRecordList.MoveNext
        End If
    Loop
    
    ' d92 Because esaChooseParentOfPieces can span multiple pages we can only
    ' check to see if there is a valid parent here
    If t_eUsage = esaChooseParentOfPieces Then
    
        ' Start again
        t_ofwRecordList.MoveFirst
        
        Do While Not t_ofwRecordList.IsAtEnd
        
            ' Get the catalogue level
            iCatalogueLevel = t_ofwRecordList.GetData(gk_sDBCOLCatalogueLevel)
               
            ' We have something that could be a parent
            If (iCatalogueLevel = 3 Or iCatalogueLevel = 4 Or iCatalogueLevel = 5) Then
                t_ofwRecordList.MoveNext
                            
                If t_ofwRecordList.IsAtEnd = False Then
                                   
                    ' If its a piece we have found a valid parent
                    If t_ofwRecordList.GetData(gk_sDBCOLCatalogueLevel) = 6 Then
                        bValidChoice = True
                        Exit Do
                    End If
                End If
            Else
                t_ofwRecordList.MoveNext
            End If
            
        Loop
    End If
    
    ' Check to see if there are any errors
    If bValidChoice = False Then
        
        Select Case t_eUsage
        
            Case esaChooseParentOfPieces:
            
                RaiseError sSource, PENoValidParentToMovePiecesFrom, _
                " The editset does not contain any pieces"
            
            Case esaChooseParentForPieces:
            
                RaiseError sSource, PENoValidParentToMovePiecesTo, _
                " The editset does not contain any valid alternative parents for pieces of " & sReference
            
            Case esaChoosePieces:
            
                RaiseError sSource, PENoLivePieces, _
                " The parent does not contain any live pieces in the editset"
        
        End Select
    
    End If
    
    If bNextPagePresent Then
        m_oFWTable.StartRow
        sParameter = "?" & gk_sParamEditSetID & "=" & t_lEditSetID _
                    & "&" & mk_sNamePosition & "=" & IIf((lStartPosition + lMaxCatPerPage) > lNumRecords, lNumRecords, (lStartPosition + lMaxCatPerPage))
        
        sHTML = m_oFWHTML.Image(mk_sNameOpenImage, gk_sImagesPath & gk_sImageScrollDown, _
                            m_ofwRefData.Search(gToolTip, ETOOLTIP.ttButtonPageDownHelp))
        sHTML = m_oFWHTML.Hyperlink(sPreviousNextPage & sParameter _
                                , sHTML, gk_sMainFrame, , _
                                m_ofwRefData.Search(gToolTip, ETOOLTIP.ttButtonPageDownHelp))
           ShowNormalStyleText sHTML, "100%", 8, False
        m_oFWTable.EndRow
    End If

    'd87 Show an ok, select all and deselect all button if showing a list of
    'pieces to move, and a hidden field that specifies the total number of
    'checkboxes
    If t_eUsage = esaChoosePieces Then
            
        sSubmit = m_oFWHTML.SubmitImage(gk_sImagesPath & gk_sImageOK, "OK", "OK")
        
        ' Select all button is script that checks all the checkboxes
        ' Deselect all button is similar thing
        Dim sSelectScript As String, sDeselectScript As String
        Dim sSelectAll As String, sDeselectAll As String
        Dim iCount As Integer
        
        For iCount = 1 To lPosition - 1
            sSelectScript = sSelectScript & gk_sFORMCheckboxName & iCount & ".checked=true;"
            sDeselectScript = sDeselectScript & gk_sFORMCheckboxName & iCount & ".checked=false;"
        Next
        
        sSelectAll = m_oFWHTML.ImageClick(gk_sImagesPath & gk_sImageSelectAll, sSelectScript)
        sDeselectAll = m_oFWHTML.ImageClick(gk_sImagesPath & gk_sImageDeselectAll, sDeselectScript)
                                   
        typCellStyle = m_oFWHTML.GetCellStyle(CELLPlainStyle)
        
        ' Split up the submit and select/deselect buttons into separate columns
        m_oFWTable.AddCellAsHTML (sSubmit), typCellStyle
        
        ' Add the hidden field
        m_oFWTable.AddCellAsHTML (m_oFWHTML.HiddenField(gk_sFORMCheckBoxCount, lPosition - 1))
        
        typCellStyle.eALIGN = HALRight
        typCellStyle.lCOLSPAN = 3  ' 5 cols - OK button - hidden field = 3
        m_oFWTable.AddCellAsHTML (sSelectAll & sDeselectAll), typCellStyle
        
    End If
    
    Exit Sub
    
ErrHandler:
    ConvertError sSource
End Sub





Private Sub ShowColumnHeading( _
    ByVal t_sHeading As String, _
    Optional ByVal t_iColSpan As Integer = 1, _
    Optional ByVal t_sColWidth As String = "")
'
'Description
'
' Show column heading (appears below the above the coloumnar data)
'
'Amendment history
'   d27     ndk     19Aug1999   Created
    
    Dim typCellStyle As CellStyle
    
    typCellStyle = m_oFWHTML.GetCellStyle(CELLColumnHeadingStyle)
    typCellStyle.lCOLSPAN = t_iColSpan
    
    If t_sColWidth <> "" Then
        typCellStyle.sWidth = t_sColWidth
    End If
    
    m_oFWTable.AddCellAsHTML _
        m_oFWHTML.FontWithStyle( _
        t_sHeading, _
        m_oFWHTML.GetFontStyle(FONTColumnHeadingStyle)), _
        typCellStyle

End Sub





Private Sub ShowColumnData( _
    ByVal t_sText As String, _
    Optional ByVal t_iColSpan As Integer = 1, _
    Optional ByVal t_sColWidth As String = "", _
    Optional ByVal t_bDisplayAsString As Boolean = False)
'
'Description
'
' Show Data in Coloumnar mode (appears below the column heading)
'
'Amendment history
'   d27     ndk     19Aug1999   Created
'   d52     ndk     22Dec1999   Replaceing carriage returns with html line breaks
'
    Dim typCellStyle As CellStyle
    
    typCellStyle = m_oFWHTML.GetCellStyle(CELLColumnDataStyle)
    typCellStyle.lCOLSPAN = t_iColSpan
    
    If t_sColWidth <> "" Then
        typCellStyle.sWidth = t_sColWidth
    End If
    
    If t_bDisplayAsString Then
       
        t_sText = m_oFWHTML.EncodeAsHTML(t_sText)
        t_sText = ReplaceString(t_sText, vbCr, gk_sHTMLLineBreak)
        m_oFWTable.AddCellAsHTML t_sText, typCellStyle
        
    Else
        m_oFWTable.AddCellAsHTML _
            m_oFWHTML.FontWithStyle( _
            t_sText, _
            m_oFWHTML.GetFontStyle(FONTColumnDataStyle)), _
            typCellStyle
    End If
End Sub





Private Sub ShowGeneralLeftHeading( _
    ByVal t_sHeading As String, _
    Optional ByVal t_sColWidth As String, _
    Optional ByVal t_iColSpan As Integer = 1)
'
'Description
'
' Shows the general label ( that appears left of the content)
'
'Amendment history
'   d27     ndk     19Aug1999   Created
'   d63     jcc     10Feb2000   removed bold on create empty edit set page

    Dim typCellStyle As CellStyle
    
    typCellStyle = m_oFWHTML.GetCellStyle(CELLGeneralLabelStyle)
    typCellStyle.lCOLSPAN = t_iColSpan
    
    If t_sColWidth <> "" Then
        typCellStyle.sWidth = t_sColWidth
    End If
    
    m_oFWTable.AddCellAsHTML _
        m_oFWHTML.FontWithStyle( _
        t_sHeading, _
        m_oFWHTML.GetFontStyle(FONTGeneralLabelStyle)), _
        typCellStyle
End Sub





Private Sub ShowGeneralRightContent( _
    ByVal t_sText As String, _
    Optional ByVal t_sColWidth As String, _
    Optional ByVal t_iColSpan As Integer = 1)
'
'Description
'
' Shows the general content ( that appears right of a label)
'
'Amendment history
'   d27     ndk     19Aug1999   Created
'   d63     jcc     10Feb2000   removed bold on create empty edit set page

    Dim typCellStyle As CellStyle
    
    typCellStyle = m_oFWHTML.GetCellStyle(CELLGeneralContentStyle)
    typCellStyle.lCOLSPAN = t_iColSpan
    
    If t_sColWidth <> "" Then
        typCellStyle.sWidth = t_sColWidth
    End If
    
    m_oFWTable.AddCellAsHTML _
        m_oFWHTML.FontWithStyle( _
        t_sText, _
        m_oFWHTML.GetFontStyle(FONTGeneralContentStyle)), _
        typCellStyle
End Sub






Private Sub ShowDetailsLine( _
    ByVal t_sDetailLine As String, _
    t_typCellStyle As CellStyle, _
    t_typFontStyle As FontStyle)
'
'Description
'
' Displayes the details according to the cellstyle and font mentioned
'
'Amendment history
'   d27     ndk     19Aug1999   Created
    
    m_oFWTable.AddCellAsHTML _
        m_oFWHTML.FontWithStyle(t_sDetailLine, t_typFontStyle), t_typCellStyle
End Sub





Private Sub ShowNormalStyleText( _
    ByVal t_sText As String, _
    Optional ByVal t_sColWidth As String = "", _
    Optional ByVal t_iColSpan = 0, _
    Optional ByVal t_bDisplayAsString As Boolean = False)
'
'Description
'
' Displays text in normal style
'
'Amendment history
'   d27     ndk     19Aug1999   Created
    
    
    Dim typCellStyle As CellStyle
    
    typCellStyle = m_oFWHTML.GetCellStyle(CELLNormalStyle)
    If t_sColWidth <> "" Then
        typCellStyle.sWidth = t_sColWidth
    End If
    
    If t_iColSpan <> 0 Then
        typCellStyle.lCOLSPAN = t_iColSpan
    End If
    If t_bDisplayAsString Then
        t_sText = m_oFWHTML.EncodeAsHTML(t_sText)
        t_sText = m_oFWHTML.FontWithStyle(t_sText, m_oFWHTML.GetFontStyle(FONTNormalStyle))
        m_oFWTable.AddCellAsHTML t_sText, typCellStyle
    Else
        m_oFWTable.AddCellAsHTML _
            m_oFWHTML.FontWithStyle(t_sText, m_oFWHTML.GetFontStyle(FONTNormalStyle)), typCellStyle
    End If
        
End Sub





Public Sub DisplayFilterForm()
'Displays the  filter form used to filter edit set entry list
'   d30     ksk     14Oct1999   Display the filtered catalogue entry list in an open tree
'   d32     ksk     14Oct1999   Made changes to display filter form
'   d78     pnp     17Jan2001   Added 'View' drop down box code to Display filter Form
'
Dim typCellStyle As CellStyle
Dim sHTML As String, sEditsetParam As String
Dim lEditSetID As Long
Dim sSource As String
Dim lStatus As Long, lLevel As Long

    lEditSetID = CLng(m_aspRequest(gk_sParamEditSetID))
    If lEditSetID <= 0 Then
        ' d2    changed the enumeration
        RaiseError sSource, EPROCATINTERNALERROR.peEditSetIDInvalue, _
                            "Editset ID not found"
    End If
    sEditsetParam = "?" & gk_sParamEditSetID & "=" & lEditSetID
    
        'Makesure that the editset ID was passes to this page
    If m_aspRequest.QueryString(gk_sParamEditSetID) = "" Then
        ' d2    changed the enumeration
        RaiseError sSource, EPROCATINTERNALERROR.peEditSetIDInvalue, _
                        "Editset ID not found"
    End If
    
    m_oboFilteredEditset.PROCATTransientInterface_GetState CStr(lEditSetID)
    
    
    lStatus = IIf(m_oboFilteredEditset.CatalogueEntryStatus = cesNONE, -1, _
                    CInt(m_oboFilteredEditset.CatalogueEntryStatus))
    lLevel = IIf(m_oboFilteredEditset.CatalogueLevelFiltered = ctNone, -1, _
                    CInt(m_oboFilteredEditset.CatalogueLevelFiltered))
    
    
    Set m_oFWTable = New FWTable
    m_oFWTable.TableStyle = m_oFWHTML.GetTableStyle(TABNoStyle)
    
    'Display the Header
    m_oFWTable.StartRow
    
    typCellStyle = m_oFWHTML.GetCellStyle(ECellStyle.CELLPageHeadingStyle)
    typCellStyle.lCOLSPAN = 2
    
    m_oFWTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                EEditorialResourceFile.erfFilterHeading), typCellStyle
                
    m_oFWTable.EndRow
    
    'Display the help text for display
    m_oFWTable.StartRow
    
    typCellStyle = m_oFWHTML.GetCellStyle(ECellStyle.CELLHelpTextStyle)
    typCellStyle.lCOLSPAN = 2
    
    m_oFWTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                EEditorialResourceFile.erfFilterHelpText), typCellStyle
    
    m_oFWTable.EndRow
    
    'Display the row with the status filter selection
    m_oFWTable.StartRow
    
    typCellStyle = m_oFWHTML.GetCellStyle(ECellStyle.CELLReadWriteLabelStyle)
    typCellStyle.sWidth = "30%"
    
    m_oFWTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                            EEditorialResourceFile.erfStatus), typCellStyle
    
    typCellStyle = m_oFWHTML.GetCellStyle(ECellStyle.CELLReadWriteDataStyle)
    typCellStyle.sWidth = "70%"
    m_oFWTable.AddCellAsHTML m_oFWHTML.Listbox(mk_sNameStatus, _
                        m_oboFilteredEditset.GetEditsetStatusList, True, lStatus), _
                        typCellStyle
    
    m_oFWTable.EndRow
    
    'Display the row with the cataloguelevel filter selection
    m_oFWTable.StartRow
    
    typCellStyle = m_oFWHTML.GetCellStyle(ECellStyle.CELLReadWriteLabelStyle)
    typCellStyle.sWidth = "30%"
    
    m_oFWTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                            EEditorialResourceFile.erfLevelName), typCellStyle
    
    typCellStyle = m_oFWHTML.GetCellStyle(ECellStyle.CELLReadWriteDataStyle)
    typCellStyle.sWidth = "70%"
    m_oFWTable.AddCellAsHTML m_oFWHTML.Listbox(mk_sNameLevel, _
                        m_oboFilteredEditset.GetCatalogueLevelList, True, lLevel), _
                        typCellStyle
    
    m_oFWTable.EndRow
    
    'Display the row with the view filter selection

    m_oFWTable.StartRow
    
    typCellStyle = m_oFWHTML.GetCellStyle(ECellStyle.CELLReadWriteLabelStyle)
    typCellStyle.sWidth = "30%"
    
    m_oFWTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                            EEditorialResourceFile.erfView), typCellStyle
    
    typCellStyle = m_oFWHTML.GetCellStyle(ECellStyle.CELLReadWriteDataStyle)
    typCellStyle.sWidth = "70%"
    m_oFWTable.AddCellAsHTML m_oFWHTML.Listbox(mk_sNameView, _
                        m_oboFilteredEditset.GetEditsetViewList, False, m_oboFilteredEditset.CatalogueViewFiltered), _
                        typCellStyle

'   d83     pnp     08Feb2001
'Display the row with the order filter selection

    m_oFWTable.StartRow
    
    typCellStyle = m_oFWHTML.GetCellStyle(ECellStyle.CELLReadWriteLabelStyle)
    typCellStyle.sWidth = "30%"
    
    m_oFWTable.AddCellAsHTML m_oFWAppData.GetDisplayInformation( _
                            EEditorialResourceFile.erfOrder), typCellStyle
    
    typCellStyle = m_oFWHTML.GetCellStyle(ECellStyle.CELLReadWriteDataStyle)
    typCellStyle.sWidth = "70%"
    m_oFWTable.AddCellAsHTML m_oFWHTML.Listbox(mk_sNameOrder, _
                        m_oboFilteredEditset.GetEditsetOrderList, False, m_oboFilteredEditset.CatalogueOrderFiltered), _
                        typCellStyle
m_oFWTable.EndRow
    
    
    'Display the Filter button
    m_oFWTable.StartRow
    typCellStyle = m_oFWHTML.GetCellStyle(ECellStyle.CELLReadWriteLabelStyle)
    typCellStyle.sWidth = "30%"
    
    m_oFWTable.AddEmptyCell
    
    typCellStyle = m_oFWHTML.GetCellStyle(ECellStyle.CELLReadWriteLabelStyle)
    typCellStyle.sWidth = "70%"
    
    sHTML = m_oFWHTML.Hyperlink(gk_sLinkDisplayEditSetDetails & sEditsetParam, _
                        m_oFWHTML.Image(mk_sNameButtonCancel, _
                                        gk_sImagesPath & gk_sImageCancel, _
                        m_ofwRefData.Search(gToolTip, _
                                        ETOOLTIP.ttButtonCancelHelp)), _
                        gk_sMainFrame, mk_sNameButtonCancel, _
                        m_ofwRefData.Search(gToolTip, ETOOLTIP.ttButtonCancelHelp))
    

    m_oFWTable.AddCellAsHTML m_oFWHTML.SubmitImage(gk_sImagesPath & gk_sImageFilter, _
                                mk_sNameFilter, _
                                m_ofwRefData.Search(EGroup.gToolTip, _
                                ETOOLTIP.ttButtonFilterHelp)) & sHTML
    
    m_oFWTable.EndRow
    
    sHTML = m_oFWHTML.FormStart(gk_sLinkDisplayEditSetDetails & sEditsetParam, gk_sMainFrame)
    sHTML = sHTML & m_oFWTable.GetHTML
    sHTML = sHTML & m_oFWHTML.EndTag(gk_sHTMLForm)
    
    m_aspResponse.Write sHTML

    m_aspResponse.Write HTMLToAddHelpButton(gk_sHelpLinkToFiltereForm)
    
    Exit Sub
    
ErrorHandler:
    m_objObjectContext.SetAbort
    ShowError sSource
End Sub




