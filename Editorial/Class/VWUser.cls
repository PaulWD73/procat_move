VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "VWUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    VWUser.cls
' System:       PROCAT, PRO
' Copyright:    (C) Quidnunc Limited
'
' Description:
'
' The User View object handling login functionality
'
' Amendment history:
'
' d1    ksk     26Apr1999       Created.
' d2    ndk     04Jun1999       Changed the login method. Completely rewritten
'                               Added StripDomainName
' d3    ndk     28Jun1999       Login(),Storing whether the logged user is administrator in session
' d4    ndk     08Jul1999       Changed name of some variables to stick to coding standards
' d5    ndk     14Jul1999       Put exit function in stripdomainname
' d6    ndk     04Aug1999       Added terminate method, user has admin role is not stored in the session
' d7    crm     09Aug1999       Moved functionality from VWUserEditSet class
'                               ShowAllUsersScreen(),DeleteUser(),DisplayUserDetail(),NewUserEntryScreen()
'                               AlphabetLinks(),StatusInfo(),UserUpdateDisplay()
' d8    crm     10Aug1999       Modified to pick up string resource via enum in EditorialResourceEnums
' d9    crm     12Aug1999       In UserUpdateDisplay() modified the rolesarray value assignment
' d10   ksk     16Aug1999       Removed references to user type
' d11   aga     06Oct1999       Modified DisplayUserDetails method,ShowAllUsersScreen, add help buttons
' d12   aga     28Oct1999       Added code to UserUpdateDisplay method to handle deletion of last procat manager
' d13   aga     29Oct1999       Handling duplication of User name, NT Logon and User initials
' d14   AGA     02Nov1999       Modified CheckPermissions
' d15   KSK     03Nov1999       Changed Checkpermissions method to used the parameter source
'                               value while raising error
' d16   SPB     01Dec1999       Read user id from session rather than NT username in ShowAllUsersScreen
' d17   CRM     10Dec1999       Added field to accept User_Ticket
' d18   aga     10Dec1999       Added validation for User_Ticket field
' d19   wdp     13Dec1999       Added GetVersion
' d20   crm     27Dec1999       Added tool tips
' d21   ndk     29Dec1999       Rechecking releasing of variables, especially if an error is deliberately raised
' d22   aga     12Jan2000       Fixing defects
' d23   crm     11Feb2000       NewUserEntryScreen made to validate if ME role is assigned
'                               does not validate if cancel button is choosen
' d24   wdp     16Feb2000       Removed need for user ticket field
' d25   wdp     24Feb2000       Added back button to StatusInfo if no OK button is displayed
' d26   tmj     06Mar2000       modified so descriptive column wider
' d27   daj     10Mar2000       Adjusted the lovely styling of the user editing screen
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit


'Interfaces
Implements ObjectControl

'Private Variables
Private m_objObjectContext As ObjectContext
Private m_aspResponse  As Response
Private m_aspRequest As Request

'Private Constants
Private Const ksLogonUser As String = "LOGON_USER"

'Moved from VWUserEditSet class
Private Const klBaseResIdentifier As Long = 28500
Private Const ksNameUserTypeList As String = "ChosenType"
Private Const k_sEditSetName As String = "Name"
Private Const k_sParagraph As String = "<P>"
Private Const klAscValueA As Long = 65
Private Const klAscValueZ As Long = 90
'd11 Permissions
Private m_oboEditorialPermission As BOEditorialPermissions
Private m_lLoggedUserID As Long

'Private Constants
Private Sub ObjectControl_Activate()
    Initialise
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    Terminate
End Sub
Private Sub Initialise()
    
    On Error GoTo ErrorHandler
    
    Dim sSource As String
    sSource = "VWUser.Initialise"
    
    'Get the object context
    Set m_objObjectContext = GetObjectContext()
    If m_objObjectContext Is Nothing Then
        RaiseError sSource, peOCUnableToGetObjectContext, "Objectcontext is nothing"
    End If
    
    'get references to ASP objects
    Set m_aspResponse = m_objObjectContext.Item(gk_sResponseObject)
    Set m_aspRequest = m_objObjectContext.Item(gk_sRequestObject)
    Set m_oboEditorialPermission = New BOEditorialPermissions
    
    'First time when a user logs, UserID will be zero. In Login method, the same line code should exist
    m_lLoggedUserID = GetDataFromSession(dsUserID, m_objObjectContext.Item(gk_sSessionObject))
    
    Exit Sub
    
ErrorHandler:
    m_objObjectContext.SetAbort
    ShowError sSource
End Sub

Private Sub Terminate()
    Set m_objObjectContext = Nothing 'release
    Set m_aspResponse = Nothing 'release
    Set m_aspRequest = Nothing 'release
    Set m_oboEditorialPermission = Nothing
End Sub
' Methods
'
Public Sub Login()
' d2    ndk     04Jun1999       Completely rewritten
' d3    ndk     28Jun1999       Storing whether the logged user is administrator in session
' d6    ndk     04Aug1999       Not Storing whether the logged user is administrator in session
' d11     aga     11Oct1999       Initializing the UserID class variable
'
' To validate login info. and permit user to proceed further
' Passes the NT Login name to the business object
'
On Error GoTo ErrorHandler
    Dim sSource As String
    
    sSource = "VWUser.Login"
    
    Dim oBOUser As BOUser
    
    Set oBOUser = New BOUser
   
    oBOUser.Initialize
  
    'Call the business object to validate the nt logon name of the user
    oBOUser.Login StripDomainName(m_aspRequest.ServerVariables(ksLogonUser))
    
    'Store the users details in the session object
    StoreDataInSession DATASTORE.dsUserID, oBOUser.UserID
    StoreDataInSession DATASTORE.dsUserInitials, oBOUser.UserInitials
    StoreDataInSession DATASTORE.dsUserName, oBOUser.UserName
    StoreDataInSession DATASTORE.dsUserNTLogon, oBOUser.UserNTLogon
    
    m_lLoggedUserID = GetDataFromSession(dsUserID, m_objObjectContext.Item(gk_sSessionObject))
    
    Set oBOUser = Nothing
    
    m_objObjectContext.SetComplete
    
    Exit Sub
    
ErrorHandler:
    m_objObjectContext.SetAbort
    ShowError sSource
End Sub


Private Function StripDomainName(ByVal t_sUserNameWithDomainName As String) As String
'  d2    ndk     04Jun1999       Created
' d5    ndk     14Jul1999       Put exit function
'
'Description :
'Strips the domain name from the user name
'Presumption : Domain Name and the user name are separated by a '/'
'
'Parameters:
'
't_sUserNameWithDomainName : NT Login Name of the user (it could contain the domain name)
'
'Return Value:
'
'User name stripped off the domain name
'
On Error GoTo ErrorHandler
    Dim sSource As String
    
    sSource = "VWUser.StripDomainName(" & t_sUserNameWithDomainName & ")"

    Dim lSlashOccursAt As Long
    lSlashOccursAt = InStr(t_sUserNameWithDomainName, "\")
    If lSlashOccursAt = 0 Then
        StripDomainName = t_sUserNameWithDomainName
    Else
        StripDomainName = Mid(t_sUserNameWithDomainName, lSlashOccursAt + 1)
    End If
    
    'Do not put set complete here
    
    Exit Function
    
ErrorHandler:
    m_objObjectContext.SetAbort
    ShowError sSource
End Function

Public Sub ShowAllUsersScreen()
'Description    :   Displays the main page where all user names are displayed,
'                   with options to view a particular user's details,delete the user
'                   or add a new user via links
'Amendment History
'   d7      CRM         Moved from VWUserEditSet class
'   d11     aga         07Oct1999       Modified. Delete button not be displayed for Managing editor
'                                                   Added security
'   d16     SPB         01Dec1999       Read user id from session rather than NT username in ShowAllUsersScreen
'   d22     aga         12Jan2000       Adding done button
'   d26     tmj         06Mar2000       Modified so descriptive column wider
'   d27     daj         10Mar2000       Adjusted the lovely styling of the user editing screen

On Error GoTo ErrorHandler

    Dim sSource As String
    
    Dim oboUserRoles As BoUserRoles
    Dim oBOUser As BOUser
    Dim ofwHTML As FWHTML
    Dim ofwTable As FWTable
    Dim ofwUserList As FWRecordList
    
    Dim typTableStyle  As TableStyle
    Dim typRowStyle As RowStyle
    Dim typCellStyle As CellStyle
        
    Dim sTempLink As String
    Dim sTempName As String
    Dim sLink As String
    
    Dim ofwRefdata As New FWReferenceData
    Dim sHTML As String
    
    sSource = "VWUser.ShowAllUsersScreen()"
    
    'Check the permissions
    CheckPermissions ESecurityCheck.EUserCanMaintainUsers, sSource, ""
    
    'Instantiate the BO's
    Set oBOUser = New BOUser
    Set oboUserRoles = New BoUserRoles
    
    oboUserRoles.UserID = m_lLoggedUserID
    'check if the user has Administrator role
    If Not (oboUserRoles.UserHasAdministratorsRole) Then
    'if not disallow entry
        StatusInfo LoadResString(EUserMaintainResource.umInvalidPrivileges), _
                   LoadResString(EUserMaintainResource.umNeedToBeAdministrator)
        Exit Sub
    End If
     
    sTempName = ""
    Set ofwHTML = New FWHTML
    Set ofwTable = New FWTable

    With typTableStyle
        .eStyle = TABCustomStyle
        .lCELLSPACING = 2
        .sWidth = "90%"
    End With
    ofwTable.TableStyle = typTableStyle

    typRowStyle = ofwHTML.GetRowStyle(RowBackgroundRed)

    ofwTable.StartRow typRowStyle

    With typCellStyle
        .eStyle = CELLCustomStyle
        .eALIGN = HALLeft
        .eBGCOLOR = COLOURWhite
        .eVALIGN = VALMiddle
    End With

    ofwTable.AddCellAsHTML ofwHTML.Font(LoadResString(EUserMaintainResource.umUserMaintenance), _
                                    COLOURBlack, 4), typCellStyle
    ofwTable.EndRow
    
    ofwTable.AddBlankRow
    
    ofwTable.StartRow typRowStyle
    ofwTable.AddCellAsHTML ofwHTML.Font(LoadResString(EUserMaintainResource.umAddUsers), _
                                    COLOURBlack), typCellStyle
    
    'Add button with link to Add a new user
    typCellStyle.eALIGN = HALRight
    typCellStyle.eVALIGN = VALMiddle
    ofwTable.AddCellAsHTML ofwHTML.ImageButton(gk_sImagesPath & gk_sImageAdd, _
                                    "AddUser", gk_sLinkNewUser, , ofwRefdata.Search(gToolTip, ETOOLTIP.ttAddNewUser)), typCellStyle
                           
    
    ofwTable.EndRow
    ofwTable.AddBlankRow
    sHTML = sHTML & k_sParagraph & ofwTable.GetHTML
    
    Set ofwUserList = oBOUser.AllUserNames
        
    'Set up the links to users based on first letter of their names
   
    sHTML = sHTML & k_sParagraph & AlphabetLinks(ofwUserList) & k_sParagraph
    
    ofwTable.Reset

    sHTML = sHTML & ofwHTML.FormStart(gk_sLinkUserDetail, "", "frmShowAllUsers", MPost)

    'Clear settings for the table
    ofwTable.Reset
    
    'Clear settings for TableStyle
    typTableStyle = ofwHTML.GetTableStyle(TABNoStyle)
    
    With typTableStyle
        .eStyle = TABCustomStyle
        .lCELLSPACING = 2
    End With

    'Set TableStyle attributes
    ofwTable.TableStyle = typTableStyle
    
    typRowStyle = ofwHTML.GetRowStyle(ROWPlainStyle)
    
    ofwTable.StartRow typRowStyle
    typCellStyle = ofwHTML.GetCellStyle(CELLColumnHeadingStyle)
    
    ofwTable.AddCellAsHTML ofwHTML.Font(LoadResString(EUserMaintainResource.umUserName), _
                                    COLOURWhite), typCellStyle
    
    ofwTable.EndRow
    
    'Body settings for RowStyle
    typRowStyle = ofwHTML.GetRowStyle(ROWPlainStyle)
    typCellStyle = ofwHTML.GetCellStyle(CELLCustomStyle)
    'Create business object BOUserMaintain and fetch the list of users
   
    ofwUserList.MoveFirst
    sTempName = ""
    sTempLink = ""
    
    'Loop thru' list of users in the RecList object
    Do Until ofwUserList.IsAtEnd
        ofwTable.StartRow typRowStyle
        With typCellStyle
            .eBGCOLOR = COLOURPink
            .eALIGN = HALLeft
            .eVALIGN = VALMiddle
        End With
        
        sTempName = ofwUserList.GetFieldValue(gk_sDBCOLUserName)
        
        typCellStyle.sWidth = "50%"
        
        'For each name find the first letter and setup a link to it,this will
        'be referenced by a set of alphabetical links at the top of the page
        
        'if the sTempName which stores the first letter of the name is the same
        'as the previous letter do not set up the link..if not then place a bookmark against
        'that name
        If UCase(Left(sTempName, 1)) <> sTempLink Then
            sTempLink = UCase(Left(sTempName, 1))
             ofwTable.AddCellAsHTML ofwHTML.Font(ofwHTML.Hyperlink(, sTempName, "", _
                                            "#" & sTempLink), COLOURBlack), typCellStyle
        Else
             ofwTable.AddCellAsHTML ofwHTML.Font(sTempName, COLOURBlack), typCellStyle
        End If
        
        'ofwTable.AddEmptyCell
        
        With typCellStyle
            .eBGCOLOR = COLOURWhite
            .eALIGN = HALCenter
            .sWidth = "10%"
        End With
        
        'Add image buttons with links to show details and delete users
        'the UserID is sent as a part of the QueryString so we can take it
        'from the destination form.
        'If user is managing editor do not show the remove button
        oboUserRoles.UserID = ofwUserList.GetFieldValue(gk_sDBCOLUserID)
        If oboUserRoles.UserHasManagingEditorRole Then
            
            ofwTable.AddCellAsHTML ofwHTML.ImageButton(gk_sImagesPath & _
                        gk_sImageEdit, "", gk_sLinkUserDetail & "?" & gk_sDBCOLUserID & _
                        "=" & ofwUserList.GetFieldValue(gk_sDBCOLUserID), , "Edit details of " & sTempName), typCellStyle
            ofwTable.AddCellAsHTML ofwHTML.Font("Managing editor", EColour.COLOURBlack), typCellStyle
        Else
            ofwTable.AddCellAsHTML ofwHTML.ImageButton(gk_sImagesPath & _
                        gk_sImageEdit, "", gk_sLinkUserDetail & "?" & gk_sDBCOLUserID & _
                        "=" & ofwUserList.GetFieldValue(gk_sDBCOLUserID), , "Edit details of " & sTempName), typCellStyle
            ofwTable.AddCellAsHTML ofwHTML.ImageButton(gk_sImagesPath & gk_sImageRemove, "", _
                        gk_sLinkDeleteUser & "?" & gk_sDBCOLUserID & _
                        "=" & ofwUserList.GetFieldValue(gk_sDBCOLUserID), , "Delete user " & sTempName), typCellStyle
        End If
        ofwUserList.MoveNext
        ofwTable.EndRow
        
    Loop
        
    sHTML = sHTML & ofwTable.GetHTML
    sHTML = sHTML & "</form>" & k_sParagraph
    
    'd22 Done button
    sHTML = sHTML & gk_sHTMLLineBreak & ofwHTML.ImageButton(gk_sImagesPath & gk_sImageDone, , _
        gk_sMaintenanceHome, gk_sMainFrame, "Go to maintenance home")
    
    'Add context sensitive help button
    sHTML = sHTML & HTMLToAddHelpButton(gk_sHelpLinkToMaintUsers)
    
    'send the HTML outside
    m_aspResponse.Write sHTML
    
    m_objObjectContext.SetComplete
    
    Set oboUserRoles = Nothing
    Set oBOUser = Nothing
    Set ofwHTML = Nothing
    Set ofwTable = Nothing
    Set ofwRefdata = Nothing
     
    Exit Sub
ErrorHandler:
    'If there is an error then abort and release the object context
    m_objObjectContext.SetAbort
    ShowError sSource
    Set oboUserRoles = Nothing
    Set oBOUser = Nothing
    Set ofwHTML = Nothing
    Set ofwTable = Nothing
    Set ofwRefdata = Nothing
End Sub
Public Sub DisplayUserDetail()
'Description    :   Displays a screen with the details of an existing user.User comes to this form
'                   from ShowUser.asp where he chooses to see the detail of a single user.
'                   Picks the UserID via a QueryString() method
'Amendment History
'd1   CRM            09Jul1999          Created
'd7   CRM            09Aug1999          Moved from VWUserEditSet
'd11  AGA            06Oct1999          Managing Editor role can not be unchecked, added help button
'                                                       Added security
'd18  AGA            11Dec1999          Added validation for user ticket
'd24   wdp     16Feb2000       Removed need for user ticket field
On Error GoTo ErrorHandler

    Dim sSource As String
    
    Dim oboUserRoles As BoUserRoles
    Dim oBOUser As BOUser
    Dim ofwHTML As FWHTML
    Dim ofwTable As FWTable
    Dim ofwUserDet As FWRecordList
    'Not used
    Dim ofwUserRoles As FWRecordList
    
    Dim bChecked As Boolean
    
    Dim typTableStyle  As TableStyle
    Dim typRowStyle As RowStyle
    Dim typCellStyle As CellStyle
    Dim lRoleID
    
    Dim sHTML As String
    
    Dim bNewUser As Boolean
    Dim lUserID As Long
    
    Dim ofwRefdata As New FWReferenceData
    
    sSource = "VWUser.DisplayUserDetail()"
    
    'Check the permissions
    CheckPermissions ESecurityCheck.EUserCanMaintainUsers, sSource, ""
    
    Set oBOUser = New BOUser
    Set oboUserRoles = New BoUserRoles
    
    'Used Now to Pick the UserID from the QueryString and assign to BO
    oBOUser.Load CLng(m_aspRequest.QueryString(gk_sDBCOLUserID))
    
    'fetch all the roles in the database
    Set ofwUserRoles = oboUserRoles.AllRoles
       
    
    Set ofwHTML = New FWHTML
    Set ofwTable = New FWTable

    With typTableStyle
        .eStyle = TABCustomStyle
        .lCELLSPACING = 2
        .sWidth = "100%"
    End With
    ofwTable.TableStyle = typTableStyle

    typRowStyle = ofwHTML.GetRowStyle(ROWPlainStyle)

    ofwTable.StartRow typRowStyle

    With typCellStyle
        .eStyle = CELLCustomStyle
        .eALIGN = HALLeft
        .sWidth = "50%"
        .lCOLSPAN = 2
        .eBGCOLOR = COLOURWhite
    End With
    
    ofwTable.AddCellAsHTML ofwHTML.Font(LoadResString(EUserMaintainResource.umDetailsof) & "  " & _
                                        oBOUser.UserName, _
                                        COLOURBlack, 4)
    ofwTable.EndRow
    ofwTable.AddBlankRow
    'Reset settings for cell style
    typCellStyle = ofwHTML.GetCellStyle(CELLNoStyle)

    typCellStyle.eStyle = CELLCustomStyle
    typCellStyle.eVALIGN = VALTop
    typCellStyle.sWidth = "80%"
    typCellStyle.lCOLSPAN = 2
    typCellStyle.eBGCOLOR = COLOURWhite
    
    ofwTable.AddCellAsHTML LoadResString(EUserMaintainResource.umModifyUserDetailsAndClick) & gk_sHTMLNonBreakSpace & _
                            ofwHTML.Bold(LoadResString(EUserMaintainResource.umOK)) & gk_sHTMLNonBreakSpace & _
                            LoadResString(EUserMaintainResource.umToUpdate)
    
    ofwTable.EndRow

    sHTML = sHTML & k_sParagraph & ofwTable.GetHTML
    
    sHTML = sHTML & ofwHTML.FormStart(gk_sLinkUpdateUser, "", "frmUserDetail", MPost, True)
    
    'Clear settings for the table
    ofwTable.Reset

    'Clear settings for TableStyle
    typTableStyle = ofwHTML.GetTableStyle(TABNoStyle)
    
    With typTableStyle
        .eStyle = TABCustomStyle
        .sWidth = "100%"
        .lCELLSPACING = 2
        .lCELLPADDING = 2
    End With

    'Set TableStyle attributes
    ofwTable.TableStyle = typTableStyle

    'Body settings for RowStyle
    typRowStyle = ofwHTML.GetRowStyle(RowBackgroundPink)

    'Field UserName
    ofwTable.StartRow typRowStyle

    With typCellStyle
        .eBGCOLOR = COLOURWhite
        .sWidth = "15%"
        .eALIGN = HALLeft
    End With
    
    'Field User Name
    ofwTable.AddCellAsHTML ofwHTML.Font(LoadResString(EUserMaintainResource.umUserName), _
                                    COLOURRed), typCellStyle
    
    With typCellStyle
        .eBGCOLOR = COLOURPink
        .eALIGN = HALLeft
        .sWidth = "87%"
    End With
    ofwTable.AddCellAsHTML ofwHTML.InputField(ITText, gk_sDBCOLUserName, gk_iDBCOLLenUserName, gk_iDBCOLLenUserName, _
                                    oBOUser.UserName & _
                                    "", VCNone, True)
    
    ofwTable.EndRow
    
    'Field UserInitials
    ofwTable.StartRow typRowStyle

    With typCellStyle
        .eBGCOLOR = COLOURWhite
        .sWidth = "30%"
        .eALIGN = HALLeft
    End With
    
    'field User Initials
    ofwTable.AddCellAsHTML ofwHTML.Font(LoadResString(EUserMaintainResource.umInitials), COLOURRed), _
                                    typCellStyle
    
    With typCellStyle
        .eBGCOLOR = COLOURPink
        .eALIGN = HALLeft
        .sWidth = "60%"
    End With
    ofwTable.AddCellAsHTML ofwHTML.InputField(ITText, gk_sDBCOLUserInitials, 5, 5, _
                                    oBOUser.UserInitials _
                                    & "", VCNone, True)
    
    ofwTable.EndRow
        
    'Field NTLogOn
    ofwTable.StartRow typRowStyle

    With typCellStyle
        .eBGCOLOR = COLOURWhite
        .sWidth = "30%"
        .eALIGN = HALLeft
    End With

    ofwTable.AddCellAsHTML ofwHTML.Font(LoadResString(EUserMaintainResource.umNTLogOn), _
                                    COLOURRed), typCellStyle
    
    With typCellStyle
        .eBGCOLOR = COLOURPink
        .eALIGN = HALLeft
        .sWidth = "60%"
    End With
    ofwTable.AddCellAsHTML ofwHTML.InputField(ITText, gk_sDBCOLUserNTlogon, 20, 20, _
                                    oBOUser.UserNTLogon _
                                    & "", VCNone, True)
    
    ofwTable.EndRow
    
    ' d17   CRM     10Dec1999
    'Field Ticket number
    'No longer needed
    'ofwTable.StartRow typRowStyle

    'With typCellStyle
    '    .eBGCOLOR = COLOURWhite
    '    .sWidth = "20%"
    '    .eALIGN = HALLeft
    'End With
    'ofwTable.AddCellAsHTML ofwHTML.Font("Ticket Number", COLOURRed), _
                                    typCellStyle
    'With typCellStyle
    '    .eBGCOLOR = COLOURPink
    '    .eALIGN = HALLeft
    '    .sWidth = "80%"
    'End With
    'd18 Added validation
    'If oBOUser.UserTicket = "" Or IsNull(oBOUser.UserTicket) Then
    'ofwTable.AddCellAsHTML ofwHTML.InputField(EInputType.ITText, gk_sDBCOLReaderUserTicket, 7, 7, , EValidCheck.VCNumber, False, False, 0, 9999999), typCellStyle
    'Else
    'ofwTable.AddCellAsHTML ofwHTML.InputField(EInputType.ITText, gk_sDBCOLReaderUserTicket, 7, 7, oBOUser.UserTicket, EValidCheck.VCNumber, False, False, 0, 9999999), typCellStyle
    'End If
    'ofwTable.EndRow
    
   
    sHTML = sHTML & k_sParagraph & ofwTable.GetHTML
    
    ofwTable.Reset
     
    typTableStyle = ofwHTML.GetTableStyle(TABNoStyle)
    With typTableStyle
        .eStyle = TABCustomStyle
        .lCELLSPACING = 2
        .sWidth = "80%"
    End With
    ofwTable.TableStyle = typTableStyle
    
    typRowStyle = ofwHTML.GetRowStyle(ROWPlainStyle)
    ofwTable.StartRow typRowStyle
    
    typCellStyle = ofwHTML.GetCellStyle(CELLNoStyle)
    
    With typCellStyle
        .eStyle = CELLCustomStyle
        .eALIGN = HALLeft
        .sWidth = "80%"
        .eBGCOLOR = COLOURWhite
    End With
    ofwTable.AddCellAsHTML _
                        ofwHTML.Font(LoadResString(EUserMaintainResource.umUserRolesCheckUnCheckToAddRemoveRoleFromUser), _
                                    COLOURBlack), typCellStyle
    ofwTable.EndRow
   
    sHTML = sHTML & k_sParagraph & ofwTable.GetHTML
    
    ofwTable.Reset
    
    With typTableStyle
        .eStyle = TABCustomStyle
        .sWidth = "100%"
        .lCELLSPACING = 2
        .lCELLPADDING = 0
    End With
    ofwTable.TableStyle = typTableStyle
    typRowStyle = ofwHTML.GetRowStyle(RowBackgroundPink)
    
    ofwUserRoles.MoveFirst
    'ofwUserDet.MoveFirst
    'Loop thru' the UserRoles RecordList and display all Roles
    Do Until ofwUserRoles.IsAtEnd
        ofwTable.StartRow typRowStyle
        With typCellStyle
            .eBGCOLOR = COLOURWhite
            .sWidth = "16%"
            .eALIGN = HALLeft
            .lCOLSPAN = 1
        End With
        ofwTable.AddCellAsHTML ofwHTML.Font(ofwUserRoles.GetFieldValue(gk_sDBCOLRoleDescription), _
                                        COLOURRed), typCellStyle
    
        With typCellStyle
            .eBGCOLOR = COLOURPink
            .eALIGN = HALLeft
            .sWidth = "65%"
            .lCOLSPAN = 1
        End With
        
        'Check if the role is present for the User
        'The output of CheckRole function decides if the CheckBox will be checked or not
        bChecked = CheckRole(oBOUser, ofwUserRoles.GetFieldValue(gk_sDBCOLRoleID))
        
        'd8  If Editor role is checked, do not show the readonly check box(image)
        If ofwUserRoles.GetFieldValue(gk_sDBCOLRoleID) = gk_lManagingEditorRoleID And bChecked = True Then
                                  
            ofwTable.AddCellAsHTML gk_sHTMLNonBreakSpace & ofwHTML.Image( _
                "", gk_sImagesPath & gk_sImageCheckBoxCheckedDisabled)
            'Show the value in hidden field
            sHTML = sHTML & ofwHTML.HiddenField(gk_sDBCOLRoleID & _
                                        ofwUserRoles.GetFieldValue(gk_sDBCOLRoleID), ofwUserRoles.GetFieldValue(gk_sDBCOLRoleID))
        Else
        
            ofwTable.AddCellAsHTML ofwHTML.Checkbox("", gk_sDBCOLRoleID & _
                                        ofwUserRoles.GetFieldValue(gk_sDBCOLRoleID), _
                                        bChecked, ofwUserRoles.GetFieldValue(gk_sDBCOLRoleID))

       End If
        ofwUserRoles.MoveNext
        ofwTable.EndRow
    Loop
    

    
    sHTML = sHTML & k_sParagraph & ofwTable.GetHTML
    
    'Clear table settings
    ofwTable.Reset
    typTableStyle = ofwHTML.GetTableStyle(TABNoStyle)
    
    
    typRowStyle = ofwHTML.GetRowStyle(ROWPlainStyle)
    
    ofwTable.StartRow typRowStyle
    
    With typCellStyle
        .eALIGN = HALLeft
        .eBGCOLOR = COLOURWhite
        .eVALIGN = VALMiddle
    End With
    
    'add buttons to confirm the changes and cancel the changes
    With ofwTable
        'this procedure calls a JavaScript function called MECheck(),Managing Editor check
        'on the client side,this requires the prescence of the MECheck.js file
        'to be present on the browser
        .AddCellAsHTML ofwHTML.SubmitImage(gk_sImagesPath & gk_sImageOK, "ok", "OK", _
                                        "return AnotherMEQuery()") & ofwHTML.ImageButton(gk_sImagesPath & gk_sImageCancel, "Cancel", _
                                        gk_sLinkShowUsers, , ofwRefdata.Search(gToolTip, ETOOLTIP.ttButtonCancelHelp)), typCellStyle
                   
        .EndRow
    End With
    
    'Have a hidden field to store the UserID
    sHTML = sHTML & ofwHTML.HiddenField(gk_sDBCOLUserID, oBOUser.UserID)
    
    'have a hidden field to store the current Managing Editor name
    sHTML = sHTML & ofwHTML.HiddenField("Hid" & gk_lManagingEditorRoleID, _
                                         oBOUser.GetManagingEditor)
    sHTML = sHTML & k_sParagraph & ofwTable.GetHTML
    
    sHTML = sHTML & "</form>"
    
    'Add context sensitive help button
    sHTML = sHTML & HTMLToAddHelpButton(gk_sHelpLinkToMaintUsers)
    
    'write the HTML to the browser
    m_aspResponse.Write sHTML
    
    Set ofwHTML = Nothing
    Set ofwTable = Nothing
    Set oBOUser = Nothing
    Set ofwUserDet = Nothing
    Set ofwUserRoles = Nothing
    Set ofwRefdata = Nothing
    
    m_objObjectContext.SetComplete
    
    Exit Sub
ErrorHandler:
    m_objObjectContext.SetAbort
    ShowError sSource
    Set ofwHTML = Nothing
    Set ofwTable = Nothing
    Set oBOUser = Nothing
    Set ofwUserDet = Nothing
    Set ofwUserRoles = Nothing
    Set ofwRefdata = Nothing
End Sub
Private Function CheckRole(oBOUser As BOUser, ByVal t_lRoleID As Long) As Boolean
'Description    :   Checks to see if a role is present for the User.
'                   Given a RoleID the function checks if it is present in the UserDetail
'                   RecorList...if present returns True.
'                   Takes the BOUser object as an argument and retrieves the Roles property
'                   for comparison on whether the particular role is present or not

'Amendment History :
'd1   CRMahesh            July 9,1999            Created
'd5   CRMahesh            July 21,1999           Modified to use the BO's UserHasRoles collection
'd7   CRMahesh            Aug 09,1999            Moved functionality from VWUserEditSet

On Error GoTo ErrorHandler
    Dim bFound As Boolean
    Dim sSource As String
    Dim vItem
    Dim vRolesArray As Variant
    Dim lIndex As Long
    
    sSource = "VWUser.CheckRole()"
    bFound = False
    
    vRolesArray = oBOUser.Roles
    
    'Loop thru the Roles property of the BO which returns a array,if found
    'set bFound to TRUE
    For lIndex = 0 To UBound(vRolesArray) - 1
           If t_lRoleID = vRolesArray(lIndex) Then
                 bFound = True
                 Exit For
           Else
                 bFound = False
           End If
    Next
    
    'Return status,this decides if the check box would be checked
    CheckRole = bFound
    
Exit Function

ErrorHandler:
    m_objObjectContext.SetAbort
    ShowError sSource
End Function
Public Sub DeleteUser()
'Description    :  Deletes the user,if not part of an existing edit set.User can be deleted
'                  from the ShowUsers.asp page
'                  If the any of the User's roles are not part of an EditSet then ShowUsers.asp
'                  page is displayed or a message is given.
'AmendmentHistory:
'd1     CRMahesh              July 16,1999    Modified to display EditSetName in status message
'd7     CRMahesh              Aug 09,1999     Moved functionality from VWUserEditSet class
'd11    aga                      Oct 11, 1999    Added security
On Error GoTo ErrorHandler

    Dim sSource As String
    
    Dim oBOEditSet As BOEditSet
    Dim oBOUser As BOUser
    Dim lEditStatus As Long
    Dim ofwHTML As FWHTML
    sSource = "VWUserEditSet.DeleteUser()"
    
    'Check the permissions
    CheckPermissions ESecurityCheck.EUserCanMaintainUsers, sSource, ""
    
    Set ofwHTML = New FWHTML
    Set oBOUser = New BOUser
    Set oBOEditSet = New BOEditSet
    
    'DeleteUser method of BO returns the RoleID,if the user's role
    'is part of an edit-set or it deletes the User and returns 0
         
    lEditStatus = oBOUser.DeleteUser(m_aspRequest.QueryString(gk_sDBCOLUserID))
    
    
    'Check if any RoleID is returned
    If lEditStatus > 0 Then
        oBOEditSet.EditsetID = lEditStatus
        oBOEditSet.Load
        StatusInfo LoadResString(EUserMaintainResource.umDeleteStatus), _
                    LoadResString(EUserMaintainResource.umUnableToCompleteRequestUserPartOfExisting) & " " & _
                    ofwHTML.Strong(oBOEditSet.EditSetName) & " " & _
                    LoadResString(EUserMaintainResource.umEditSet), gk_sLinkShowUsers
    Else
        m_aspResponse.Redirect gk_sLinkShowUsers
    End If
    
    Set ofwHTML = Nothing
    Set oBOUser = Nothing
    Set oBOEditSet = Nothing
    
    m_objObjectContext.SetComplete
        
    Exit Sub
ErrorHandler:
    m_objObjectContext.SetAbort
    ShowError sSource
    Set ofwHTML = Nothing
    Set oBOUser = Nothing
    Set oBOEditSet = Nothing
End Sub

Public Sub UserUpdateDisplay()
'Description    :  Updates the user information.
'                  Picks data from either NewUser.asp or UserDetail.asp.
'                  If it finds a UserID via the Request method then
'                  its a modifiation or its a new user which has to be entered
'Amendment History :
' d2   CRMahesh            July 13,1999         method UserUpdate()
' d5   CRMahesh            July 21,1999         Whole functionality to check current roles
'                                               moved to BO.
'                                               Should call ResetRoles to clear the collection and
'                                               Add the choosen roles to the UserHasRoles
'                                               property and call update method on BO
'd7     CRMahesh           Aug 09,1999          Moved functionality from VWUserEditSet class
'd9     crm                 12Aug1999            modified the rolesarray value assigned
'd11    aga                Oct 11, 1999         Added security
'd12    aga                28Oct1999            Added code to handle deletion of last procat manager
'd13    aga                29Oct1999            Handling duplication of User name, NT Logon and User initials
'd24    wdp                16Feb2000            Removed need for user ticket field
On Error GoTo ErrorHandler

    Dim sSource As String
    
    Dim oBOUser As BOUser
    Dim ofwHTML As FWHTML
    Dim sUserName As String
    Dim sNTLogOn As String
    
    Dim lRoleID As Long
    Dim sUserInitials As String
    Dim sErrStatus As String
    Dim bMEditorIsThere As Boolean
    Dim lRoleArray() As Long
    Dim vItem As Variant
    Dim lIndex As Long
    
    
    sSource = "VWUserEditSet.UserUpdateDisplay()"
    
    'Check the permissions
    CheckPermissions ESecurityCheck.EUserCanMaintainUsers, sSource, ""
    
    Set ofwHTML = New FWHTML
    Set oBOUser = New BOUser
    
    'Collect all the required values from the form
    sUserName = m_aspRequest.Form(gk_sDBCOLUserName)
    sNTLogOn = m_aspRequest.Form(gk_sDBCOLUserNTlogon)
    sUserInitials = m_aspRequest.Form(gk_sDBCOLUserInitials)
    
    'User has not entered required fields
    If Trim(sUserName) = "" Or Trim(sNTLogOn) = "" Or Trim(sUserInitials) = "" Or _
                IsNumeric(sUserName) Or IsNumeric(sUserInitials) Or IsNumeric(sNTLogOn) Then
                
        StatusInfo LoadResString(EUserMaintainResource.umIncomplete), _
                    LoadResString(EUserMaintainResource.umYouhavenotenteredrequiredfields)
        Exit Sub
    End If
    
    'Try to get the UserID which is hidden field in the UserDetail form
    '..check the UserDetail method of this class,assign the properties
    With oBOUser
        'If the userid is greater than 0 the BO fetches the details from Database,
        'assign the new properties
        .UserID = m_aspRequest.Form(gk_sDBCOLUserID)
        .UserName = sUserName
        .UserInitials = sUserInitials
        .UserNTLogon = sNTLogOn
        
    End With
       
        'Enumerate items on form
        For Each vItem In m_aspRequest.Form
        'Check if the User Role CheckBox has come
            If InStr(1, vItem, gk_sDBCOLRoleID, vbTextCompare) > 0 Then
            'If its a Role Check box then add it to UserHasRoles property
                    lIndex = lIndex + 1
                    ReDim Preserve lRoleArray(lIndex)
                    lRoleArray(UBound(lRoleArray) - 1) = CLng(Mid(vItem, _
                        Len(gk_sDBCOLRoleID) + 1))
            End If
        Next
        ' d9    crm     12Aug1999    modified the rolesarray value
        'if the array has not been redimed assign null to the roles property
        'of the business object.
        oBOUser.Roles = IIf(lIndex > 0, lRoleArray, Null)
        
            'Update the BO, if it returns false....retrive the status from BO
            If Not (oBOUser.Save) Then
                    StatusInfo LoadResString(EUserMaintainResource.umUserUpdateStatus), _
                                LoadResString(EUserMaintainResource.umRoles) & _
                                ofwHTML.Strong(oBOUser.Status) & _
                                LoadResString(EUserMaintainResource.umIsPartOfEditSetRemoveUser), _
                                                gk_sLinkShowUsers
                
            Else
                        'if proper redirect to AllUsers page
                        m_aspResponse.Redirect gk_sLinkShowUsers
                
            End If

    'destroy object reference
    Set oBOUser = Nothing
    
    'commit and release ObjectContext reference
    m_objObjectContext.SetComplete
    
    Exit Sub
ErrorHandler:
    'd12
    If Err.Number = EPROCATINTERNALERROR.peCanNotDeleteLastProcatManager Then
        'someone is trying to delete the last procat manager, show a message
        PROCATEditorialErrorHandling EPROCATINTERNALERROR.peCanNotDeleteLastProcatManager, _
            sSource, "Cannot delete the only PROCAT Manager"
        m_objObjectContext.SetComplete
        Exit Sub
    End If
    'd13
    If Err.Number = EPROCATINTERNALERROR.peDuplicateUserName Then
        'Duplicate user name, show a message
        PROCATEditorialErrorHandling EPROCATINTERNALERROR.peDuplicateUserName, _
            sSource, "User name already exists"
        m_objObjectContext.SetComplete
        Exit Sub
    End If
    If Err.Number = EPROCATINTERNALERROR.peDuplicateUserNTLogon Then
        'Duplicate user NT Logon, show a message
        PROCATEditorialErrorHandling EPROCATINTERNALERROR.peDuplicateUserNTLogon, _
            sSource, "User NT Logon already exists"
        m_objObjectContext.SetComplete
        Exit Sub
    End If
    If Err.Number = EPROCATINTERNALERROR.peDuplicateUserInitials Then
        'Duplicate user initials, show a message
        PROCATEditorialErrorHandling EPROCATINTERNALERROR.peDuplicateUserInitials, _
            sSource, "User Initials already exists"
        m_objObjectContext.SetComplete
        Exit Sub
    End If
     m_objObjectContext.SetAbort
     ShowError sSource
     Set oBOUser = Nothing
End Sub
Public Sub NewUserEntryScreen()
'Description    :   Displays empty screen to add an user and details
'Amendment History
'd1      crm     09Jul1999          Created
'd7      crm     09Aug1999          Moved functionality from VWUserEditSet class
'd11     aga     08Oct1999          Added help button, added security
'd18     AGA     11Dec1999          Added validation for user ticket
'd23     crm     11Feb2000          NewUserEntryScreen made to validate if ME role is assigned
'                                   does not validate if cancel button is choosen
'd24     wdp     16Feb2000       Removed need for user ticket field
On Error GoTo ErrorHandler

    Dim sSource As String
    
    Dim oBOUser As BOUser
    Dim ofwHTML As FWHTML
    Dim ofwTable As FWTable
    Dim oboUserRoles As BoUserRoles
    
    Dim ofwUserRoles As FWRecordList
    
    Dim bChecked As Boolean
    
    Dim typTableStyle  As TableStyle
    Dim typRowStyle As RowStyle
    Dim typCellStyle As CellStyle
    Dim lRoleID
    Dim sHTML As String
    
    Dim bNewUser As Boolean
    Dim lUserID As Long
    
    sSource = "VWUserEditSet.NewUserEntryScreen()"
    
    'Check the permissions
    CheckPermissions ESecurityCheck.EUserCanMaintainUsers, sSource, ""
    
    Set oBOUser = New BOUser
    Set oboUserRoles = New BoUserRoles
    
    Set ofwUserRoles = oboUserRoles.AllRoles
    
    Set ofwHTML = New FWHTML
    Set ofwTable = New FWTable

    With typTableStyle
        .eStyle = TABCustomStyle
        .lCELLSPACING = 2
        .sWidth = "50%"
    End With
    ofwTable.TableStyle = typTableStyle

    typRowStyle = ofwHTML.GetRowStyle(RowBackgroundRed)

    ofwTable.StartRow typRowStyle

    With typCellStyle
        .eStyle = CELLCustomStyle
        .eALIGN = HALLeft
        .sWidth = "50%"
        .lCOLSPAN = 2
        .eBGCOLOR = COLOURWhite
    End With
    
    ofwTable.AddCellAsHTML ofwHTML.Bold(LoadResString(EUserMaintainResource.umNewUser)), typCellStyle
    
    ofwTable.EndRow
    ofwTable.AddBlankRow
    'Reset settings for cell style
    typCellStyle = ofwHTML.GetCellStyle(CELLNoStyle)

    typCellStyle.eStyle = CELLCustomStyle
    typCellStyle.eVALIGN = VALTop
    typCellStyle.sWidth = "80%"
    typCellStyle.lCOLSPAN = 2
    typCellStyle.eBGCOLOR = COLOURWhite
    
    ofwTable.AddCellAsHTML LoadResString(EUserMaintainResource.umAddUserDetailsAndClick) & gk_sHTMLNonBreakSpace & _
                           ofwHTML.Bold(LoadResString(EUserMaintainResource.umOK)) & gk_sHTMLNonBreakSpace & _
                           LoadResString(EUserMaintainResource.umToUpdate)
    
    ofwTable.EndRow

    sHTML = sHTML & k_sParagraph & ofwTable.GetHTML
    
    sHTML = sHTML & ofwHTML.FormStart(gk_sLinkUpdateUser, "", "frmUserDetail", MPost, True)
    
    'Clear settings for the table
    ofwTable.Reset

    'Clear settings for TableStyle
    typTableStyle = ofwHTML.GetTableStyle(TABNoStyle)
    
    With typTableStyle
        .eStyle = TABCustomStyle
        .sWidth = "100%"
        .lCELLSPACING = 2
        .lCELLPADDING = 2
    End With

    'Set TableStyle attributes
    ofwTable.TableStyle = typTableStyle

    'Body settings for RowStyle
    typRowStyle = ofwHTML.GetRowStyle(RowBackgroundPink)

    'Field UserName
    ofwTable.StartRow typRowStyle

    With typCellStyle
        .eBGCOLOR = COLOURWhite
        .sWidth = "15%"
        .eALIGN = HALLeft
    End With
    ofwTable.AddCellAsHTML ofwHTML.Font(LoadResString(EUserMaintainResource.umUserName), COLOURRed), _
                                    typCellStyle
    
    With typCellStyle
        .eBGCOLOR = COLOURPink
        .eALIGN = HALLeft
        .sWidth = "85%"
    End With
    
    
    ofwTable.AddCellAsHTML ofwHTML.InputField(ITText, gk_sDBCOLUserName, gk_iDBCOLLenUserName, gk_iDBCOLLenUserName, , VCNone, True)
    ofwTable.EndRow
    
    'Field UserInitials
    ofwTable.StartRow typRowStyle

    With typCellStyle
        .eBGCOLOR = COLOURWhite
        .sWidth = "20%"
        .eALIGN = HALLeft
    End With
    ofwTable.AddCellAsHTML ofwHTML.Font(LoadResString(EUserMaintainResource.umInitials), COLOURRed), _
                            typCellStyle

    With typCellStyle
        .eBGCOLOR = COLOURPink
        .eALIGN = HALLeft
        .sWidth = "80%"
    End With
    
    
    ofwTable.AddCellAsHTML ofwHTML.InputField(ITText, gk_sDBCOLUserInitials, gk_iDBCOLLenUserInitials, gk_iDBCOLLenUserInitials, , VCNone, _
                                    True)
    ofwTable.EndRow
        
    'Field NTLogOn
    ofwTable.StartRow typRowStyle

    With typCellStyle
        .eBGCOLOR = COLOURWhite
        .sWidth = "20%"
        .eALIGN = HALLeft
    End With
    ofwTable.AddCellAsHTML ofwHTML.Font(LoadResString(EUserMaintainResource.umNTLogOn), COLOURRed), _
                                    typCellStyle

    With typCellStyle
        .eBGCOLOR = COLOURPink
        .eALIGN = HALLeft
        .sWidth = "80%"
    End With
    
   
    ofwTable.AddCellAsHTML ofwHTML.InputField(ITText, gk_sDBCOLUserNTlogon, gk_iDBCOLLenUserNTLogon, gk_iDBCOLLenUserNTLogon, , VCNone, _
                                    True)
    ofwTable.EndRow
    
    ' d17   CRM     10Dec1999
    'Field Ticket number
    'no longer needed
    'ofwTable.StartRow typRowStyle

    'With typCellStyle
    '    .eBGCOLOR = COLOURWhite
    '    .sWidth = "20%"
    '    .eALIGN = HALLeft
    'End With
    'ofwTable.AddCellAsHTML ofwHTML.Font("Ticket Number", COLOURRed), _
                                    typCellStyle
    'With typCellStyle
    '    .eBGCOLOR = COLOURPink
    '    .eALIGN = HALLeft
    '    .sWidth = "80%"
    'End With
    'd18 Added validation
    'ofwTable.AddCellAsHTML ofwHTML.InputField(EInputType.ITText, gk_sDBCOLReaderUserTicket, 7, 7, , EValidCheck.VCNumber, False, False, 0, 9999999), typCellStyle
    'ofwTable.EndRow
    
    typRowStyle = ofwHTML.GetRowStyle(RowBackgroundPink)
   
    sHTML = sHTML & k_sParagraph & ofwTable.GetHTML
    
    ofwTable.Reset
     
    typTableStyle = ofwHTML.GetTableStyle(TABNoStyle)
    With typTableStyle
        .eStyle = TABCustomStyle
        .lCELLSPACING = 2
        .sWidth = "100%"
    End With
    ofwTable.TableStyle = typTableStyle
    
    typRowStyle = ofwHTML.GetRowStyle(RowBackgroundRed)
    ofwTable.StartRow typRowStyle
    
    typCellStyle = ofwHTML.GetCellStyle(CELLNoStyle)
    
    With typCellStyle
        .eStyle = CELLCustomStyle
        .eALIGN = HALLeft
        .sWidth = "100%"
        .eBGCOLOR = COLOURWhite
    End With
    ofwTable.AddCellAsHTML ofwHTML.Font(LoadResString(EUserMaintainResource.umUserRolesCheckUnCheckToAddRemoveRoleFromUser), _
                                COLOURBlack), typCellStyle
                                    
    ofwTable.EndRow
    
    sHTML = sHTML & k_sParagraph & ofwTable.GetHTML
    
    ofwTable.Reset
    
    With typTableStyle
        .eStyle = TABCustomStyle
        .sWidth = "100%"
        .lCELLSPACING = 2
        .eBGCOLOR = COLOURWhite
    End With
    ofwTable.TableStyle = typTableStyle
    typRowStyle = ofwHTML.GetRowStyle(RowBackgroundPink)
    
    ofwUserRoles.MoveFirst
    
    'Show User roles,loop thru the UserRoles RecList
    Do Until ofwUserRoles.IsAtEnd
        ofwTable.StartRow typRowStyle
        
        With typCellStyle
            .eBGCOLOR = COLOURWhite
            .sWidth = "17%"
            .eALIGN = HALLeft
            .lCOLSPAN = 1
        End With
        
        ofwTable.AddCellAsHTML ofwHTML.Font(ofwUserRoles.GetFieldValue(gk_sDBCOLRoleDescription), COLOURRed), _
                                        typCellStyle
    
        With typCellStyle
            .eBGCOLOR = COLOURPink
            .eALIGN = HALLeft
            .sWidth = "83%"
            .lCOLSPAN = 1
        End With
               
        'for each role palce a check box,concatenate the roleid to the chekbox name
        'so that the name can be just trimmed to get the roleid from the
        'view class
        ofwTable.AddCellAsHTML ofwHTML.Checkbox("", gk_sDBCOLRoleID & _
                        ofwUserRoles.GetFieldValue(gk_sDBCOLRoleID), False, _
                        ofwUserRoles.GetFieldValue(gk_sDBCOLRoleID)), typCellStyle
       
        ofwUserRoles.MoveNext
        ofwTable.EndRow
    Loop
    
    ofwTable.AddBlankRow
    sHTML = sHTML & k_sParagraph & ofwTable.GetHTML
    
    'Clear table settings
    ofwTable.Reset
    typTableStyle = ofwHTML.GetTableStyle(TABNoStyle)
    
    typRowStyle = ofwHTML.GetRowStyle(ROWPlainStyle)
    
    ofwTable.StartRow typRowStyle
    
    With typCellStyle
        .eALIGN = HALCenter
        .eBGCOLOR = COLOURWhite
        .eVALIGN = VALMiddle
    End With
    
    'add buttons for OK and cancel
    With ofwTable
        'TODO
         .AddCellAsHTML ofwHTML.SubmitImage(gk_sImagesPath & gk_sImageOK, "ok", "OK", _
                                "return AnotherMEQuery()") & ofwHTML.ImageButton(gk_sImagesPath & gk_sImageCancel, "Cancel", _
                                gk_sLinkShowUsers), typCellStyle
        .EndRow
    End With
    
    'have hidden field for storing the Managing Editor Name , this is required for the
    'client side script
    sHTML = sHTML & ofwHTML.HiddenField("Hid" & gk_lManagingEditorRoleID, _
                            oBOUser.GetManagingEditor)
    sHTML = sHTML & k_sParagraph & ofwTable.GetHTML
    
    sHTML = sHTML & "</form>" & k_sParagraph
    
    'Add context sensitive help button
    sHTML = sHTML & HTMLToAddHelpButton(gk_sHelpLinkToMaintUsers)
    
    m_aspResponse.Write sHTML
    
    Set oBOUser = Nothing
   
    Set ofwUserRoles = Nothing
    Set ofwHTML = Nothing
    Set ofwTable = Nothing
    
    m_objObjectContext.SetComplete
    
    Exit Sub
ErrorHandler:
    m_objObjectContext.SetAbort
    ShowError sSource
    Set oBOUser = Nothing
    Set ofwUserRoles = Nothing
    Set ofwHTML = Nothing
    Set ofwTable = Nothing
End Sub
Private Sub StatusInfo(t_sTitle As String, t_sMessage As String, Optional t_sMoveToPage)
'Description    :  A generic function to display status messages
'Arguments      :  t_sTitle - The title to be displayed
'                  t_sMessgae - The message
'                  t_sMoveToPage - The page to which the 'OK' button should point to.
'Amendment History
'd2     crm         July 13,1999    created
'd7     crm         Aug 09,1999     Moved from VWUserEditSet class
'd25    wdp         24Feb2000       Added back button to StatusInfo if no OK button is displayed

On Error GoTo ErrorHandler

    Dim sSource As String
    
    
    Dim lEditStatus As Long
    Dim ofwHTML As FWHTML
    Dim ofwTable As FWTable
    
    Dim typTableStyle  As TableStyle
    Dim typRowStyle As RowStyle
    Dim typCellStyle As CellStyle
        
    Dim sHTML As String
    
    sSource = "VWUserEditSet.StatusInfo()"
        
    Set ofwHTML = New FWHTML
    Set ofwTable = New FWTable

    With typTableStyle
        .eStyle = TABCustomStyle
        .lCELLSPACING = 2
        .sWidth = "80%"
    End With
    ofwTable.TableStyle = typTableStyle

    typRowStyle = ofwHTML.GetRowStyle(ROWPlainStyle)

    ofwTable.StartRow typRowStyle

    With typCellStyle
        .eStyle = CELLCustomStyle
        .eALIGN = HALLeft
        .sWidth = "50%"
        .lCOLSPAN = 2
        .eBGCOLOR = COLOURWhite
    End With
    
    ofwTable.AddCellAsHTML ofwHTML.Bold(t_sTitle), typCellStyle

    ofwTable.EndRow
    ofwTable.AddBlankRow

    'Reset settings for cell style
    typCellStyle = ofwHTML.GetCellStyle(CELLNoStyle)
  
    ofwTable.AddCellAsHTML t_sMessage
    
    ofwTable.EndRow
    ofwTable.AddBlankRow

    sHTML = sHTML & k_sParagraph & ofwTable.GetHTML
    
    'if a link is not required then do not have the OK Image Button
    If Not IsMissing(t_sMoveToPage) Then
        sHTML = sHTML & ofwHTML.ImageButton(gk_sImagesPath & gk_sImageOK, "StatusPage", CStr(t_sMoveToPage))
    Else
        'add a back button
        sHTML = sHTML & BackButtonHTML("")
    End If
    
    m_aspResponse.Write sHTML
     
    Set ofwHTML = Nothing
    Set ofwTable = Nothing
     
    m_objObjectContext.SetComplete
     
    Exit Sub
ErrorHandler:
    m_objObjectContext.SetAbort
    ShowError sSource
    Set ofwHTML = Nothing
    Set ofwTable = Nothing
End Sub
Private Function AlphabetLinks(t_ofwUserList As FWRecordList) As String
'Description    :   To setup the links to user names
'Amendment History
' d6    crm     July 28,1999        created
'
On Error GoTo ErrorHandler
    Dim oHypCol As New Collection
    Dim iAsc As Integer
    Dim sTempLink As String
    Dim ofwHTML As New FWHTML
    Dim sLink As String
    Dim vItem As Variant
    
    'load a collection object with alphabets,and assign a key too
    For iAsc = klAscValueA To klAscValueZ
        oHypCol.Add Chr(iAsc) & " ", Chr(iAsc)
    Next
    
    t_ofwUserList.MoveFirst
    
    'loop thru the names of the users and for each first-letter,remove
    'the element from the collection i.e the alphabet and replace it with
    'a hyperlink element
    
    Do Until t_ofwUserList.IsAtEnd
        sTempLink = UCase(Left(t_ofwUserList.GetFieldValue(gk_sDBCOLUserName), 1))
        oHypCol.Remove sTempLink
        oHypCol.Add ofwHTML.Hyperlink("#" & sTempLink, sTempLink, ""), sTempLink
        t_ofwUserList.MoveNext
    Loop
    
    'loop thru the elements in collection,now there will be a hyperlink if a name
    'starting with that letter is present,concatenate all of them and return as a string
     For iAsc = klAscValueA To klAscValueZ
        sLink = sLink & oHypCol.Item(Chr(iAsc))
     Next
       
    AlphabetLinks = sLink
    
    Set ofwHTML = Nothing
Exit Function
ErrorHandler:
    Resume Next

End Function

Private Sub CheckPermissions(t_eFunctionality As ESecurityCheck, _
                         t_sSource As String, _
                        t_sDescription As String)
'Subroutine that checks and raises an error if the user has no rights to
' the passed function
'
'd11     aga     11Oct1999       Copied from VWCatalogueBulkUpdate and modified
'd14     aga     02Nov1999       Modified due to change in Check method
'd15     KSK     03Nov1999       Changed Checkpermissions method to used the parameter source
'                                 value while raising error
    Dim eCheckRights As ESecurityCheckResult
    
    m_oboEditorialPermission.UserID = m_lLoggedUserID
    
    eCheckRights = m_oboEditorialPermission.Check(t_eFunctionality)
    
    If eCheckRights <> _
            ESecurityCheckResult.ESRSecurityCheckSuccessful Then
    
        RaiseError t_sSource, _
            GetErrorEnumForSecurityCheckResult(eCheckRights), _
            GetDescriptionForSecurityCheckError(GetErrorEnumForSecurityCheckResult(eCheckRights))
    End If
        
End Sub

Public Function GetVersion() As String
'   Get version info
'
'   d19 WDP     13-Dec-1999     Created
On Error GoTo ErrorHandler
    
    Dim sSource As String
    
    sSource = "VWUser.GetVersion"
    
    GetVersion = gk_sEditorialReleaseVersion
    
    Exit Function

ErrorHandler:
    ShowError sSource
End Function

