VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BOUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    BOUser.cls
' System:       PROCAT, PRO
' Copyright:    (C) Quidnunc Limited
'
' Description:  The User business object
'
' Amendment history:
'
'   d1      ksk     27May1999      Created.
'   d2      ndk     04Jun1999      Changed the login method. Completely rewritten
'                                  Storing UserType and UserInitials in state information
'   d3      ndk     04Jun1999      Completely changed - Authentication is now done by picking up the users NT logon name
'   d4      ndk     14Jun1999      Added function GetAllUserNames,GetUserDetails
'   d5      ndk     28Jun199       Added function UserHasAdministratorsRole
'   d6      ndk     08Jul999       Changed the names of methods to adhere to naming conventions
'                                  GetUserDetails changed to UserDetails
'                                  GetAllUserNames changed to AllUserNames
'                                  Using constants in place of hard coded coloumn names
'   d7      crm     27Jul1999      Moved functionality from BOUserEditSet class
'                                  Not Using UserDetails(),read user details from properties of class
'                                  Added Load() to load the details of a user
'                                  Added 'Let' property procedures for all properties
'                                  Added AddUser(),AddRoleToUser(),RemoveRoleFormUser(),
'                                  DeleteUser(),IsUserRoleInEditSet(),Save(),Status
'                                  UpdateUser()
'   d8      crm     29Jul1999      Added GetManagingEditor()
'   d9      crm     04Aug1999      Modified Login() method to set the UserHasAdministratorRole property
'                                  Removed check on assigning values to properties to see
'                                  if they have changed their values and then assign them,and removed
'                                  the private boolean property to hold that information
'   d10      ksk    09Aug1999      Changed the CreateDelimiter function call as
'                                  Session ID is passes as a seperate ID in Transient
'                                  manager class
'   d11      crm    10Aug1999      Removed AllRoles(),using BOUserRoles directly in View class
'                                  Removed UserTypeID property,removed EditSetDesc()
'                                  Removed UserHasAdministratorRole Property
'   d12      crm    12Aug1999      Modified enumerating the role array
'   d10      ksk    16Aug1999      Removing the User type field from the business object
'   d11      aga    08Oct1999      Minor modifications in case of the role Managing Editor
'                                            (Modified DeleteUser method - Managing editor can't be deleted directly)
'   d12      aga    29Oct1999      User_ID is now IDENTITY column in database: Removing Max(ID) query in AddUser
'                                           Handling duplication of User name, NT Logon and User initials
'                                           Replacing Hardcoded constants to DBContants
'   d13      crm    10Dec1999      Added property & updated methods to handle User_Ticket
'   d14      aga    11Dec1999      Taken care of nullable UserTicket: modified the datatype of the variable to variant from long
'   d15      WDP    16Feb2000      Removed need for user_ticket
'   d16      TMJ    17Mar2000      Removed gk_sMessageDelimiter from GetManagingEditor
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
'
'Interfaces
Implements PROCATTransientInterface
Implements ObjectControl

'
'Property Variables
'
Private m_lUserID As Long
Private m_sUserNTLogon As String        'NT Logon of the user, picked up from request.servervariables
Private m_sUserName As String
Private m_sUserInitials As String
'd7      crm     27July1999
Private m_sUserTypeName As String       'Property to hold UserTypeName

'd7      crm     27Jul1999
Private Const k_sUserTypename As String = "user_type_name"
   'Const for User TypeID
Private m_sStatus As String

Private m_oboUserRoles As BoUserRoles   'BOUserRoles class instance reference
Private m_vRolesArray As Variant


'
' Properties
'
Public Property Get Name()
    Name = "BOUser"
End Property
Public Property Get UserID() As Long
    UserID = m_lUserID
End Property
Public Property Get UserTypeName() As String
    UserTypeName = m_sUserTypeName
End Property
Public Property Get UserNTLogon() As String
    UserNTLogon = m_sUserNTLogon
End Property

Public Property Get UserName() As String
    UserName = m_sUserName
End Property
Public Property Get UserInitials() As String
    UserInitials = m_sUserInitials
End Property

Public Property Get Status() As String
'Description    :   Return the status
'Amendment History  :
'd7     CRM    27Jul1999    Copied from BOUserEditSet class
    Status = m_sStatus
End Property


'Let property procedures copied from BOUserEditSet
'
Public Property Let UserName(ByVal t_sUserName As String)
On Error GoTo ErrorHandler

    Dim sSource As String
    sSource = "BOUser.UserName"
    
        m_sUserName = t_sUserName
       
Exit Property
ErrorHandler:
    ConvertError sSource
End Property

Public Property Let UserInitials(ByVal t_sUserInitials As String)
On Error GoTo ErrorHandler

    Dim sSource As String
    sSource = "BOUser.UserInitials"
    
        m_sUserInitials = t_sUserInitials
      
Exit Property
ErrorHandler:
    ConvertError sSource
End Property
Public Property Let UserNTLogon(ByVal t_sUserNTLogon As String)
On Error GoTo ErrorHandler

    Dim sSource As String
    sSource = "BOUser.UserNTLogon"
    
        m_sUserNTLogon = t_sUserNTLogon
       
Exit Property
ErrorHandler:
    ConvertError sSource
End Property


Public Property Let UserID(ByVal t_lUserID As Long)
On Error GoTo ErrorHandler
    Dim sSource As String
       
    sSource = "BOUser.UserID"
  
        m_lUserID = t_lUserID
    
    Exit Property
ErrorHandler:
    ConvertError sSource
End Property
Public Property Get Roles() As Variant
'Description    :   Property returns reference to BOUserRoles Object
'                   which gives access to User's roles related functionality
'   d7      CRM     27Jul1999       Created
       Roles = m_vRolesArray
End Property

Public Property Let Roles(ByVal RolesArray As Variant)
'Description    :   Assign the array of roleid's to the private variable

m_vRolesArray = RolesArray
End Property



' Methods
Public Sub Initialize()
    m_lUserID = 0
    Set m_oboUserRoles = New BoUserRoles
    m_sUserNTLogon = ""
    m_sUserName = ""
    m_sUserInitials = ""
    m_sStatus = ""
    m_sUserTypeName = ""
End Sub
Public Sub Terminate()
Set m_oboUserRoles = Nothing
End Sub


Public Sub Login(ByVal t_sUserNTLogon As String)
' Description:
'
' To validate login info by checking in the database
' The users name is picked from his NT login name
' If the name is empty or if it is not present in the database, the user
' is not permitted to enter the system
' Amendment history:
'
' d1    ksk     26Apr1999       Created.
' d2    ndk     05May1999       Changes made to reflect the changed in the database field names
' d3    ndk     04Jun1999       Completely changed - Authenticatin is now done by picking up the users NT logon name
' d9    crm     04Aug1999        Modified Login() method to set the UserHasAdministratorRole property
' d13   CRM        10Dec1999     Modified to handle User Ticket
' d15   WDP    16Feb2000      Removed need for user_ticket
On Error GoTo ErrorHandler

    Dim ofwDBHelper As FWDBHelper
    Dim oFWRecordlist As FWRecordList
   
    Dim sSource As String
    Dim sSQL As String
    
    Set ofwDBHelper = New FWDBHelper
    
    sSource = "BOUser.Login() for " & t_sUserNTLogon

    If LTrim(RTrim(t_sUserNTLogon)) = "" Then        'No user name is passed
        'This can happen when iis settings is set to allow anonymous
        RaiseError sSource, EPROCATINTERNALERROR.peDBInvalidUser, "User name is empty"
    End If
    
    'Construct Sql to search the database
    sSQL = "SELECT " & _
                    gk_sDBCOLUserID & "," & _
                    gk_sDBCOLUserName & "," & _
                    gk_sDBCOLUserInitials & _
                " FROM " & _
                    gk_sDBTABLEUsers & "(nolock)" & _
                " WHERE " & _
                    gk_sDBCOLUserNTlogon & " = " & Quotes(t_sUserNTLogon)
    
    
            
    'Search the database for the name and password
    Set oFWRecordlist = ofwDBHelper.FindBySql(sSQL)

    'See if a record is found matching the condition
    If oFWRecordlist Is Nothing Then
        'Resultset could not be created
        RaiseError sSource, EPROCATINTERNALERROR.peDBInvalidUser, "User query returned empty"
    End If
    If oFWRecordlist.RecordCount = 0 Then
        'No user with that username and password
        RaiseError sSource, EPROCATINTERNALERROR.peDBInvalidUser, "No User with that name"
    End If
    
    'Get the users info from the table
    m_lUserID = oFWRecordlist.GetFieldValue(gk_sDBCOLUserID)
    m_sUserNTLogon = t_sUserNTLogon
    m_sUserName = oFWRecordlist.GetFieldValue(gk_sDBCOLUserName)
    m_sUserInitials = oFWRecordlist.GetFieldValue(gk_sDBCOLUserInitials)
    
         
    
    'Relase the objects
    Set ofwDBHelper = Nothing
    Set oFWRecordlist = Nothing
   
    Exit Sub
    
ErrorHandler:
    ConvertError sSource
End Sub

Public Function AllUserNames() As FWRecordList
' Description:
'
' Returns names of all the users in user table
' This used to display the names of users when a user wants to view
' the recordsets of other users
'
' Amendment history:
'
' d1   NDK        14Jun1999     Created

On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim sSQL As String
    Dim ofwDBHelper As FWDBHelper
    
      
    sSource = "BOUser.GetAllUserNames()"
    
    
    Set ofwDBHelper = New FWDBHelper
    
    'Construct the sql
    sSQL = "SELECT " & _
                   gk_sDBCOLUserID & "," & _
                   gk_sDBCOLUserNTlogon & "," & _
                   gk_sDBCOLUserName & "," & _
                   gk_sDBCOLUserInitials & _
               " FROM " & _
                    gk_sDBTABLEUsers & "(nolock)" & _
               " ORDER BY  " & _
                   gk_sDBCOLUserName
           
       
       
    'Run the sql to get a list of all users
    Set AllUserNames = ofwDBHelper.FindBySql(sSQL)
    
    'Release the objects
    Set ofwDBHelper = Nothing
    Exit Function
ErrorHandler:
    ConvertError sSource
End Function


Private Sub Class_Initialize()
Initialize
End Sub

Private Sub ObjectControl_Activate()
'Create the BOUserRoles object
 Set m_oboUserRoles = New BoUserRoles
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
Terminate
End Sub

Public Sub PROCATTransientInterface_DeleteState(Optional t_sInstanceReference As String)
'   d10      ksk    Changed to relect the change in FWTransientManager
    Dim sInstanceReference As String
    Dim ofwTransientManager As FWTransientManager

    sInstanceReference = CreateDelimitedString(Me.Name, t_sInstanceReference)

    Set ofwTransientManager = New FWTransientManager
    ofwTransientManager.DeleteData GetObjectContext.Item(gk_sSessionObject).SessionID, _
                                    sInstanceReference

    LogMessage "Deleted BOUser state" & sInstanceReference, EProfile.pComplete

End Sub
'
Public Sub PROCATTransientInterface_GetState(Optional t_sInstanceReference As String)
'   d10      ksk    Changed to relect the change in FWTransientManager
'   d15      WDP    16Feb2000      Removed need for user_ticket
    Dim sInstanceReference As String
    Dim sSerialisedString As String
    Dim ofwTransientManager As FWTransientManager

    sInstanceReference = CreateDelimitedString(Me.Name, t_sInstanceReference)

    Set ofwTransientManager = New FWTransientManager
    sSerialisedString = ofwTransientManager.GetData(GetObjectContext.Item(gk_sSessionObject).SessionID, _
                                            sInstanceReference)
    If Trim(sSerialisedString) = "" Then
        LogMessage "BOUser state not found " & sInstanceReference, EProfile.pProfileInformation
    Else
        m_lUserID = CLng(GetDataFromSerialisedString(sSerialisedString))
        m_sUserNTLogon = GetDataFromSerialisedString(sSerialisedString)
        m_sUserName = GetDataFromSerialisedString(sSerialisedString)
        m_sUserInitials = GetDataFromSerialisedString(sSerialisedString)
        
        LogMessage "Read BOUser state" & sInstanceReference, EProfile.pComplete
    End If
End Sub
'
Public Sub PROCATTransientInterface_SetState(Optional t_sInstanceReference As String)
'   d10      ksk    Changed to relect the change in FWTransientManager
'   d15      WDP    16Feb2000      Removed need for user_ticket
    Dim sInstanceReference As String
    Dim sSerialisedString As String
    Dim ofwTransientManager As FWTransientManager

    sInstanceReference = CreateDelimitedString(Me.Name, t_sInstanceReference)

    sSerialisedString = ""
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, CStr(m_lUserID))
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, m_sUserNTLogon)
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, m_sUserName)
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, m_sUserInitials)
    
    Set ofwTransientManager = New FWTransientManager
    ofwTransientManager.AddData GetObjectContext.Item(gk_sSessionObject).SessionID, _
                                sInstanceReference, sSerialisedString

    LogMessage "Stored BOUser state" & sInstanceReference, EProfile.pComplete

End Sub

Public Sub Load(ByVal t_lUserID As Long)
'Description        :   Loads a user details from the database and assigns
'                       the values to the private properties

'Amendment History  :
'
' d7    27Jul1999   CRM     Created
' d13  CRM        10Dec1999     Modified to handle User Ticket
' d14  AGA        11Dec1999     modified to handle nullable user_ticket field
' d15  WDP        16Feb2000      Removed need for user_ticket
On Error GoTo ErrorHandler
    Dim sSource As String
    Dim ofwUserDet As FWRecordList
    Dim vRoleArray() As Variant
    Dim lIndex As Long
    sSource = "BOUser.Load"
       
    m_lUserID = t_lUserID
        
    'Check if its a new user
    If m_lUserID <> 0 Then
        'Fetch the User's details,based on the UserID
        Set ofwUserDet = DetailsOfUser(m_lUserID)
        
        'assign the values to the private variables
        m_sUserName = ofwUserDet.GetFieldValue(gk_sDBCOLUserName)
        m_sUserInitials = ofwUserDet.GetFieldValue(gk_sDBCOLUserInitials)
        m_sUserNTLogon = ofwUserDet.GetFieldValue(gk_sDBCOLUserNTlogon)
        
        'Assign the userID to the BOUser object
        m_oboUserRoles.UserID = m_lUserID
        
        'read user details and populate roleid's in an array
        Do Until ofwUserDet.IsAtEnd
            lIndex = lIndex + 1
            ReDim Preserve vRoleArray(lIndex)
                                                'check if the value of the roleid column is null,
                                                'if so assign 0 there instead of the value or else
                                                'there will be an error thrown up
            vRoleArray(UBound(vRoleArray) - 1) = IIf(IsNull(ofwUserDet.GetFieldValue(gk_sDBCOLRoleID)), 0, _
                                                    ofwUserDet.GetFieldValue(gk_sDBCOLRoleID))
            ofwUserDet.MoveNext
        Loop
        m_vRolesArray = vRoleArray
    End If
    
    Set ofwUserDet = Nothing
    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub

Private Function DetailsOfUser(ByVal t_lUserID As Long) As FWRecordList
'Description    :   Returns  the details for a specific user
'Arguments      :   t_lUserID ------ UserID of the user
' d7    CRMahesh          July 27,1999      Copied from BOUserEditSet class
'd12     aga         29Oct1999   Replacing Hardcoded constants to DBContants
' d13  CRM        10Dec1999     Modified to handle User Ticket
' d14  AGA        11Dec1999     modified to handle nullable user_ticket field
' d15  WDP        16Feb2000      Removed need for user_ticket
On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim sSQL As String
    Dim ofwDBHelper As FWDBHelper
    
    sSource = "BOUser.DetailsOfUser()"
    
    Set ofwDBHelper = New FWDBHelper
    
    'Construct the sql
    'Gets data from Users,User_Type,User_Role tables,outer-join on Users & User_Role
    'so that even users with no roles can be retrieved
        
          sSQL = "SELECT " & gk_sDBTABLEUsers & "." & gk_sDBCOLUserID & ", " & _
               gk_sDBTABLEUsers & "." & gk_sDBCOLUserNTlogon & ", " & _
               gk_sDBTABLEUsers & "." & gk_sDBCOLUserName & ", " & _
               gk_sDBTABLEUsers & "." & gk_sDBCOLUserInitials & ", " & _
               gk_sDBTABLEUserRole & "." & gk_sDBCOLRoleID & _
               " FROM " & gk_sDBTABLEUsers & ", " & gk_sDBTABLEUserRole & _
               " WHERE " & _
               gk_sDBTABLEUsers & "." & gk_sDBCOLUserID & " *= " & gk_sDBTABLEUserRole & "." & gk_sDBCOLUserID & _
               " AND " & gk_sDBTABLEUsers & "." & gk_sDBCOLUserID & " = " & t_lUserID
               
    'Run the sql and get the recordlist of users
    Set DetailsOfUser = ofwDBHelper.FindBySql(sSQL)
        
    'Release the objects
    Set ofwDBHelper = Nothing
    
    Exit Function
    
ErrorHandler:
   
    ConvertError sSource

End Function
Private Function AddUser(ByVal t_sUserNTLogon As String, ByVal t_sUserName As String, ByVal t_sUserInitials As String) As Long
'Description    :   Adds a new user to the database and returns the new UserID
'Arguments      :   t_sUserTypeID - User's Type
'                   t_sUserNTLogon - NT LogOn
'                   t_sUserInitials - User's Initials
'
'Amendment      :
'   d7      crm     27Jul1999       Copied from BOUserEditSet class
'   d12     aga    29Oct1999      User_ID is now IDENTITY column in database: Removed Max(ID) query
'                                          Handling duplication of User name, NT Logon and User initials
'                                          Replacing Hardcoded constants to DBContants
'   d14     AGA    11Dec1999      modified to handle nullable user_ticket field
'   d15      WDP    16Feb2000      Removed need for user_ticket

On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim sSQL As String
    Dim ofwDBHelper As FWDBHelper
    Dim oFWRecordlist As FWRecordList
    Dim lMaxID As Long
    sSource = "BOUser.AddUser()"
    
    Set ofwDBHelper = New FWDBHelper
    Set oFWRecordlist = New FWRecordList
    
    'Check for the duplication of User name, NT logon, Initials here
''''''''''''''''''''''''
    'Check whether the NT Logon Name is unique
    sSQL = "SELECT * FROM " & gk_sDBTABLEUsers _
                & " WHERE " & gk_sDBCOLUserNTlogon & "=" & QuoteValueWithQuotes(t_sUserNTLogon)
    Set oFWRecordlist = ofwDBHelper.FindBySql(sSQL)
    If oFWRecordlist.RecordCount <> 0 Then
        'Raise the error, if NT logon already exists
        RaiseError sSource, peDuplicateUserNTLogon, "Duplicate User NT Logon"
    End If
    
    'Check whether the User Name is unique
    sSQL = "SELECT * FROM " & gk_sDBTABLEUsers _
                & " WHERE " & gk_sDBCOLUserName & "=" & QuoteValueWithQuotes(t_sUserName)
    Set oFWRecordlist = ofwDBHelper.FindBySql(sSQL)
    If oFWRecordlist.RecordCount <> 0 Then
        'Raise the error, if User name already exists
        RaiseError sSource, peDuplicateUserName, "Duplicate User name"
    End If
    
    'Check whether the User initials is unique
    sSQL = "SELECT * FROM " & gk_sDBTABLEUsers _
                & " WHERE " & gk_sDBCOLUserInitials & "=" & QuoteValueWithQuotes(t_sUserInitials)
    Set oFWRecordlist = ofwDBHelper.FindBySql(sSQL)
    If oFWRecordlist.RecordCount <> 0 Then
        'Raise the error, if user initials already exists
        RaiseError sSource, peDuplicateUserInitials, "Duplicate User intials"
    End If
    

'AGA : Taken care of nullable UserTicket
        
    'Insert new User: Construct the sql
        sSQL = "INSERT INTO " & gk_sDBTABLEUsers _
            & " ( " & gk_sDBCOLUserNTlogon & ", " & gk_sDBCOLUserName & ", " & _
            gk_sDBCOLUserInitials & ")" & _
            " VALUES(" & QuoteValueWithQuotes(t_sUserNTLogon) & "," & _
            QuoteValueWithQuotes(t_sUserName) & "," & QuoteValueWithQuotes(t_sUserInitials) & ")"
                  
    'Run the sql and insert the new user details and get the IDENTITY
    lMaxID = ofwDBHelper.InsertAndReturnIdentity(sSQL)
        
    'Release the objects
    Set ofwDBHelper = Nothing
    Set oFWRecordlist = Nothing
    
    AddUser = lMaxID
    Exit Function
    
ErrorHandler:
    ConvertError sSource
End Function


Public Function DeleteUser(ByVal t_lUserID As Long) As Long
'Description    :   Deletes user,if associated role is not part of an edit set,if role is part of
'                    an editset then returns the roleid,Requires UserID as the argument
'Amendment History  :
'd7      CRM        27Jul1999   Copied from BOUserEditSet
'd11     aga        07Oct1999   Added functionality for not deleting user if he is a managing editor
'd12     aga        29Oct1999   Replacing Hardcoded constants to DBContants
'd17     tmj        21Mar2000   Added error checking so covers all tables referenced by users

On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim sSQL As String
    Dim ofwDBHelper As FWDBHelper
    Dim ofwUserRoles As FWRecordList
    Dim bFound As Boolean
    Dim lRoleID As Long
    Dim oboUserRoles As BoUserRoles
    
    sSource = "BOUser.DeleteUser()"
    
    Set ofwDBHelper = New FWDBHelper
    'Check if the user has managing editor role
    Set oboUserRoles = New BoUserRoles
    oboUserRoles.UserID = t_lUserID
    If oboUserRoles.UserHasManagingEditorRole Then
        'Do not delete the user if managing editor
        DeleteUser = 0
    Else
        'Construct the sql
        sSQL = "SELECT " & gk_sDBCOLEditSetID & ", " & gk_sDBCOLRoleID _
                & " FROM " & gk_sDBTABLEUserRoleOnEditset & " WHERE " & _
                gk_sDBCOLUserID & " = " & t_lUserID
           
        'Run the sql and get the recordlist of user roles
        Set ofwUserRoles = ofwDBHelper.FindBySql(sSQL)
            
        'Check if the User has any role in the EditSet
        If ofwUserRoles.RecordCount > 0 Then
            'Roles are present,return the first RoleID
            DeleteUser = ofwUserRoles.GetFieldValue(gk_sDBCOLEditSetID)
        Else
            'Checking to ensure that there are no values that refer to user in audit trail
            sSQL = "SELECT " & gk_sDBCOLEditSetID & " FROM " & _
                   gk_sDBTableAuditTrailUserRoleOnEditSet & " WHERE " & _
                   gk_sDBCOLUserID & " = " & t_lUserID
                   
            'Run the sql and get the recordlist of user roles
            Set ofwUserRoles = ofwDBHelper.FindBySql(sSQL)
                
            'Check if the User has any role in the EditSet
            If ofwUserRoles.RecordCount > 0 Then
                'If any records exist then raise an error
                RaiseError sSource, PEUserInAuditTrail, _
                "Cannot Delete User As they are refered to in the audit trail"
            End If
            
            'Checking to ensure that there are no values that refer to user in
            'audit trail for editset history
            sSQL = "SELECT " & gk_sDBCOLEditSetID & " FROM " & _
                   gk_sDBTABLEAuditTrailEditSetHistory & " WHERE " & _
                   gk_sDBCOLUserID & " = " & t_lUserID
                   
            'Run the sql and get the recordlist of user roles
            Set ofwUserRoles = ofwDBHelper.FindBySql(sSQL)
                
            'Check if the User has any role in the EditSet history
            If ofwUserRoles.RecordCount > 0 Then
                'If any records exist then raise an error
                RaiseError sSource, PEUserInAuditTrail, _
                "Cannot Delete User As they are refered to in the audit trail"
            End If
                        
            'Checking to ensure that there are no values that refer to user in place tables
            sSQL = "SELECT " & gk_sDBCOLPlaceReferenceID & " FROM " & _
                   gk_sDBTABLEPlace & " WHERE " & _
                   gk_sDBCOLAFCreated_By & " = " & t_lUserID
                   
            'Run the sql and get the recordlist of user roles
            Set ofwUserRoles = ofwDBHelper.FindBySql(sSQL)
                
            'Check if the User has any entries in the place table
            If ofwUserRoles.RecordCount > 0 Then
                'If any records exist then raise an error
                RaiseError sSource, PEUserCreatedIndexTerm, _
                "Cannot Delete User As they are created a place index"
            End If
                
            'Checking to ensure that there are no values that refer to user in
            'corporate body tables
            sSQL = "SELECT " & gk_sDBCOLCorporateBodyReferenceID & " FROM " & _
                   gk_sDBTABLECorporateBody & " WHERE " & _
                   gk_sDBCOLAFCreated_By & " = " & t_lUserID
                   
            'Run the sql and get the recordlist of user roles
            Set ofwUserRoles = ofwDBHelper.FindBySql(sSQL)
                
            'Check if the User has any entries in the corporate body table
            If ofwUserRoles.RecordCount > 0 Then
                'If any records exist then raise an error
                RaiseError sSource, PEUserCreatedIndexTerm, _
                "Cannot Delete User As they created a corporate body index"
            End If
                                
            'Checking to ensure that there are no values that refer to user in person tables
            sSQL = "SELECT " & gk_sDBCOLPersonReferenceID & " FROM " & _
                   gk_sDBTABLEPerson & " WHERE " & _
                   gk_sDBCOLAFCreated_By & " = " & t_lUserID
                   
            'Run the sql and get the recordlist of user roles
            Set ofwUserRoles = ofwDBHelper.FindBySql(sSQL)
                
            'Check if the User has any entries in the person table
            If ofwUserRoles.RecordCount > 0 Then
                'If any records exist then raise an error
                RaiseError sSource, PEUserCreatedIndexTerm, _
                "Cannot Delete User As they created a person index"
            End If
                
            'Checking to ensure that there are no values that refer to user in
            'subject tables
            sSQL = "SELECT " & gk_sDBCOLSubjectReferenceID & " FROM " & _
                   gk_sDBTABLESubject & " WHERE " & _
                   gk_sDBCOLAFCreated_By & " = " & t_lUserID
                   
            'Run the sql and get the recordlist of user roles
            Set ofwUserRoles = ofwDBHelper.FindBySql(sSQL)
                
            'Check if the User has any entries in the corporate body table
            If ofwUserRoles.RecordCount > 0 Then
                'If any records exist then raise an error
                RaiseError sSource, PEUserCreatedIndexTerm, _
                "Cannot Delete User As they created a subject index"
            End If
                                
            'No roles in EditSet
            'Delete the details in User_Role table first
            sSQL = "DELETE " & gk_sDBTABLEUserRole & " WHERE " & gk_sDBCOLUserID & _
                   " = " & t_lUserID
            ofwDBHelper.Execute sSQL
            'Delete the entry in the Users table
            sSQL = "DELETE " & gk_sDBTABLEUsers & " WHERE " & gk_sDBCOLUserID & _
                   " = " & t_lUserID
            ofwDBHelper.Execute sSQL
            DeleteUser = 0
        End If
    End If
    'Close objects
    Set ofwUserRoles = Nothing
    Set ofwDBHelper = Nothing

    
    Exit Function
    
ErrorHandler:
    ConvertError sSource
End Function

Private Sub UpdateUser(ByVal t_lUserID As Long, ByVal t_sUserNTLogon As String, ByVal t_sUserName As String, ByVal t_sUserInitials As String)
'Description    :   Updates a User's detail in the database
'Arguments      :   t_sUserTypeID - User's Type
'                   t_sUserNTLogon - NT LogOn
'                   t_sUserInitials - User's Initials
'Amendment History  :
'd7      CRM    27Jul1999       Copied from BOUserEditSet class
'                               Removed UserTypeID as an argument,update
'                               statement does not modify UserTypeID anymore
'd12     aga    29Oct1999       Handling duplication of User name, NT Logon and User initials
'                                        Replacing Hardcoded constants to DBContants
'd13     crm    10Dec1999       Modified to update UserTicket field
'd15      WDP    16Feb2000      Removed need for user_ticket
On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim sSQL As String
    Dim ofwDBHelper As FWDBHelper
    Dim oFWRecordlist As FWRecordList
    Dim lMaxID As Long
    sSource = "BOUser.UpdateUser()"
    
    Set ofwDBHelper = New FWDBHelper
    
    Set oFWRecordlist = New FWRecordList
    
    'Check for the duplication of User name, NT logon, Initials here

    'Check whether the NT Logon Name is unique
    sSQL = "SELECT * FROM " & gk_sDBTABLEUsers _
                & " WHERE " & gk_sDBCOLUserNTlogon & "=" & QuoteValueWithQuotes(t_sUserNTLogon) _
                & " AND " & gk_sDBCOLUserID & "<>" & t_lUserID
    Set oFWRecordlist = ofwDBHelper.FindBySql(sSQL)
    If oFWRecordlist.RecordCount <> 0 Then
        'Raise the error, if NT logon already exists
        RaiseError sSource, peDuplicateUserNTLogon, "Duplicate User NT Logon"
    End If
    
    'Check whether the User Name is unique
    sSQL = "SELECT * FROM " & gk_sDBTABLEUsers _
                & " WHERE " & gk_sDBCOLUserName & "=" & QuoteValueWithQuotes(t_sUserName) _
                & " AND " & gk_sDBCOLUserID & "<>" & t_lUserID
    Set oFWRecordlist = ofwDBHelper.FindBySql(sSQL)
    If oFWRecordlist.RecordCount <> 0 Then
        'Raise the error, if User name already exists
        RaiseError sSource, peDuplicateUserName, "Duplicate User name"
    End If
    
    'Check whether the User initials is unique
    sSQL = "SELECT * FROM " & gk_sDBTABLEUsers _
                & " WHERE " & gk_sDBCOLUserInitials & "=" & QuoteValueWithQuotes(t_sUserInitials) _
                & " AND " & gk_sDBCOLUserID & "<>" & t_lUserID
    Set oFWRecordlist = ofwDBHelper.FindBySql(sSQL)
    If oFWRecordlist.RecordCount <> 0 Then
        'Raise the error, if user initials already exists
        RaiseError sSource, peDuplicateUserInitials, "Duplicate User intials"
    End If
    

    'Construct the sql to update the user
        sSQL = "UPDATE " & gk_sDBTABLEUsers & " SET " _
                & gk_sDBCOLUserNTlogon & "=" & QuoteValueWithQuotes(t_sUserNTLogon) & _
               ", " & gk_sDBCOLUserName & "=" & QuoteValueWithQuotes(t_sUserName) & _
               ", " & gk_sDBCOLUserInitials & "=" & QuoteValueWithQuotes(t_sUserInitials)
    
    sSQL = sSQL & " WHERE " & gk_sDBCOLUserID & "=" & t_lUserID
               
    'Run the sql update the user details
    ofwDBHelper.Execute sSQL
    
    'Release the objects
    
    Set ofwDBHelper = Nothing
    Set oFWRecordlist = Nothing
    
    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub

Public Function Save() As Boolean
'Description    :   Method which either Saves the details of an existing user or
'                   adds a new user.
'Amendment History :
'd7     CRM    27Jul1999    Copied from BOUserEditSet class and renamed to Save
'                           Call to UpdateUser function does not have UserTypeID as
'                           argument anymore
'd13      crm    10Dec1999  Modified to send User Ticket while updating
'd15      WDP    16Feb2000      Removed need for user_ticket
On Error GoTo ErrorHandler

    Dim sSource As String
    Dim lNewUserID As Long
    Dim vItem
    Dim ofwUserRoles As FWRecordList
    Dim bRoleIsThere As Boolean
    Dim lTempRoleID As Long
    Dim sRoleNotRemoved As String
    Dim bIncomplete As Boolean
    Dim lRoleID As Long
    Dim lIndex As Long
    
    sSource = "BOUser.Save"
    'Check if a UserID was received,if it is then it is greater than 0 or its a new user
    If m_lUserID > 0 Then
            UpdateUser m_lUserID, m_sUserNTLogon, m_sUserName, m_sUserInitials
        
        'Get a list of all the roles
        Set ofwUserRoles = m_oboUserRoles.AllRoles
         
        m_oboUserRoles.UserID = m_lUserID
        
        'Loop thru all roles and check against the UserHasRoles collection
        'if the RoleID is present
        Do Until ofwUserRoles.IsAtEnd
            bRoleIsThere = False
            lRoleID = ofwUserRoles.GetFieldValue(gk_sDBCOLRoleID)
            
            '   d12      crm    12Aug1999      Modified enumerating the role array
            
            'For each user role compare to see if role checked is received
            'there is the possibility of no roles choosen and the array variable
            'being empty.
            If Not IsNull(m_vRolesArray) Then
            For lIndex = LBound(m_vRolesArray) To UBound(m_vRolesArray) - 1
                If lRoleID = m_vRolesArray(lIndex) Then
                         ' If role is there then set flag
                         bRoleIsThere = True
                         lTempRoleID = lRoleID
                End If
            Next
            End If

            If bRoleIsThere Then
                 'If role is there and if its a Managing Editor,perform a check to see
                 'if there are other Managing Editor(s)
                 If lRoleID = gk_lManagingEditorRoleID Then
                    m_sStatus = GetManagingEditor
                 End If
                 'Add role for the user
                 m_oboUserRoles.AddRoleToUser lTempRoleID
                
            Else
                'Remove role,function returns EditSetID if the role is part of edit set
                If Not (m_oboUserRoles.RemoveRoleFromUser(lRoleID)) Then
                        'Set flag if role could not be removed
                        bIncomplete = True
                        'Get the description of the Role from the RoleID for status
                        sRoleNotRemoved = sRoleNotRemoved & _
                                            ofwUserRoles.GetFieldValue(gk_sDBCOLRoleDescription) & ", "
                End If
            End If
            ofwUserRoles.MoveNext
        Loop
         'If sRoleNotRemoved is not empty then remove the extra ,
            If Trim(sRoleNotRemoved) <> "" Then
                sRoleNotRemoved = Left(sRoleNotRemoved, Len(sRoleNotRemoved) - 2)
            End If
            'Check if all changes were  successful,inform if incomplete or goto main page
            If bIncomplete Then
               m_sStatus = m_sStatus & sRoleNotRemoved
                Save = False
            Else
                Save = True
            End If
    Else
        'all values provided thru UI except UserTypeID which is a constant for this module
        m_lUserID = AddUser(m_sUserNTLogon, m_sUserName, m_sUserInitials)
        m_oboUserRoles.UserID = m_lUserID
        
        '   d12      crm    12Aug1999      Modified enumerating the role array
            'loop thru array of roleids and save them
        If Not IsNull(m_vRolesArray) Then
            For lIndex = LBound(m_vRolesArray) To UBound(m_vRolesArray) - 1
            m_oboUserRoles.AddRoleToUser m_vRolesArray(lIndex)
            Next
        End If
        Save = True
    End If
    
        Set ofwUserRoles = Nothing
    
Exit Function
ErrorHandler:

    ConvertError sSource
    End Function



Public Function GetManagingEditor() As String
'Description    :   Checks for another Managing Editor in the system
'                   if there is one then returns the name or an empty
'                   string
'Amendment History  :
'
'   d8      crm     29Jul1999   created
'                               Checks for any previous Managaing Editor role
'                               assigned
'   d12    aga     29Oct1999    Replacing Hardcoded constants to DBContants
'   d16    tmj     17Mar2000    Removed gk_sMessageDelimiter from end of managing editor value

On Error GoTo ErrorHandler

    Dim sSource As String
    Dim sSQL As String
    Dim ofwDBHelper As FWDBHelper
    Dim ofwMngEditor As FWRecordList
    Dim sMngEditorName As String
    
    
    sSource = "BOUser.GetManagingEditor()"
    Set ofwDBHelper = New FWDBHelper
    
    'Get the names of users with a Managing Editor RoleID
    sSQL = "SELECT " & gk_sDBCOLUserName & " FROM " & gk_sDBTABLEUsers & ", " & gk_sDBTABLEUserRole & _
           " WHERE " & gk_sDBTABLEUsers & "." & gk_sDBCOLUserID & " = " & gk_sDBTABLEUserRole & "." & gk_sDBCOLUserID & _
           " AND " & gk_sDBTABLEUserRole & "." & gk_sDBCOLRoleID & " = " & gk_lManagingEditorRoleID
    
    Set ofwMngEditor = ofwDBHelper.FindBySql(sSQL)
        
    'if there is a Managing Editor role then
    'find their names(this is a possibility only in the beginning before this
    'module is setup..then we ensure that only one Managing Editor role
    'is present at any given time
    If ofwMngEditor.RecordCount > 0 Then
        Do Until ofwMngEditor.IsAtEnd
            sMngEditorName = sMngEditorName & _
                             ofwMngEditor.GetFieldValue(gk_sDBCOLUserName) & ", "
            ofwMngEditor.MoveNext
        Loop
        
        'Trim the last comma
        sMngEditorName = Left(sMngEditorName, Len(sMngEditorName) - 2)
    End If
    
    
    GetManagingEditor = sMngEditorName '& gk_sMessageDelimiter
    Set ofwDBHelper = Nothing
    Set ofwMngEditor = Nothing
    
Exit Function

ErrorHandler:
ConvertError sSource
End Function

