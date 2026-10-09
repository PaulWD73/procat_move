VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BoUserRoles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    BOUserRoles.cls
' System:       PROCAT, PRO
' Copyright:    (C) Quidnunc Limited
'
' Description:  User and Roles Class
'               Contains methods to return list of role, valid roles for a user, valid user for roles etc.
'
' Amendment history:
'   d1   ndk      26Jul1999      Created
'                                AllRoles,RolesAssignableToEditSet,ValidUsersForTheRole were move from BoEditSet
'                                and slight modifications were made to contents and function names
'                                Added UserRoles property
'   d2   ndk      28Jul1999      Added UserHasManagingEditorRole
'                                Moved UserHasAdministratorRole from BoEditSet
'   d3   crm      02Aug1999      Added AddRoleToUser(),RemoveRoleFromUser(),IsUserRoleInEditSet()
'                                copied from BOUserEditSet.Removed UserRoles property
'   d4   ndk      01Sep1999      All references to column default_user_for_role removed
'   d5   aga      06Oct1999      Modified AddRoleToUser to check for Managing Editor role can be assigned to one user
'                                        modified  RemoveRoleFromUser - Managing Editor role cannont be removed
'                                         Added UserHasAFEditorRole method
'   d6   aga      28Oct1999     modified  RemoveRoleFromUser - Last Procat Editor role cannont be removed
'                                       Replaced hardcoded constants with DBConstants
'   d7   ndk      11Nov1999     Added method UserHasRole
'   d8   ndk      12Nov1999     Updated method RolesAssignableToEditSet to filter out few more roles
'   d9   ndk      23Dec1999     Modified ValidUsersForTheRole,RolesAssignableToEditSet to return the list in alphabetical order
'   d10  dtm      01feb2000     Created UserHasOneOfMultipleRoles
'   d41  dtm      10feb2000     Modified RolesAssignableToEditSet so that Accessions role is not returned
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit


'Private property variables

Private m_lRoleID As Long
Private m_lUserID As Long

'Public property methods

Public Property Let UserID(ByVal t_lUserID As Long)
' Description:
'
' to assign the userid whose details are to be retreived
'
' Amendment history:
'
' d1   ndk      26Jul1999      Created


    m_lUserID = t_lUserID


End Property

Public Property Let RoleID(ByVal t_lRoleID As Long)
' Description:
'
' to assign the roleid whose details are to be retreived
'
' Amendment history:
'
' d1   ndk      26Jul1999      Created


    m_lRoleID = t_lRoleID

End Property

Public Function AllRoles() As FWRecordList
' Description:
'
' Returns the list of roles
' Data is returned as a FWRecordlist
'
' Amendment history:
'
' d1   ndk      26Jul1999      Copied from BoEditSet
'
On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim sSql As String
    Dim oFWDBHelper As FWDBHelper
    
    sSource = "BoUserRoles.Roles()"
    
    Set oFWDBHelper = New FWDBHelper
    
    'Construct the sql
    sSql = "SELECT " & _
            gk_sDBCOLRoleID & " , " & _
            gk_sDBCOLRoleDescription & " , " & _
            gk_sDBCOLRoleLevel & _
            " FROM " & _
            gk_sDBTABLERole
            
       
    'Run the sql and get the recordlist of roles
    Set AllRoles = oFWDBHelper.FindBySql(sSql)
        
    'Release the objects
    Set oFWDBHelper = Nothing

    
    Exit Function
    
ErrorHandler:
    ConvertError sSource
End Function

Public Function RolesAssignableToEditSet() As FWRecordList
 
' Description:
'
' Returns the list of roles
' Does not return the role of administrator
' Data is returned as a FWRecordlist
'
' Amendment history:
'
' d1   ndk      26Jul1999      Copied from BoEditSet
' d8   ndk      12Nov1999      Filtering off few more roles
' d9   ndk      23Dec1999      Modified to return the list in alphabetical order
' d11  dtm      10feb2000      Modified so that Accessions role is not returned
'
On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim sSql As String
    Dim oFWDBHelper As FWDBHelper
    
    sSource = "BoUserRoles.RolesAssignableToEditSet()"
    
    Set oFWDBHelper = New FWDBHelper
    
    ' d11  dtm      10feb2000      Modified so that Accessions role is not returned
    'Construct the sql
    sSql = "SELECT " & _
                gk_sDBCOLRoleID & " , " & _
                gk_sDBCOLRoleDescription & " , " & _
                gk_sDBCOLRoleLevel & _
            " FROM " & _
                gk_sDBTABLERole & _
            " WHERE " & _
                "NOT " & gk_sDBCOLRoleID & " = " & gk_lAdministratorRoleID & " AND " & _
                "NOT " & gk_sDBCOLRoleID & " = " & gk_lManagingEditorRoleID & " AND " & _
                "NOT " & gk_sDBCOLRoleID & " = " & gk_lESFAEditorRoleID & " AND " & _
                "NOT " & gk_sDBCOLRoleID & " = " & gk_lPSFAEditorRoleID & " AND " & _
                "NOT " & gk_sDBCOLRoleID & " = " & gk_lLeafletEditorRoleID & " AND " & _
                "NOT " & gk_sDBCOLRoleID & " = " & gk_lPopularSearchEditorRoleID & " AND " & _
                "NOT " & gk_sDBCOLRoleID & " = " & gk_lAccessionRoleID & _
            " ORDER BY " & gk_sDBCOLRoleDescription
    
    'Run the sql and get the recordlist of roles
    Set RolesAssignableToEditSet = oFWDBHelper.FindBySql(sSql)
        
    'Release the objects
    Set oFWDBHelper = Nothing

    Exit Function
    
ErrorHandler:
    ConvertError sSource
End Function

Public Function ValidUsersForTheRole() As FWRecordList
' Description:
'
' Returns the list valid users for the role stored in property RoleID
' Data is returned as a FWRecordlist
'
' Amendment history:
'
' d1   ndk      26Jul1999      Copied from BoEditSet
' d4   ndk      01Sep1999      All references to column default_user_for_role removed
' d9   ndk      23Dec1999      Modified to return the list in alphabetical order
    
    Dim sSource As String
    Dim sSql As String
    Dim oFWDBHelper As FWDBHelper
    
    sSource = "BoUserRoles.ValidUsersForTheRole()"
    
    Set oFWDBHelper = New FWDBHelper
    
    'Construct the sql
    sSql = "SELECT " & _
                " U." & gk_sDBCOLUserID & " , " & _
                " U." & gk_sDBCOLUserName & _
            " FROM " & _
                gk_sDBTABLEUsers & " U," & _
                gk_sDBTABLEUserRole & " UR " & _
            " WHERE " & _
                " U." & gk_sDBCOLUserID & " = " & "UR." & gk_sDBCOLUserID & " AND " & _
                " UR." & gk_sDBCOLRoleID & " = " & m_lRoleID & _
            " ORDER BY " & " U." & gk_sDBCOLUserName
            
                
    'Run the sql and get the recordlist of user names
    Set ValidUsersForTheRole = oFWDBHelper.FindBySql(sSql)
        
    'Release the objects
    Set oFWDBHelper = Nothing

    Exit Function
    
ErrorHandler:
    ConvertError sSource
End Function

Public Function ValidRolesForTheUser() As FWRecordList
' Description:
'
' Returns the list valid roles for the user stored in property UserID
' Data is returned as a FWRecordlist
'
' Amendment history:
'
' d1   ndk      26Jul1999      Created
    
    Dim sSource As String
    Dim sSql As String
    Dim oFWDBHelper As FWDBHelper
    
    sSource = "BoUserRoles.ValidRolesForTheUser()"
    
    Set oFWDBHelper = New FWDBHelper
    
    'Construct the sql
    sSql = "SELECT " & _
            " R." & gk_sDBCOLRoleID & " , " & _
            " R." & gk_sDBCOLRoleDescription & " , " & _
            " R." & gk_sDBCOLRoleLevel & _
            " FROM " & _
                gk_sDBTABLERole & " R " & ", " & gk_sDBTABLEUserRole & " UR" & _
            " WHERE " & _
                " R." & gk_sDBCOLRoleID & " = " & "UR." & gk_sDBCOLRoleID & " AND " & _
                " UR." & gk_sDBCOLUserID & " = " & m_lUserID
                
    'Run the sql and get the recordlist role details
    Set ValidRolesForTheUser = oFWDBHelper.FindBySql(sSql)
        
    'Release the objects
    Set oFWDBHelper = Nothing

   
    Exit Function
    
    
ErrorHandler:
    ConvertError sSource
End Function
'
Public Sub AddRoleToUser(ByVal t_lRoleID As Long)
'Description    :   Adds a role to an user,updates the User_Role table
'Arguments      :
'                   t_lRoleID - RoleID
'
'Amendment History  :
'd7      crm                Copied from BOUserEditSet class
'd5      aga     06Oct1999      Modified completely to update managing editor role ( to be assigned to one person only)
'd6      aga     28Oct1999      Replaced hardcoded constants with DBConstants

    On Error GoTo ErrorHandler
        
    Dim sSource As String
    Dim sSql As String
    Dim oFWDBHelper As FWDBHelper
    Dim lRecCount As Long
    Dim ofwUserRole As FWRecordList
    Dim ofwMngEdit As FWRecordList
    Dim bUpdate As Boolean
               
    sSource = "BOUser.AddRoleToUser()"
    
    Set oFWDBHelper = New FWDBHelper
    
    If t_lRoleID = gk_lManagingEditorRoleID Then
        'Remove the editor role from all users
        sSql = " DELETE " _
            & gk_sDBTABLEUserRole _
            & " WHERE " & _
                   gk_sDBCOLRoleID & " = " & gk_lManagingEditorRoleID
        oFWDBHelper.Execute sSql
        
        'Add the Editor role to this user
        sSql = "INSERT INTO " & gk_sDBTABLEUserRole & " (" & gk_sDBCOLUserID & "," & gk_sDBCOLRoleID & ") VALUES ( " & _
                   m_lUserID & "," & gk_lManagingEditorRoleID & ")"
        oFWDBHelper.Execute sSql
    Else
        'Construct SQL to find if User and Role exist
        sSql = "SELECT * FROM " & gk_sDBTABLEUserRole _
            & "  WHERE " & gk_sDBCOLUserID & " = " & m_lUserID & " and " & gk_sDBCOLRoleID & "=" & t_lRoleID
        Set ofwUserRole = oFWDBHelper.FindBySql(sSql)
        'See if the User-Role is present
        If ofwUserRole.RecordCount = 0 Then
            'Construct the sql
            sSql = "INSERT INTO " & gk_sDBTABLEUserRole & " (" & gk_sDBCOLUserID & "," & gk_sDBCOLRoleID & ") VALUES ( " & _
                   m_lUserID & "," & t_lRoleID & ")"
                
            oFWDBHelper.Execute sSql
        End If
    End If
    
    
    'Release the objects
    Set ofwUserRole = Nothing
    Set oFWDBHelper = Nothing
    Set ofwMngEdit = Nothing
    
    Exit Sub
        
ErrorHandler:
        ConvertError sSource
End Sub
Public Function RemoveRoleFromUser(ByVal t_lRoleID As Long) As Boolean
'Description    :   Removes the specified role from the user,if role is not part of existing edit set
'                   If the user
'Amendment History  :
'd7      CRM    27Jul1999    Copied from BOUserEditSet class
'd5      AGA    06Oct1999   Role cannot be deleted if the roles is managing editor
'd6      AGA    28Oct1999   Last Procat Editor role cannont be removed
'                                    Replaced hardcoded constants with DBConstants

On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim sSql As String
    Dim oFWDBHelper As New FWDBHelper
    Dim oFWRecordlist As New FWRecordList
    Dim lRoleID As Long
    
    sSource = "BOUser.RemoveRoleFromUser()"
    
    'd5  If RoleId is of ManagingEditor do not remove it
    If t_lRoleID <> gk_lManagingEditorRoleID Then
        'Check if the User is part of an Edit Set,if a part then IsRoleInEditSet returns RoleID
        lRoleID = IsUserRoleInEditSet(m_lUserID, t_lRoleID)
        If lRoleID > 0 Then
            RemoveRoleFromUser = False
        Else
            'd6
            If t_lRoleID = gk_lAdministratorRoleID Then
                'If the role is of Procat Manager, check if the user is last procat manager or not
                sSql = "SELECT * FROM " & gk_sDBTABLEUserRole & " WHERE " & gk_sDBCOLRoleID & "=" & t_lRoleID
                Set oFWRecordlist = oFWDBHelper.FindBySql(sSql)
                If oFWRecordlist.RecordCount > 1 Then
                    'More than 1 Procat Manager should exist to remove this role from a user
                    sSql = "DELETE " & gk_sDBTABLEUserRole & " WHERE " & gk_sDBCOLUserID & "= " & m_lUserID & " and " & _
                       gk_sDBCOLRoleID & " = " & t_lRoleID
                    oFWDBHelper.Execute sSql
                    RemoveRoleFromUser = True
                Else
                    'Raise error, and handle this in view class
                    RaiseError sSource, peCanNotDeleteLastProcatManager, "Cannot delete the only PROCAT manager"
                End If
            Else
                sSql = "DELETE " & gk_sDBTABLEUserRole & " WHERE " & gk_sDBCOLUserID & " = " & m_lUserID & " and " & _
                       gk_sDBCOLRoleID & " = " & t_lRoleID
                Set oFWDBHelper = New FWDBHelper
                oFWDBHelper.Execute sSql
                RemoveRoleFromUser = True
            End If
            
        End If
    Else
        'Don't remove managing editor, but return true
        RemoveRoleFromUser = True
    End If
    Set oFWDBHelper = Nothing
    Set oFWRecordlist = Nothing
    Exit Function
    
ErrorHandler:
    
    ConvertError sSource
End Function

Private Function IsUserRoleInEditSet(ByVal t_lUserID As Long, ByVal t_lRoleID As Long) As Long
'Description    :   Checks if user's role is part of an edit set,if it is then returns theRoleID
'                   Takes the UserID and the RoleID as arguments and returns an EditSetID if
'                   the user in that particular role is part of an active EditSet
'Amendment      :
'   d7      CRM       27Jul1999       Copied from BoUserEditSet class
'   d6      aga        28Oct1999     Replaced hardcoded constants with DBConstants
On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim sSql As String
    Dim oFWDBHelper As FWDBHelper
    Dim ofwUserRoles As FWRecordList
    Dim bFound As Boolean
    Dim lRoleID As Long
    
    sSource = "BOUser.IsUserRoleInEditSet()"
    
    Set oFWDBHelper = New FWDBHelper
    
    'Construct the sql
    sSql = "SELECT " & gk_sDBCOLEditSetID & " FROM " & gk_sDBTABLEUserRoleOnEditset & " WHERE " & _
           gk_sDBCOLUserID & "  = " & t_lUserID & " and " & gk_sDBCOLRoleID & " = " & t_lRoleID
       
    'Run the sql and get the recordlist of users
    Set ofwUserRoles = oFWDBHelper.FindBySql(sSql)
     
    If ofwUserRoles.RecordCount > 0 Then
        IsUserRoleInEditSet = ofwUserRoles.GetFieldValue(gk_sDBCOLEditSetID)
    Else
        IsUserRoleInEditSet = 0
    End If
    
    'Release the objects
    Set ofwUserRoles = Nothing
    Set oFWDBHelper = Nothing
    
    Exit Function
    
ErrorHandler:
    ConvertError sSource
End Function
Public Function UserHasManagingEditorRole() As Boolean
' Description:
'
' Checks in the user_role table whether the userid set as property UserID
' has the role of managing editor
'
' Amendment history:
'
'   d2   ndk      28Jul1999      Created
'
On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim sSql As String
    Dim oFWDBHelper As FWDBHelper
    Dim oFWRecordlist As FWRecordList
    
    sSource = "BOUserRoles.UserHasManagingEditorRole()"
    
    Set oFWDBHelper = New FWDBHelper
    
    'Construct the sql
    sSql = "SELECT " & _
                   " * " & _
               " FROM " & _
                    gk_sDBTABLEUserRole & "(nolock)" & _
               " WHERE " & _
                   gk_sDBCOLUserID & " = " & m_lUserID & " AND " & _
                   gk_sDBCOLRoleID & " = " & gk_lManagingEditorRoleID
       
    'Run the sql
    Set oFWRecordlist = oFWDBHelper.FindBySql(sSql)
    
    If oFWRecordlist.RecordCount > 0 Then
        'The user has an administrator role
        UserHasManagingEditorRole = True
    Else
        ' No he does not have an administrator role
        UserHasManagingEditorRole = False
    End If
    
    'Release the objects
    Set oFWRecordlist = Nothing
    Set oFWDBHelper = Nothing
    
    Exit Function
ErrorHandler:
    ConvertError sSource
End Function

Public Function UserHasAdministratorsRole() As Boolean
' Description:
'
' Checks in the user_role table whether the userid set as property UserID
' has the role of Administrator
'
' Amendment history:
'
'   d2      ndk     28Jun199       moved from BoUser
'
On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim sSql As String
    Dim oFWDBHelper As FWDBHelper
    Dim oFWRecordlist As FWRecordList
    
    sSource = "BOUser.UserHasAdministratorsRole()"
    
    Set oFWDBHelper = New FWDBHelper
    
    'Construct the sql
    sSql = "SELECT " & _
                   " * " & _
               " FROM " & _
                    gk_sDBTABLEUserRole & "(nolock)" & _
               " WHERE " & _
                   gk_sDBCOLUserID & " = " & m_lUserID & " AND " & _
                   gk_sDBCOLRoleID & " = " & gk_lAdministratorRoleID
       
    'Run the sql
    Set oFWRecordlist = oFWDBHelper.FindBySql(sSql)
    
    If oFWRecordlist.RecordCount > 0 Then
        'The user has an administrator role
        UserHasAdministratorsRole = True
    Else
        ' No he does not have an administrator role
        UserHasAdministratorsRole = False
    End If
    
    'Release the objects
    
    Set oFWRecordlist = Nothing
    Set oFWDBHelper = Nothing
    
    Exit Function

ErrorHandler:
    ConvertError sSource
End Function

Public Function UserHasAFEditorRole() As Boolean
' Description:
'
' Checks in the user_role table whether the userid set as property UserID
' has the role of AF  editor
'
' Amendment history:
'
'   d5   aga      11Oct1999      Created
'
On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim sSql As String
    Dim oFWDBHelper As FWDBHelper
    Dim oFWRecordlist As FWRecordList
    
    sSource = "BOUserRoles.UserHasAFEditorRole()"
    
    Set oFWDBHelper = New FWDBHelper
    
    'Construct the sql
    sSql = "SELECT " & _
                   " * " & _
               " FROM " & _
                    gk_sDBTABLEUserRole & "(nolock)" & _
               " WHERE " & _
                   gk_sDBCOLUserID & " = " & m_lUserID & " AND " & _
                   gk_sDBCOLRoleID & " = " & gk_lAFEditorRoleID
       
    'Run the sql
    Set oFWRecordlist = oFWDBHelper.FindBySql(sSql)
    
    If oFWRecordlist.RecordCount > 0 Then
        'The user has an administrator role
        UserHasAFEditorRole = True
    Else
        ' No he does not have an administrator role
        UserHasAFEditorRole = False
    End If
    
    'Release the objects
    Set oFWRecordlist = Nothing
    Set oFWDBHelper = Nothing
    
    Exit Function
ErrorHandler:
    ConvertError sSource
End Function




Public Function UserHasRole(ByVal t_lUserID As Long, ByVal t_lRoleID As Long) As Boolean
' Description:
'
' Checks in the user_role table whether the user t_lUserID has the
' role t_lRoleID
'
' Amendment history:
'
'   d7   ndk      11Nov1999     Created
'
On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim sSql As String
    Dim oFWDBHelper As FWDBHelper
    Dim oFWRecordlist As FWRecordList
    
    sSource = "BOUserRoles.UserHasRole()"
    
    Set oFWDBHelper = New FWDBHelper
    
    'Construct the sql
    sSql = "SELECT " & _
                   " * " & _
               " FROM " & _
                    gk_sDBTABLEUserRole & "(nolock)" & _
               " WHERE " & _
                   gk_sDBCOLUserID & " = " & t_lUserID & " AND " & _
                   gk_sDBCOLRoleID & " = " & t_lRoleID
       
    'Run the sql
    Set oFWRecordlist = oFWDBHelper.FindBySql(sSql)
    
    If oFWRecordlist.RecordCount > 0 Then
        'The user has the role
        UserHasRole = True
    Else
        ' No he does not that role
        UserHasRole = False
    End If
    
    'Release the objects
    Set oFWRecordlist = Nothing
    Set oFWDBHelper = Nothing
    
    Exit Function
ErrorHandler:
    ConvertError sSource
End Function



Public Function UserHasOneOfMultipleRoles(ByVal t_lUserID As Long, ByVal t_sRoleIDs As String) As Boolean
' Description:
'
' Checks in the user_role table whether the user t_lUserID has the
' any of the roles in the comma-separated string t_sRoleIDs
'
' Amendment history:
'
'   d10  David Manning  01-feb-2000  Created UserHasOneOfMultipleRoles
'
On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim sSql As String
    Dim oFWDBHelper As FWDBHelper
    Dim oFWRecordlist As FWRecordList
    
    sSource = "BOUserRoles.UserHasOneOfMultipleRoles()"
    
    Set oFWDBHelper = New FWDBHelper
    
    'Construct the sql
    sSql = "SELECT " & _
                   " * " & _
               " FROM " & _
                    gk_sDBTABLEUserRole & "(nolock)" & _
               " WHERE " & _
                   gk_sDBCOLUserID & " = " & t_lUserID & " AND " & _
                   gk_sDBCOLRoleID & " IN ( " & t_sRoleIDs & ")"
       
    'Run the sql
    Set oFWRecordlist = oFWDBHelper.FindBySql(sSql)
    
    If oFWRecordlist.RecordCount > 0 Then
        'The user does have one of the roles
        UserHasOneOfMultipleRoles = True
    Else
        ' No they do not have any of those roles
        UserHasOneOfMultipleRoles = False
    End If
    
    'Release the objects
    Set oFWRecordlist = Nothing
    Set oFWDBHelper = Nothing
    
    Exit Function
ErrorHandler:
    ConvertError sSource
End Function

