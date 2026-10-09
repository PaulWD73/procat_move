VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BOEditSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    BOEditSet.cls
' System:       PROCAT, PRO
' Copyright:    (C) Quidnunc Limited
'
' Description:  Editset business objects
'
' Amendment history:
'   d1      NDK     11May1999       Copied from functional Prototype
'   d2      NDK     11May1999       Added function GetEditSetListsForAllUsers
'   d3      NDK     16Jun1999       Added UpdateEditSetStatus
'   d4      NDK     22Jun1999       Added GetEditSetStatus,RoleCanChangeStatus
'   d5      NDK     23Jun1999       Added GetUsersAndRolesForEditSet,GetRoles,RemoveRoleFromEditSet
'                                   GetValidUsersForTheRole,AddUpdateRoleOnEditSet,UserExistsOnEditSet,RoleExistsOnEditSet
'   d6      NDK     28Jun1999       Reorganised UpdateEditSetStatus, added error handling
'   d7      ndk     30Jun1999       Changed GetValidUsersForTheRole , now returning coloumn default_user_for_role also
'                                   Added GetContributor1ForNewEditSet, GetManagingEditorForNewEditSet
'                                   Added CreateEmptyEditSet,GetEditSetTypes
'   d8      ndk     01Jul1999       Added GetEditSetHistory, GetEditSetDetails
'   d9      ndk     06Jul999        Changed the names of methods to adhere to naming conventipns
'
'                                   GetEditListsForUser changed to EditSetListForUser
'                                   GetEditSetListsForAllUsers changed to EditSetListForAllUsers
'                                   GetEditSetStatus changed to EditSetStatus
'                                   GetUsersAndRolesForEditSet changed to UsersAndRolesForEditSet
'                                   GetRoles changed to Roles
'                                   GetEditSetDetails changed to EditSetDetails
'                                   GetEditSetHistory changed to EditSetHistory
'                                   GetEditSetTypes changed to EditSetTypes
'                                   GetManagingEditorForNewEditSet to ManagingEditorForNewEditSet
'                                   GetContributor1ForNewEditSet to Contributor1ForNewEditSet
'                                   GetValidUsersForTheRole to ValidUsersForTheRole
'
'                                   Now using constants in place of hardcoded coloumn names
'   d10     ndk     07Jul1999       Added NextPreviousStagesForEditSet, MoveEditSetToDifferentStage,IsChecked
'   d10     ndk     09Jul1999       Added AssignableRoles
'   d11     ndk     12Jul1999       Added EditsetEntriesList (coded by krish)
'   d12     ndk     12Jul1999       created CreateNewCatalogueByCopying,CatalogueDetails
'                                   CreateNewCatalogueFromScratch,
'   d13     ndk     15Jul1999       Replace quotes with QuoteValueWithQuotes, added provision to handel new table lettercode_sortedtitle
'   d14     ndk     16Jul1999       Usercanchangestage - corrected
'   d15     ndk     16Jul1999       Changed CreateCatalogueEntryFromScratch to insert title in title table
'   d16     ndk     16Jul1999       Added EditSetEntriesCount
'   d17     ndk     20Jul1999       Changed commenting and description of sSource varaible
'   d18     ndk     26July1999      Moved methods ValidUsersForTheRole,Roles,AssignableRoles to BoUserRole
'                                   Move permission checking/security related methods to BOEditorialPermissions
'                                   Removed GetCatalogueDetails,EditSetEntriesCount,EditSetDetails,CreateEmptyEditSet
'                                   Added properties
'                                   Added Load, CreateNewEmpty, CurrentStageOfEditSet,UserPresentOnEditSet, UsersRoleOnEditSet
'                                   RolePresentOnEditSet, RolesUserOnEditSet
'   d19     ksk     02Aug1999       Added code for Release and Abandon Edit set. After these methods
'                                   are executed BOReleaseContoller is notified of the release so that
'                                   appropriate actions can be taken
'   d20     ksk     05Aug1999       Changed the name of the SP for releaseing catalogues
'   d21     ndk     17Aug1999       Changed the calling parameters for the sp in methods EditSetListForUser and EditSetListForAllUsers, Moved AddCataloguesIntoEditSet to BOBasicCatalogueEntry
'   d22     ndk     13Oct1999       Persmission check added
'   d23     ndk     28Oct1999       Changes because the return type of Check method changed
'                                   Duplicate editsetname checking done
'   d24     ndk     02Nov1999       Validation to check that a user can not have two roles
'   d25     ksk     10Nov1999       Code to Abandon an editset
'   d26     NDK     11Nov1999       Abandon removed from list of stages shown in drop down
'   d27     NDK     11Nov1999       Changed the way deleteeditsetentry method of bodeletecatalogue was called
'   d28     ndk     15Nov1999       Release code removed
'                                   Changed the code to use pre-release wherever release was used
'   d29     ndk     18Nov1999       Minor change
'   d30     ksk     18Nov1999       Changed the create to set the edit set id Property
'   d31     ndk     22Nov1999       In function GetCatalogueIDsInEditset now not returning readonly entries
'   d32     ndk     22Nov1999       Added method ValidUsersForTheRoleOnEditSet
'   d33     ndk     22Nov1999       Changed the sql in ValidUsersForTheRoleOnEditSet
'   d34     ndk     22Dec1999       Added method RolesAssignableByUserInEditSet
'   d35     ndk     23Dec1999       Modified ValidUsersForTheRoleOnEditSet,RolesAssignableByUserInEditSet to return the list in alphabetical order
'   d36     ndk     28Dec1999       Updating date created field when roles are added/updated to editset
'   d37     ndk     29Dec1999       Rechecking releasing of variables, especially if an error is deliberately raised
'   d38     crm     13Jan2000       Added optional arg to fetch only 'New Accession' edit
'                                   set type if it is a new import
'   d39     jcc     09Feb2000       Add comments data to insert into edit set history on creating empty edit set
'   d40     dtm     10feb2000       Modified RolesAssignableByUserInEditSet so that Accessions role is not returned
'   d41     dtm     10feb2000       Modified RolesAssignableByUserInEditset to allow managing editor to assign editor role when in edit stage
'   d42     crm     14Feb2000       Replace the bracket with a comma in CreateEmptyEditSet
'   d43     crm     14Feb2000       Added a comma to CreateEmptyEditSet
'   d44     wdp     23Feb2000       Altered user roles functionality to use currency flag to indicate
'                                   users involvement rather than deletion and readdition
'   d45     tmj     09Mar2000       Added Method GetCreator, and moved 2 methods from VWEditset
'   d46     daj     30Mar2000       Changed to get Druid DSN connection string from registry
'   d47     tmj     13Apr2000       Added Abandon Deleting methods
'   d48     daj     06Jul2000       Fixed Tom's abandon methods so they work! ;-)  Added DeleteDetailsWithLive
'   d49     daj     18Jul2000       Changed EditSetTypes to be more aware of the usage
'   d50     daj     08Aug2000       Added restrictions to EditSetTypes so that updates only use major amendments
'   d51     gdb     07Feb2001       Created - EditSetListWithoutEditor
'   d52     map     14Mar2001       Added extra error handling in DeleteDetailsWithLive
'   d53     alm     09Oct2001       Added a delete type of editset
'   d54     map     22Oct2001       Added code to handle EAD & Druid scheduling
'   d55     ALM     23Oct2001       Sets the Edit Set Type property on loading edit set details
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Option Explicit

'private business objects and framework classes
Private m_oFWDBHelper As FWDBHelper

'Private property variables
Private m_lEditsetID As Long
Private m_vEditSetName As Variant
Private m_lEditSetTypeID As Long
Private m_vEditSetTypeName As Variant
Private m_lContributor1UserId As Long
Private m_vContributor1UserName As Variant
Private m_lCurrentStageID As Integer
Private m_vCurrentStageName As Variant
Private m_lDaysInCurrentStage As Integer
Private m_vCreationDate As Variant
Private m_vDescription As Variant

'Count of catalogue entries
Private m_iTotalDepartMents As Integer
Private m_iTotalDivisions As Integer
Private m_iTotalSeries As Integer
Private m_iTotalSubSeries As Integer
Private m_iTotalSubSubSeries As Integer
Private m_iTotalPieces As Integer
Private m_iTotalItems As Integer

'   d38      crm     13Jan2000
Private Const mk_lESNewAccessionTypeID As Long = 3

' d50 - Added
Private Const mk_lESMajorAmendmentTypeID As Long = 1
Private Const mk_lESMinorAmendmentTypeID As Long = 2

' d53 - Added
Private Const mk_lESDeleteTypeID As Long = 4


Private Sub Initialise()
    Set m_oFWDBHelper = New FWDBHelper
End Sub
Private Sub Terminate()
    Set m_oFWDBHelper = Nothing
End Sub
Private Sub Class_Initialize()
    Initialise
End Sub
Private Sub Class_Terminate()
    Terminate
End Sub
Public Property Get Name() As String
    Name = "BOEditSet"
End Property
Public Property Let EditsetID(ByVal t_lEditSetID As Long)
    m_lEditsetID = t_lEditSetID
End Property
Public Property Get EditsetID() As Long
    EditsetID = m_lEditsetID
End Property
Public Property Let EditSetName(ByVal t_vEditSetName As Variant)
    m_vEditSetName = t_vEditSetName
End Property
Public Property Get EditSetName() As Variant
    EditSetName = m_vEditSetName
End Property
Public Property Let EditSetTypeID(ByVal t_lEditSetTypeID As Long)
    m_lEditSetTypeID = t_lEditSetTypeID
End Property
Public Property Get EditSetTypeID() As Long
    EditSetTypeID = m_lEditSetTypeID
End Property
Public Property Get EditSetTypeName() As Variant
    EditSetTypeName = m_vEditSetTypeName
End Property
Public Property Let Contributor1UserID(ByVal t_lContributor1UserID As Long)
    m_lContributor1UserId = t_lContributor1UserID
End Property
Public Property Get Contributor1UserID() As Long
    Contributor1UserID = m_lContributor1UserId
End Property
Public Property Get Contributor1UserName() As Variant
    Contributor1UserName = m_vContributor1UserName
End Property
Public Property Get CurrentStageID() As Long
    CurrentStageID = m_lCurrentStageID
End Property
Public Property Get CurrentStageName() As Variant
    CurrentStageName = m_vCurrentStageName
End Property
Public Property Get DaysInCurrentStage() As Long
    DaysInCurrentStage = m_lDaysInCurrentStage
End Property
Public Property Get CreationDate() As Variant
    CreationDate = m_vCreationDate
End Property
Public Property Get TotalDepartments() As Long
    TotalDepartments = m_iTotalDepartMents
End Property
Public Property Get TotalDivisions() As Long
    TotalDivisions = m_iTotalDivisions
End Property
Public Property Get TotalSeries() As Long
    TotalSeries = m_iTotalSeries
End Property
Public Property Get TotalSubSeries() As Long
    TotalSubSeries = m_iTotalSubSeries
End Property
Public Property Get TotalSubSubSeries() As Long
    TotalSubSubSeries = m_iTotalSubSubSeries
End Property
Public Property Get TotalPieces() As Long
    TotalPieces = m_iTotalPieces
End Property
Public Property Get TotalItems() As Long
    TotalItems = m_iTotalItems
End Property
Public Property Let Description(ByVal t_vDescription As Variant)
    m_vDescription = t_vDescription
End Property
Public Property Get Description() As Variant
    Description = m_vDescription
End Property




Public Function EditSetListForAllUsers() As FWRecordList
' Description:
'
' Returns a list all edit sets for all the users
'
' Amendment history:
'
' d01   NDK     14Jun1999       Created
' d21   NDK     17Aug1999       Changed the calling parameters for the sp

On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim sSQL As String
    
    sSource = Me.Name & ".EditSetListForAllUsers()"
    
    'Construct the sql
    sSQL = "sp_Edit_GetEditSetListForAllUsers " & _
            "@t_iContributor1RoleID  = " & gk_lContributor1RoleID
            
    'Run the sp and get the recordlist containing recordlists and details
    Set EditSetListForAllUsers = m_oFWDBHelper.FindBySql(sSQL)
    
    Exit Function
    
ErrorHandler:
    ConvertError sSource
End Function



Public Function EditSetListForUser(ByVal t_lUserID As Long) As FWRecordList
' Description:
'
' Returns a list all edit sets for the user whose id is passed as parameter t_lUserId
'
' Amendment history:
'
' d01   NDK     11May1999    Added
' d21   NDK     17Aug1999    Changed the calling parameters for the sp
' d23   NDK     28Oct1999    Changed the calling parameters for the sp

On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim sSQL As String
    
    sSource = Me.Name & ".EditSetListForUser"
    
    'Construct the sql
    sSQL = "sp_Edit_GetEditSetListForOneUser " & _
            "@t_iUserID = " & t_lUserID & "," & _
            "@t_iContributor1RoleID  = " & gk_lContributor1RoleID
       
    'Run the sp and get the recordlist containing recordlists and details
    Set EditSetListForUser = m_oFWDBHelper.FindBySql(sSQL)
    
    Exit Function
ErrorHandler:
    ConvertError sSource
End Function

Public Function EditSetListWithoutEditor() As FWRecordList
' Description:
'
' Returns a list all edit sets that are in editorial stage without an
' editor assigned to them.
'
' Amendment history:
'
' d51   GDB     07Feb2001       Created


On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim sSQL As String
    
    sSource = Me.Name & ".EditSetListForAllUsers()"
    
    'Construct the sql
    sSQL = "sp_Edit_GetEditSetInEditorialStageWithoutEditors " & _
           gk_lContributor1RoleID
            
    'Run the sp and get the recordlist containing recordlists and details
    Set EditSetListWithoutEditor = m_oFWDBHelper.FindBySql(sSQL)
    
    Exit Function
    
ErrorHandler:
    ConvertError sSource
End Function

Public Sub UpdateEditSetStatus( _
            ByVal t_lLoggedUserID As Long, _
            ByVal t_lEditSetID As Long, _
            ByVal t_lAssignedUser As Long, _
            ByVal t_bNewStatus As Boolean)
            
' Description:
'
' Updates the database with changes made to the completion status of the editsets
' This is called when the users checks/unchecks the status of editsets
'
' Amendment history:
'
'   d3      NDK     16Jun1999       Created
'   d18     NDK     26Jul1999       Now updates only one editset and user at a time

On Error GoTo ErrorHandler

    Dim sSQL As String, sSource As String
    Dim eSecurityCheckResultStatus As ESecurityCheckResult
    Dim oBOEditorialPermissions As BOEditorialPermissions
    
    sSource = Me.Name & ".UpdateEditSetStatus()"
        
    Set oBOEditorialPermissions = New BOEditorialPermissions
        
    oBOEditorialPermissions.UserID = t_lLoggedUserID    'User trying to change
    oBOEditorialPermissions.EditsetID = t_lEditSetID    'Edit set id
    oBOEditorialPermissions.OtherUserID = t_lAssignedUser   'who is the originally assigned user
    
    'Check whether user has permissions to change the status
    eSecurityCheckResultStatus = oBOEditorialPermissions.Check(EUserCanChangeCompletionStatusForOtherUser)
    
    If eSecurityCheckResultStatus <> ESecurityCheckResult.ESRSecurityCheckSuccessful Then
        RaiseError sSource, _
            GetErrorEnumForSecurityCheckResult(eSecurityCheckResultStatus), _
            GetDescriptionForSecurityCheckError(GetErrorEnumForSecurityCheckResult(eSecurityCheckResultStatus))
    End If
        
    'Construct the sql to update table user_role_on_editset
    sSQL = "UPDATE " & gk_sDBTABLEUserRoleOnEditset & vbCr & _
            " SET " & gk_sDBCOLComplete & " = " & IIf(t_bNewStatus, "1", "0") & vbCr & _
            " WHERE " & vbCr & _
            gk_sDBCOLEditSetID & " = " & t_lEditSetID & " AND " & vbCr & _
            gk_sDBCOLUserID & " = " & t_lAssignedUser
            
    'Update the table
    m_oFWDBHelper.Execute sSQL
    
    LogMessage "Completion status changed to :" & IIf(t_bNewStatus, "1", "0") & " for Editset :" & t_lEditSetID & " user :" & t_lAssignedUser, pProfileInformation
    
    Set oBOEditorialPermissions = Nothing
    
    Exit Sub

ErrorHandler:

    Set oBOEditorialPermissions = Nothing
    
    ConvertError sSource
End Sub

Public Function UsersAndRolesForEditSet(ByVal t_lEditSetID As Long) As FWRecordList

' Description:
'
' Returns the list of users and their roles on the editset passed as parameter t_lEditSetID
' Data is returned as a FWRecordlist
'
' Amendment history:
'
'   d5      NDK     23Jun1999       Created
On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim sSQL As String
    
    sSource = Me.Name & ".UsersAndRolesForEditSet(" & t_lEditSetID & ")"
    
    'Construct the sql
    sSQL = "sp_Edit_GetUsersAndRolesForEditSet " & _
            "@t_iEditSetID = " & t_lEditSetID & " , " & _
            "@iStageAbandonID = " & gk_lStageAbandonID & " , " & _
            "@iStagePreReleaseID = " & gk_lStagePreReleaseID

    'Run the sp and get the recordlist containing users and roles
    Set UsersAndRolesForEditSet = m_oFWDBHelper.FindBySql(sSQL)

    Exit Function
    
ErrorHandler:
    ConvertError sSource
End Function


Public Sub RemoveRoleFromEditSet( _
            ByVal t_lEditSetID As Long, _
            ByVal t_lRoleToBeRemoved As Long, _
            ByVal t_lLoggedUserID As Long)
' Description:
'
' Removes a role from the editset
'
' Amendment history:
'
'   d5      NDK     23Jun1999       Created
'   d23     NDK     28Oct1999       Not accepting parameters t_lUserHoldingRole
'   d44     wdp     23Feb2000       Altered user roles functionality to use currency flag to indicate
'                                   users involvement rather than deletion and readdition

On Error GoTo ErrorHandler

    Dim sSource As String
    Dim sSQL As String
    Dim eSecurityCheckResultStatus As ESecurityCheckResult
    Dim oBOEditorialPermissions As BOEditorialPermissions
    
    sSource = Me.Name & ".RemoveRoleFromEditSet"

    'Create instances
    Set oBOEditorialPermissions = New BOEditorialPermissions
    
    oBOEditorialPermissions.UserID = t_lLoggedUserID
    oBOEditorialPermissions.RoleToBeRemovedID = t_lRoleToBeRemoved
    oBOEditorialPermissions.EditsetID = t_lEditSetID
    
    'check the permissions
    eSecurityCheckResultStatus = oBOEditorialPermissions.Check(EUserCanRemoveRoleFromEditSet)
    
    'if not permitted raise an error
    If eSecurityCheckResultStatus <> ESecurityCheckResult.ESRSecurityCheckSuccessful Then
        RaiseError sSource, _
            GetErrorEnumForSecurityCheckResult(eSecurityCheckResultStatus), _
            GetDescriptionForSecurityCheckError(GetErrorEnumForSecurityCheckResult(eSecurityCheckResultStatus))
    End If
        
    'remove the role
    sSQL = "UPDATE " & gk_sDBTABLEUserRoleOnEditset & _
            " SET " & gk_sDBCOLCurrencyIndicator & "=0, " & _
            gk_sDBCOLDateRemoved & "=GETDATE()" & _
            " WHERE " & _
                gk_sDBCOLEditSetID & " = " & t_lEditSetID & " AND " & _
                gk_sDBCOLRoleID & " = " & t_lRoleToBeRemoved & " AND " & _
                gk_sDBCOLCurrencyIndicator & "=1"
                                
    'Execute the sql
    m_oFWDBHelper.Execute sSQL
    
    LogMessage "Role :" & t_lRoleToBeRemoved & " removed from editset :" & t_lEditSetID & " by user :" & t_lLoggedUserID, pProfileInformation
    
    Set oBOEditorialPermissions = Nothing
    
    Exit Sub

ErrorHandler:

    Set oBOEditorialPermissions = Nothing
    
    ConvertError sSource

End Sub

Public Sub AddUpdateRoleOnEditSet(ByVal t_lEditSetID As Long, _
                                ByVal t_RoleToBeAdded As Long, _
                                ByVal t_lUserIDToBeAdded As Long, _
                                ByVal t_lLoggedUserID As Long)
' Description:
'
' Adds a role to the edit set, if the role already exists then the existing user with the role
' is replaced with the new user
'
' Amendment history:
'
'   d5      NDK     23Jun1999       Created
'   d23     NDK     28Oct1999       Security check changed
'   d24     ndk     02Nov1999       Validation to check that a user can not have two roles
'   d36     ndk     28Dec1999       Updating date created field when roles are added/updated to editset
'
'   d44     wdp     23Feb2000       Altered user roles functionality to use currency flag to indicate
'                                   users involvement rather than deletion and readdition
'
On Error GoTo ErrorHandler

    Dim sSource As String, sSQL As String
    Dim eSecurityCheckResultStatus As ESecurityCheckResult
    Dim oBOEditorialPermissions As BOEditorialPermissions
    
    sSource = Me.Name & ".AddUpdateRoleOnEditSet"

    'Create instances
    Set oBOEditorialPermissions = New BOEditorialPermissions
    
    oBOEditorialPermissions.UserID = t_lLoggedUserID
    oBOEditorialPermissions.RoleToBeAddedID = t_RoleToBeAdded
    oBOEditorialPermissions.EditsetID = t_lEditSetID
    oBOEditorialPermissions.OtherUserID = t_lUserIDToBeAdded
    
    'check the permissions
    eSecurityCheckResultStatus = oBOEditorialPermissions.Check(EUserCanAddRoleToEditSet)
    
    'if not permitted raise an error
    If eSecurityCheckResultStatus <> ESecurityCheckResult.ESRSecurityCheckSuccessful Then
        RaiseError sSource, _
            GetErrorEnumForSecurityCheckResult(eSecurityCheckResultStatus), _
            GetDescriptionForSecurityCheckError(GetErrorEnumForSecurityCheckResult(eSecurityCheckResultStatus))
    End If
    
    'Make current user role inactive
    sSQL = "UPDATE " & _
                gk_sDBTABLEUserRoleOnEditset & _
            " SET " & _
                gk_sDBCOLCurrencyIndicator & "=0, " & _
                gk_sDBCOLDateRemoved & "=GETDATE()" & _
            " WHERE " & _
                gk_sDBCOLEditSetID & " = " & t_lEditSetID & " AND " & _
                gk_sDBCOLRoleID & " = " & t_RoleToBeAdded & " AND " & _
                gk_sDBCOLCurrencyIndicator & "=1"
    
    m_oFWDBHelper.Execute sSQL
    
    If HasHadRoleOnEditSet(t_RoleToBeAdded, t_lEditSetID, t_lUserIDToBeAdded) Then
        'do an update on the active flag
        sSQL = "UPDATE " & gk_sDBTABLEUserRoleOnEditset & " SET " & _
                gk_sDBCOLCurrencyIndicator & "=1, " & _
                gk_sDBCOLDateRemoved & "=NULL WHERE " & _
                gk_sDBCOLEditSetID & "=" & t_lEditSetID & " AND " & _
                gk_sDBCOLUserID & "=" & t_lUserIDToBeAdded & " AND " & _
                gk_sDBCOLRoleID & "=" & t_RoleToBeAdded
                
    Else
        sSQL = "INSERT " & _
                    gk_sDBTABLEUserRoleOnEditset & "(" & _
                    gk_sDBCOLRoleID & ", " & _
                    gk_sDBCOLEditSetID & ", " & _
                    gk_sDBCOLUserID & ", " & _
                    gk_sDBCOLCurrencyIndicator & "," & _
                    gk_sDBCOLDateCreated & "," & _
                    gk_sDBCOLComplete & ") " & _
                " VALUES (" & _
                    t_RoleToBeAdded & "," & t_lEditSetID & "," & t_lUserIDToBeAdded & ",1 " & "," & "GetDate()" & ",0)"
    End If
    'Execute the sql to add or update the role details
    m_oFWDBHelper.Execute sSQL

    LogMessage "Role :" & t_RoleToBeAdded & " assigned to user :" & t_lUserIDToBeAdded & _
        " on editset :" & t_lEditSetID & " by user :" & t_lLoggedUserID, pProfileInformation
    
    Set oBOEditorialPermissions = Nothing
    
    Exit Sub

ErrorHandler:

    Set oBOEditorialPermissions = Nothing
    
    ConvertError sSource

End Sub


Public Function EditSetTypes(Optional t_eUsage As ECreateEditSetUsage = cesNormal) As FWRecordList
'
' Description:
'
' Returns the list of editsettypes to be used in a list box
'
' Amendment history:
'
'   d7      ndk     30Jun1999       Created
'   d38     crm     13Jan2000       Added optional arg and code to fetch only 'New Accession' edit
'                                   set type if it is a new import
'   d49     daj     18Jul2000       Changed EditSetTypes to be more aware of the usage
'   d50     daj     08Aug2000       Added restrictions to EditSetTypes so that updates only use major amendments
'   d54     map     22Oct2001       Added DRUID Load case
On Error GoTo ErrorHandler

    Dim sSource As String
    Dim sSQL As String
    
    sSource = Me.Name & ".EditSetTypes()"
    
    Select Case t_eUsage
        ' d54 - Added cesDruidLoad
        Case cesImport, cesDRUIDLoad
            sSQL = "SELECT " & _
                            gk_sDBCOLEditSetType & " AS " & gk_sDBCOLID & " , " & _
                            gk_sDBCOLEditSetTypeDescription & " AS " & gk_sDBCOLValue & _
                        " FROM " & _
                            gk_sDBTABLEEditSetType & _
                        " WHERE " & _
                            gk_sDBCOLEditSetType & " = " & mk_lESNewAccessionTypeID
        Case cesUpdate
            sSQL = "SELECT " & _
                            gk_sDBCOLEditSetType & " AS " & gk_sDBCOLID & " , " & _
                            gk_sDBCOLEditSetTypeDescription & " AS " & gk_sDBCOLValue & _
                        " FROM " & _
                            gk_sDBTABLEEditSetType & _
                        " WHERE " & _
                            gk_sDBCOLEditSetType & " = " & mk_lESMajorAmendmentTypeID
        
        'd53 - Added
        Case cesDelete
            sSQL = "SELECT " & _
                            gk_sDBCOLEditSetType & " AS " & gk_sDBCOLID & " , " & _
                            gk_sDBCOLEditSetTypeDescription & " AS " & gk_sDBCOLValue & _
                        " FROM " & _
                            gk_sDBTABLEEditSetType
                            
        'd53 - Edited to exclude 'delete'
        Case cesNormal, Default
            sSQL = "SELECT " & _
                            gk_sDBCOLEditSetType & " AS " & gk_sDBCOLID & " , " & _
                            gk_sDBCOLEditSetTypeDescription & " AS " & gk_sDBCOLValue & _
                        " FROM " & _
                            gk_sDBTABLEEditSetType & _
                        " WHERE " & _
                            gk_sDBCOLEditSetType & " != " & mk_lESDeleteTypeID
    End Select
    
    'Run the sql and get the recordlist of roles
    Set EditSetTypes = m_oFWDBHelper.FindBySql(sSQL)
    
    Exit Function

ErrorHandler:
    ConvertError sSource

End Function

Public Sub CreateNewEmpty(ByVal t_lUserIDCreatingTheEditSet As Long)
                            
'
' Description:
'
' Creates a new empty edit set using properties ( this replaces CreateEmptyEditSet method)
'
' Amendment history:
'
'   d1      ndk     27July999       Created
'   d23     ndk     28Oct1999       Duplicate editsetname checking done
'   d36     ndk     28Dec1999       Updating date created field when roles are added/updated to editset
'   d39     jcc     09Feb2000       Add comments data to insert into edit set history on creating empty edit set
'   d42     crm     14Feb2000       Replace the bracket with a comma in CreateEmptyEditSet
On Error GoTo ErrorHandler

    Dim sSource As String
    Dim sSQL As String
    Dim lNewEditSetID As Long, oFWRecordlist As FWRecordList
    
    sSource = Me.Name & ".CreateNewEmpty"
    
    'Check whether an editset with such a name already exists
    sSQL = _
            " SELECT * from " & gk_sDBTABLEEditset & _
            " WHERE UPPER(LTRIM(RTRIM(" & gk_sDBCOLName & "))) = " & UCase(Trim(QuoteValueWithQuotes(EditSetName)))
    
    Set oFWRecordlist = m_oFWDBHelper.FindBySql(sSQL)
    
    If oFWRecordlist.RecordCount <> 0 Then
        RaiseError sSource, EPROCATINTERNALERROR.PEDuplicateEditSetName, _
            "Editset by name : " & EditSetName & " already exists"
    End If
    
    'Insert the edit set into the edit set table
    sSQL = "INSERT " & _
                gk_sDBTABLEEditset & "(" & _
                gk_sDBCOLName & "," & _
                gk_sDBCOLEditSetType & "," & _
                gk_sDBCOLDescription & ")" & _
            " VALUES (" & _
                QuoteValueWithQuotes(EditSetName) & "," & _
                EditSetTypeID & "," & _
                QuoteValueWithQuotes(Description) & ")"
                    
    'Insert and  Get the newly inserted edit set id
    lNewEditSetID = m_oFWDBHelper.InsertAndReturnIdentity(sSQL)
    
    'd30     ksk     18Nov1999       Changed the create to set the edit set id Property
    m_lEditsetID = lNewEditSetID

'   d39     jcc     09Feb2000       Add comments data to insert into edit set history on creating empty edit set
'   d42     crm     14Feb2000       Replace the bracket with a comma
'   d43     crm     14Feb2000       Added a comma
    'Insert details into edit set history
    sSQL = "INSERT " & _
                gk_sDBTABLEEditsetHistory & "(" & _
                gk_sDBCOLEditSetID & "," & _
                gk_sDBCOLEditSetStage & "," & _
                gk_sDBCOLEditToSetStage & "," & _
                gk_sDBCOLEditTranstitionDate & "," & _
                gk_sDBCOLUserID & "," & _
                gk_sDBCOLComment & ")" & _
            " VALUES (" & _
                lNewEditSetID & "," & _
                gk_lStageContributionID & "," & _
                gk_lStageContributionID & "," & _
                "GetDate()" & "," & _
                t_lUserIDCreatingTheEditSet & "," & _
                QuoteValueWithQuotes(Description) & ")"
    
    m_oFWDBHelper.Execute sSQL
    
    'Insert 1 rows in user_role_on_edit set for Contributor 1
    
    sSQL = "INSERT " & _
                gk_sDBTABLEUserRoleOnEditset & "(" & _
                gk_sDBCOLEditSetID & "," & _
                gk_sDBCOLRoleID & "," & _
                gk_sDBCOLUserID & "," & _
                gk_sDBCOLDateCreated & "," & _
                gk_sDBCOLComplete & ")" & _
            " VALUES (" & _
                lNewEditSetID & "," & _
                gk_lContributor1RoleID & "," & _
                Contributor1UserID & "," & _
                "GetDate()" & "," & _
                "0" & ")"

    'Add row
    m_oFWDBHelper.Execute sSQL
    
    LogMessage "New editset :" & lNewEditSetID & " created by :" & t_lUserIDCreatingTheEditSet & " with " & _
        Contributor1UserID & " as contributor1", pProfileInformation
    
    Set oFWRecordlist = Nothing
    
    Exit Sub

ErrorHandler:

    Set oFWRecordlist = Nothing
    
    ConvertError sSource
End Sub


Public Function EditSetHistory(t_lEditSetID As Long) As FWRecordList
'
' Description:
'
' Returns the history of the editset, like the stages it went
' through , days spent in stage
'
' Amendment history:
'
'   d8      ndk     01Jul1999       Created
On Error GoTo ErrorHandler

    Dim sSource As String
    Dim sSQL As String
    
    sSource = Me.Name & ".EditSetHistory(" & t_lEditSetID & ")"
    
    'construct the sql
    sSQL = "sp_Edit_GetEditSetHistory " & _
            "@t_iEditSetID = " & t_lEditSetID
    
    'Execute the sp
    Set EditSetHistory = m_oFWDBHelper.FindBySql(sSQL)
    
    If EditSetHistory.GetFieldValue(1) = -1 Then
        RaiseError sSource, EPROCATINTERNALERROR.PEInvalidEditSetHistory, _
            "Invalid data found in edit set history for editset : " & t_lEditSetID
    End If
    
    Exit Function

ErrorHandler:
    ConvertError sSource
End Function

Public Function Load()
' Description:
'
' Loads the details of the editset whose id is set in property EditSetID
'
' Amendment history:
'

'   d8      NDK     27Jul1999      Created
'   d29     ndk     18Nov1999      Minor change
'   d55     ALM     23Oct2001      Sets the Edit Set Type property
On Error GoTo ErrorHandler

    Dim oFWRecordlist As FWRecordList
    Dim sSource As String
    Dim sSQL As String
    
    sSource = Me.Name & ".Load"
    
    'Construct Sql to search the database
    'construct the sql
    sSQL = "sp_Edit_GetEditSetDetails " & _
            "@t_iEditSetID = " & EditsetID & " , " & _
            "@t_iRoleIDForContributor1  = " & gk_lContributor1RoleID
    
    'Execute the sp
    Set oFWRecordlist = m_oFWDBHelper.FindBySql(sSQL)
    
    If oFWRecordlist.RecordCount = 0 Then
        RaiseError sSource, EPROCATINTERNALERROR.peEditSetIDInvalue, "Editset not present :" & EditsetID
    End If
    
    m_vEditSetName = oFWRecordlist.GetFieldValue(gk_sDBCOLEditsetName)
    m_vContributor1UserName = oFWRecordlist.GetFieldValue(gk_sDBCOLTempContributor1)
    m_vCreationDate = oFWRecordlist.GetFieldValue(gk_sDBCOLTempEditSetCreationDate)
    m_lDaysInCurrentStage = oFWRecordlist.GetFieldValue(gk_sDBCOLTempDaysInCurrentStage)
    m_lCurrentStageID = oFWRecordlist.GetFieldValue(gk_sDBCOLStageID)
    m_vCurrentStageName = oFWRecordlist.GetFieldValue(gk_sDBCOLEditsetStageDescription)
    
    'd55
    m_lEditSetTypeID = oFWRecordlist.GetFieldValue(gk_sDBCOLEditSetType)
    
    'The count of entries
    m_iTotalDepartMents = oFWRecordlist.GetFieldValue(gk_sDBCOLTempLetterCodeCount)
    m_iTotalDivisions = oFWRecordlist.GetFieldValue(gk_sDBCOLTempDivisionCount)
    m_iTotalSeries = oFWRecordlist.GetFieldValue(gk_sDBCOLTempClassCount)
    m_iTotalSubSeries = oFWRecordlist.GetFieldValue(gk_sDBCOLTempHeaderCount)
    m_iTotalSubSubSeries = oFWRecordlist.GetFieldValue(gk_sDBCOLTempSubHeaderCount)
    m_iTotalPieces = oFWRecordlist.GetFieldValue(gk_sDBCOLTempPieceCount)
    m_iTotalItems = oFWRecordlist.GetFieldValue(gk_sDBCOLTempItemCount)

    'Release the objects
    Set oFWRecordlist = Nothing
    
    Exit Function
    
ErrorHandler:
        
    Set oFWRecordlist = Nothing

    ConvertError sSource

End Function


Public Function NextPreviousStagesForEditSet( _
                ByVal t_lStageID As Long, _
                ByVal t_bMoveToNextStage As Boolean) As FWRecordList
' Description:
'
' Returns the list of valid next/previous stages for an editset
'
' Amendment history:
'
'   d8      NDK     01Jul1999      Created
'   d26     NDK     11Nov1999      Abandon removed from list of stages shown in drop down

On Error GoTo ErrorHandler

    Dim sSource As String
    Dim sSQL As String
    
    sSource = Me.Name & ".NextPreviousStagesForEditSet"
    
    'construct the sql
    
    If t_bMoveToNextStage Then
        'Stages for the next stage
        sSQL = " SELECT " & _
                    " ESST." & gk_sDBCOLEditSetStage1 & " AS " & gk_sDBCOLID & ", " & _
                    " ESS." & gk_sDBCOLEditsetStageDescription & " AS " & gk_sDBCOLValue & _
                " FROM " & _
                    gk_sDBTABLEEditSetStageTransition & " ESST , " & _
                    gk_sDBTABLEEditSetStage & " ESS " & _
                " WHERE " & _
                    " ESST." & gk_sDBCOLEditSetStage1 & " = " & "ESS." & gk_sDBCOLStageID & " AND " & _
                    " ESST." & gk_sDBCOLEditSetStage & " = " & t_lStageID & " AND " & _
                    " ESST." & gk_sDBCOLEditSetStage1 & " <> " & gk_lStageAbandonID & " AND " & _
                    " ESST." & gk_sDBCOLForward_Transition & " =  1 "
                    
          'To exclude the first transtition (contributor1 to contributor1)
          sSQL = sSQL & " AND NOT ESST." & gk_sDBCOLEditSetStage1 & " = " & " ESST." & gk_sDBCOLEditSetStage
                    
    Else
        'Stages for the previous stage
        sSQL = " SELECT " & _
                    " ESST." & gk_sDBCOLEditSetStage1 & " AS " & gk_sDBCOLID & ", " & _
                    " ESS." & gk_sDBCOLEditsetStageDescription & " AS " & gk_sDBCOLValue & _
                " FROM " & _
                    gk_sDBTABLEEditSetStageTransition & " ESST , " & _
                    gk_sDBTABLEEditSetStage & " ESS " & _
                " WHERE " & _
                    " ESST." & gk_sDBCOLEditSetStage1 & " = " & "ESS." & gk_sDBCOLStageID & " AND " & _
                    " ESST." & gk_sDBCOLEditSetStage & " = " & t_lStageID & " AND " & _
                    " ESST." & gk_sDBCOLEditSetStage1 & " <> " & gk_lStageAbandonID & " AND " & _
                    " ESST." & gk_sDBCOLForward_Transition & " = 0 "
    End If
    
        
    'Execute the sql
    Set NextPreviousStagesForEditSet = m_oFWDBHelper.FindBySql(sSQL)

    Exit Function
ErrorHandler:
    ConvertError sSource
End Function


Public Function EditsetEntriesList(t_lEditSetID As Long) As FWRecordList
'   d11     ndk     12Jul1999       Created(coded by krish)
'Returns the list of catalogues in an editset
On Error GoTo ErrorHandler
                
    Dim sSQL As String
    Dim sSource As String
    
    sSource = Me.Name & ".EditsetEntriesList(" & t_lEditSetID & ")"
    
    
    sSQL = "execute sp_Edit_CatalogueList @t_iEditsetID = " & t_lEditSetID
    
    
    Set EditsetEntriesList = m_oFWDBHelper.FindBySql(sSQL)

    Exit Function
ErrorHandler:
    ConvertError sSource
End Function

Public Function UsersRoleOnEditSet( _
                    ByVal t_lUserID As Long, _
                    ByVal t_lEditSetID As Long) As Long
    
' Description: Returns the role of the user on the editset, if the user
'              does not have a role it returns -1
'
' Amendment history:
'   d18      NDK     26July1999       Created
'
'   d44     wdp     23Feb2000       Altered user roles functionality to use currency flag to indicate
'                                   users involvement rather than deletion and readdition
    On Error GoTo ErrorHandler
    
    Dim sSource As String, sSQL As String
    Dim oFWRecordlist As FWRecordList
    
    sSource = Me.Name & ".UsersRoleOnEditSet()"
    
    sSQL = _
        " SELECT " & gk_sDBCOLRoleID & " FROM " & _
            gk_sDBTABLEUserRoleOnEditset & _
        " WHERE " & _
            gk_sDBCOLUserID & " = " & t_lUserID & " AND " & _
            gk_sDBCOLEditSetID & " = " & t_lEditSetID & " AND " & _
            gk_sDBCOLCurrencyIndicator & "=1"
    
    'Execute to create a recordset
    Set oFWRecordlist = m_oFWDBHelper.FindBySql(sSQL)
    
    If oFWRecordlist.RecordCount = 0 Then
        'No, the user does not have a role on the editset
        UsersRoleOnEditSet = -1
    Else
        'OK, user has a role on the editset
        UsersRoleOnEditSet = oFWRecordlist.GetFieldValue(gk_sDBCOLRoleID)
    End If
    
    'Release the obejects
    Set oFWRecordlist = Nothing
    
    Exit Function
ErrorHandler:
    Set oFWRecordlist = Nothing

    ConvertError sSource
End Function

Public Function RolesUserOnEditSet( _
                    ByVal t_lRoleID As Long, _
                    ByVal t_lEditSetID As Long) As Long

' Description: Returns the user who has the role of on the editset, if the role
'              does not have a user assigned it returns -1
'
' Amendment history:
'   d18      NDK     26July1999       Created
'
'   d44     wdp     23Feb2000       Altered user roles functionality to use currency flag to indicate
'                                   users involvement rather than deletion and readdition
    On Error GoTo ErrorHandler

    Dim sSource As String, sSQL As String
    Dim oFWRecordlist As FWRecordList
    
    sSource = Me.Name & ".RolesUserOnEditSet()"
    
    
    sSQL = _
            " SELECT " & gk_sDBCOLUserID & " FROM " & _
                gk_sDBTABLEUserRoleOnEditset & _
            " WHERE " & _
                gk_sDBCOLRoleID & " = " & t_lRoleID & " AND " & _
                gk_sDBCOLEditSetID & " = " & t_lEditSetID & " AND " & _
                gk_sDBCOLCurrencyIndicator & "=1"
    
    'Execute to create a recordset
    Set oFWRecordlist = m_oFWDBHelper.FindBySql(sSQL)
    
    If oFWRecordlist.RecordCount = 0 Then
        'No, the user does not have a role on the editset
        RolesUserOnEditSet = -1
    Else
        'OK, user has a role on the editset
        RolesUserOnEditSet = oFWRecordlist.GetFieldValue(gk_sDBCOLUserID)
    End If
    
    'Release the obejects
    Set oFWRecordlist = Nothing
    
    Exit Function
ErrorHandler:
    Set oFWRecordlist = Nothing

    ConvertError sSource
End Function

Public Function UserPresentOnEditSet( _
                    ByVal t_lUserID As Long, _
                    ByVal t_lEditSetID As Long) As Boolean
    
    
' Description:  Checks whether the user has a role on the editset
'
' Amendment history:
'   d18      NDK     26July1999       Created
'
    Dim sSource As String

    On Error GoTo ErrorHandler

    sSource = Me.Name & ".UserPresentOnEditSet()"
    
    'Get the user's role on the editset
    If UsersRoleOnEditSet(t_lUserID, t_lEditSetID) > 0 Then 'The users has a role
        UserPresentOnEditSet = True
    Else
        UserPresentOnEditSet = False    'The users does not have a role
    End If
    
    Exit Function
ErrorHandler:
    ConvertError sSource
End Function

Public Function RolePresentOnEditSet( _
                    ByVal t_lRoleID As Long, _
                    ByVal t_lEditSetID As Long) As Boolean

' Description: Checks whether the role is present on the editset
'
' Amendment history:
'   d18      NDK     26July1999       Created
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String

    sSource = Me.Name & ".RolePresentOnEditSet()"
    
    If RolesUserOnEditSet(t_lRoleID, t_lEditSetID) > 0 Then
        RolePresentOnEditSet = True
    Else
        RolePresentOnEditSet = False
    End If

    Exit Function
ErrorHandler:
    ConvertError sSource
End Function

Public Function HasHadRoleOnEditSet( _
                    ByVal t_lRoleID As Long, _
                    ByVal t_lEditSetID As Long, _
                    ByVal t_lUserID As Long) As Boolean

' Description: Checks whether user has already had role on edit set
'              (entries are kept in the user_role_on_edit_set table
'               but marked as deleted)
'
' Amendment history:
'   d44      WDP     23Feb1999       Created
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String

    sSource = Me.Name & ".HasHadRoleOnEditSet()"
    
    Dim sSQL As String
    sSQL = " SELECT * FROM " & gk_sDBTABLEUserRoleOnEditset & _
            " WHERE " & gk_sDBCOLUserID & "=" & t_lUserID & " AND " & _
            gk_sDBCOLRoleID & "=" & t_lRoleID & " AND " & _
            gk_sDBCOLEditSetID & "=" & t_lEditSetID
            
    Dim ofwHelper As New FWDBHelper
    Dim ofwResult As New FWRecordList
    
    Set ofwResult = ofwHelper.FindBySql(sSQL)
    If ofwResult.RecordCount > 0 Then
        HasHadRoleOnEditSet = True
    Else
        HasHadRoleOnEditSet = False
    End If
    
    Exit Function
ErrorHandler:
    ConvertError sSource
End Function


Public Function CurrentStageOfEditSet(ByVal t_lEditSetID As Long) As Long
' Description:
'
'Returns the current stage of the editset t_lEditSetId
'
' Amendment history:
'   d18      NDK     26July1999       Created
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String, sSQL As String
    Dim oFWRecordlist As FWRecordList
    
    sSource = Me.Name & ".CurrentStageOfEditSet"

    sSQL = "SELECT " & gk_sDBCOLEditToSetStage & _
                " FROM " & gk_sDBTABLEEditsetHistory & _
                " WHERE " & _
                        gk_sDBCOLEditSetID & " = " & t_lEditSetID & " AND " & _
                        gk_sDBCOLEditTranstitionDate & " = (" & _
                        "SELECT Max(" & gk_sDBCOLEditTranstitionDate & ")" & _
                            " FROM " & gk_sDBTABLEEditsetHistory & _
                            " WHERE " & _
                                    gk_sDBCOLEditSetID & " = " & t_lEditSetID & ")"
            
    
    Set oFWRecordlist = m_oFWDBHelper.FindBySql(sSQL)
    
    CurrentStageOfEditSet = oFWRecordlist.GetFieldValue(gk_sDBCOLEditToSetStage)

    Set oFWRecordlist = Nothing

    Exit Function

ErrorHandler:
    Set oFWRecordlist = Nothing

    ConvertError sSource
End Function
Public Sub MoveEditSetToDifferentStage( _
            ByVal t_lEditSetID As Long, _
            ByVal t_lCurrentStageID As Long, _
            ByVal t_lStageEditSetIsToBeMoved As Long, _
            ByVal t_bMoveToNextStage As Boolean, _
            ByVal t_lLoggedUserID As Long, _
            ByVal t_sComment As String)

' Description:
'
' Moves the edit set from one stage to another stage
'
' Amendment history:
'

'   d8      NDK     01Aug1999     Created
'   d19     ksk     02Aug1999     Added extra code to check for Abandon and release
'                                 of an editset
'   d23     ndk     28Oct1999     changed security check
'   d28     ndk     15Nov1999       Release code removed
On Error GoTo ErrorHandler

    Dim eSecurityCheckResultStatus As ESecurityCheckResult
    Dim sSource As String, sSQL As String
    Dim oBOEditorialPermissions As BOEditorialPermissions
    
    sSource = Me.Name & ".MoveEditSetToDifferentStage"
    
    'create instances
    Set oBOEditorialPermissions = New BOEditorialPermissions
    
    oBOEditorialPermissions.UserID = t_lLoggedUserID
    oBOEditorialPermissions.EditsetID = t_lEditSetID
    oBOEditorialPermissions.StageToBeMovedFromID = t_lCurrentStageID
    oBOEditorialPermissions.StageToBeMovedToID = t_lStageEditSetIsToBeMoved
    oBOEditorialPermissions.ForwardTransition = t_bMoveToNextStage
    
    'check the permissions
    eSecurityCheckResultStatus = oBOEditorialPermissions.Check(EUserCanMoveEditSetFromCurrentStageToDifferentStage)
    
    'if not permitted raise an error
    If eSecurityCheckResultStatus <> ESecurityCheckResult.ESRSecurityCheckSuccessful Then
        RaiseError sSource, _
            GetErrorEnumForSecurityCheckResult(eSecurityCheckResultStatus), _
            GetDescriptionForSecurityCheckError(GetErrorEnumForSecurityCheckResult(eSecurityCheckResultStatus))
    End If
        
    'd25 KSK removed abandone from the list
    
    'd28 ' release code removed
    sSQL = _
        "INSERT " & gk_sDBTABLEEditsetHistory & "(" & _
            gk_sDBCOLEditSetID & "," & _
            gk_sDBCOLEditSetStage & "," & _
            gk_sDBCOLEditToSetStage & "," & _
            gk_sDBCOLEditTranstitionDate & "," & _
            gk_sDBCOLUserID & "," & _
            gk_sDBCOLComment & ")" & _
        " VALUES (" & _
            t_lEditSetID & "," & _
            t_lCurrentStageID & "," & _
            t_lStageEditSetIsToBeMoved & "," & _
            "GetDate()" & "," & _
            t_lLoggedUserID & "," & _
            QuoteValueWithQuotes(t_sComment) & ")"
                
    'insert the row to change stage
    m_oFWDBHelper.Execute sSQL

    LogMessage "Editset :" & t_lEditSetID & " moved from stage : " & t_lCurrentStageID & " to " & t_lStageEditSetIsToBeMoved & _
        " by user : " & t_lLoggedUserID, pProfileInformation

    Set oBOEditorialPermissions = New BOEditorialPermissions

    Exit Sub
ErrorHandler:
    Set oBOEditorialPermissions = New BOEditorialPermissions

    ConvertError sSource
End Sub

Private Function GetSQLForCatalogueListInEditset(t_eCatalogueLevel As ECATALOGUELEVEL, _
                                            t_lEditSetID As Long) As String
'   d19     ksk     02Aug1999
' Function to generate the an SQL to return all the catalogues (ID ) of a level in an edit set
' Parameter t_eCatalogueLevel - Level of the Catalogue to be released
'           t_lEditSetId - Editset to that contains the catalogue to be released
'

    Select Case t_eCatalogueLevel
    Case ECATALOGUELEVEL.ctLetterCode
        GetSQLForCatalogueListInEditset = "Select " & gk_sDBCOLLetterCodeID & _
                                                " as " & gk_sDBCOLCatalogueID & _
                                            ", 1 as " & gk_sDBCOLCatalogueLevel & _
                                            " from " & gk_sDBTABLEEditsetLettercode & " (Nolock)" & _
                                            " Where Edit_Set_ID = " & CStr(t_lEditSetID)
    Case ECATALOGUELEVEL.ctDivision
        GetSQLForCatalogueListInEditset = "Select " & gk_sDBCOLDivisionID & _
                                                " as " & gk_sDBCOLCatalogueID & _
                                            ", 2 as " & gk_sDBCOLCatalogueLevel & _
                                            " from " & gk_sDBTABLEEditsetDivision & " (Nolock)" & _
                                            " Where Edit_Set_ID = " & CStr(t_lEditSetID)
    Case ECATALOGUELEVEL.ctClass
        GetSQLForCatalogueListInEditset = "Select " & gk_sDBCOLClassID & _
                                                " as " & gk_sDBCOLCatalogueID & _
                                            ", 3 as " & gk_sDBCOLCatalogueLevel & _
                                            " from " & gk_sDBTABLEEditsetClass & " (Nolock)" & _
                                            " Where Edit_Set_ID = " & CStr(t_lEditSetID)
    Case ECATALOGUELEVEL.ctHeaders
        GetSQLForCatalogueListInEditset = "Select " & gk_sDBCOLHeaderID & _
                                                " as " & gk_sDBCOLCatalogueID & _
                                            ", 4 as " & gk_sDBCOLCatalogueLevel & _
                                            " from " & gk_sDBTABLEEditsetHeader & " (Nolock)" & _
                                            " Where Edit_Set_ID = " & CStr(t_lEditSetID)
    Case ECATALOGUELEVEL.ctSubHeaders
        GetSQLForCatalogueListInEditset = "Select " & gk_sDBCOLSubHeaderID & _
                                                " as " & gk_sDBCOLCatalogueID & _
                                            ", 5 as " & gk_sDBCOLCatalogueLevel & _
                                            " from " & gk_sDBTABLEEditsetSubheader & " (Nolock)" & _
                                            " Where Edit_Set_ID = " & CStr(t_lEditSetID)
    Case ECATALOGUELEVEL.ctPiece
        GetSQLForCatalogueListInEditset = "Select " & gk_sDBCOLPieceID & _
                                                " as " & gk_sDBCOLCatalogueID & _
                                            ", 6 as " & gk_sDBCOLCatalogueLevel & _
                                            " from " & gk_sDBTABLEEditsetPiece & " (Nolock)" & _
                                            " Where Edit_Set_ID = " & CStr(t_lEditSetID)
    Case ECATALOGUELEVEL.ctItem
        GetSQLForCatalogueListInEditset = "Select " & gk_sDBCOLItemID & _
                                                " as " & gk_sDBCOLCatalogueID & _
                                            ", 7 as " & gk_sDBCOLCatalogueLevel & _
                                            " from " & gk_sDBTABLEEditsetItem & " (Nolock)" & _
                                            " Where Edit_Set_ID = " & CStr(t_lEditSetID)
    End Select

End Function

Private Function GetSPForCatalogueRelease(t_eCatalogueLevel As ECATALOGUELEVEL, _
                                            t_lEditSetID As Long, _
                                            t_lCatalogueID As Long) As String
'   d19     ksk     02Aug1999
'   d20     ksk     05Aug1999   Changed name of SP
' Function to generate the SP to release a catalogue of a level in an editset
'Parameter t_eCatalogueLevel - Level of the Catalogue to be released
'           t_lEditSetId - Editset to that contains the catalogue to be released
'           t_lCatalogueID - Catalogue to be released
'
    Select Case t_eCatalogueLevel
    Case ECATALOGUELEVEL.ctLetterCode
        GetSPForCatalogueRelease = "Execute sp_Edit_ReleaseALettercodeFromEditSet " & _
                                        " @t_iEditsetID = " & CStr(t_lEditSetID) & _
                                        ", @t_iCatalogueID  = " & CStr(t_lCatalogueID)
    Case ECATALOGUELEVEL.ctDivision
        GetSPForCatalogueRelease = "Execute sp_Edit_ReleaseADivisionFromEditSet " & _
                                        " @t_iEditsetID = " & CStr(t_lEditSetID) & _
                                        ", @t_iCatalogueID  = " & CStr(t_lCatalogueID)
    Case ECATALOGUELEVEL.ctClass
        GetSPForCatalogueRelease = "Execute sp_Edit_ReleaseAClassFromEditSet " & _
                                        " @t_iEditsetID = " & CStr(t_lEditSetID) & _
                                        ", @t_iCatalogueID  = " & CStr(t_lCatalogueID)
    Case ECATALOGUELEVEL.ctHeaders
        GetSPForCatalogueRelease = "Execute sp_Edit_ReleaseAHeaderFromEditSet " & _
                                        " @t_iEditsetID = " & CStr(t_lEditSetID) & _
                                        ", @t_iCatalogueID  = " & CStr(t_lCatalogueID)
    Case ECATALOGUELEVEL.ctSubHeaders
        GetSPForCatalogueRelease = "Execute sp_Edit_ReleaseASubheaderFromEditSet " & _
                                        " @t_iEditsetID = " & CStr(t_lEditSetID) & _
                                        ", @t_iCatalogueID  = " & CStr(t_lCatalogueID)
    Case ECATALOGUELEVEL.ctPiece
        GetSPForCatalogueRelease = "Execute sp_Edit_ReleaseAPieceFromEditSet " & _
                                        " @t_iEditsetID = " & CStr(t_lEditSetID) & _
                                        ", @t_iCatalogueID  = " & CStr(t_lCatalogueID)
    Case ECATALOGUELEVEL.ctItem
        GetSPForCatalogueRelease = "Execute sp_Edit_ReleaseAItemFromEditSet " & _
                                        " @t_iEditsetID = " & CStr(t_lEditSetID) & _
                                        ", @t_iCatalogueID  = " & CStr(t_lCatalogueID)
    End Select

End Function





Public Sub AbandonEditset()
'Abandons the current editset
'   d25     ksk     10Nov1999 Created
'   d27     NDK     11Nov1999       Changed the way deleteeditsetentry method of bodeletecatalogue was called
'   d48     daj     06Jul2000       Fixed Tom's abandon methods so they work! ;-)
'
On Error GoTo ErrHandler

    Dim sSource As String
    Dim oboDelCatalogue As BODeleteCatalogue
    Dim i As Integer
    Dim lCatalogueID As Long, eLevel As ECATALOGUELEVEL
    Dim ofwIDList As FWRecordList
    Dim oboNotifyController As BOReleaseController
    
    sSource = Me.Name & ".AbandonEditset"
    
    If m_lEditsetID <= 0 Then
        RaiseError sSource, EPROCATINTERNALERROR.peEditSetIDInvalue, _
                "Editset ID not found"
    End If
    
    'Notify the Release controller to take necessary steps
    Set oboNotifyController = New BOReleaseController
    
    oboNotifyController.ReleaseObject = ERELEASEOBJECT.roEditSet
    oboNotifyController.ObjectID = m_lEditsetID
    oboNotifyController.Abandon

        
    'This procedure calls an SP that deletes all catalogues as it is being
    ' used by release too. But it is OK since there will be none to delete
    ' when called from abandon
    DeleteDetailsWithLive m_lEditsetID
    
    'Log message
    LogMessage "Abandon for Editset ID " & m_lEditsetID & " complete", _
                                        EProfile.pProfileInformation
    
    Set oboDelCatalogue = Nothing
    Set ofwIDList = Nothing
    Set oboNotifyController = Nothing
    
    Exit Sub
ErrHandler:
    Set oboDelCatalogue = Nothing
    Set ofwIDList = Nothing
    Set oboNotifyController = Nothing

    ConvertError sSource
End Sub


Private Function GetCatalogueIDsInEditset(t_eCatalogueLevel As ECATALOGUELEVEL) _
                            As FWRecordList
'Returns the list of catalogue ID's in the t_eCatalogueLevel
'   d25     ksk     10Nov1999 Created
'   d31     ndk     22Nov1999 Now not returning readonly entries,
'                             querying the extension table
'
    Dim sSQL As String
    
    sSQL = "SELECT " & GetColumnName(ECATALOGUECOLUMNTYPE.cctIDColumn, _
                t_eCatalogueLevel) & " AS " & gk_sDBCOLCatalogueID
    
    
    sSQL = sSQL & " FROM " & GetTableName(ECATALOGUETABLETYPE.cttExtensionTable, _
                            ESectionType.SECeditorial, t_eCatalogueLevel) & _
            " (nolock)"
                            
    sSQL = sSQL & " WHERE " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID)
    
    sSQL = sSQL & " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & ECATALOGUEENTRYSTATUS.cesReadeOnly

    'Force to get all the records
    Set GetCatalogueIDsInEditset = m_oFWDBHelper.FindBySql(sSQL, 0)

End Function


Private Sub DeleteDetails(t_lEditSetID As Long)
'   d19     ksk     02Aug1999
'Physically remove the details from the database
' Parameter t_lEditsetID is the Edit set to be physically removed from the database
'
' Remove all the details related to the editset from the database
'
On Error GoTo ErrorHandler
Dim sSQL As String
Dim sSource As String
    
    sSource = Me.Name & ".DeleteDetails()"


    'Deletes all catalogue entries (in all levels) and then deletes the editset detail
    sSQL = "sp_Edit_DeleteEditSet @t_iEditsetID = " & CStr(t_lEditSetID)
    m_oFWDBHelper.Execute sSQL
    
    LogMessage "Editset " & CStr(t_lEditSetID) & " was removed from PROCAT", pProfileInformation
    
    Exit Sub
    
ErrorHandler:
    ConvertError sSource
End Sub

Private Sub DeleteDetailsWithLive(t_lEditSetID As Long)
' Removes the edit set details (like DeleteDetails), but also removes the live entries
' associated with the edit set entries if they have a Live_Flag of 2 (newly created or
' accessioned)
'
'   d48     daj     06Jul2000       Created
'   d52     map     14Mar2001       Added error handling to catch deleting entries that exist in other edit sets

On Error GoTo ErrorHandler
    Dim sSQL As String
    Dim sSource As String
    sSource = Me.Name & ".DeleteDetails()"

    ' Deletes all catalogue entries (in all levels) and then deletes the editset details
    sSQL = "sp_Edit_DeleteEditSetAndLive @t_iEditsetID = " & CStr(t_lEditSetID)
    m_oFWDBHelper.Execute sSQL
    
    LogMessage "Editset " & CStr(t_lEditSetID) & " was removed from PROCAT", pProfileInformation
    
    Exit Sub
    
ErrorHandler:
    'd52 map
    If Err.Number = gk_lErrorNumberDatabase And InStr(Err.Description, gk_sErrorDescriptionEntryInAnotherEditSet) > 0 Then
        RaiseError sSource, peEntryInAnotherEditSet, Err.Description
        Exit Sub
    End If

    ConvertError sSource
End Sub


Public Function ValidUsersForTheRoleOnEditSet( _
    ByVal t_lRoleID As Long, _
    ByVal t_lEditSetID As Long) As FWRecordList
    
' Description:
'
' Returns the list valid users for the role stored in property RoleID
' Data is returned as a FWRecordlist. It does not return the list of users
' who already possess that role on the editset. This is the difference between the
' method BOUserRoles.ValidUsersForTheRole  and this method
'
' Amendment history:
'
'   d32     ndk     22Nov1999       Created
'   d33     ndk     22Nov1999       Changed the sql
'   d35     ndk     23Dec1999       Modified to return the list in alphabetical order
'   d44     wdp     23Feb2000       Altered user roles functionality to use currency flag to indicate
'                                   users involvement rather than deletion and readdition
'
    Dim sSource As String, sSQL As String
    
    sSource = Me.Name & ".ValidUsersForTheRoleOnEditSet()"
    
    'Construct the sql
   sSQL = _
    "SELECT " & _
        " U." & gk_sDBCOLUserID & " , " & _
        " U." & gk_sDBCOLUserName & _
    " FROM " & _
        gk_sDBTABLEUsers & " U," & _
        gk_sDBTABLEUserRole & " UR " & _
    " WHERE " & _
        " U." & gk_sDBCOLUserID & " = " & "UR." & gk_sDBCOLUserID & " AND " & _
        " UR." & gk_sDBCOLRoleID & " = " & t_lRoleID & " AND " & _
        " U." & gk_sDBCOLUserID & " NOT IN (" & _
            " SELECT " & _
                gk_sDBCOLUserID & _
            " FROM " & _
                gk_sDBTABLEUserRoleOnEditset & _
            " WHERE " & gk_sDBCOLEditSetID & " =" & t_lEditSetID & _
            " AND " & gk_sDBCOLCurrencyIndicator & "=1)" & _
    " ORDER BY " & " U." & gk_sDBCOLUserName
            
    'Run the sql and get the recordlist of user names
    Set ValidUsersForTheRoleOnEditSet = m_oFWDBHelper.FindBySql(sSQL)
        
    Exit Function
ErrorHandler:
    ConvertError sSource
End Function


Public Function RolesAssignableByUserInEditSet( _
    ByVal t_lUserID As Long, _
    ByVal t_lEditSetID As Long) As FWRecordList
 
' Description:
'
' Returns the list of roles that can be assigned by the user in the editset
' The list is constructed considering the role of the user, stage of editset etc
'
' Amendment history:
'
'   d34     ndk     22Dec1999       Created
'   d35     ndk     23Dec1999       Modified to return the list in alphabetical order
'   d40     dtm     10feb2000       Modified so that Accessions role is not returned
'   d41     dtm     10feb2000       Modified to allow managing editor to assign editor role when in edit stage
'
On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim sSQL As String, sWhereRoleCanBeAssigned As String, sWhereManagingEditor As String
    Dim lUsersRoleOnEditSet As Long, lCurrentStageOFEditSet As Long
    Dim oboUserRoles As BoUserRoles
    
    sSource = Me.Name & ".RolesAssignableToEditSet()"
    
    'initialise
    Set oboUserRoles = New BoUserRoles
    
    'if admin
    If oboUserRoles.UserHasRole(t_lUserID, gk_lAdministratorRoleID) Then
        Set RolesAssignableByUserInEditSet = oboUserRoles.RolesAssignableToEditSet
    Else
        '   d40     dtm     10feb2000       Modified so that Accessions role is not returned
        '   d41     dtm     10feb2000       Modified to allow managing editor to assign editor role when in edit stage
        ' remember that user may have other roles as well, so a user who is ManEd may be able to assign other roles in addition to editor

        'get the users role on editset
        lUsersRoleOnEditSet = UsersRoleOnEditSet(t_lUserID, t_lEditSetID)

        'get the current stage of editset
        lCurrentStageOFEditSet = CurrentStageOfEditSet(t_lEditSetID)

        'Construct the sql

        If oboUserRoles.UserHasRole(t_lUserID, gk_lManagingEditorRoleID) Then
            sWhereManagingEditor = "R." & gk_sDBCOLRoleID & " = " & gk_lEditorRoleID
        Else
            sWhereManagingEditor = "1=0"
        End If

        sWhereRoleCanBeAssigned = _
               "     NOT R." & gk_sDBCOLRoleID & " IN (" & _
                         gk_lAdministratorRoleID & ", " & gk_lManagingEditorRoleID & ", " & gk_lESFAEditorRoleID & ", " & _
                         gk_lPSFAEditorRoleID & ", " & gk_lLeafletEditorRoleID & ", " & gk_lPopularSearchEditorRoleID & ", " & _
                         gk_lAccessionRoleID & _
               "         ) " & _
               "     AND R." & gk_sDBCOLRoleLevel & ">=" & _
               "         (SELECT " & gk_sDBCOLRoleLevel & _
               "          FROM " & gk_sDBTABLERole & _
               "          WHERE " & gk_sDBCOLRoleID & "=" & lUsersRoleOnEditSet & ")"


         sSQL = " SELECT " & _
                "R." & gk_sDBCOLRoleID & " , " & _
                "R." & gk_sDBCOLRoleDescription & " , " & _
                "R." & gk_sDBCOLRoleLevel & _
            " FROM " & _
                gk_sDBTABLERole & " R, " & _
                gk_sDBTABLEStageRoleInvolvedIn & " SRII " & _
            " WHERE " & _
            "     R." & gk_sDBCOLRoleID & "=" & "SRII." & gk_sDBCOLRoleID & _
            " AND SRII." & gk_sDBCOLStageID & "=" & lCurrentStageOFEditSet & _
            " AND ((" & sWhereRoleCanBeAssigned & ") OR (" & sWhereManagingEditor & "))" & _
            " ORDER BY " & "R." & gk_sDBCOLRoleDescription


        'Run the sql and get the recordlist of roles
        Set RolesAssignableByUserInEditSet = m_oFWDBHelper.FindBySql(sSQL)
    End If
    
    Set oboUserRoles = Nothing
    
    Exit Function
    
ErrorHandler:
    Set oboUserRoles = Nothing
    
    ConvertError sSource
End Function

Public Function GetCreator(ByVal t_lEditSetID As Long) As String
'
' Description:
'
' Gets the name of the first contributor, this is just a wrapper around the SQL
'
' Amendment history:
'
'   d45   tmj   09Mar2000   Created

    On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim oFWRecordlist As FWRecordList
    Dim ofwSearch As FWDBHelper

    sSource = Me.Name & ".DisplayGetCreator"
    
    Dim sSQL As String
    
    sSQL = "SELECT U." & gk_sDBCOLUserName & " FROM " _
           & gk_sDBTABLEUserRoleOnEditset & " AS UROE, " & gk_sDBTABLEUsers _
           & " AS U WHERE UROE." & gk_sDBCOLUserID & " = U." _
           & gk_sDBCOLUserID & " AND UROE." & gk_sDBCOLEditSetID & " = " _
           & t_lEditSetID & " AND UROE." & gk_sDBCOLDateCreated & " = " _
           & "(SELECT MIN(" & gk_sDBCOLDateCreated & ") FROM " _
           & gk_sDBTABLEUserRoleOnEditset & " WHERE " & gk_sDBCOLEditSetID _
           & " = " & t_lEditSetID & ")"

    Set oFWRecordlist = m_oFWDBHelper.FindBySql(sSQL)

    ' Check to see that we get at least one entry
    If oFWRecordlist Is Nothing Or oFWRecordlist.RecordCount = 0 Or _
       oFWRecordlist.RecordCount <> 1 Then
        RaiseError sSource, PEGeneralInternalError, _
                     "Edit set has incorrect number of entries in PROCAT_Accession"
    End If

    ' Get the user name returned
    GetCreator = CStr(oFWRecordlist.GetData(gk_sDBCOLUserName))

Exit Function

ErrorHandler:
    ConvertError sSource
End Function

Public Function GetFilenameForEditSet(ByVal t_lEditSetID As Long) As String
' Description:
'
' Function to use the argument t_lEditSetID to find the relevant entry
' in PROCAT_Accession in the Druid database.  Returns the filename used
' to generate the edit set - if it exists.
'
' Amendment History:
'   d45     tmj     09Mar2000   Moved from VWEditset, created by daj
'   d46     daj     30Mar2000       Changed to get Druid DSN connection string from registry
On Error GoTo ErrorHandler

    Dim sSource As String
    Dim sFilename As String
    Dim sSQL As String
    Dim ofwResult As FWRecordList
    
    sSource = Me.Name & ".GetFilenameForEditSet"

    ' The SQL to get the filename back
    sSQL = "SELECT " & gk_sDBCOLImportedFilename & _
        " FROM " & gk_sDBTABLEPROCATAccession & _
        " WHERE " & gk_sDBCOLEditSetID & " = " & t_lEditSetID
        
    ' Execute the SQL
    ' d46 - Changed to use registry
    Set ofwResult = m_oFWDBHelper.FindBySql(sSQL, , loadFromRegistry(pDRUIDDSNString))
    
    ' Check to see that we get at least one entry
    If ofwResult Is Nothing Or ofwResult.RecordCount = 0 Then
        GetFilenameForEditSet = ""
        Exit Function
    End If
    
    If ofwResult.RecordCount <> 1 Then
        RaiseError sSource, PEGeneralInternalError, "Edit set has incorrect number of entries in PROCAT_Accession"
    End If

    ' Get the filename returned
    sFilename = ofwResult.GetData(gk_sDBCOLImportedFilename)
    
    If IsNull(sFilename) Or sFilename = "" Then
        sFilename = ""
    End If
    
    GetFilenameForEditSet = sFilename
    
    Exit Function
    
ErrorHandler:
    ConvertError sSource
End Function



Public Function GetBatchIDForEditSet(ByVal t_lEditSetID As Long) As Long
' Description:
'
' Function to use the argument t_lEditSetID to find the relevant entry
' in PROCAT_Accession in the Druid database.  Returns the batch ID used
' to generate the edit set - if it exists.
'
' Amendment History:
'   d45     tmj     09Mar2000   Moved from VWEditset, created by daj
'   d46     daj     30Mar2000       Changed to get Druid DSN connection string from registry

On Error GoTo ErrorHandler

    Dim sSource As String
    Dim lBatchID As Long
    Dim sSQL As String
    Dim ofwResult As FWRecordList
    
    sSource = Me.Name & ".GetFilenameForEditSet"

    ' The SQL to get the batch ID back
    sSQL = "SELECT " & gk_sDBCOLBatchID & _
        " FROM " & gk_sDBTABLEPROCATAccession & _
        " WHERE " & gk_sDBCOLEditSetID & " = " & t_lEditSetID
        
    ' Execute the SQL
    ' d46 - Changed to use registry
    Set ofwResult = m_oFWDBHelper.FindBySql(sSQL, , loadFromRegistry(pDRUIDDSNString))
    
    ' Check to see that we get at least one entry; if not return
    If ofwResult Is Nothing Or ofwResult.RecordCount = 0 Then
        GetBatchIDForEditSet = 0
        Exit Function
    End If
    
    If ofwResult.RecordCount <> 1 Then
        RaiseError sSource, PEGeneralInternalError, "Edit set has incorrect number of entries in PROCAT_Accession"
    End If

    ' Get the batch ID returned
    lBatchID = ofwResult.GetData(gk_sDBCOLBatchID)
    
    If IsNull(lBatchID) Or lBatchID = 0 Then
        lBatchID = 0
    End If
    
    GetBatchIDForEditSet = lBatchID
    
    Exit Function
    
ErrorHandler:
    ConvertError sSource
End Function

Private Sub ResetAllRecordStatus()
'   Function Resets record status of all entries in Editset
'
' Amendment History:
'   d47 tmj     13Apr2000       Created

    Dim sSource As String
    Dim sSQL As String
    Dim oFWRecordlist As FWRecordList
    
    On Error GoTo ErrorHandler
    
    'Select All pieces in editset
    sSQL = "SELECT " & gk_sDBCOLPieceID & " FROM " & gk_sDBTABLEEditsetPiece & _
           " WHERE " & gk_sDBCOLEditSetID & " = " & m_lEditsetID
           
    Set oFWRecordlist = m_oFWDBHelper.FindBySql(sSQL)
    
    While Not oFWRecordlist.IsAtEnd
    
        'basic query
        sSQL = " sp_edit_reset_record_status " & m_lEditsetID & ", " & _
                        CStr(oFWRecordlist.GetFieldValue(gk_sDBCOLPieceID)) & ", " & _
                        ECATALOGUELEVEL.ctPiece
    
        m_oFWDBHelper.Execute sSQL

        oFWRecordlist.MoveNext
    
    Wend
    
    'Select All items in editset
    sSQL = "SELECT " & gk_sDBCOLItemID & " FROM " & gk_sDBTABLEEditsetItem & _
           " WHERE " & gk_sDBCOLEditSetID & " = " & m_lEditsetID
           
    Set oFWRecordlist = m_oFWDBHelper.FindBySql(sSQL)
    
    While Not oFWRecordlist.IsAtEnd
    
        'basic query
        sSQL = " sp_edit_reset_record_status " & m_lEditsetID & ", " & _
                        CStr(oFWRecordlist.GetFieldValue(gk_sDBCOLItemID)) & ", " & _
                        ECATALOGUELEVEL.ctItem
    
        m_oFWDBHelper.Execute sSQL

        oFWRecordlist.MoveNext
    
    Wend
    
Exit Sub
        
ErrorHandler:
    ConvertError sSource
End Sub
