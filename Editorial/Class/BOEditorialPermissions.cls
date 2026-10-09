VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BOEditorialPermissions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    BOEditorialPermissions.cls
' System:       PROCAT, PRO
' Copyright:    (C) Quidnunc Limited
'
' Description:
'
' Business object to validate usage of functionality by the user
'
' Amendment history:
'
' d1    ksk     26Ju11999       Created.
' d2    ksk     27Ju11999       Coded with template
' d3    ndk     28Jul1999       Completely re-written
' d4    ndk     04Aug1999       Abandon stage not filtered out from valid forward transitions
' d5    ndk     13Aug1999       Admin and Mng Editor role considered in function UserCanChangeCompletionStatusForOtherUser
' d6    ndk     16Aug1999       Interchanged order of parameters in Role1CanChangeCompletionStatusForRole2
'                               Added empty function to check for permission to create a catalogue entry
' d7    ndk     01Oct1999       Added new enums and empty stubs, changed some functionality
' d8    aga     11Oct1999       Added UserCanMaintainAuthorityFiles method
' d9    ndk     27Oct1999       Now the check method returns an enum (instead of the earlier true/false) so that
'                               the exact reason for the failure of 'check' is known
'                               Added new enum ESecurityCheckResult
'                               Added new enums to check whether buttons like 'filter','add entry' etc should be enabled/disabled
'                               Removed Role1CanChangeCompletionStatusForRole2
' d10   ndk     01Nov1999       Change in logic while checking when af editor/proof reader are tried to be replaced in method UserCanAddRoleToEditSet
' d11   ndk     02Nov1999       Validation to check that a user can not have two roles
' d12   ndk     11Nov1999       All activities can now be performed if their role is
'                               role is valid for that stage (and they have rights)
'                               Put code for some empty functinos
' d13   ksk     12Nov1999       Added the security for importing
' d14   ndk     12Nov1999       Minor change in Role1CanRemoveRole2
' d15   ndk     15Nov1999       Change in the way me rights are handled while adding roles
' d16   ndk     15Nov1999       Minor changes to the previous change
' d17   wdp     16Nov1999       Moved enums to FWTypes
' d18   ndk     17Nov1999       Added security check for delete catalogue entry
' d19   ndk     18Nov1999       Update RoleCanOnlyBeReplaced
' d20   ndk     22Nov1999       Administrator allowed to check other users editset completion status (this was done earlier, was reintroduced)
' d21   ndk     22Nov1999       Mng editors role interpretation changed slightly, af editor and  proof reader now
'                               handled as other roles
' d23   ndk     23Nov1999       Error corrected in UserCanAddRoleToEditSet
' d24   ksk     13Dec1999       Addec codeo to seperate importing and loading in to edit set for accessioning
' d25   dtm     02Feb2000       Modified Check()
'                               Modified UserCanImportFromEAD
'                               Modified UserCanLoadIntoEditset
'                               Corrected typo of 'Catalogye' to 'Catalogue'
' d26   dtm     10Feb2000       Editor no longer in Pre-Release stage but is still allowed to send back to Checking
' d27   crm     14Feb2000       Can not move ES to PreRelease if there are still
'                               candidate terms
' d28   crm     15Feb2000       Changed implementation for d27
' d29   tmj     13Mar2000       Removed checking for Proof Reader when moving from Edit stage
' d30   tmj     20Mar2000       Changed so Editors can Remove Proof Readers and AF Editors
' d31   tmj     25Mar2000       Changed so managing editor cannot edit roles in Contribution
' d32   daj     14Jul2000       Added check for updating/exporting from EAD
' d33   gdb     08Feb2001       Added check if user is managing editor in Check
' d34   gdb     08Feb2001       Created UserIsManagingEditor()
' d35   map     22Oct2001       Added code to handle EAD & Druid scheduling
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

'Private framework and business objects

Private m_ofwRecordlist As FWRecordList
Private m_oFWDBHelper As FWDBHelper
Private m_oBOEditSet As BOEditSet
Private m_oboUserRoles As BoUserRoles

'Private strings
'Private property variables

Private m_lEditsetID As Long
Private m_lUserID As Long
Private m_lRoleToBeAddedID As Long
Private m_lRoleToBeRemovedID As Long
Private m_lStageToBeMovedFromID As Long
Private m_lStageToBeMovedToID As Long
Private m_lOtherUserID As Long
Private m_bIsForwardTransition As Boolean


Private Const gk_lCandidateTerm As Long = 4



Private Sub Class_Initialize()
    Initialise
End Sub


Private Sub Class_Terminate()
    Terminate
End Sub
Public Property Get Name() As String
    Name = "BOEditorialPermissions"
End Property
Public Property Let EditsetID(ByVal t_lEditSetID As Long)
    m_lEditsetID = t_lEditSetID
End Property
Public Property Let UserID(ByVal t_lUserID As Long)
    m_lUserID = t_lUserID
End Property
Public Property Let RoleToBeAddedID(ByVal t_lRoleToBeAddedID As Long)
    m_lRoleToBeAddedID = t_lRoleToBeAddedID
End Property
Public Property Let RoleToBeRemovedID(ByVal t_lRoleToBeRemovedID As Long)
    m_lRoleToBeRemovedID = t_lRoleToBeRemovedID
End Property
Public Property Let StageToBeMovedFromID(ByVal t_lStageToBeMovedFromID As Long)
    m_lStageToBeMovedFromID = t_lStageToBeMovedFromID
End Property
Public Property Let StageToBeMovedToID(ByVal t_lStageToBeMovedToID As Long)
    m_lStageToBeMovedToID = t_lStageToBeMovedToID
End Property
Public Property Let ForwardTransition(ByVal t_bIsForwardTransition As Boolean)
    m_bIsForwardTransition = t_bIsForwardTransition
End Property
Public Property Let OtherUserID(ByVal t_lOtherUserID As Long)
    m_lOtherUserID = t_lOtherUserID
End Property



Private Sub Initialise()
    Dim sSource As String
    sSource = Me.Name & ".Initialise"
    
    Set m_oFWDBHelper = New FWDBHelper
    Set m_oBOEditSet = New BOEditSet
    Set m_oboUserRoles = New BoUserRoles
    
End Sub



Private Sub Terminate()
    Dim sSource As String
    
    sSource = Me.Name & ".Terminate"
    
    Set m_oFWDBHelper = Nothing
    Set m_ofwRecordlist = Nothing
    Set m_oBOEditSet = Nothing
    Set m_oboUserRoles = Nothing
    
End Sub



Public Sub Reset()
    m_lEditsetID = 0
    m_lUserID = 0
    m_lRoleToBeAddedID = 0
    m_lRoleToBeRemovedID = 0
    m_lStageToBeMovedFromID = 0
    m_lStageToBeMovedToID = 0
    m_lOtherUserID = 0
    m_bIsForwardTransition = 0
End Sub



Public Function Check(ByVal eCheckType As ESecurityCheck) As ESecurityCheckResult
' Description:  Performs security check, depending on the enum value appropriate function is called
'               Before calling this function the requisite properties should be set
'
' Amendment history:
'   d3    NDK            28-Jul-1999  Created
'   d25   David Manning  02-Feb-2000  function call for EUserCanPrepareDRUIDData permissions
'                                       changed from UserCanImportFromEAD replaced with UserCanLoadIntoEditset
'   d27   crm            14-Feb-2000    Can not move ES to PreRelease if there are still
'                                        candidate terms
' d32   daj     14Jul2000       Added check for updating/exporting from EAD
' d33   gdb     08Feb2001       Added Case EUserIsManagingEditor in Check
'   d35  map     17Oct2001       Added Case EUserCanEditScheduledJobs check
    On Error GoTo ErrorHandler
    
    Dim sSource As String

    sSource = Me.Name & ".Check()"
    
    Select Case eCheckType
    
        Case EUserCanEditRolesOnEditSet
                'Check whether a user can edit roles (to enable/disable 'edit roles' button)
            
                'Properties required to be set
                'UserID,EditSetID
            
            Check = UserCanEditRolesOnEditSet
            
        Case EUserCanAddCatalogueEntryToEditSet
                'Check whether a user can add a catalogue entry to editset 'to enable/disable 'add entry ' button and validate when an entry is added
            
                'Properties required to be set
                'UserID,EditSetID
            
            Check = UserCanAddCatalogueEntryToEditSet
            
        Case EUserCanCreateCatalogueEntryOnEditSet
                'Check whether a user can create a catalogue entry on editset 'to enable/disable 'create entry ' button and validate when an entry is created
            
                'Properties required to be set
                'UserID,EditSetID
            
            Check = UserCanCreateCatalogueEntryOnEditSet
            
        Case EUserCanEditCatalogueEntryOnEditSet
                'Check whether a user can edit a catalogue entry on editset 'to enable/disable 'edit/move' button when catalogue details are displayed and validate when an entry is edited
            
                'Properties required to be set
                'UserID,EditSetID
            
            Check = UserCanEditCatalogueEntryOnEditSet
            
        Case EUserCanDeleteCatalogueEntryOnEditSet
                'Check whether a user can delete catalogue entry on editset 'to enable/disable 'delete' button when list of catalogues under an editset are displayed
            
                'Properties required to be set
                'UserID,EditSetID
            
            Check = UserCanDeleteCatalogueEntryOnEditSet
            
        Case ESecurityCheck.EUserCanMoveEditSetFromCurrentStage
                'Check whether a user can move editset from the current stage 'to enable/disable 'next stage, previous stage'
            
                'Properties required to be set
                'UserID, EditSetID
            Check = UserCanMoveEditSetFromCurrentStage
            
        Case ESecurityCheck.EUserCanSetFilterOnEditSet
                'Check whether a user can set a filter on editset entries being viewed 'to enable/disable 'filter' button
            
                'Properties required to be set
                'UserID, EditSetID
            Check = UserCanSetFilterOnEditSet
            
        Case ESecurityCheck.EUserCanAbandonEditSet
                'Check whether a user can abandon the editset  'to enable/disable 'Abandon' button
            
                'Properties required to be set
                'UserID, EditSetID
            Check = UserCanAbandonEditSet
            
        Case ESecurityCheck.EUserCanAddRoleToEditSet
                'Check whether a user can add the  role to the editset
            
                'Properties required to be set
                'UserID,EditSetID, RoleToBeAddedID,OtherUserID
            
                Check = UserCanAddRoleToEditSet
        
        Case ESecurityCheck.EUserCanRemoveRoleFromEditSet
                'Check whether a user can remove the role from editset
            
                'Properties required to be set
                'UserID,EditSetID, RoleToBeRemovedID
            
                Check = UserCanRemoveRoleFromEditSet
        
        Case ESecurityCheck.EUserCanMoveEditSetFromCurrentStageToDifferentStage
                'Check whether a user can move editset to a specified stage
            
                'Properties required to be set
                'UserID, EditSetID, m_bIsForwardTransition,m_lStageToBeMovedFromID,m_lStageToBeMovedToID
            Check = UserCanMoveEditSetFromCurrentStageToDifferentStage
        
        Case ESecurityCheck.EUserCanChangeCompletionStatusForOtherUser
                'Check whether a user can move editset from the current stage
            
                'Properties required to be set
                'UserID, EditSetID, OtherUserID
            
            Check = UserCanChangeCompletionStatusForOtherUser
            
        Case ESecurityCheck.EUserCanBeContributor1
            
                'Check whether a user can be contributor1
            
                'Properties required to be set
                'UserID
            
            Check = UserCanBeContributor1
 '
            
        Case ESecurityCheck.EUserCanMaintainAuthorityFiles
                'Check whether the user can maintain authority files
            
                'Properties required to be set
                'UserID
            
            Check = UserCanMaintainAuthorityFiles
        
        Case ESecurityCheck.EUserCanMaintainLeaflets
                'Check whether the user can maintain leaflets
            
                'Properties required to be set
                'UserID
            
            Check = UserCanMaintainLeafLets
        
         
        Case ESecurityCheck.EUserCanMaintainPopularSearches
                'Check whether the user can maintain popular searches
            
                'Properties required to be set
                'UserID
            
            Check = UserCanMaintainPopularSearches
        
        Case ESecurityCheck.EUserCanMaintainPrintedSFA
                'Check whether the user can maintain printed sfa
        
                'Properties required to be set
                'UserID
            
            Check = UserCanMaintainPrintedSFA
        
        Case ESecurityCheck.EUserCanMaintainElectronicSFA
                'Check whether the user can maintain electronic printed sfa
        
                'Properties required to be set
                'UserID
            
            Check = UserCanMaintainElectronicSFA
        
        Case ESecurityCheck.EUserCanMaintainLookUps
                'Check whether the user can maintain look ups
        
                'Properties required to be set
                'UserID
            
            Check = UserCanMaintainMaintainLookUps
        
        Case ESecurityCheck.EUserCanMaintainUsers
                'Check whether the user can maintain users
        
                'Properties required to be set
                'UserID
            
            Check = UserCanMaintainUsers
        
        Case ESecurityCheck.EUserCanBulkUpdate
            
            Check = UserCanBulkupdate
            
        Case EUserCanImportFromEAD
                ' importing from EAD
            Check = UserCanImportFromEAD
            
        Case EUserCanPrepareDRUIDData
                ' creating edit set from data imported by DRUID
            Check = UserCanLoadIntoEditset
        
        Case EUserCanLoadIntoEditset
                ' creating edit set from data imported from EAD
            Check = UserCanLoadIntoEditset
                
        Case EUserCanUpdateFromEAD
            ' Update the catalogue from an EAD file
            Check = UserCanUpdateFromEAD
            
        Case EUserCanExportToEAD
            ' Export catalogue entries to an EAD file
            Check = UserCanExportToEAD
        ' d33 - Added to check if user is the managing editor
        Case EUserIsManagingEditor
            Check = UserIsManagingEditor
            
        ' d35 - Added to check if user can enable/disable scheduled jobs
        Case EUserCanEditScheduledJobs
            Check = UserIsScheduledJobEditor
            
    End Select
    
    Exit Function
ErrorHandler:
    ConvertError sSource
End Function



Private Function RoleBelongsToStage( _
                    ByVal t_lRoleID As Long, _
                    ByVal t_lStageID As Long) As Boolean

' Description:  Checks whether the role belongs to a stage by
'               searching the table stage_role_involved_in
'
' Amendment history:
'   d3      NDK     28July1999       Created
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String, sSQL As String

    sSource = Me.Name & ".RoleBelongsToStage()"
   
    sSQL = "SELECT " & gk_sDBCOLStageID & " FROM " & _
               gk_sDBTABLEStageRoleInvolvedIn & " SRII " & _
                "WHERE " & _
                "SRII." & gk_sDBCOLStageID & " = " & t_lStageID & " AND " & _
                "SRII." & gk_sDBCOLRoleID & " = " & t_lRoleID
    
    'Execute to create a recordset
    Set m_ofwRecordlist = m_oFWDBHelper.FindBySql(sSQL)
    
    If m_ofwRecordlist.RecordCount = 0 Then
        'No, the role does not belong to the current stage
        RoleBelongsToStage = False
    Else
        'OK, belongs to the stage
        RoleBelongsToStage = True
    End If
    
    Exit Function
    
ErrorHandler:
    ConvertError sSource
End Function




Private Function Role1LevelHigherOrEqualToRole2Level( _
                    ByVal t_lRole1ID As Long, _
                    ByVal t_lRole2ID As Long) As Boolean
                    
' Description: Checks whether Role1 is higher (is above) or equal to Role2
'
' Amendment history:
'   d3      NDK     28July1999       Created
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String

    sSource = Me.Name & ".Role1LevelHigherOrEqualToRole2Level()"
    
    'Get the role levels of the two roles and compare them
    
    If RoleLevelOfRole(t_lRole1ID) <= RoleLevelOfRole(t_lRole2ID) Then
        Role1LevelHigherOrEqualToRole2Level = True
    Else
        Role1LevelHigherOrEqualToRole2Level = False
    End If

    Exit Function
    
ErrorHandler:
    ConvertError sSource
End Function



Private Function Role1CanRemoveRole2( _
                    ByVal t_lRole1ID As Long, _
                    ByVal t_lRole2ID As Long) As ESecurityCheckResult

' Description: Checks whether Role1 can remove Role2
'
' Amendment history:
'   d3      NDK     28Jul1999       Created
'   d4      ndk     27Oct1999       Changed to return ESecurityCheckResult of data
'  d14      ndk     12Nov1999       Minor change
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String

    sSource = Me.Name & ".Role1CanRemoveRole2()"
    
    'Check the role levels
    If Not Role1LevelHigherOrEqualToRole2Level(t_lRole1ID, t_lRole2ID) Then
        
        Role1CanRemoveRole2 = _
            ESecurityCheckResult.ESRRoleIsNotHigherOrEqualToOtherRole
    
    Else    'Role1 is higher than Role2
        
        'Check whether Role1 belongs to all stages Role2 belongs
        If Role1BelongsToAllStagesRole2Belongs(t_lRole1ID, t_lRole2ID) Then
            Role1CanRemoveRole2 = ESecurityCheckResult.ESRSecurityCheckSuccessful
        Else
            Role1CanRemoveRole2 = ESecurityCheckResult.ESRRoleDoesNotBelongToAllStagesOtherRoleBelongs
        End If
    End If
    
    Exit Function
    
ErrorHandler:
    ConvertError sSource
End Function



Private Function Role1CanAddRole2( _
                    ByVal t_lRole1ID As Long, _
                    ByVal t_lRole2ID As Long) As Boolean

' Description: Checks whether Role1 can add Role2
'
' Amendment history:
'   d3      NDK     28July1999       Created
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String, sSQL As String

    sSource = Me.Name & ".Role1CanAddRole2()"
    
    'Check the role levels


    If Not Role1LevelHigherOrEqualToRole2Level(t_lRole1ID, t_lRole2ID) Then
        
        Role1CanAddRole2 = False
    
    Else    'Role1 is higher than Role2
        
        'Check whether Role1 belongs to all stages Role2 belongs
        
        If Not Role1BelongsToAllStagesRole2Belongs(t_lRole1ID, t_lRole2ID) Then
            Role1CanAddRole2 = False
        Else
            Role1CanAddRole2 = True
        End If
    End If
    
    Exit Function
    
ErrorHandler:
    ConvertError sSource
End Function



Private Function Role1BelongsToAllStagesRole2Belongs( _
                    ByVal t_lRole1ID As Long, _
                    ByVal t_lRole2ID As Long) As Boolean
                    
' Description: Checks whether Role1 belongs to all the stages Role2 belongs
'
' Amendment history:
'   d3      NDK     28July1999       Created
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String, sSQL As String

    sSource = Me.Name & ".Role1BelongsToAllStagesRole2Belongs()"
                    
    Dim oFWStagesForRole1 As FWRecordList
    Dim oFWStagesForRole2 As FWRecordList
    Dim lCurrentStageBeingChecked As Long
    
    'Now get the stage/stages where the role2 belongs
                
    sSQL = "SELECT " & _
                gk_sDBCOLStageID & _
            " FROM " & _
                    gk_sDBTABLEStageRoleInvolvedIn & _
            " WHERE " & _
                gk_sDBCOLRoleID & " = " & t_lRole2ID
            
    'Execute to create a list of stages for Role2
    Set oFWStagesForRole2 = m_oFWDBHelper.FindBySql(sSQL)
    
    Role1BelongsToAllStagesRole2Belongs = True  'Initialise
    
    'Loop through the stages of role2
    Do While Not oFWStagesForRole2.IsAtEnd
        
        lCurrentStageBeingChecked = oFWStagesForRole2.GetFieldValue(gk_sDBCOLStageID)
        
        'Check whether role1 is involved in the stage : lCurrentStageBeingChecked
        sSQL = "SELECT " & _
                gk_sDBCOLStageID & _
            " FROM " & _
                    gk_sDBTABLEStageRoleInvolvedIn & _
            " WHERE " & _
                gk_sDBCOLRoleID & " = " & t_lRole1ID & " AND " & _
                gk_sDBCOLStageID & " = " & lCurrentStageBeingChecked
        
        Set oFWStagesForRole1 = m_oFWDBHelper.FindBySql(sSQL)
        
        If oFWStagesForRole1.RecordCount = 0 Then   'No it does not belong to this stage
        
            Role1BelongsToAllStagesRole2Belongs = False
            
            Exit Do
        
        End If
        
        oFWStagesForRole2.MoveNext  'Next stage of Role2
        
    Loop

    'Release the objects
    
    Set oFWStagesForRole1 = Nothing
    Set oFWStagesForRole2 = Nothing

    Exit Function
    
ErrorHandler:
    ConvertError sSource
End Function



Public Function RoleCanOnlyBeReplaced( _
        ByVal t_lRoleID As Long) As Boolean

' Description:  Returns true/or false after checking whether the role passed as parameter can only be replaced
'               this role can not be removed from the editset
'
' Amendment history:
'   d3      NDK     28July1999      Created
'   d19     ndk     18Nov1999       Removed managing editor
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String

    sSource = Me.Name & ".RoleCanOnlyBeReplaced()"
    
    Select Case t_lRoleID
        Case gk_lContributor1RoleID         'Contributor1
            RoleCanOnlyBeReplaced = True
        Case gk_lEditorRoleID               'Editor
            RoleCanOnlyBeReplaced = True
        Case Else
            RoleCanOnlyBeReplaced = False
    End Select
    
    Exit Function
    
ErrorHandler:
    ConvertError sSource
End Function



Public Function RoleNeverToBeAddedToEditSet( _
        ByVal t_lRoleID As Long) As Boolean

' Description:  Returns true/or false after checking whether the role passed as parameter
'               can be added to an editset
'
' Amendment history:
'   d3      NDK     28July1999       Created
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String

    sSource = Me.Name & ".RoleNeverToBeAddedToEditSet()"
    
    Select Case t_lRoleID
        Case gk_lAdministratorRoleID
            RoleNeverToBeAddedToEditSet = True
        Case gk_lManagingEditorRoleID
            RoleNeverToBeAddedToEditSet = True
        Case Else
            RoleNeverToBeAddedToEditSet = False
    End Select
    
    Exit Function
    
ErrorHandler:
    ConvertError sSource
End Function



Private Function RoleCanMoveEditSetFromCurrentStage( _
                    ByVal t_lRoleID As Long, _
                    ByVal t_lStageID As Long) As ESecurityCheckResult

' Description:  Returns true/or false after checking whether the role passed as parameter can move the editset
'               from the current stage
'
' Amendment history:
'   d3      NDK     28Jul1999       Created
'   d4      ndk     27Oct1999       Changed to return ESecurityCheckResult of data
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String

    sSource = Me.Name & ".RoleCanMoveEditSetFromCurrentStage()"

    'Check whether the role belongs to the stage
    If Not RoleBelongsToStage(t_lRoleID, t_lStageID) Then   'Role is not valid for the current stage
        RoleCanMoveEditSetFromCurrentStage = _
            ESecurityCheckResult.ESRRoleDoesNotBelongToStage
    Else
        'Check if it is the highest role
        If Not RoleLevelOfRole(t_lRoleID) = gk_lHighestRoleLevelForEditSetStage Then  'Not highest role
            RoleCanMoveEditSetFromCurrentStage = _
                ESecurityCheckResult.ESRRoleIsNotTheHighestRoleForThisStage
        Else
            RoleCanMoveEditSetFromCurrentStage = _
                ESecurityCheckResult.ESRSecurityCheckSuccessful 'Can move from current stage
        End If
    End If

   Exit Function
    
ErrorHandler:
    ConvertError sSource
End Function
                    


Private Function ValidForwardTransition() As Boolean

'Description:  Returns true when the transtion is a valid forward transition
'
'Properties required to be set
'StageToBeMovedFrom , StageToBeMovedTo
'
' Amendment history:
'   d3      NDK     28July1999          Created
'   d4      NDK     04Aug1999           Abandon stage not filtered out from valid forward transitions

    On Error GoTo ErrorHandler
    
    Dim sSource As String, sSQL As String

    sSource = Me.Name & ".ValidForwardTransition()"
    
    'Construct a sql to check whether the combination is a valid forward transition
    
    'Contribution1 to contribution1 is excluded (basically same stage to same stage is excluded)
    'any stage to abandon is excluded
    
    sSQL = "SELECT " & "*" & _
                " FROM " & _
                    gk_sDBTABLEEditSetStageTransition & _
                " WHERE " & _
                    gk_sDBCOLForward_Transition & " =  1 " & " AND " & _
                    gk_sDBCOLEditSetStage & " = " & m_lStageToBeMovedFromID & " AND " & _
                    gk_sDBCOLEditSetStage1 & " = " & m_lStageToBeMovedToID & " AND " & _
                    " NOT " & gk_sDBCOLEditSetStage1 & " = " & gk_sDBCOLEditSetStage

    'run the sql
    Set m_ofwRecordlist = m_oFWDBHelper.FindBySql(sSQL)
       
    If m_ofwRecordlist.RecordCount <> 0 Then
        ValidForwardTransition = True
    Else
        ValidForwardTransition = False
    End If

    Exit Function
    
ErrorHandler:
    ConvertError sSource
End Function



Private Function ValidReverseTransition() As Boolean

'Description:  Returns true when the transtion is a valid forward transition
'
'Properties required to be set
'StageToBeMovedFrom , StageToBeMovedTo
'
'Amendment history:
'   d3      NDK     28July1999       Created
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String, sSQL As String

    sSource = Me.Name & ".ValidForwardTransition()"
    
    'Construct a sql to check whether the combination is a valid reverse transition
    
    'any stage to abandon is excluded

    sSQL = "SELECT " & "*" & _
                " FROM " & _
                    gk_sDBTABLEEditSetStageTransition & _
                " WHERE " & _
                    gk_sDBCOLForward_Transition & " =  0 " & " AND " & _
                    gk_sDBCOLEditSetStage & " = " & m_lStageToBeMovedFromID & " AND " & _
                    gk_sDBCOLEditSetStage1 & " = " & m_lStageToBeMovedToID & " AND " & _
                    " NOT " & gk_sDBCOLEditSetStage1 & " = " & gk_lStageAbandonID

    'run the sql
    Set m_ofwRecordlist = m_oFWDBHelper.FindBySql(sSQL)
       
    If m_ofwRecordlist.RecordCount <> 0 Then
        ValidReverseTransition = True
    Else
        ValidReverseTransition = False
    End If

    Exit Function
    
ErrorHandler:
    ConvertError sSource
End Function



Private Function RoleLevelOfRole( _
                    ByVal t_lRoleID As Long) As Long
                    
' Description:  Returns the role level of the role passed as parameter
'
' Amendment history:
'   d3      NDK     28July1999       Created
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String, sSQL As String

    sSource = Me.Name & ".RoleLevelOfRole()"
   
    'Construct a sql to get the role level
    sSQL = "SELECT " & gk_sDBCOLRoleLevel & _
                " FROM " & gk_sDBTABLERole & _
                " WHERE " & _
                gk_sDBCOLRoleID & " = " & t_lRoleID
                
    'run the sql
    Set m_ofwRecordlist = m_oFWDBHelper.FindBySql(sSQL)
    
    'Get the role level
    RoleLevelOfRole = m_ofwRecordlist.GetFieldValue(gk_sDBCOLRoleLevel)
                    
    Exit Function
    
ErrorHandler:
    ConvertError sSource
End Function



Private Function UserCanMoveEditSetFromCurrentStage() As ESecurityCheckResult

' Description: Checks whether the user can move the edit set from the current stage
'              This method is used to enable/disable 'next stage, previous stage' buttons
'
'User should have a role on the editset
'Role should be valid for the current stage
'His role should be highest role at that stage
'
'Properties required to be set
'UserID, EditSetID
'
' Amendment history:
'   d3      NDK     28Jul1999       Created
'   d4      ndk     27Oct1999       Changed to return ESecurityCheckResult of data
'   d12     ndk     11Nov1999       All activities can now be performed if their role is
'                                   role is valid for that stage (and they have rights)
'   d26     dtm     10Feb2000       Editor no longer in Pre-Release stage but is still allowed to send back to Checking
'   d27     crm     14Feb2000      from Checking to PreRelease is possible only if the
'                                  EditSet does not have Candidate terms
'   d28     crm     15Feb2000       Changed implementation for d27
    On Error GoTo ErrorHandler

    Dim lUsersRoleOnEditSet As Long, lCurrentStageOFEditSet As Long, sSource As String

    sSource = Me.Name & ".UserCanMoveEditSetFromCurrentStage()"
    
    'Administrator
    m_oboUserRoles.UserID = m_lUserID
    If m_oboUserRoles.UserHasAdministratorsRole Then    'Administrator can move any editset from any stage to any stage
        UserCanMoveEditSetFromCurrentStage = ESecurityCheckResult.ESRSecurityCheckSuccessful
        Exit Function
    End If
    
    'Get the role of the user on the editset
    lUsersRoleOnEditSet = m_oBOEditSet.UsersRoleOnEditSet(m_lUserID, m_lEditsetID)
    
    If lUsersRoleOnEditSet <= 0 Then   'User does not have a role on the ediset
        UserCanMoveEditSetFromCurrentStage = _
            ESecurityCheckResult.ESRUserDoesNotHaveRoleOrHisRoleNotPermittedToPerformThisOperation
        Exit Function
    End If
    
    'Get the current stage of the editset
    lCurrentStageOFEditSet = m_oBOEditSet.CurrentStageOfEditSet(m_lEditsetID)

    ' d26   dtm     10Feb2000       Editor no longer in Pre-Release stage but is still allowed to send back to Checking
    If ((lCurrentStageOFEditSet = gk_lStagePreReleaseID) And (lUsersRoleOnEditSet = gk_lEditorRoleID)) Then
        UserCanMoveEditSetFromCurrentStage = ESecurityCheckResult.ESRSecurityCheckSuccessful
        Exit Function
    End If
    
    ' d27   crm     14Feb2000      from Checking to PreRelease is possible only if the
    '                              EditSet does not have Candidate terms
    If m_lStageToBeMovedFromID = gk_lStageCheckingID And m_lStageToBeMovedToID = gk_lStagePreReleaseID Then
        If Not EditSetHasCandidateTerms Then
            UserCanMoveEditSetFromCurrentStage = ESecurityCheckResult.ESRSecurityCheckSuccessful
            Exit Function
        Else
            UserCanMoveEditSetFromCurrentStage = ESecurityCheckResult.ESREditSetHasCandidateTerms
            Exit Function
        End If
    End If
    
    If Not RoleBelongsToStage(lUsersRoleOnEditSet, lCurrentStageOFEditSet) Then
        UserCanMoveEditSetFromCurrentStage = _
            ESecurityCheckResult.ESRUserDoesNotHaveRoleOrHisRoleNotPermittedToPerformThisOperation
        Exit Function
    End If
    
    'delegate further validation to other function
    UserCanMoveEditSetFromCurrentStage = RoleCanMoveEditSetFromCurrentStage(lUsersRoleOnEditSet, lCurrentStageOFEditSet)

    Exit Function
    
ErrorHandler:
    ConvertError sSource
End Function



Private Function UserCanMoveEditSetFromCurrentStageToDifferentStage() As ESecurityCheckResult

' Description: Checks whether the user can move the edit set from the current stage
'              This method is used when actually an editset is tried to be moved
'
'User should have a role on the editset
'Role should be valid for the current stage
'His role should be highest role at that stage
'And some additional rules for abandon and release and other rules
'
'Properties required to be set
'UserID, EditSetID, m_bIsForwardTransition,m_lStageToBeMovedFromID,m_lStageToBeMovedToID
'
' Amendment history:
'   d3      NDK     28Jul1999       Created
'   d4      ndk     27Oct1999       Changed to return ESecurityCheckResult of data
'   d12     ndk     11Nov1999       All activities can now be performed if their role is
'                                   role is valid for that stage (and they have rights)
'   d26     dtm     10Feb2000       Editor no longer in Pre-Release stage but is still allowed to send back to Checking
'   d27     crm     14Feb2000       Can not move ES to PreRelease if there are still
'                                   candidate terms
'   d28     crm     15Feb2000       Changed implementation for d27
'   d29     tmj     13Mar2000       Removed checking for Proof Reader before moving into checking stage
    On Error GoTo ErrorHandler
    
    Dim lUsersRoleOnEditSet As Long, lCurrentStageOFEditSet As Long, sSource As String

    sSource = Me.Name & ".UserCanMoveEditSetFromCurrentStageToDifferentStage()"

    If m_bIsForwardTransition Then    'Forward transition
        If Not ValidForwardTransition Then
            UserCanMoveEditSetFromCurrentStageToDifferentStage = _
                ESecurityCheckResult.ESRNotAValidForwardTransition
            Exit Function
        End If
    Else                            'Reverse transition
        If Not ValidReverseTransition Then
            UserCanMoveEditSetFromCurrentStageToDifferentStage = _
                ESecurityCheckResult.ESRNotAValidReverseTransition
            Exit Function
        End If
    End If
    
    
    'd29 Commented out, so check is no longer performed
'    'Can not move the editset from editstage to next stage (forward transition) until
'    'a proof reader is assigned
'    If m_bIsForwardTransition And m_lStageToBeMovedFromID = gk_lStageEditID Then
'        If Not m_oBOEditSet.RolePresentOnEditSet(gk_lProofReaderRoleID, m_lEditsetID) Then
'            'Role of Proof Reader is not present
'            UserCanMoveEditSetFromCurrentStageToDifferentStage = _
'                ESecurityCheckResult.ESRNoProofReaderPresentOnEditSet
'            Exit Function
'        End If
'    End If
    
    'Administrator
    m_oboUserRoles.UserID = m_lUserID
    If m_oboUserRoles.UserHasAdministratorsRole Then
        UserCanMoveEditSetFromCurrentStageToDifferentStage = _
            ESecurityCheckResult.ESRSecurityCheckSuccessful    'Administrator can move any editset from any stage to any stage
        'TODO
        UserCanMoveEditSetFromCurrentStageToDifferentStage = EditSetHasCandidateTerms
        Exit Function
    End If
    
    'Get the role of the user on the editset
    lUsersRoleOnEditSet = m_oBOEditSet.UsersRoleOnEditSet(m_lUserID, m_lEditsetID)
    
    If Not lUsersRoleOnEditSet > 0 Then   'User does not have a role on the ediset
        UserCanMoveEditSetFromCurrentStageToDifferentStage = _
            ESecurityCheckResult.ESRUserDoesNotHaveRoleOrHisRoleNotPermittedToPerformThisOperation
        Exit Function
    End If
    
    'If abandoning an editset - Only contributo1 and admin can abandon (admin taken care at top of this function)
    If m_lStageToBeMovedToID = gk_lStageAbandonID Then
        If Not lUsersRoleOnEditSet = gk_lContributor1RoleID Then
            UserCanMoveEditSetFromCurrentStageToDifferentStage = _
                ESecurityCheckResult.ESRUserDoesNotHaveRoleOrHisRoleNotPermittedToPerformThisOperation
            Exit Function
        End If
    End If
    
    'Get the users role and current stage on editset
    lCurrentStageOFEditSet = m_oBOEditSet.CurrentStageOfEditSet(m_lEditsetID)

    ' d26   dtm     10Feb2000       Editor no longer in Pre-Release stage but is still allowed to send back to Checking
    If ((lCurrentStageOFEditSet = gk_lStagePreReleaseID) And (m_lStageToBeMovedToID = gk_lStageCheckingID) And (lUsersRoleOnEditSet = gk_lEditorRoleID)) Then
        UserCanMoveEditSetFromCurrentStageToDifferentStage = ESecurityCheckResult.ESRSecurityCheckSuccessful
        Exit Function
    End If
    ' d27   crm     14Feb2000      from Checking to PreRelease is possible only if the
    '                              EditSet does not have Candidate terms
        
    
    If Not RoleBelongsToStage(lUsersRoleOnEditSet, lCurrentStageOFEditSet) Then
        UserCanMoveEditSetFromCurrentStageToDifferentStage = _
            ESecurityCheckResult.ESRUserDoesNotHaveRoleOrHisRoleNotPermittedToPerformThisOperation
        Exit Function
    End If
    
    'delegate further validation to another function
    UserCanMoveEditSetFromCurrentStageToDifferentStage = _
            RoleCanMoveEditSetFromCurrentStage(lUsersRoleOnEditSet, m_lStageToBeMovedFromID)
                
    'TODO
        UserCanMoveEditSetFromCurrentStageToDifferentStage = EditSetHasCandidateTerms
                
    Exit Function
    
ErrorHandler:
    ConvertError sSource
End Function



Private Function UserCanRemoveRoleFromEditSet() As ESecurityCheckResult

'Description: Checks whether the user can remove the role from the editset
'This is called once the action is attempted.
'Enabling/disabling the 'edit roles' URL is done by UserCanEditRolesOnEditSet
'
'User should belong to the editset
'
'The user who is trying to remove should have the role whose levels is greater than the role level of the role
'he is trying to remove
'
'The user who is trying to remove should have the role which belongs to all the stages the role he is trying to
'remove belongs
'
'Properties required to be set
'UserID,EditSetID, RoleToBeRemovedID
'
' Amendment history:
'   d3      NDK     28Jul1999       Created
'   d4      ndk     27Oct1999       Changed to return ESecurityCheckResult of data
'   d12     ndk     11Nov1999       All activities can now be performed if their role is
'                                   role is valid for that stage (and they have rights)
'   d30     tmj     20Mar2000       Commented out all logic that prevents the removal of
'                                   AFEditors and Proof Readers
'
    On Error GoTo ErrorHandler

    Dim sSource As String, lCurrentStageOFEditSet As Long, lUsersRoleOnEditSet As Long
    
    sSource = Me.Name & ".UserCanRemoveRoleFromEditSet()"
    
    'Get the current stage of the editset
    lCurrentStageOFEditSet = m_oBOEditSet.CurrentStageOfEditSet(m_lEditsetID)
    
    'Check whether the role is 'replace only  eg: cont1, editor
    If RoleCanOnlyBeReplaced(m_lRoleToBeRemovedID) Then       'These roles can not be removed
        UserCanRemoveRoleFromEditSet = ESecurityCheckResult.ESRTryingToRemoveReplaceOnlyRole
        Exit Function
    End If
    
    'Administrator ( The Admin role takes precedence over his other roles on the editset)
    m_oboUserRoles.UserID = m_lUserID
    If m_oboUserRoles.UserHasAdministratorsRole Then
        UserCanRemoveRoleFromEditSet = ESecurityCheckResult.ESRSecurityCheckSuccessful       'Administrator can add/remove any role
        Exit Function
    End If
    
    'Get the users role and current stage on editset
    lUsersRoleOnEditSet = m_oBOEditSet.UsersRoleOnEditSet(m_lUserID, m_lEditsetID)
    lCurrentStageOFEditSet = m_oBOEditSet.CurrentStageOfEditSet(m_lEditsetID)

    If Not lUsersRoleOnEditSet > 0 Then
        'User does not have a role on the ediset
        'Managing editor is not considered here as he can not remove any role
        UserCanRemoveRoleFromEditSet = ESecurityCheckResult.ESRUserDoesNotHaveRoleOrHisRoleNotPermittedToPerformThisOperation
        Exit Function
    End If
    
    If Not RoleBelongsToStage(lUsersRoleOnEditSet, lCurrentStageOFEditSet) Then
        UserCanRemoveRoleFromEditSet = ESecurityCheckResult.ESRUserDoesNotHaveRoleOrHisRoleNotPermittedToPerformThisOperation
        Exit Function
    End If
    
'd30
'    'Proof reader can only replace himself
'    If m_lRoleToBeRemovedID = gk_lProofReaderRoleID Then
'        If lUsersRoleOnEditSet = gk_lProofReaderRoleID Then
'            UserCanRemoveRoleFromEditSet = ESecurityCheckResult.ESRSecurityCheckSuccessful
'        Else
'            UserCanRemoveRoleFromEditSet = ESecurityCheckResult.ESRProofReaderTryingToChangeOtherRole
'        End If
'        Exit Function
'    End If
'
'    'AF Editor can only replace himself
'    If m_lRoleToBeRemovedID = gk_lAFEditorRoleID Then
'        If lUsersRoleOnEditSet = gk_lAFEditorRoleID Then
'            UserCanRemoveRoleFromEditSet = ESecurityCheckResult.ESRSecurityCheckSuccessful
'        Else
'            UserCanRemoveRoleFromEditSet = ESecurityCheckResult.ESRAFEditorTryingToChangeOtherRole
'        End If
'        Exit Function
'    End If
        
    'delegate further validation to other function - Check whether role1 can remove role2
    UserCanRemoveRoleFromEditSet = _
        Role1CanRemoveRole2(lUsersRoleOnEditSet, m_lRoleToBeRemovedID)
    
    Exit Function
ErrorHandler:
    ConvertError sSource
End Function
            


Private Function UserCanAddRoleToEditSet() As ESecurityCheckResult

' Description: Checks whether the user can add a specific role to the editset.
' This is called once the action is attempted.
' Enabling/disabling the 'edit roles' URL is done by UserCanEditRolesOnEditSet
'
' User should belong to the editset
'
' The user who is trying to add should have the role whose levels is greater than the role level of the role
' he is trying to add
'
' The user who is trying to add should have the role which belongs to all the stages the role he is trying to
' add belongs
'
' Properties required to be set
' UserID, lEditSetID, RoleToBeAddedID, OtherUserID


' Amendment history:
' d3      NDK     28Jul1999       Created
' d4      ndk     27Oct1999       Changed to return ESecurityCheckResult of data
' d10     ndk     01Nov1999       Change in logic while checking when af editor/proof reader are tried to be replaced
' d11     ndk     02Nov1999       Validation to check that a user can not have two roles
' d15     ndk     15Nov1999       Change in the way me rights are handled while adding roles
' d16     ndk     15Nov1999       Minor changes to the previous change
' d21     ndk     22Nov1999       Mng editors role interpretation changed slightly, af editor and  proof reader now
'                                 handled as other roles
' d23     ndk     23Nov1999       Error corrected

    On Error GoTo ErrorHandler
    
    Dim sSource As String, lCurrentStageOFEditSet As Long, lUsersRoleOnEditSet As Long

    sSource = Me.Name & ".UserCanAddRoleToEditSet()"

    If m_oBOEditSet.UserPresentOnEditSet(m_lOtherUserID, m_lEditsetID) Then
        UserCanAddRoleToEditSet = ESecurityCheckResult.ESRUserBeingAddedAlreadyHasARoleOnEditSet
        Exit Function
    End If

    'Check whether it is a role that should never be added
    If RoleNeverToBeAddedToEditSet(m_lRoleToBeAddedID) Then
        UserCanAddRoleToEditSet = ESecurityCheckResult.ESRTryingToAddRoleNeverToBeAdded
        Exit Function
    End If
    
    'Administrator ( The Admin role takes precedence over his other roles on the editset)
    m_oboUserRoles.UserID = m_lUserID
    If m_oboUserRoles.UserHasAdministratorsRole Then
        UserCanAddRoleToEditSet = ESecurityCheckResult.ESRSecurityCheckSuccessful   'Administrator can add/remove any role
        Exit Function
    End If
    
    'get the current stage of editset
    lCurrentStageOFEditSet = m_oBOEditSet.CurrentStageOfEditSet(m_lEditsetID)
    
    'if the user has a me role
    If m_oboUserRoles.UserHasRole(m_lUserID, gk_lManagingEditorRoleID) Then
        'and he is trying to add an editor and the current stage is edit stage
        If m_lRoleToBeAddedID = gk_lEditorRoleID And _
            lCurrentStageOFEditSet = gk_lStageEditID Then
                'Return successful and exit the function
            UserCanAddRoleToEditSet = ESecurityCheckResult.ESRSecurityCheckSuccessful
            Exit Function
        End If
        'If the control come shere, he is a me and he may be adding other role or
        'he may not be in the appropriate stage to add the editor
        'now treat him as any other user
        'this happens as the control goes out of this if block
    End If
        
    'Get the users role and current stage on editset
    lUsersRoleOnEditSet = m_oBOEditSet.UsersRoleOnEditSet(m_lUserID, m_lEditsetID)

    If Not RoleBelongsToStage(lUsersRoleOnEditSet, lCurrentStageOFEditSet) Then
        UserCanAddRoleToEditSet = ESecurityCheckResult.ESRUserDoesNotHaveRoleOrHisRoleNotPermittedToPerformThisOperation
        Exit Function
    End If
    
    'For all other users
    If Role1CanAddRole2(lUsersRoleOnEditSet, m_lRoleToBeAddedID) Then   'Check whether Role1 can add Role2
        UserCanAddRoleToEditSet = ESecurityCheckResult.ESRSecurityCheckSuccessful
    Else
        UserCanAddRoleToEditSet = _
            ESecurityCheckResult.ESRUserDoesNotHaveRoleOrHisRoleNotPermittedToPerformThisOperation
    End If
    
    Exit Function
    
ErrorHandler:
    ConvertError sSource
End Function

            

Private Function UserCanChangeCompletionStatusForOtherUser() As ESecurityCheckResult

' Description:
'
' Checks whether the one user can change the status (mark as complete/in complete) the edit set
' where the user OtherUser is supposed to change the status (assigned user)
'
'Properties required to be set
'UserID, EditSetID, OtherUserID
'
' Amendment history:
'   d3  NDK     28Jul1999   Created
'   d5  NDK     13Aug1999   Admin and Mng Editor role considered
'   d9  NDK     27Oct1999   Modified to return ESecurityCheckResult data type
'                           New rule - user can update his own checkbox (except admin)
'  d20  NDK     22Nov1999   Administrator allowed to check other users editset completion status (this was done earlier, was reintroduced)
'
    On Error GoTo ErrorHandler

    Dim sSource As String, sSQL As String
    Dim lRoleOfOtherUser As Long, lCurrentStageOFEditSet As Long
   
    sSource = Me.Name & ".UserCanChangeCompletionStatusForOtherUser()"

    'Admin can update anybody's status
    If m_oboUserRoles.UserHasRole(m_lUserID, gk_lAdministratorRoleID) Then
        UserCanChangeCompletionStatusForOtherUser = ESecurityCheckResult.ESRSecurityCheckSuccessful
        Exit Function
    End If

    'get the current stage of editset
    lCurrentStageOFEditSet = m_oBOEditSet.CurrentStageOfEditSet(m_lEditsetID)

    'role of user whose checkbox is checked on editset
    lRoleOfOtherUser = m_oBOEditSet.UsersRoleOnEditSet(m_lOtherUserID, m_lEditsetID)
    
    If lRoleOfOtherUser <= 0 Then        'Does not have a role on editset
        UserCanChangeCompletionStatusForOtherUser = _
            ESecurityCheckResult.ESROtherUserDoesNotHaveRoleOnEditSet
        Exit Function
    End If
    
    'Only those users checkboxes can be checked whose role is valid for current stage
    If Not RoleBelongsToStage(lRoleOfOtherUser, lCurrentStageOFEditSet) Then
        UserCanChangeCompletionStatusForOtherUser = ESecurityCheckResult.ESROpertionNotValidAtThisStage
        Exit Function
    End If

    If m_lUserID = m_lOtherUserID Then       'user trying to change his own checkbox
        UserCanChangeCompletionStatusForOtherUser = _
            ESecurityCheckResult.ESRSecurityCheckSuccessful
    Else
        UserCanChangeCompletionStatusForOtherUser = _
            ESecurityCheckResult.ESRUserTryingToChangeOtherUsersCheckBox
    End If
    
    Exit Function

ErrorHandler:
    ConvertError sSource
End Function



Private Function UserCanEditRolesOnEditSet() As ESecurityCheckResult

' Description:
'
' Checks whether the user is able to edit roles generally on an editset.
' this is called to enable/disable the 'edit roles' URL.  Some users can only edit certain roles:
' this is contained within UserCanAddRoleToEditSet()
'
'Properties required to be set
'UserID
'EditSetID
'
' Amendment history:
' d9    ndk     27Oct1999       Created
' d12   ndk     11Nov1999       All activities can now be performed if their role is
'                               role is valid for that stage (and they have rights)
' d31   tmj     25Mar2000       Changed so managing editor can only  edit roles
'                               when not in Contribution
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String, lUsersRoleOnEditSet As Long, lCurrentStageOFEditSet As Long

    sSource = Me.Name & ".UserCanEditRolesOnEditSet()"
    
    If m_oBOEditSet.CurrentStageOfEditSet(m_lEditsetID) = gk_lStagePreReleaseID Then
        UserCanEditRolesOnEditSet = ESecurityCheckResult.ESROpertionNotValidAtThisStage
        Exit Function
    End If
    
    m_oboUserRoles.UserID = m_lUserID
    If m_oboUserRoles.UserHasAdministratorsRole Then        'Admin can edit roles
        UserCanEditRolesOnEditSet = ESecurityCheckResult.ESRSecurityCheckSuccessful
        Exit Function
    End If
    
    m_oboUserRoles.UserID = m_lUserID
    
    lCurrentStageOFEditSet = m_oBOEditSet.CurrentStageOfEditSet(m_lEditsetID)
    
    If m_oboUserRoles.UserHasManagingEditorRole _
                 And lCurrentStageOFEditSet <> 1 Then        'Managing editor
        UserCanEditRolesOnEditSet = ESecurityCheckResult.ESRSecurityCheckSuccessful
        Exit Function
    End If
    
    'get the users role and stage
    lUsersRoleOnEditSet = m_oBOEditSet.UsersRoleOnEditSet(m_lUserID, m_lEditsetID)
    lCurrentStageOFEditSet = m_oBOEditSet.CurrentStageOfEditSet(m_lEditsetID)
    
    If RoleBelongsToStage(lUsersRoleOnEditSet, lCurrentStageOFEditSet) Then
        UserCanEditRolesOnEditSet = ESecurityCheckResult.ESRSecurityCheckSuccessful
    Else
        UserCanEditRolesOnEditSet = ESecurityCheckResult.ESRUserDoesNotHaveRoleOrHisRoleNotPermittedToPerformThisOperation
    End If
    
    Exit Function
ErrorHandler:
    ConvertError sSource
End Function



Private Function UserCanAddCatalogueEntryToEditSet() As ESecurityCheckResult
' Description:
'
' Checks whether the user can add a catalogue entry to the editset
'
'Properties required to be set
'UserID
'EditSetID
'
' Amendment history:
' d9    ndk     27Oct1999       Created
' d12   ndk     11Nov1999       All activities can now be performed if their role is
'                               role is valid for that stage (and they have rights)
'
    On Error GoTo ErrorHandler
    
    Dim lUsersRoleOnEditSet As Long, lCurrentStageOFEditSet As Long, sSource As String

    sSource = Me.Name & ".UserCanAddCatalogueEntryToEditSet()"
    
    'Get the users role and current stage on editset
    lUsersRoleOnEditSet = m_oBOEditSet.UsersRoleOnEditSet(m_lUserID, m_lEditsetID)
    lCurrentStageOFEditSet = m_oBOEditSet.CurrentStageOfEditSet(m_lEditsetID)

    If Not RoleBelongsToStage(lUsersRoleOnEditSet, lCurrentStageOFEditSet) Then
        UserCanAddCatalogueEntryToEditSet = ESecurityCheckResult.ESRUserDoesNotHaveRoleOrHisRoleNotPermittedToPerformThisOperation
        Exit Function
    End If
    
    'cont1, cont2, cont3 and editor can add catalogue entry to the edit set
    If lUsersRoleOnEditSet = gk_lContributor1RoleID Or _
        lUsersRoleOnEditSet = gk_lContributor2RoleID Or _
        lUsersRoleOnEditSet = gk_lContributor3RoleID Or _
        lUsersRoleOnEditSet = gk_lEditorRoleID Then
        
        UserCanAddCatalogueEntryToEditSet = ESecurityCheckResult.ESRSecurityCheckSuccessful
    Else
        UserCanAddCatalogueEntryToEditSet = ESecurityCheckResult.ESRUserDoesNotHaveRoleOrHisRoleNotPermittedToPerformThisOperation
    End If
    
    Exit Function
ErrorHandler:
    ConvertError sSource
End Function



Private Function UserCanCreateCatalogueEntryOnEditSet()
' Description:
'
' Checks whether the user can create a catalogue entry on the editset
'
'Properties required to be set
'UserID
'EditSetID
'
' Amendment history:
' d9    ndk     27Oct1999       Created
' d12   ndk     11Nov1999       All activities can now be performed if their role is
'                               role is valid for that stage (and they have rights)
'
    On Error GoTo ErrorHandler
    
    Dim lUsersRoleOnEditSet As Long, lCurrentStageOFEditSet As Long, sSource As String

    sSource = Me.Name & ".UserCanCreateCatalogueEntryOnEditSet()"
    
    'Get the users role and current stage on editset
    lUsersRoleOnEditSet = m_oBOEditSet.UsersRoleOnEditSet(m_lUserID, m_lEditsetID)
    lCurrentStageOFEditSet = m_oBOEditSet.CurrentStageOfEditSet(m_lEditsetID)

    If Not RoleBelongsToStage(lUsersRoleOnEditSet, lCurrentStageOFEditSet) Then
        UserCanCreateCatalogueEntryOnEditSet = ESecurityCheckResult.ESRUserDoesNotHaveRoleOrHisRoleNotPermittedToPerformThisOperation
        Exit Function
    End If
    
    'cont1, cont2, cont3 and editor can create catalogue entry on the edit set
    If lUsersRoleOnEditSet = gk_lContributor1RoleID Or _
        lUsersRoleOnEditSet = gk_lContributor2RoleID Or _
        lUsersRoleOnEditSet = gk_lContributor3RoleID Or _
        lUsersRoleOnEditSet = gk_lEditorRoleID Then
        
        UserCanCreateCatalogueEntryOnEditSet = ESecurityCheckResult.ESRSecurityCheckSuccessful
    Else
        UserCanCreateCatalogueEntryOnEditSet = ESecurityCheckResult.ESRUserDoesNotHaveRoleOrHisRoleNotPermittedToPerformThisOperation
    End If
    
    Exit Function
ErrorHandler:
    ConvertError sSource
End Function



Private Function UserCanEditCatalogueEntryOnEditSet()
' Description:
'
' Checks whether the user can edit a catalogue entry on the editset
'
'Properties required to be set
'UserID
'EditSetID
'
' Amendment history:
' d9    ndk     27Oct1999       Created
' d12   ndk     11Nov1999       All activities can now be performed if their role is
'                               role is valid for that stage (and they have rights)
'
    On Error GoTo ErrorHandler
    
    Dim lUsersRoleOnEditSet As Long, lCurrentStageOFEditSet As Long, sSource As String

    sSource = Me.Name & ".UserCanEditCatalogueEntryOnEditSet()"
    
    'Get the users role and current stage on editset
    lUsersRoleOnEditSet = m_oBOEditSet.UsersRoleOnEditSet(m_lUserID, m_lEditsetID)
    lCurrentStageOFEditSet = m_oBOEditSet.CurrentStageOfEditSet(m_lEditsetID)

    If Not RoleBelongsToStage(lUsersRoleOnEditSet, lCurrentStageOFEditSet) Then
        UserCanEditCatalogueEntryOnEditSet = ESecurityCheckResult.ESRUserDoesNotHaveRoleOrHisRoleNotPermittedToPerformThisOperation
        Exit Function
    End If
    
    'cont1, cont2, cont3, editor, proof reader and af editor can edit catalogue entry
    If lUsersRoleOnEditSet = gk_lContributor1RoleID Or _
        lUsersRoleOnEditSet = gk_lContributor2RoleID Or _
        lUsersRoleOnEditSet = gk_lContributor3RoleID Or _
        lUsersRoleOnEditSet = gk_lEditorRoleID Or _
        lUsersRoleOnEditSet = gk_lProofReaderRoleID Or _
        lUsersRoleOnEditSet = gk_lAFEditorRoleID Then
        
        UserCanEditCatalogueEntryOnEditSet = ESecurityCheckResult.ESRSecurityCheckSuccessful
    Else
        UserCanEditCatalogueEntryOnEditSet = ESecurityCheckResult.ESRUserDoesNotHaveRoleOrHisRoleNotPermittedToPerformThisOperation
    End If
    
    Exit Function
ErrorHandler:
    ConvertError sSource
End Function



Private Function UserCanSetFilterOnEditSet() As ESecurityCheckResult
' Description:
'
' Checks whether the user can set a filter on an editset
'
'Properties required to be set
'UserID
'EditSetID
'
' Amendment history:
' d9    ndk     27Oct1999       Created
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String

    sSource = Me.Name & ".UserCanSetFilterOnEditSet()"
    
    'Anybody can set a filter on editset
    UserCanSetFilterOnEditSet = ESecurityCheckResult.ESRSecurityCheckSuccessful
    
    Exit Function
ErrorHandler:
    ConvertError sSource

End Function



Private Function UserCanAbandonEditSet() As ESecurityCheckResult
' Description:
'
' Checks whether the user can abandon an editset
' Only Administrator and Contributor1 can abandon an editset
'
'Properties required to be set
'UserID
'EditSetID
'
' Amendment history:
' d9    ndk     27Oct1999       Created
' d12   ndk     11Nov1999       All activities can now be performed if their role is
'                               role is valid for that stage (and they have rights)
'
    On Error GoTo ErrorHandler
    
    Dim lUsersRoleOnEditSet As Long, lCurrentStageOFEditSet As Long, sSource As String

    sSource = Me.Name & ".UserCanAbandonEditSet()"
    
    'Admin can abandon an editset
    m_oboUserRoles.UserID = m_lUserID
    If m_oboUserRoles.UserHasAdministratorsRole Then
        UserCanAbandonEditSet = ESecurityCheckResult.ESRSecurityCheckSuccessful
        Exit Function
    End If
    
    'Get the users role and current stage on editset
    lUsersRoleOnEditSet = m_oBOEditSet.UsersRoleOnEditSet(m_lUserID, m_lEditsetID)
    lCurrentStageOFEditSet = m_oBOEditSet.CurrentStageOfEditSet(m_lEditsetID)

    If Not RoleBelongsToStage(lUsersRoleOnEditSet, lCurrentStageOFEditSet) Then
        UserCanAbandonEditSet = ESecurityCheckResult.ESRUserDoesNotHaveRoleOrHisRoleNotPermittedToPerformThisOperation
        Exit Function
    End If
    
    'Contributor1 can abandon an editset
    If m_oBOEditSet.UsersRoleOnEditSet(m_lUserID, m_lEditsetID) = gk_lContributor1RoleID Then
        UserCanAbandonEditSet = ESecurityCheckResult.ESRSecurityCheckSuccessful
    Else
        UserCanAbandonEditSet = ESecurityCheckResult.ESRUserDoesNotHaveRoleOrHisRoleNotPermittedToPerformThisOperation
    End If
    
    Exit Function
ErrorHandler:
    ConvertError sSource
End Function



Private Function UserCanMaintainAuthorityFiles() As ESecurityCheckResult
'
'Checks whether the user can maintain authority files
'
' d8    aga     11Oct1999       Created
' d9    ndk     27Oct1999       Modified to return ESecurityCheckResult data type
    
    On Error GoTo ErrorHandler
    
    Dim sSource As String

    sSource = Me.Name & ".UserCanMaintainAuthorityFiles()"
    
    m_oboUserRoles.UserID = m_lUserID
    
    If m_oboUserRoles.UserHasAFEditorRole Then
        UserCanMaintainAuthorityFiles = ESecurityCheckResult.ESRSecurityCheckSuccessful
    Else
        UserCanMaintainAuthorityFiles = ESecurityCheckResult.ESRUserDoesNotHaveRoleOrHisRoleNotPermittedToPerformThisOperation
    End If

    Exit Function
ErrorHandler:
    ConvertError sSource
End Function



Private Function UserCanMaintainMaintainLookUps() As ESecurityCheckResult
'
' Checks whether the user can maintain lookups
'
'Properties required to be set
'UserID
'EditSetID
'
' Amendment history:
' d9    ndk     27Oct1999       Modified to return ESecurityCheckResult data type (Created date not known)
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String

    sSource = Me.Name & ".UserCanMaintainMaintainLookUps()"
    
    'Get the users role
    m_oboUserRoles.UserID = m_lUserID
    
    If m_oboUserRoles.UserHasAdministratorsRole Then
        UserCanMaintainMaintainLookUps = ESecurityCheckResult.ESRSecurityCheckSuccessful
    Else
        UserCanMaintainMaintainLookUps = ESecurityCheckResult.ESRUserDoesNotHaveRoleOrHisRoleNotPermittedToPerformThisOperation
    End If

    Exit Function
ErrorHandler:
    ConvertError sSource
End Function



Private Function UserCanMaintainUsers() As ESecurityCheckResult
' Description:
'
' Checks whether the user can maintain users
'
'Properties required to be set
'UserID
'EditSetID
'
' Amendment history:
' d9    ndk     27Oct1999       Modified to return ESecurityCheckResult data type (Created date not known)
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String

    sSource = Me.Name & ".UserCanMaintainUsers()"
    
    m_oboUserRoles.UserID = m_lUserID
    
    If m_oboUserRoles.UserHasAdministratorsRole Then
        UserCanMaintainUsers = ESecurityCheckResult.ESRSecurityCheckSuccessful
    Else
        UserCanMaintainUsers = ESecurityCheckResult.ESRUserDoesNotHaveRoleOrHisRoleNotPermittedToPerformThisOperation
    End If

    Exit Function
ErrorHandler:
    ConvertError sSource
End Function



Private Function UserCanBulkupdate() As ESecurityCheckResult
' Description:
'
' Checks whether the user can perform bulk update
'
'Properties required to be set
'UserID
'EditSetID
'
' Amendment history:
' d9    ndk     27Oct1999       Modified to return ESecurityCheckResult data type (Created date not known)
' d12   ndk     11Nov1999       All activities can now be performed if their role is
'                               role is valid for that stage (and they have rights)
'
    On Error GoTo ErrorHandler
    
    Dim lUsersRoleOnEditSet As Long, lCurrentStageOFEditSet As Long, sSource As String

    sSource = Me.Name & ".UserCanBulkupdate()"
    
    'Get the users role and current stage on editset
    lUsersRoleOnEditSet = m_oBOEditSet.UsersRoleOnEditSet(m_lUserID, m_lEditsetID)
    lCurrentStageOFEditSet = m_oBOEditSet.CurrentStageOfEditSet(m_lEditsetID)

    If Not RoleBelongsToStage(lUsersRoleOnEditSet, lCurrentStageOFEditSet) Then
        UserCanBulkupdate = ESecurityCheckResult.ESRUserDoesNotHaveRoleOrHisRoleNotPermittedToPerformThisOperation
        Exit Function
    End If
    
    If lUsersRoleOnEditSet = gk_lContributor1RoleID Or _
        lUsersRoleOnEditSet = gk_lContributor2RoleID Or _
        lUsersRoleOnEditSet = gk_lContributor3RoleID Or _
        lUsersRoleOnEditSet = gk_lEditorRoleID Then
        UserCanBulkupdate = ESecurityCheckResult.ESRSecurityCheckSuccessful
    Else
        UserCanBulkupdate = ESecurityCheckResult.ESRUserDoesNotHaveRoleOrHisRoleNotPermittedToPerformThisOperation
    End If
    
    Exit Function
ErrorHandler:
    ConvertError sSource
End Function



Public Function UserCanBeContributor1() As Boolean

' Description:
'
' Checks whether the user can don the role of contributor1
'
'Properties required to be set
'UserID
'
' Amendment history:
'   d3      NDK     28July1999       Created
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String

    sSource = Me.Name & ".UserCanBeContributor1()"
    
    'Get the list of users who can be contributor1
    m_oboUserRoles.RoleID = gk_lContributor1RoleID
    Set m_ofwRecordlist = m_oboUserRoles.ValidUsersForTheRole
    
    'See whether the current user is a part of it
    If m_oboUserRoles.ValidUsersForTheRole.Find(gk_sDBCOLUserID, m_lUserID) Then
        UserCanBeContributor1 = True
    Else
        UserCanBeContributor1 = False
    End If
        
    Exit Function
ErrorHandler:
    ConvertError sSource
End Function



Private Function UserCanMaintainLeafLets() As ESecurityCheckResult
' Description:
'
' Checks whether the user can maintain leaflets
'
'Properties required to be set
'UserID
'EditSetID
'
' Amendment history:
' d12   ndk     11Nov1999       Created
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String

    sSource = Me.Name & ".UserCanMaintainLeafLets()"
    
    If m_oboUserRoles.UserHasRole(m_lUserID, gk_lLeafletEditorRoleID) Then
        UserCanMaintainLeafLets = ESecurityCheckResult.ESRSecurityCheckSuccessful
    Else
        UserCanMaintainLeafLets = ESecurityCheckResult.ESRUserDoesNotHaveRoleOrHisRoleNotPermittedToPerformThisOperation
    End If
    
    Exit Function
ErrorHandler:
    ConvertError sSource
End Function



Private Function UserCanMaintainPopularSearches() As ESecurityCheckResult
' Description:
'
' Checks whether the user can maintain popular searches
'
'Properties required to be set
'UserID
'EditSetID
'
' Amendment history:
' d12   ndk     11Nov1999       Created
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String

    sSource = Me.Name & ".UserCanMaintainPopularSearches()"

    If m_oboUserRoles.UserHasRole(m_lUserID, gk_lPopularSearchEditorRoleID) Then
        UserCanMaintainPopularSearches = ESecurityCheckResult.ESRSecurityCheckSuccessful
    Else
        UserCanMaintainPopularSearches = ESecurityCheckResult.ESRUserDoesNotHaveRoleOrHisRoleNotPermittedToPerformThisOperation
    End If

    Exit Function
ErrorHandler:
    ConvertError sSource
End Function



Private Function UserCanMaintainPrintedSFA() As ESecurityCheckResult
' Description:
'
' Checks whether the user can printed sfas
'
'Properties required to be set
'UserID
'EditSetID
'
' Amendment history:
' d12   ndk     11Nov1999       Created
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String

    sSource = Me.Name & ".UserCanMaintainPrintedSFA()"

    If m_oboUserRoles.UserHasRole(m_lUserID, gk_lPSFAEditorRoleID) Then
        UserCanMaintainPrintedSFA = ESecurityCheckResult.ESRSecurityCheckSuccessful
    Else
        UserCanMaintainPrintedSFA = ESecurityCheckResult.ESRUserDoesNotHaveRoleOrHisRoleNotPermittedToPerformThisOperation
    End If

    Exit Function
ErrorHandler:
    ConvertError sSource
End Function



Private Function UserCanMaintainElectronicSFA() As ESecurityCheckResult
' Description:
'
' Checks whether the user can maintain electronic sfas
'
'Properties required to be set
'UserID
'EditSetID
'
' Amendment history:
' d12   ndk     11Nov1999       Created
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String

    sSource = Me.Name & ".UserCanMaintainElectronicSFA()"
    
    If m_oboUserRoles.UserHasRole(m_lUserID, gk_lESFAEditorRoleID) Then
        UserCanMaintainElectronicSFA = ESecurityCheckResult.ESRSecurityCheckSuccessful
    Else
        UserCanMaintainElectronicSFA = ESecurityCheckResult.ESRUserDoesNotHaveRoleOrHisRoleNotPermittedToPerformThisOperation
    End If

    Exit Function
ErrorHandler:
    ConvertError sSource
End Function



Private Function UserCanImportFromEAD() As ESecurityCheckResult
'Used to check if the user can import from EAD files into the ILDB and to prepare the DRUID data
'
' d13   ksk     12Nov1999       Added the security for importing
' d25   David Manning  02-Feb-2000  Modified to use UserHasOneOfMultipleRoles for accession/contributor roles

On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim bHasRole As Boolean
    
    sSource = Me.Name & ".UserCanImportFromEAD()"
    
    bHasRole = False
    
    bHasRole = m_oboUserRoles.UserHasOneOfMultipleRoles(m_lUserID, _
            CStr(gk_lAccessionRoleID) & "," & _
            CStr(gk_lEditorRoleID) & "," & _
            CStr(gk_lContributor1RoleID) & "," & _
            CStr(gk_lContributor2RoleID) & "," & _
            CStr(gk_lContributor3RoleID))
       
    If bHasRole Then
        UserCanImportFromEAD = ESecurityCheckResult.ESRSecurityCheckSuccessful
    Else
        UserCanImportFromEAD = ESecurityCheckResult.ESRUserDoesNotHaveRoleOrHisRoleNotPermittedToPerformThisOperation
    End If
    
    Exit Function
ErrorHandler:
    ConvertError sSource
End Function

Private Function UserCanUpdateFromEAD() As ESecurityCheckResult
' Used to check whether the user has the role required to update from EAD
'
' d32   daj     14Jul2000       Created

On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim bHasRole As Boolean
    
    sSource = Me.Name & ".UserCanUpdateFromEAD()"
    
    bHasRole = False
    
    bHasRole = m_oboUserRoles.UserHasOneOfMultipleRoles(m_lUserID, _
        CStr(gk_lAdministratorRoleID) & ", " & _
        CStr(gk_lContributor1RoleID))
       
    If bHasRole Then
        UserCanUpdateFromEAD = ESecurityCheckResult.ESRSecurityCheckSuccessful
    Else
        UserCanUpdateFromEAD = ESecurityCheckResult.ESRUserDoesNotHaveRoleOrHisRoleNotPermittedToPerformThisOperation
    End If
    
    Exit Function
ErrorHandler:
    ConvertError sSource
End Function

Private Function UserCanExportToEAD() As ESecurityCheckResult
' Used to check whether the user has the role required to export EAD
'
' d32   daj     14Jul2000       Created

On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim bHasRole As Boolean
    
    sSource = Me.Name & ".UserCanExportToEAD()"
    
    bHasRole = False
    
    ' User must have Accession role
    bHasRole = m_oboUserRoles.UserHasRole(m_lUserID, gk_lAccessionRoleID)
       
    If bHasRole Then
        UserCanExportToEAD = ESecurityCheckResult.ESRSecurityCheckSuccessful
    Else
        UserCanExportToEAD = ESecurityCheckResult.ESRUserDoesNotHaveRoleOrHisRoleNotPermittedToPerformThisOperation
    End If
    
    Exit Function
ErrorHandler:
    ConvertError sSource
End Function



Private Function UserCanLoadIntoEditset() As ESecurityCheckResult
' Used to check if the user can used the imported record to load into an edit set.
' (aka Create Edit set from loaded data)
'
' d24   ksk     13Dec1999       Addec codeo to seperate importing and loading in to edit set for accessioning
' d25   David Manning  02-Feb-2000  Modified to use UserHasOneOfMultipleRoles for speed

On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim bHasRole As Boolean
    
    sSource = Me.Name & ".UserCanLoadIntoEditset()"
    
    bHasRole = False
    
    bHasRole = m_oboUserRoles.UserHasOneOfMultipleRoles(m_lUserID, _
        CStr(gk_lAdministratorRoleID) & ", " & _
        CStr(gk_lContributor1RoleID))
       
    If bHasRole Then
        UserCanLoadIntoEditset = ESecurityCheckResult.ESRSecurityCheckSuccessful
    Else
        UserCanLoadIntoEditset = ESecurityCheckResult.ESRUserDoesNotHaveRoleOrHisRoleNotPermittedToPerformThisOperation
    End If
    
    Exit Function
ErrorHandler:
    ConvertError sSource
End Function




Private Function UserCanDeleteCatalogueEntryOnEditSet()
' Description:
'
' Checks whether the user can delete an entry on the editset
'
'Properties required to be set
'UserID
'EditSetID
'
' Amendment history:
' d18   ndk     17Nov1999       Created
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String, eSecurityCheckStatus As ESecurityCheckResult
    
    sSource = Me.Name & ".UserCanDeleteCatalogueEntryOnEditSet()"
    
    'If the user is both able to create and add catalogue entries
    'then he is permitted to delete entries
    
    'can he add
    eSecurityCheckStatus = UserCanAddCatalogueEntryToEditSet
    
    If eSecurityCheckStatus <> ESecurityCheckResult.ESRSecurityCheckSuccessful Then
        UserCanDeleteCatalogueEntryOnEditSet = eSecurityCheckStatus
        Exit Function
    End If
    
    'can he create
    eSecurityCheckStatus = UserCanCreateCatalogueEntryOnEditSet
    
    If eSecurityCheckStatus <> ESecurityCheckResult.ESRSecurityCheckSuccessful Then
        UserCanDeleteCatalogueEntryOnEditSet = eSecurityCheckStatus
        Exit Function
    End If
    
    'he can both add and create
    UserCanDeleteCatalogueEntryOnEditSet = ESecurityCheckResult.ESRSecurityCheckSuccessful
    
    Exit Function
ErrorHandler:
    ConvertError sSource
End Function


Private Function EditSetHasCandidateTerms() As ESecurityCheckResult
' Description:
'
' Checks whether the edit set has candidate terms
'
'Properties required to be set
'UserID
'EditSetID
'
' Amendment history:
'   d27   crm            14-Feb-2000   Created
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim ofwIndexTermList As FWRecordList
    Dim ofwDBHelper As New FWDBHelper
    Dim sSQL As String
    
    sSource = Me.Name & ".EditSetHasCandidateTerms()"
    
    If m_lStageToBeMovedFromID = gk_lStageCheckingID And m_lStageToBeMovedToID = gk_lStagePreReleaseID Then
        
        sSQL = "SP_Edit_GetEditSetIndexTermsDetails @tlEditSetId = " & m_lEditsetID _
                                    & ", @tlContributorID = " & gk_lContributor1RoleID _
                                    & ", @tlStageCheckingID = " & gk_lStageCheckingID _
                                    & ", @tlRejectedID      = " & gk_lCandidateTerm
        Set ofwIndexTermList = ofwDBHelper.FindBySql(sSQL)
        If ofwIndexTermList.RecordCount > 0 Then
                EditSetHasCandidateTerms = ESREditSetHasCandidateTerms
        Else
                EditSetHasCandidateTerms = ESRSecurityCheckSuccessful
        End If
    End If
    
    Set ofwIndexTermList = Nothing
    Set ofwDBHelper = Nothing
Exit Function
ErrorHandler:
    ConvertError sSource
End Function

Private Function UserIsScheduledJobEditor() As ESecurityCheckResult
' Used to check whether the user can edit scheduled jobs
'
' d35   map     17Oct2001       Created

On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim bHasRole As Boolean
    
    sSource = Me.Name & ".UserIsManagingEditor()"
    
    bHasRole = False
    
    ' User must have Admistrator role
    bHasRole = m_oboUserRoles.UserHasRole(m_lUserID, gk_lAdministratorRoleID)
       
    If bHasRole Then
        UserIsScheduledJobEditor = ESecurityCheckResult.ESRSecurityCheckSuccessful
    Else
        UserIsScheduledJobEditor = ESecurityCheckResult.ESRUserDoesNotHaveRoleOrHisRoleNotPermittedToPerformThisOperation
    End If
    
    Exit Function
ErrorHandler:
    ConvertError sSource
End Function

Private Function UserIsManagingEditor() As ESecurityCheckResult
' Used to check whether the user has the Managing editor role
'
' d34   gdb     08Feb2001       Created

On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim bHasRole As Boolean
    
    sSource = Me.Name & ".UserIsManagingEditor()"
    
    bHasRole = False
    
    ' User must have Managing Editor role
    bHasRole = m_oboUserRoles.UserHasRole(m_lUserID, gk_lManagingEditorRoleID)
       
    If bHasRole Then
        UserIsManagingEditor = ESecurityCheckResult.ESRSecurityCheckSuccessful
    Else
        UserIsManagingEditor = ESecurityCheckResult.ESRUserDoesNotHaveRoleOrHisRoleNotPermittedToPerformThisOperation
    End If
    
    Exit Function
ErrorHandler:
    ConvertError sSource
End Function


