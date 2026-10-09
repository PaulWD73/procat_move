Attribute VB_Name = "HardcodedIDConstants"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    HardcodedIDConstants.cls
' System:       PROCAT, PRO
' Copyright:    (C) Quidnunc Limited
'
' Description:  Used to store the hardcoded ID values of various flags etc.
'
' Amendment history:
'   d1      ksk     06Aug1999       Created
'   d2      ksk     06Aug1999       Moved constants from Editotial Constants
'                                   and added constants for Editset entry flag.
'   d3      GS      09Aug1999       Renamed the constants names since they were misspelt.
'   d4      ndk     30Sep1999       Added defn. for edit stage and reorganised the file
'                                   Add few more roles and stage ids
'   d5      ksk     08Oct1999       Added contributor2 & 3 role constants
'   d6      ksk     12Nov1999       Added role constant for all roles
'                                   Added Default Constants for closure status, closure code,
'                                   closure type and record status
'   d7      ndk     16Nov1999       value of gk_lStageReleaseID set to gk_lStagePreReleaseID
'                                   constant gk_lStageReleaseID removed
'   d8      ksk     13Dec1999       Added Accession as a role constant
'	d9		daj		24Feb2000		Removed default record status/closure code constants; moved
'									to Constants.bas as they are used in the Reader
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'Roles
Public Const gk_lContributor1RoleID As Long = 1         'Role Id of contributor1
Public Const gk_lContributor2RoleID As Long = 2         'Role Id of contributor2
Public Const gk_lContributor3RoleID As Long = 3         'Role Id of contributor3
Public Const gk_lEditorRoleID As Long = 4               'Role Id of Editor
Public Const gk_lAFEditorRoleID As Long = 5             'Role Id of A F Editor
Public Const gk_lProofReaderRoleID As Long = 6          'Role Id of Proof Reader
Public Const gk_lAdministratorRoleID As Long = 7        'Role Id of administrator
Public Const gk_lManagingEditorRoleID As Long = 8       'Role Id of Managing Editor

'   d6      ksk     12Nov1999
Public Const gk_lESFAEditorRoleID As Long = 9        'Role Id of Electronic SFA
Public Const gk_lPSFAEditorRoleID As Long = 10        'Role Id of Printed SFA
Public Const gk_lLeafletEditorRoleID As Long = 11        'Role Id of Leaflet editor
Public Const gk_lPopularSearchEditorRoleID As Long = 12        'Role Id of Popular searches
Public Const gk_lAccessionRoleID As Long = 13        'Role Id of Accession

'Stages for editset
Public Const gk_lStageContributionID As Long = 1    'Stage Id for contribution stage
Public Const gk_lStageEditID As Long = 2            'Stage Id for edit stage
Public Const gk_lStageCheckingID As Long = 3        'Stage id for checking stage
Public Const gk_lStagePreReleaseID As Long = 4      'Stage id for pre-release stage
Public Const gk_lStageAbandonID  As Long = 5        'Stage id for abandon stage

'Authority file term status
Public Const gk_lApprovedComplete As Long = 1
Public Const gk_lApprovedInComplete As Long = 2
'Public Const gk_lRejected As Long = 3 moved to shared constants jes
Public Const gk_lNotApproved As Long = 4
Public Const gk_lInvisible As Long = 5

'Editset entry status
Public Const gk_lEditsetEntryCreatedID                As Long = 1
Public Const gk_lEditsetEntryAccessionedID            As Long = 2
Public Const gk_lEditsetEntryCreatedAndModifiedID     As Long = 3
Public Const gk_lEditsetEntryAccessionedAndModifiedID As Long = 4



