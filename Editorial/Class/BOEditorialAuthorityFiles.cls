VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BOEditorialAuthorityFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    BOEditorialAuthorityFiles.cls
' System:       PRO, PROCAT
' Copyright:    (C) Quidnunc Limited
'
' Description:
'           This is the business object used to implement business logic of
'           Authority Files Maintenance.
'
' Amendment history:
'   d1  Geedha Sivanadian     19July1999   Created
'   d2  Geedha Sivanadian     23July1999   Removed the restriction to retrieve only the index terms raised
'                                          through edit_sets that are assigned to logged user.
'   d3  Geedha Sivanadian     03Aug1999    Added a method to return authority status values.
'   d4  AGA                   11Oct1999    Minor modifications to "Not used Term" related methods
'   d5  ndk                   15Nov1999    Changed the code to use pre-release wherever release was used
'   d6  ksk                   23Nov1999    Changed query to use max in the list
'   d7  aga                   07Dec1999    Modifications for new specs
'   d8  aga                   13Dec1999    Specifying the UserId for getting the editset list
'   d9  aga                   15Dec1999    Minor modifications
'   d10 crm                   11Feb2000    Modifications to GetLeafletsList &
'                                          GetLeafletIndexTerms
'   d11 d11 crm               11Feb2000
'   d12 crm                   14Feb2000    Changed parameter for GetEditSetList
'                                          GetEditSetIndexTerms
'   d13 daj                   09May2000    Added extra parameter to SP call to get back
'                                           the Contributor1 rather than the AFEditor

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Class constants

' Instance variables
Dim m_ofwSQLExecuteHelper As FWDBHelper


' ObjectContext object and ASP objects can be accessed (if necessary)
Private m_objContext As Object
Private m_aspResponse As Response

Private Const mk_iMaxRecordsToShow As Integer = 32000

Private Const gk_lCandidateTerm As Long = 4




'
' Public Methods
'

Public Function GetLeafLetsList() As FWRecordList
'   Retrieve all non approved index terms raised through various leaflets
'
'
'   d1  Geedha Sivanadian     20July1999   Created
'   d7  aga                   07Dec1999    Changed the sp
'   d11 crm                   11Feb2000    Changed parameter, sending only candidate
'                                          status value so that the SP can check for
'                                          everything which does not have an candidate
'                                          status
'NOTE: Though there are two args to the SP @tlNotApprovedID and @tlRejectedID
'      only one is being used, which explains why both have the same arg values.
'      the number of args has been kept the same to ensure backward compatibility
'      with any other code which might use the sp.There is no code which uses this
'      but arg kept as such for safety.The sp uses only the @tlNotApprovedID arg.
On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim sSQL As String
    
    sSource = "BOEditorialAuthorityFiles.GetLeafLetsList()"
                       

    'd7
    'sSQL = "SP_Edit_GetLeafLetsList @tlNotApprovedID     = " & gk_lApprovedComplete _
                                    & ", @tlRejectedID     = " & gk_lRejected
    'd11
    sSQL = "SP_Edit_GetLeafLetsList @tlNotApprovedID     = " & gk_lCandidateTerm _
                                    & ", @tlRejectedID     = " & gk_lCandidateTerm
                
    Set GetLeafLetsList = m_ofwSQLExecuteHelper.FindBySql(sSQL, mk_iMaxRecordsToShow)
    
    LogMessage sSource & ": Retrieved the list of LeafSets", pComplete
    
Exit Function
ErrorHandler:
    ConvertError sSource
End Function

Public Function GetEditSetsList(t_lUserID As Long) As FWRecordList
'   Retrieve all non approved index terms raised through various edit sets
'
'
'   d1  Geedha Sivanadian     20July1999   Created
'   d2  Geedha Sivanadian     23July1999   Removed the user_id parameter.
'   d7  aga                   07Dec1999    Changed the sp
'   d8  aga                   13Dec1999    Specifying the UserId for getting the editset list
'   d9  aga                   15Dec1999    Changed the parameter
'   d12 crm                   14Feb2000    Changed parameter.
'   d13 daj                   09May2000    Added extra parameter to SP call to get back
'                                           the Contributor1 rather than the AFEditor
'
On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim sSQL As String
    
    sSource = "BOEditorialAuthorityFiles.GetEditSetsList()"
                       
    ' d12 crm                   14Feb2000
    'The change in args is to list only candidate terms
    ' d13 - Added Contibutor1 role ID
    sSQL = "SP_Edit_GetEditSetsList @tlUserID = " & t_lUserID & _
        ", @tlAFEditorID = " & gk_lAFEditorRoleID & _
        ", @tlContributor1ID = " & gk_lContributor1RoleID & _
        ", @tlRejectedID      = " & gk_lCandidateTerm
                                
    Set GetEditSetsList = m_ofwSQLExecuteHelper.FindBySql(sSQL, mk_iMaxRecordsToShow)
    
    LogMessage sSource & ": Retrieved the list of EditSets", pComplete
    
Exit Function
ErrorHandler:
    ConvertError sSource
End Function


Public Function GetInvisibleIndexTermsCount() As FWRecordList
'   Retrieve all index terms where authority_status = "Invisible"
'
'
'   d1  Geedha Sivanadian     29July1999   Created
'   d7  AGA                   07Dec1999   Changed sp name from SP_Edit_GetLiveTerms to SP_Edit_GetTermsCountWithStatus

On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim sSQL As String
    
    sSource = "BOEditorialAuthorityFiles.GetInvisibleIndexTermsCount()"
    'd7
    sSQL = "SP_Edit_GetTermsCountWithStatus @tlAuthorityStatusId = " & gk_lInvisible
            
    Set GetInvisibleIndexTermsCount = m_ofwSQLExecuteHelper.FindBySql(sSQL)
    
    LogMessage sSource & ": Retrieved invisible index terms count", pComplete
    
Exit Function
ErrorHandler:
    ConvertError sSource
End Function

Public Function GetEditSetIndexTerms(tlEditSetId As Long) As FWRecordList
'   Retrieve all not approved index terms for an editset.
'
'
'   d1  Geedha Sivanadian     30July1999   Created
'   d7  aga                   07Dec1999    Changed the sp
'   d12 crm                   14Feb2000    Changed argument so that sp can list only
'                                          Candidate terms

On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim sSQL As String
    
    sSource = "BOEditorialAuthorityFiles.GetEditSetIndexTerms() Edit Set Id:" & tlEditSetId
    'd7
    'sSQL = "SP_Edit_GetEditSetIndexTermsDetails @tlEditSetId = " & tlEditSetId _
                                & ", @tlContributorID = " & gk_lContributor1RoleID _
                                & ", @tlStageCheckingID = " & gk_lStageCheckingID _
                                & ", @tlRejectedID      = " & gk_lRejected
                
    sSQL = "SP_Edit_GetEditSetIndexTermsDetails @tlEditSetId = " & tlEditSetId _
                                & ", @tlContributorID = " & gk_lContributor1RoleID _
                                & ", @tlStageCheckingID = " & gk_lStageCheckingID _
                                & ", @tlRejectedID      = " & gk_lCandidateTerm
    Set GetEditSetIndexTerms = m_ofwSQLExecuteHelper.FindBySql(sSQL)
    
    LogMessage sSource & ": Retrieved index terms for editset:" & tlEditSetId, pComplete
    
Exit Function
ErrorHandler:
    ConvertError sSource
End Function

Public Function GetLeafLetIndexTerms(tlLeafLetId As Long) As FWRecordList
'   Retrieve all non approved index terms for a leaflet
'
'
'   d1  Geedha Sivanadian     2 Aug 1999   Created
'   d7  aga                   07Dec1999    Changed the sp
'   d11 crm                   11Feb2000    Changed parameter, sending only candidate
'                                          status value so that the SP can check for
'                                          everything which does not have an candidate
'                                          status
'NOTE: Though there are two args to the SP @tlNotApprovedID and @tlRejectedID
'      only one is being used, which explains why both have the same arg values.
'      the number of args has been kept the same to ensure backward compatibility
'      with any other code which might use the sp.There is no code which uses this
'      but arg kept as such for safety.The sp uses only the @tlNotApprovedID arg.
On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim sSQL As String
    
    sSource = "BOEditorialAuthorityFiles.GetLeafLetIndexTerms() Leaflet Id:" & tlLeafLetId
                       
    'd7
    'sSQL = "SP_Edit_GetLeafLetIndexTermDetails @tlLeafLetReferenceId = " & tlLeafLetId _
                                        & ", @tlNotApprovedStatus = " & gk_lApprovedComplete _
                                        & ", @tlRejectedID = " & gk_lRejected
    'd11
    sSQL = "SP_Edit_GetLeafLetIndexTermDetails @tlLeafLetReferenceId = " & tlLeafLetId _
                                        & ", @tlNotApprovedStatus = " & gk_lCandidateTerm _
                                        & ", @tlRejectedID = " & gk_lCandidateTerm
    Set GetLeafLetIndexTerms = m_ofwSQLExecuteHelper.FindBySql(sSQL)
    
    LogMessage sSource & ": Retrieved index terms for Leafet:" & tlLeafLetId, pComplete
    
Exit Function
ErrorHandler:
    ConvertError sSource
End Function


Public Function GetIncompleteIndexTermsDetails() As FWRecordList
'   Retrieve all Incomplete index terms.
'
'
'   d1  Geedha Sivanadian     2 Aug 1999   Created
'   d7  AGA                   07Dec1999   Changed sp name from SP_Edit_GetLiveTermsDetails

On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim sSQL As String
    
    sSource = "BOEditorialAuthorityFiles.GetIncompleteIndexTermsDetails()"
    'd7
    sSQL = "SP_Edit_GetTermsDetailsWithStatus @tlAuthorityStatusID     = " & gk_lApprovedInComplete
            
    Set GetIncompleteIndexTermsDetails = m_ofwSQLExecuteHelper.FindBySql(sSQL)
    
    LogMessage sSource & ": Retrieved incomplete index terms", pComplete
    
Exit Function
ErrorHandler:
    ConvertError sSource
End Function

Public Function GetInvisibleIndexTermsDetails() As FWRecordList
'   Retrieve all Invisible index terms.
'
'
'   d1  Geedha Sivanadian     2 Aug 1999   Created
'   d7  AGA                   07Dec1999   Changed sp name from SP_Edit_GetLiveTermsDetails

On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim sSQL As String
    
    sSource = "BOEditorialAuthorityFiles.GetInvisibleIndexTermsDetails()"
    'd7
    sSQL = "SP_Edit_GetTermsDetailsWithStatus @tlAuthorityStatusID     = " & gk_lInvisible
            
    Set GetInvisibleIndexTermsDetails = m_ofwSQLExecuteHelper.FindBySql(sSQL)
    
    LogMessage sSource & ": Retrieved invisible index terms", pComplete
    
Exit Function
ErrorHandler:
    ConvertError sSource
End Function


Public Function GetNotUsedIndexTermsCount() As FWRecordList
'   Retrieve all not used index terms where authority_status = "Not Approved"
'
'   d1  Geedha Sivanadian     29July1999   Created
'   d4  AGA                         11Oct1999   Changed the type from 'NOT APPROVED' to 'APPROVED COMPLETE'
'                                                        NOTE: Not used terms are "Approved Complete terms" which have no links to
'                                                           ES_PC, PC and SFA tables
                                                        

On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim sSQL As String
    
    sSource = "BOEditorialAuthorityFiles.GetNotUsedIndexTermsCount()"
                       
    sSQL = "SP_Edit_GetNotUsedTerms @tlNotApprovedID = " & gk_lApprovedComplete
            
    Set GetNotUsedIndexTermsCount = m_ofwSQLExecuteHelper.FindBySql(sSQL)
    
    LogMessage sSource & ": Retrieved NotUsed index terms count", pComplete
    
Exit Function
ErrorHandler:
    ConvertError sSource
End Function

Public Function GetIncompleteIndexTermsCount() As FWRecordList
'   Retrieve all index terms where authority_status = "Approved Incomplete"
'
'
'   d1  Geedha Sivanadian     29Jul1999   Created
'   d7  AGA                   07Dec1999   Changed sp name from SP_Edit_GetLiveTerms

On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim sSQL As String
    
    sSource = "BOEditorialAuthorityFiles.GetIncompleteIndexTermsCount()"
                       
    sSQL = "SP_Edit_GetTermsCountWithStatus @tlAuthorityStatusId = " & gk_lApprovedInComplete
            
    Set GetIncompleteIndexTermsCount = m_ofwSQLExecuteHelper.FindBySql(sSQL)
    
    LogMessage sSource & ": Retrieved incomplete index terms count", pComplete
    
Exit Function
ErrorHandler:
    ConvertError sSource
End Function

'd3
Public Function GetAuthorityStatusValues() As FWRecordList
'   Retrieve all authority status values from authority status table.
'
'
'   d1  Geedha Sivanadian     3Aug1999   Created

On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim sSQL As String
    
    sSource = "BOEditorialAuthorityFiles.GetAuthorityStatusValues()"
                       
    sSQL = "SELECT Authority_Status " & gk_sDBCOLID & " , " & _
            "Authority_Status_Description " & gk_sDBCOLValue & _
            " FROM authority_status WHERE currency_indicator = 1"
    
    Set GetAuthorityStatusValues = m_ofwSQLExecuteHelper.FindBySql(sSQL)
    
    LogMessage sSource & ": Retrieved the list authority status values", pComplete
    
Exit Function
ErrorHandler:
    ConvertError sSource
End Function
Public Function GetNonCandidateNonPreferredTermsCount() As FWRecordList
'   returns the count of non-candidate, non preferred index terms
'
'   d7  AGA     07Dec1999       Created
                                                        

On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim sSQL As String
    
    sSource = "BOEditorialAuthorityFiles.GetNonCandidateNonPreferredTermsCount()"
                       
    sSQL = "SP_Edit_GetNonCandidateNonPreferredTerms @tlCandidateTypeID  = " & gk_lNotApproved _
            & ", @tlRejectedID = " & gk_lRejected
            
    Set GetNonCandidateNonPreferredTermsCount = m_ofwSQLExecuteHelper.FindBySql(sSQL)
    
    LogMessage sSource & ": Retrieved NotUsed index terms count", pComplete
    
Exit Function
ErrorHandler:
    ConvertError sSource
End Function
Public Function GetNonCandidateNonPreferredTermsDetails() As FWRecordList
'   retrieves all non-candidate, non preferred index terms
'
'   d7  AGA     07Dec1999       Created
On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim sSQL As String
    Dim lNumberOfRecords As Long
    
    sSource = "BOEditorialAuthorityFiles.GetNonCandidateNonPreferredTermsDetails()"
                       
    sSQL = "SP_Edit_GetNonCandidateNonPreferredTermDetails @tlCandidateTypeID     = " & gk_lNotApproved _
            & ", @tlRejectedID = " & gk_lRejected
    'Get the number of record
    lNumberOfRecords = GetNonCandidateNonPreferredTermsCount().GetFieldValue(gk_sDBCOLTempCounts)
    Set GetNonCandidateNonPreferredTermsDetails = m_ofwSQLExecuteHelper.FindBySql(sSQL, lNumberOfRecords)
    
    LogMessage sSource & ": Retrieved invisible index terms", pComplete
    
Exit Function
ErrorHandler:
    ConvertError sSource
End Function
Public Function GetNotUsedPreferredTermsCount() As FWRecordList
'   Retrieve count of not used preferred index terms which are not used anywhere.
'
'   d7  AGA     07Dec1999       Created
                                         

On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim sSQL As String
    
    sSource = "BOEditorialAuthorityFiles.GetNotUsedPreferredTermsCount()"
                       
    sSQL = "SP_Edit_GetNotUsedPreferredTerms @tlCandidateTypeID      = " & gk_lNotApproved _
            & ", @tlRejectedID = " & gk_lRejected
            
    Set GetNotUsedPreferredTermsCount = m_ofwSQLExecuteHelper.FindBySql(sSQL, mk_iMaxRecordsToShow)
    
    LogMessage sSource & ": Retrieved not used preferred index terms", pComplete
    
Exit Function
ErrorHandler:
    ConvertError sSource
End Function
Public Function GetNotUsedPreferredTermsDetails() As FWRecordList
'   retrieves all not used preferred index terms
'
'   d7  AGA     07Dec1999       Created
On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim sSQL As String
    Dim lNumberOfRecords As Long
    
    sSource = "BOEditorialAuthorityFiles.GetNotUsedPreferredTermsDetails()"
                       
    sSQL = "SP_Edit_GetNotUsedPreferredIndexTermDetails @tlCandidateTypeID = " & gk_lNotApproved _
            & ", @tlRejectedID = " & gk_lRejected
    'Get the number of record
    lNumberOfRecords = GetNotUsedPreferredTermsCount().GetFieldValue(gk_sDBCOLTempCounts)
    Set GetNotUsedPreferredTermsDetails = m_ofwSQLExecuteHelper.FindBySql(sSQL, lNumberOfRecords)
    
    LogMessage sSource & ": Retrieved invisible index terms", pComplete
    
Exit Function
ErrorHandler:
    ConvertError sSource
End Function

Public Function GetNotUsedCandidateTermsCount() As FWRecordList
'   Retrieve count of not used candidate index terms which are not used anywhere.
'
'   d7  AGA     07Dec1999       Created
                                         

On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim sSQL As String
    
    sSource = "BOEditorialAuthorityFiles.GetNotUsedCandidateTermsCount()"
                       
    sSQL = "SP_Edit_GetNotUsedCandidateTerms @tlCandidateTypeID      = " & gk_lNotApproved _
            & ", @tlRejectedID = " & gk_lRejected
            
    Set GetNotUsedCandidateTermsCount = m_ofwSQLExecuteHelper.FindBySql(sSQL, mk_iMaxRecordsToShow)
    
    LogMessage sSource & ": Retrieved not used candidate index terms", pComplete
    
Exit Function
ErrorHandler:
    ConvertError sSource
End Function
Public Function GetNotUsedCandidateTermsDetails() As FWRecordList
'   retrieves all not used candidate index terms
'
'   d7  AGA     07Dec1999       Created
On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim sSQL As String
    Dim lNumberOfRecords As Long
    
    sSource = "BOEditorialAuthorityFiles.GetNotUsedCandidateTermsDetails()"
                       
    sSQL = "SP_Edit_GetNotUsedCandidateTermDetails @tlCandidateTypeID = " & gk_lNotApproved _
            & ", @tlRejectedID = " & gk_lRejected
    'Get the number of record
    lNumberOfRecords = GetNotUsedCandidateTermsCount().GetFieldValue(gk_sDBCOLTempCounts)
    Set GetNotUsedCandidateTermsDetails = m_ofwSQLExecuteHelper.FindBySql(sSQL, lNumberOfRecords)
    
    LogMessage sSource & ": Retrieved invisible index terms", pComplete
    
Exit Function
ErrorHandler:
    ConvertError sSource
End Function

'' Class initialisation
'
Public Function GetTermsDetailsWithStatus(t_lStatus As Long) As FWRecordList
'   Retrieve all Incomplete index terms.
'
'
'   d7  AGA                   07Dec1999   Created

On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim sSQL As String
    
    sSource = "BOEditorialAuthorityFiles.GetTermsDetailsWithStatus()"
                       
    sSQL = "SP_Edit_GetTermsDetailsWithStatus @tlAuthorityStatusID     = " & t_lStatus
            
    Set GetTermsDetailsWithStatus = m_ofwSQLExecuteHelper.FindBySql(sSQL)
    
    LogMessage sSource & ": Retrieved incomplete index terms", pComplete
    
Exit Function
ErrorHandler:
    ConvertError sSource
End Function

Public Function GetTermsCountWithStatus(t_lStatus As Long) As FWRecordList
'   Retrieve the count of terms with status.
'
'
'   d7  AGA                   07Dec1999   Created

On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim sSQL As String
    
    sSource = "BOEditorialAuthorityFiles.GetTermsCountWithStatus()"
                       
    sSQL = "SP_Edit_GetTermsCountWithStatus @tlAuthorityStatusID     = " & t_lStatus
            
    Set GetTermsCountWithStatus = m_ofwSQLExecuteHelper.FindBySql(sSQL)
    
    LogMessage sSource & ": Retrieved invisible index terms", pComplete
    
Exit Function
ErrorHandler:
    ConvertError sSource
End Function



Private Sub Class_Initialize()
    Initialise
End Sub
Private Sub Initialise()
    Set m_objContext = GetObjectContext()
    Set m_aspResponse = m_objContext.Item("Response")
        
    ' Specific initialisation
    Set m_ofwSQLExecuteHelper = New FWDBHelper
    
End Sub
Private Sub Class_Terminate()
    Set m_ofwSQLExecuteHelper = Nothing
End Sub
