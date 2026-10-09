VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BOPopSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    BOPopSearch.cls
' System:       PRO, PROCAT
' Copyright:    (C) Quidnunc Limited
'
' Description:  Represents a popular search. Used by Editorial
'               to create and update popular search details
'               The BOPopSearchLoader class should be used
'               to load pop search details.
'
'
'
' Amendment history:
'   d1  WDP     14-Oct-1999     Created stub
'   d2  CRM     09-Dec-1999     Added Save ,Load , Delete etc..methods
'   d3  WDP     13-Dec-1999     Added IndexInVerity
'   d4  WDP     27-Jan-2000     Split indexing into creation of data files
'                               (perfromed by BOPopularSearchIndexer) and
'                               the indexing itself (performed by BOIndexer -
'                               possibly remotely)'
'   d5  WDP     01-Feb-2000     Select saved searches based on NT login name
'                               rather than ticket number
'   d6  WDP     01-Feb-2000     Use overspill table for long search strings
'   d7  CRM     04-Feb-2000     Allowed IsReaderSearch field to be updateable
'   d8  TMJ     17-Mar-2000     Added checking to ensure a new name does not
'                               overwrite an old one
'   d9  WDP     29-Mar-2000     Altered IndexInVerity for deletes
'   d10 MAP     17-Jan-2001     Changed SavedSearches for replication solution
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit


' Class constants

' Instance variables

'The name used by editorial staff to identitfy the search
Private m_sName As String

'The search description to display for reader users
'and what searches on popular search data is performed over
Private m_sDescription As String

'Flag to indicate whether the search is available to reader users
Private m_bIsReaderSearch As Boolean

'The serialised search string (a serialisation of a BOSearch object)
Private m_sSearchString As String

' ObjectContext object and ASP objects can be accessed (if necessary)
Private m_objContext As Object

'Flag to indicate if a search is new (inserted into DB)
Private m_bIsNew As Boolean

'
' Public Methods
'
Public Property Get SearchString() As String
    SearchString = m_sSearchString
End Property

Public Property Get Name() As String
    Name = m_sName
End Property

Public Property Get Description() As String
    Description = m_sDescription
End Property

Public Property Get IsReaderSearch() As Boolean
    IsReaderSearch = m_bIsReaderSearch
End Property

Public Property Let SearchString(t_sSearchString As String)
    m_sSearchString = t_sSearchString
End Property

Public Property Let Name(t_sName As String)
    m_sName = t_sName
End Property

Public Property Let Description(t_sDescription As String)
    m_sDescription = t_sDescription
End Property

Public Property Let IsReaderSearch(t_bIsReaderSearch As Boolean)
    m_bIsReaderSearch = t_bIsReaderSearch
End Property

Public Sub Load(t_sName As String)
'   Description
'   Loads the popular search from the popular search table, delegates to BOPopSearchLoader
'
'   t_sName   -   the popular search id
'
'   d2  CRM    09Dec1999   Created
On Error GoTo ErrorHandler
    Dim sSource As String
    Dim oboPopSearch As New BOPopSearchLoader
    
    sSource = "BOPopSearch.Load()"
    
    If Len(Trim(t_sName)) > 0 Then
        'if we get a search id
        With oboPopSearch
            'load and fetch the details
            .LoadData t_sName
            m_sName = .Name
            m_sDescription = .Description
            m_bIsReaderSearch = .IsReaderSearch
            m_sSearchString = .SearchString
        End With
    End If
    
    'release object references
    Set oboPopSearch = Nothing
    
Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub

Public Sub Save(Optional ByVal t_bUpdate As Boolean = False)
'   Description
'   Saves the popular search to the popular search table
'
'   d2  CRM     09Dec1999   Created
'   d3  WDP     13Dec1999   Update m_bIsNew if new search
'                           so Verity indexing knows whether to
'                           insert or update
'   d6  WDP     01Feb2000   Use overspill table
'   d7  CRM     04Feb2000   Allowed IsReaderSearch field to be updateable
'   d10 TMJ     17Mar2000   Added error checking
On Error GoTo ErrorHandler
    Dim sSource As String
    Dim sSQL As String
    Dim ofwDBHelper As New FWDBHelper
    Dim ofwRecList As FWRecordList
    Dim sSearchStringPart As String, sSearchStringRemaining As String
    
    sSearchStringPart = Left(m_sSearchString, gk_lSearchStringMaxLen)
    sSearchStringRemaining = Mid(m_sSearchString, gk_lSearchStringMaxLen + 1)
    
    sSource = "BOPopSearch.Save()"
            
    If Len(Trim(m_sName)) > 0 Then
        'first check if it is an existing search
        sSQL = "SELECT " & _
            gk_sDBCOLPSDescription & ", " & _
            gk_sDBCOLPSSearchString & ", " & _
            gk_sDBCOLPSSearchName & ", " & _
            gk_sDBCOLPSIsReaderSearch & _
            " FROM " & gk_sDBTABLEPopularSearch & _
            " WHERE " & gk_sDBCOLPSSearchName & "=" & QuoteValueWithQuotes(m_sName)
        Set ofwRecList = ofwDBHelper.FindBySql(sSQL)
                
        'd10 added this to catch the error
        If ofwRecList.RecordCount = 1 And t_bUpdate = False Then
            RaiseError sSource, EPROCATINTERNALERROR.PEDupPopSearchName, "The name is already in use, " _
                       & "please choose another one"
            Exit Sub
        End If
                
        If ofwRecList.RecordCount = 0 Then
            'this is a new search
            m_bIsNew = True
            sSQL = "INSERT INTO " & gk_sDBTABLEPopularSearch & "(" & _
                    gk_sDBCOLPSSearchName & "," & _
                    gk_sDBCOLPSDescription & "," & _
                    gk_sDBCOLPSSearchString & "," & _
                    gk_sDBCOLPSIsReaderSearch & ")" & _
                    " VALUES ( " & _
                    QuoteValueWithQuotes(m_sName) & "," & _
                    QuoteValueWithQuotes(m_sDescription) & "," & _
                    QuoteValueWithQuotes(sSearchStringPart) & "," & _
                    BooleanAsID(m_bIsReaderSearch) & ")"
            'Insert the new Search
            ofwDBHelper.Execute sSQL
            
            
            '+put rest of long searches in overspill table
            If Not sSearchStringRemaining = "" Then
                Dim lOrder As Integer
                lOrder = 1
                While Not sSearchStringRemaining = ""
                    sSearchStringPart = Left(sSearchStringRemaining, gk_lSearchStringMaxLen)
                    sSearchStringRemaining = Mid(sSearchStringRemaining, gk_lSearchStringMaxLen + 1)
                    
                    sSQL = "INSERT INTO " & gk_sDBTABLEPopularSearchOverspill & _
                            "(" & gk_sDBCOLPSSearchName & ", " & _
                            gk_sDBCOLPSSearchString & ", " & _
                            gk_sDBCOLPopularSearchOverspillOrder & ") " & _
                            " VALUES (" & QuoteValueWithQuotes(m_sName) & ", " & _
                            QuoteValueWithQuotes(sSearchStringPart) & ", " & lOrder & ")"
                    ofwDBHelper.Execute sSQL
                    lOrder = lOrder + 1
                    
                Wend
            End If
        End If
        'if this is a existing search
        If ofwRecList.RecordCount = 1 Then
            m_bIsNew = False
            'update the search
            '   d7  CRM     04-Feb-2000
            sSQL = "UPDATE " & gk_sDBTABLEPopularSearch & _
                   " SET " & gk_sDBCOLPSDescription & " = " & _
                   QuoteValueWithQuotes(m_sDescription) & _
                   " ," & gk_sDBCOLPSSearchString & "=" & QuoteValueWithQuotes(m_sSearchString) & _
                   " , " & gk_sDBCOLPSIsReaderSearch & "=" & BooleanAsID(m_bIsReaderSearch) & _
                   " WHERE " & gk_sDBCOLPSSearchName & " = " & _
                   QuoteValueWithQuotes(m_sName)
            ofwDBHelper.Execute sSQL
            'delete old overspill
            sSQL = "DELETE " & gk_sDBTABLEPopularSearchOverspill & " WHERE " & _
                    gk_sDBCOLPSSearchName & "=" & QuoteValueWithQuotes(m_sName)
            ofwDBHelper.Execute sSQL
            '+put rest of long searches in overspill table
            If Not sSearchStringRemaining = "" Then
                Dim lOrder2 As Integer
                lOrder2 = 1
                While Not sSearchStringRemaining = ""
                    sSearchStringPart = Left(sSearchStringRemaining, gk_lSearchStringMaxLen)
                    sSearchStringRemaining = Mid(sSearchStringRemaining, gk_lSearchStringMaxLen + 1)
                    
                    sSQL = "INSERT INTO " & gk_sDBTABLEPopularSearchOverspill & _
                            "(" & gk_sDBCOLPSSearchName & ", " & _
                            gk_sDBCOLPSSearchString & ", " & _
                            gk_sDBCOLPopularSearchOverspillOrder & ") " & _
                            " VALUES (" & QuoteValueWithQuotes(m_sName) & ", " & _
                            QuoteValueWithQuotes(sSearchStringPart) & ", " & lOrder2 & ")"
                    ofwDBHelper.Execute sSQL
                    lOrder = lOrder + 1
                    
                Wend
            End If
        End If
    End If
        
    'release object references
    Set ofwDBHelper = Nothing
    Set ofwRecList = Nothing
    
Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub


Public Sub IndexInVerity(Optional t_eIndexAction As Variant)
'   Description
'   Saves the popular search to the popular search table
'
'   d3  WDP    13Dec1999        Created
'   d4  WDP    27-Jan-2000      Split indexing into creation of data files
'                               (perfromed by BOPopularSearchIndexer) and
'                               the indexing itself (performed by BOIndexer -
'                               possibly remotely)
'   d9  WDP     29-Mar-2000     Altered IndexInVerity for deletes
On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim oboVerIndexer
    sSource = "BOPopSearch.Save()"
            
    Dim eIndexAction As EIndexingAction
    
    'Inserts and updates read the action from whether the
    'search is new or not, deletes will pass it as a parameter
    If IsMissing(t_eIndexAction) Then
        If m_bIsNew = False Then
            eIndexAction = IDXUpdate
        Else
            eIndexAction = IDXInsert
        End If
    Else
        eIndexAction = IDXDelete
    End If
    
    Dim oboIndexer As New BOPopularSearchIndexer
    Dim sActionLog As String
    
    'Firstly we create SGML and BLK data files
    'the indexing actions are returned in the logfile
    sActionLog = oboIndexer.IndexPopSearch(m_sName, eIndexAction)
    
    Set oboVerIndexer = CreateObject("PROCATEdVerIdx.BOIndexer")
    'The Verity server may be remote so we make a (D)COM call
    'to trigger the actual indexing from the logfile
    oboVerIndexer.IndexFromIndexLog sActionLog, False, True
    Set oboVerIndexer = Nothing
        
Exit Sub
ErrorHandler:
    Set oboVerIndexer = Nothing
    ConvertError sSource
End Sub

'
' Class initialisation
'

Private Sub Initialise()
    Set m_objContext = GetObjectContext()
    
    ' Specific initialisation
    
End Sub

Private Sub Class_Initialize()
    Initialise
End Sub

Private Sub Class_Terminate()
    ' specific clear up
End Sub

Public Function SavedSearches(ByVal t_lLoggedUserID As Long) As Variant
'   Description
'   Returns a recordlist containing the saved searches for the particular user
'
'
'   d2  CRM    09Dec1999   Created
'   d5  WDP    1-Feb-2000  Select searches based on NT username rather than ticket number

On Error GoTo ErrorHandler
    Dim sSource As String
    Dim sSQL As String
    Dim ofwDBHelper As New FWDBHelper
    Dim ofwRecList As FWRecordList
    Dim lUserReaderID As Long
    
    sSource = "BOPopSearch.SavedSearches()"
    
    'first get the User's Reader ID based on the User_Ticket number in the
    'Reader_User table
    sSQL = "SELECT " & gk_sDBCOLReaderUserID & _
           " FROM " & gk_sDBTABLEUsers & "," & gk_sDBTABLEReaderUser & _
           " WHERE " & _
           gk_sDBTABLEReaderUser & "." & gk_sDBCOLReaderUserNTLogin & "=" & _
           gk_sDBTABLEUsers & "." & gk_sDBCOLUserNTlogon & " AND " & _
           gk_sDBTABLEUsers & "." & gk_sDBCOLUserID & "=" & t_lLoggedUserID & " AND " & _
           gk_sDBTABLEReaderUser & "." & gk_sDBCOLReaderUserType & "=" & DORISUTStaff
              
    Set ofwRecList = ofwDBHelper.FindBySql(sSQL)
    If ofwRecList.RecordCount = 1 Then
        lUserReaderID = ofwRecList.GetFieldValue(gk_sDBCOLReaderUserID)
    End If
    
    'now get the saved searches for the USer
    'd10 MAP Added AND clause in statement to only retrieve non-deleted searches
    sSQL = "SELECT * FROM " & gk_sDBTABLESavedSearch & _
           " WHERE " & gk_sDBCOLReaderUserID & " = " & lUserReaderID & _
           " AND " & gk_sDBCOLSavedSearchDeleted & " = 0"
    Set ofwRecList = ofwDBHelper.FindBySql(sSQL)
        
    'return the record list
    Set SavedSearches = ofwRecList
    
    'release object references
    Set ofwRecList = Nothing
    Set ofwDBHelper = Nothing
    
Exit Function
ErrorHandler:
    ConvertError sSource
End Function

Public Function PopularSearches(ByVal t_lLoggedUserID As Long) As Variant
'   Description
'   Returns a recordlist containing the Popular searches for the particular user
'
'
'   d2  CRM    09Dec1999   Created

On Error GoTo ErrorHandler
    Dim sSource As String
    Dim sSQL As String
    Dim ofwDBHelper As New FWDBHelper
    Dim ofwRecList As FWRecordList
    Dim lUserReaderID As Long
    
    sSource = "BOPopSearch.PopularSearches()"
    
    'now get the popular searches
    sSQL = "SELECT * FROM " & gk_sDBTABLEPopularSearch
    Set ofwRecList = ofwDBHelper.FindBySql(sSQL)
        
    'return the record list
    Set PopularSearches = ofwRecList
    
    'release object references
    Set ofwRecList = Nothing
    Set ofwDBHelper = Nothing
    
Exit Function
ErrorHandler:
    ConvertError sSource
End Function

Public Sub Delete(ByVal t_sPSName As String)
'   Description
'   Deletes the popular search from the popular search table
'
'   d2  CRM    10Dec1999   Created
'   d6  WDP    1Feb2000    Delete from overspill table too

On Error GoTo ErrorHandler
    Dim sSource As String
    Dim sSQL As String
    Dim ofwDBHelper As New FWDBHelper
        
    sSource = "BOPopSearch.Delete()"
    
            
    If Len(Trim(t_sPSName)) > 0 Then
        sSQL = "DELETE " & gk_sDBTABLEPopularSearch & _
               " WHERE " & gk_sDBCOLPSSearchName & "=" & QuoteValueWithQuotes(t_sPSName)
        ofwDBHelper.Execute sSQL
        sSQL = "DELETE " & gk_sDBTABLEPopularSearchOverspill & _
               " WHERE " & gk_sDBCOLPSSearchName & "=" & QuoteValueWithQuotes(t_sPSName)
        ofwDBHelper.Execute sSQL
    End If
    
    Set ofwDBHelper = Nothing
        
Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub
