VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "VWEntrySearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    VWReaderSearch.cls
' System:       PRO, PROCAT
' Copyright:    (C) Quidnunc Limited
'
' Description:  View class to display search forms, run searches
'               and show search results
'
' Amendment history:
'   d1  WDP     8-Jul-1999   Created
'   d2  WDP     12-Jul-1999  Extending existing work (parsing restrictions
'                            Ranges of results, prev/next, transient storage)
'   d3  WDP     13-Jul-1999  Extended further (search context, sort options)
'   d4  WDP     15-Jul-1999  Extended further (sort type in results)
'   d5  WDP     30-Jul-1999  Fixed defect in GetBasicSearchFields
'   d6  WDP     12-Aug-1999  Added help button in HeaderHTML
'   d7  SPB     13-Aug-1999  Change target of link in AddTableRow
'                            to be DisplayCatDetailsSearch.asp?<ref parts>
'                            Added GetTargetURL function.
'   d8  SPB     13-aug-1999  Added ExecuteLastSearch so user can run
'                            last stored search from anywhere.
'   d9  WDP     13-Aug-1999  Added functionality to cope with different
'                            types of results (list, summary or too many)
'   d10 JES     30-Sep-1999  Added BuildPopularSearch
'   d11 JES     04-Oct-1999  Added search type setting to build methods.
'                            Stopped the popular search displaying a refine button
'   d12 JES     19-OCT-1999  Parsed search term in basic search
'                            Added leaflet search
'   d13 WDP     19-Oct-1999  Moved FT term parsing to BOVeritySearch
'                            and fixed display defect in DisplayResults
'   d14 JES     27-Oct-1999  Added BuildAdvancedSearch
'   d15 JES     28-Oct-1999  Updated GetRefineURL to cope with different search types
'   d16 JES     1-Nov-1999   Added restrictions to advanced search
'   d17 JES     2-Nov-1999   Added RefineadvancedsearchURL
'   d18 JES     3-Nov-1999   Added sort type to advanced search
'   d19 WDP     3-Nov-1999   Added some go to cat entry stuff and
'                            altered ParseRestricitons to do proper
'                            restriction checking
'   d20 WDP     4-Nov-1999   Altered leaflet results member to be
'                            generic ILeafletResults type and added
'                            index search
'   d21 JES     5-Nov-1999   Added operator precedence to advanced search
'   d22 WDP     11-Nov-1999  Added BuildSavedSearch
'   d23 JES     22-Nov-1999  Added usage switches so editorial can use it
'   d24 JES     25-Nov-1999  Altered display functions to produce editorial display
'                            on editorial usage.
'   d25 WDP     2-Dec-1999   Added AddSaveButton and linked to it for saved searches
'   d26 JES     2-Dec-1999    Added LogStats() to send message to stats server
'   d27 JES     6-Dec-1999    Added reader link parameter "access method?" for stats logging
'   d28 JES     7-Jan-2000   Fixed some defects
'   d29 WDP     20-Jan-2000  Fixed defects, changed behaviour for no results
'   d30 WDP     1-Feb-2000   Searches given new number (for trans storage)
'                            rather than storing old restrictions when
'                            going between summary levels
'   d31 DTM     09-Feb-2000  Created gk_sImageSelectAll for AddResultButtons
'   d32 JCC     10-Feb-2000  Added white background
'   d33 WDP     22-Feb-2000  Added dates in search results for editorial
'   d34 DAJ     09-Mar-2000  Added deselect all button to the search results
'   d35 WDP     17-Mar-2000  Fixed date defects in BuildBasicSearch
'   d36 DAJ     05-May-2000  Removed the Refine button for GoTo search results
'   d37 DAJ     08-May-2000  Modified Basic Search to use 'yyyy' format for dates
'   d38 WDP     24-May-2000  Removed ParseVQL call in ExecuteLeafletSearch as this
'                            is done in the BO
'   d39 WDP     25-May-2000  Read # of editorial results to display from FWAppData
'   d40 WDP     31-May-2000  Read results behaviour boundary constants from registry
'   d41 DAJ     06-Jun-2000  Fixed date conversion problem in Basic search
'   d42 DAJ     07-Jun-2000  Added extra help text for when a reference search returns too many results
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit



Implements ObjectControl

Private Const gk_sMaxEditorialResultCount As Long = 200

Private m_objContext As Object
Private m_aspResponse As Response
Private m_aspRequest As Request
Private m_aspServer As Server

'Search object
Private m_boSearch As BOSearch

'Search instance number
Private m_sSearchNumber As String

'Results object (can be verity or database results)
Private m_oiResults As New ISearchResults

'HTML object for reference
Private m_fwHTML As FWHTML

'Strings to store old restrictions for
'navigating through results summaries
Private m_sClassLevelOldRestr As String
Private m_sLetCodeLevelOldRestr As String

' Constants
'!temporary - use resource file
Private Const gk_sDISPLAYRefHeader As String = "PRO Reference"
Private Const gk_sDISPLAYTitleHeader As String = "Title/Scope and Content"
Private Const gk_sDISPLAYDatesHeader As String = "Covering Dates"
Private Const gk_sDISPLAYFormerRefHeader As String = "Former Ref."
Private Const gk_sDISPLAYHitsHeader As String = "Hits"
Private Const gk_sDISPLAYAddHeader As String = "Add"
Private Const gk_sDISPLAYLevels As String = "Levels"

'Results table constants
Private Const mk_sTableWidth As String = "100%"
Private Const mk_lTablePadding As Long = 1
Private Const mk_lTableSpacing As Long = 2

'The number of results to display at one time
Private Const mk_lResultsToDisplay As Long = 20

'The leaflet results object
Private m_objLeafResults As ILeafletResults

'Table to display the leaflet results in
Private m_fwLeafTable As FWTable

'Set if leaflet search has been carried out
Private m_bDisplayLeaflets As Boolean

'The usage
Private m_eUsage As ESearchUsage

'Row counter for adding checkboxes to editorial version
Private m_lRowCount As Long
'Stub name field for editorial checkboxes
Private Const mk_sCheckboxName As String = "CHKBX"
Private Const mk_sNameCatalogueSearch As String = "FCS"
Private Const mk_sNameCheckBoxCount As String = "CBN"

'User stats message
Private m_boMessage As BOMessage

'Maximum number of results
' no longer needed - read from registry
'Private Const mk_lMaxResults As Long = 3000
'Private Const mk_lMaxSingleClassResults As Long = 200

Public Sub SetSearchUsage(ByVal t_sSearchUsage As String)
'   Sets the usage for the search
'   d23 JES 22-Nov-1999 Created
On Error GoTo ErrorHandler
    Dim sSource As String
    sSource = "VWReaderSearch.SetSearchUsage"
    If LCase(t_sSearchUsage) = "reader" Then
        m_eUsage = ESearchReader
        
    ElseIf LCase(t_sSearchUsage) = "editorial" Then
        m_eUsage = ESearchEditorial
    Else
        sSource = sSource & ": Invalid search usage " & t_sSearchUsage
        GoTo ErrorHandler
    End If
Exit Sub
ErrorHandler:
    m_objContext.SetComplete
    ShowError sSource
End Sub

Public Sub GetSearchFromNumber(ByVal t_sSearchNumber As String)
'   Build a search from transient store given its instance number
'
'   d4  WDP    15-Jul-1999      Created

On Error GoTo ErrorHandler
    
    Dim sSource As String
    sSource = "VWReaderSearch.GetSearchFromNumber"
    Set m_boSearch = New BOSearch
    m_sSearchNumber = t_sSearchNumber
    m_boSearch.PROCATTransientInterface_GetState t_sSearchNumber
    
Exit Sub
ErrorHandler:
    If Err.Number = PETMInvalidInstanceReference Then
        'convert the error into a "cannot get search from transient store error"
        Err.Description = "Search has expired in transient store"
        Err.Number = peUSERSearchNoLongerInTransientStore
    End If
    ' Only use SetAbort if this method did an update,
    ' (and then there should have been no transient storage).
    ' Otherwise use SetComplete
    m_objContext.SetComplete
    
    ShowError sSource
End Sub


Public Sub SetSearchSortType(ByVal t_sSortType As String)
'   Set the search sort type - used for altering the
'   sort type of a search after it has been run
'
'   d4  WDP    15-Jul-1999      Created

On Error GoTo ErrorHandler
    
    Dim sSource As String
    sSource = "VWReaderSearch.SetSearchSortType"
    
    Select Case t_sSortType
        Case gk_sFORMSORTRelevance:
            m_boSearch.SortType = SORTRelevance
        Case gk_sFORMSORTReference:
            m_boSearch.SortType = SORTReference
        Case gk_sFORMSORTFormerRef:
            m_boSearch.SortType = SORTFormerReference
        Case gk_sFORMSORTCoveringDates:
            m_boSearch.SortType = SORTDate
        Case Else:
            '+ ADD ERROR HANDLING
            sSource = sSource & ": Invalid sort type."
    End Select
    
    'We have altered the search so need to restore it
    m_boSearch.PROCATTransientInterface_SetState m_sSearchNumber
    
    
    
Exit Sub
ErrorHandler:
    ' Only use SetAbort if this method did an update,
    ' (and then there should have been no transient storage).
    ' Otherwise use SetComplete
    m_objContext.SetComplete
    
    ShowError sSource
End Sub


Public Sub SetRestriction(ByVal t_sLettercode As String, _
                        ByVal t_sDivision As String, _
                        ByVal t_sClass As String)
'   Set the search breadth restriction - used when searching
'   within a summary entry
'
'   d9  WDP    17-Aug-1999      Created
'   d10 WDP    31-Aug-1999      Altered parameters to take
'                               separate class, div and lettercode

On Error GoTo ErrorHandler
    
    Dim sSource As String
    sSource = "VWReaderSearch.SetRestriction"
    Dim sNewRestriction As String
    
    'Dim l As Long
    'For l = 1 To m_boSearch.NumberOfBRestrictions
        'sRestr = sRestr & m_boSearch.GetBreadthRestriction(l)
    'Next
       
    'move the restrictions to the 'old' collection
    'm_boSearch.MoveBRestrictionsToOld
    'Add the new restriction
    
    sNewRestriction = t_sLettercode
    If Not t_sClass = "" Then
        sNewRestriction = sNewRestriction & " " & t_sClass
    End If
    
    m_boSearch.ClearBreadthRestrictions
    m_boSearch.AddBreadthRestriction (sNewRestriction)
    
    'Store the unrefined search number in the old search number
    'member of the search object so it can be recovered
    m_boSearch.OldSearchNumber = m_sSearchNumber
    'We have altered the search so need to restore it
    'under a new search number
    m_sSearchNumber = GetInstanceNumber()
    m_boSearch.PROCATTransientInterface_SetState m_sSearchNumber
    
Exit Sub
ErrorHandler:
    ' Only use SetAbort if this method did an update,
    ' (and then there should have been no transient storage).
    ' Otherwise use SetComplete
    m_objContext.SetComplete
    
    ShowError sSource
End Sub


Public Sub GetLastRestriction()
'   alter the search to have its previous set of
'   restrictions
'
'   d9  WDP    17-Aug-1999      Created
'   d30 WDP    1-Feb-2000       Use old search number member to
'                               get old unrestricted search

On Error GoTo ErrorHandler
    
    Dim sSource As String
    sSource = "VWReaderSearch.GetLastRestriction"
    
    'Get the old restrictions and make them current
    'm_boSearch.GetBRestrictionsFromOld
    
    Dim sOldSearchNumber As String
    sOldSearchNumber = m_boSearch.OldSearchNumber
    
    'recover the old search
    m_boSearch.PROCATTransientInterface_GetState sOldSearchNumber
    
    'We have altered the search so need to restore it
    'm_boSearch.PROCATTransientInterface_SetState m_sSearchNumber
    
Exit Sub
ErrorHandler:
    ' Only use SetAbort if this method did an update,
    ' (and then there should have been no transient storage).
    ' Otherwise use SetComplete
    m_objContext.SetComplete
    
    ShowError sSource
End Sub


Public Function GetBasicSearchFields(ByVal t_sSearchNumber As String) As String
'   Get a basic search object from storage and
'   get the display field values from it as a
'   concatenated string suitable to appen to
'   the BasicSearch.asp URL
'
'   d3  WDP     13-Jul-1999      Created
'   d5  WDP     30-Jul-1999      Fixed so can handle double quotes in
'                                the fields (using CodeQuotes)
'   d37 DAJ     08-May-2000  Modified Basic Search to use 'yyyy' format for dates

On Error GoTo ErrorHandler
    
    Dim sSource As String
    sSource = "VWReaderSearch.GetBasicSearchFields"
                       
    Dim sRes As String
    'The search object instance number
    Dim sSearchNumber As String
    'The basic search fields
    Dim sSearchTerm As String
    Dim sFirstDate As String
    Dim sLastDate As String
    Dim sRestriction As String
    Dim sSortType As String
    

    Set m_boSearch = New BOSearch
    'Get the search from storage
    m_boSearch.PROCATTransientInterface_GetState t_sSearchNumber
    
    'The search term will be in four terms (title, scope-content
    'admin history and arrangement) so we could look for any of these
    'terms - wlog we look for title term
    sSearchTerm = FindInTerm(gk_sDITitle, m_boSearch.RootSearchTerm)
    
    ' d37 - Use 'yyyy' format for Basic Search dates
    'Look for a first date term
    sFirstDate = Format(FindInTerm(gk_sDIFirstDate, m_boSearch.RootSearchTerm), gk_sReaderSearchBasicSearchDateFormat)
    'Look for a lastdate term
    sLastDate = Format(FindInTerm(gk_sDILastDate, m_boSearch.RootSearchTerm), gk_sReaderSearchBasicSearchDateFormat)
    
    'Get the breadth restrictions
    Dim l As Long
    For l = 1 To m_boSearch.NumberOfBRestrictions
        If Not sRestriction = "" Then
            sRestriction = sRestriction & ", "
        End If
        sRestriction = sRestriction & m_boSearch.GetBreadthRestriction(l)
    Next
    
    'Get the sort type
    Select Case m_boSearch.SortType
        Case SORTReference:
            sSortType = gk_sFORMSORTReference
        Case SORTRelevance:
            sSortType = gk_sFORMSORTRelevance
        Case SORTDate:
            sSortType = gk_sFORMSORTCoveringDates
        Case SORTFormerReference:
            sSortType = gk_sFORMSORTFormerRef
        Case Else:
            sSource = sSource & ": Invalid search sort type."
            GoTo ErrorHandler
    End Select
    
    Set m_fwHTML = New FWHTML
    sRes = "?" & gk_sFORMBASICSearchTerm & "=" & URLEncode(m_fwHTML.CodeQuotes(sSearchTerm)) & "&" & _
            gk_sFORMBASICFirstDate & "=" & URLEncode(m_fwHTML.CodeQuotes(sFirstDate)) & "&" & _
            gk_sFORMBASICLastDate & "=" & URLEncode(m_fwHTML.CodeQuotes(sLastDate)) & "&" & _
            gk_sFORMBASICRestriction & "=" & URLEncode(m_fwHTML.CodeQuotes(sRestriction)) & "&" & _
            gk_sFORMBASICSortType & "=" & URLEncode(m_fwHTML.CodeQuotes(sSortType))
                       
    GetBasicSearchFields = sRes
    
    m_objContext.SetComplete
    
Exit Function
ErrorHandler:
    ' Only use SetAbort if this method did an update,
    ' (and then there should have been no transient storage).
    ' Otherwise use SetComplete
    m_objContext.SetComplete
    
    ShowError sSource
End Function

Public Sub DisplayResults()
'   Generate the HTML from the search results member
'   calls methods for different sections of the page
'   '+NOTE could use helper class for some parts as reuse
'   in editorial is possible.
'
'   d1  WDP     8-Jul-1999   Created
'   d9  WDP     13-Aug-1999  Added call to TooManyResultsHTML
'   d12 JES     18-Oct-1999  Now all results are enclosed in an outer table
'                            to avoid problems with displaying leaflets
'                            Added leaflet table to results
'   d29 WDP     20-Jan-2000  Altered behaviour for no results

On Error GoTo ErrorHandler
    
    Dim sSource As String
    sSource = "VWReaderSearch.DisplayResults"
                           
    '+ use string buffer instead?
    'The HTML to display
    Dim sHTML As String
    Set m_fwHTML = New FWHTML
    
    'd12
    Dim ofwOuterTable As New FWTable
    Dim typTabStyle As TableStyle
    typTabStyle.eBGCOLOR = COLOURWhite
    
    ofwOuterTable.TableStyle = typTabStyle
    'Add the page start and header to the outer table
    
    ofwOuterTable.StartRow
    ofwOuterTable.AddCellAsHTML PageStartHTML() & HeaderHTML()
    ofwOuterTable.EndRow
    
    ofwOuterTable.AddBlankRow

    If m_bDisplayLeaflets = True And Not m_objLeafResults Is Nothing Then
        'Add the leaflet result summary
        If m_objLeafResults.ResultCount > 0 Then
            ofwOuterTable.StartRow
            ofwOuterTable.AddCellAsHTML AddLeafSummary
            ofwOuterTable.EndRow
            
            ofwOuterTable.AddBlankRow
            
            'display the leaflets
            Dim sLeafHTML As String
            ofwOuterTable.StartRow
            ofwOuterTable.AddCellAsHTML AddLeafList
            ofwOuterTable.EndRow
            
            ofwOuterTable.AddBlankRow
        End If
    End If
        
    
    'Only show #hits on catalogue if we have hits on the catalogue
    'Or if we have no leaflet hits either
    If m_objLeafResults Is Nothing Then
        m_bDisplayLeaflets = False
    Else
        If m_objLeafResults.ResultCount = 0 Then
            m_bDisplayLeaflets = False
        End If
    End If
    
    If Not (m_bDisplayLeaflets = True And m_oiResults.Hits = 0) Then
        'Show search context and result summary
        sHTML = sHTML & ContextHTML()
    End If
        
    'Get results table
    sHTML = sHTML & ResultsHTML()
    
    'End page
    sHTML = sHTML & PageEndHTML()
    
    ' Set the cookie to hold search number and page number
    ' (range min for page), so user can get back to
    ' the search results from anywhere
    If m_eUsage = ESearchReader Then
        m_aspResponse.Cookies(gk_sPROCATCookieName).Item(gk_sPROCATCookieCurrentSearch) = m_sSearchNumber
        m_aspResponse.Cookies(gk_sPROCATCookieName).Item(gk_sPROCATCookieCurrentSearchPage) = m_oiResults.RangeMin
    End If
    
    'd12 add main results to the outer table
    ofwOuterTable.StartRow
    ofwOuterTable.AddCellAsHTML sHTML
    ofwOuterTable.EndRow
    
    sHTML = ofwOuterTable.GetHTML()
    
    'Output result
    m_aspResponse.Write sHTML
      
    m_objContext.SetComplete
    
Exit Sub
ErrorHandler:
    ' Only use SetAbort if this method did an update,
    ' (and then there should have been no transient storage).
    ' Otherwise use SetComplete
    m_objContext.SetComplete
    
    ShowError sSource
End Sub


Public Sub ExecuteLastSearch()

'   Wrapper round ExecuteSearch called from details screens etc
'   so that we can rerun last search.
'
'   d1  WDP     13-Aug-1999   Created

On Error GoTo ErrorHandler
    Dim sSource As String
    sSource = "VWReaderSearch.ExecuteLastSearch"
        
    'Get the search instance number and page number from cookie
    Dim sRangeMin As String
    Dim lRangeMin As Long
    m_sSearchNumber = m_aspRequest.Cookies(gk_sPROCATCookieName).Item(gk_sPROCATCookieCurrentSearch)
    sRangeMin = m_aspRequest.Cookies(gk_sPROCATCookieName).Item(gk_sPROCATCookieCurrentSearchPage)
    
    If m_sSearchNumber = "" Then
        ' Havent had one before
        '+ Improve the error handling, eg redirect to home
        RaiseError sSource, PEUSERNoSearchToRun, "No search to run"
    End If
    
    If sRangeMin = "" Or (Not IsNumeric(sRangeMin)) Then
        ' Can't retrieve result number to start from
        lRangeMin = 1
    Else
        lRangeMin = CLng(sRangeMin)
    End If
    
    LogMessage sSource & ": Cookie: Search=" & m_sSearchNumber & ", Page=" & lRangeMin, pComplete
    
    
    ' Get the old search out of transient storage,
    ' and call executeSearch as if we are calling next
    Set m_boSearch = New BOSearch
    m_boSearch.PROCATTransientInterface_GetState (m_sSearchNumber)
    
    ExecuteSearch lRangeMin
Exit Sub
ErrorHandler:
    ' Only use SetAbort if this method did an update,
    ' (and then there should have been no transient storage).
    ' Otherwise use SetComplete
    m_objContext.SetComplete
    
    ShowError sSource
End Sub

Public Sub ExecuteSearch(Optional t_lRangeMin As Long = -1)
'   Execute the search, the parameter passed specifies
'   (along with the class constant mk_lResultsToDisplay)
'   which results we are after. If this is -1 then we are
'   getting more results from an existing search so we
'   must reload the search from storage and get the range to
'   to search from from the asp request
'
'   d1  WDP     8-Jul-1999   Created
'   d2  WDP     12-Jul-1999  Added Range parameter
'   d20 WDP     4-Nov-1999   Altered to only do leaflet searches
'                            for basic and index searches
'   d23 JES     19-Nov-1999  Added usage switches
'   d26 JES     2-DEC-1999   Added stats message
'   d39 WDP     25-May-2000  Read # of editorial results to display from FWAppData

On Error GoTo ErrorHandler
    Dim sSource As String
    sSource = "VWReaderSearch.ExecuteSearch"
    Dim oboExec As New BOSearchExecutive
    
    If m_eUsage = ESearchEditorial Then
        Dim ofwAppData As New FWApplicationData
        oboExec.RangeMin = 1
        oboExec.RangeMax = ofwAppData.GetApplicationConfigurationInfomation(pSearchEditorialBound)
    
    ElseIf m_eUsage = ESearchReader Then
    
        If t_lRangeMin = -1 Then
            'If no range min was passed:
            'We need to reload the search and the range from the request
            Dim lRangeMin As Long
            
            'Get the search instance number and build the search from it
            m_sSearchNumber = m_aspRequest(gk_sFORMSearchNumber)
            Set m_boSearch = New BOSearch
            m_boSearch.PROCATTransientInterface_GetState (m_sSearchNumber)
            
                'Get the range to search from from the request and
                'set it for the search executive
                lRangeMin = m_aspRequest(gk_sFORMResultMarker)
                oboExec.RangeMin = lRangeMin
                oboExec.RangeMax = lRangeMin + mk_lResultsToDisplay - 1
                
                'Execute the leaflet search if this is the first page
                'and we are doing an index or basic search
                If lRangeMin < mk_lResultsToDisplay And _
                    (m_boSearch.SearchType = Basic Or m_boSearch.SearchType = Index) Then
                    If m_boSearch.SearchType = Basic Then
                        ExecuteLeafletSearch
                    Else
                        ExecuteLeafletIndexSearch
                    End If
                End If

        Else
            'If the range minimum was passed:
            'We use that to set the range
            oboExec.RangeMin = t_lRangeMin
            oboExec.RangeMax = t_lRangeMin + mk_lResultsToDisplay - 1
            
            'Execute the leaflet search if this is the first page
            'and we are doing an index or basic search
            If lRangeMin < mk_lResultsToDisplay And _
                (m_boSearch.SearchType = Basic Or m_boSearch.SearchType = Index) Then
                If m_boSearch.SearchType = Basic Then
                    ExecuteLeafletSearch
                Else
                    ExecuteLeafletIndexSearch
                End If
            End If
            
            
        End If
    End If
    
    
    'Check we have a search to execute
    If m_boSearch Is Nothing Then
        sSource = sSource & ": No search to run."
        GoTo ErrorHandler
    End If
        
    
    'Execute the search and get the results
    oboExec.Search = m_boSearch
    oboExec.Usage = m_eUsage
    oboExec.RunSearch
    Set m_oiResults = New ISearchResults
    Set m_oiResults = oboExec.Results
    
    'Send a user stats message to stats server
    If m_eUsage = ESearchReader And m_boSearch.SendStats And Not m_boMessage Is Nothing Then
        
        LogStats
        
        'Clear the flag so we don't resend the message
        m_boSearch.SendStats = False

    End If
    
    
    
Exit Sub
ErrorHandler:
    ' Only use SetAbort if this method did an update,
    ' (and then there should have been no transient storage).
    ' Otherwise use SetComplete
    m_objContext.SetComplete
    
    ShowError sSource
End Sub

Public Sub ExecuteLeafletSearch()
'   Performs a verity free text search for leaflets
'   Only runs when no constraints have been specified
'   Only called from the basicsearchresults.asp page , not from
'   moreresults.asp
'
'   d12 JES     19-OCT-1999
'   d38  WDP     24-May-2000     Removed call to ParseVQL - this is
'                               done in BOVerityLeafletSearch

On Error GoTo ErrorHandler
    Dim sSource As String
    sSource = "VWReaderSearch.ExecuteLeafletSearch"
    
    Dim sSearchTerm As String
    Dim sRest As String
    
    ' Get fields from Request
    sSearchTerm = m_aspRequest(gk_sFORMBASICSearchTerm)
    sRest = m_aspRequest(gk_sFORMBASICFirstDate)
    sRest = sRest & m_aspRequest(gk_sFORMBASICLastDate)
    sRest = sRest & m_aspRequest(gk_sFORMBASICRestriction)
    
    'Get the fields from the search object
    'The search term will be in four terms (title, scope-content
    'admin history and arrangement) so we could look for any of these
    'terms - wlog we look for title term
    sSearchTerm = FindInTerm(gk_sDITitle, m_boSearch.RootSearchTerm)
    'Look for a first date term
    sRest = FindInTerm(gk_sDIFirstDate, m_boSearch.RootSearchTerm)
    'Look for a lastdate term
    sRest = sRest & FindInTerm(gk_sDILastDate, m_boSearch.RootSearchTerm)
    
    'Get the breadth restrictions
    Dim l As Long
    For l = 1 To m_boSearch.NumberOfBRestrictions
        If Not sRest = "" Then
            sRest = sRest & ", "
        End If
        sRest = sRest & m_boSearch.GetBreadthRestriction(l)
    Next
    
    
    'If there are any restrictions then don't display
    If Not (sRest = "") Then
        Exit Sub
    End If
    
    'If this is not a basic search then leave
    If Not (m_boSearch.SearchType = Basic) Then
        Exit Sub
    End If
    
    'If there are previous restriction then this is not the highest level of the results
    If m_boSearch.HasOldRestrictions = True Then
        Exit Sub
    End If
    
    
    'Initialise BO and load the leaflet list
    Dim objLeafResults As New BOVerityLeafletResults
    Dim objLeafSearch As New BOVerityLeafletSearch
    Dim sLeafResults() As String
    
    'Execute the leaflet search
    objLeafSearch.ExecuteLeafletSearch sSearchTerm
        
    'Register the search results with the results object
    objLeafResults.ResultCount = objLeafSearch.DocsFound
    objLeafSearch.GetResults sLeafResults
    objLeafResults.Results = sLeafResults
        
    Set m_objLeafResults = objLeafResults
        
    'Destroy the search object
    Set objLeafSearch = Nothing
    
    m_bDisplayLeaflets = True
    
Exit Sub
ErrorHandler:
    ' Only use SetAbort if this method did an update,
    ' (and then there should have been no transient storage).
    ' Otherwise use SetComplete
    m_objContext.SetComplete
    
    ConvertError sSource
    
End Sub


Public Sub ExecuteLeafletIndexSearch()
'   Performs an index search for leaflets
'
'   d20 WDP     4-Nov-1999  Created

On Error GoTo ErrorHandler
    Dim sSource As String
    sSource = "VWReaderSearch.ExecuteLeafletIndexSearch"
    
    Dim lAFTermID As Long
    Dim eTermType As EAFTermType
    Dim sTerm As String
    
    sTerm = FindInTerm(gk_sDIPerson, m_boSearch.RootSearchTerm)
    eTermType = AFPerson
    
    If sTerm = "" Then
        sTerm = FindInTerm(gk_sDIPlace, m_boSearch.RootSearchTerm)
        eTermType = AFPlace
    End If
    
    If sTerm = "" Then
        sTerm = FindInTerm(gk_sDISubject, m_boSearch.RootSearchTerm)
        eTermType = AFSubject
    End If
    
    If sTerm = "" Then
        sTerm = FindInTerm(gk_sDICorporateName, m_boSearch.RootSearchTerm)
        eTermType = AFCorporateBody
    End If
    
    If sTerm = "" Then
        RaiseError sSource, PEGeneralInternalError, "Cannot find search term."
    End If
    
    lAFTermID = CLng(Mid(sTerm, 3, Len(sTerm) - 2))
    
    'Run the search over leaflets
    Dim oboLeafIdxSearch As New BOLeafletIndexSearch
    Dim oboLeafResults As New BODBLeafletResults
    
    oboLeafIdxSearch.Execute lAFTermID, eTermType
    oboLeafResults.Results = oboLeafIdxSearch.Results
    
    Set m_objLeafResults = oboLeafResults
    
    If m_objLeafResults.ResultCount = 0 Then
        Set m_objLeafResults = Nothing
    End If
    
    m_bDisplayLeaflets = True
    
Exit Sub
ErrorHandler:
    ' Only use SetAbort if this method did an update,
    ' (and then there should have been no transient storage).
    ' Otherwise use SetComplete
    m_objContext.SetComplete
    
    ShowError sSource
    
End Sub



Private Function AddLeafSummary() As String
'   d12 JES 19-OCT-1999 Created

    Dim sHTML As String
    If m_objLeafResults.ResultCount = 1 Then
        sHTML = "There was <B>1</B> leaflet that matched your search."
    Else
        sHTML = "There were <B>" & m_objLeafResults.ResultCount & "</B> leaflets that matched your search."
    End If
    
    sHTML = m_fwHTML.Font(sHTML, t_iSize:=2)
    AddLeafSummary = sHTML
    
End Function
Private Function AddLeafList() As String
'   builds html for the leaflet list
'
'   d12  JES     19OCT1999      Created

On Error GoTo ErrorHandler
    Dim sSource As String
    sSource = "VWReaderSearch.AddLeafList"
    
    
    Dim typPinkRow As RowStyle
    Dim typWhiteCell As CellStyle
    Dim typButtonsCell As CellStyle
    Dim typTableStyle As TableStyle
    
    Dim iRowCount As Integer
    Dim iRowsToAdd As Integer
    
    
    'Set table, row and cell styles
    typTableStyle.eALIGN = HALLeft
    '+ remove hardcoded values
    typTableStyle.lCELLPADDING = 1
    typTableStyle.lCELLSPACING = 2
    typTableStyle.sWidth = "100%"
    
    typPinkRow = m_fwHTML.GetRowStyle(RowBackgroundPink)
    typWhiteCell = m_fwHTML.GetCellStyle(CELLWhiteStyle)
    
    m_fwLeafTable.TableStyle = typTableStyle
    m_fwLeafTable.DefaultRowStyle = typPinkRow
    
    'Add the column headers to the table
    AddLeafTableHeader
    
    iRowsToAdd = m_objLeafResults.ResultCount
        
    For iRowCount = 1 To iRowsToAdd
        m_fwLeafTable.StartRow
        
        
        m_fwLeafTable.AddCellAsHTML AddLeafLink(m_objLeafResults.GetFieldValue(gk_sDILeafletTitle), _
                                                m_objLeafResults.GetFieldValue(gk_sDIDBID))
        
        m_fwLeafTable.EndRow
        m_objLeafResults.MoveNext
    Next

       
    AddLeafList = m_fwLeafTable.GetHTML

    Exit Function

ErrorHandler:
    ' Only use SetAbort if this method did an update,
    ' (and then there should have been no transient storage).
    ' Otherwise use SetComplete
    m_objContext.SetComplete
    
    ShowError sSource
    
End Function

Private Function AddLeafHeader() As String
'   Description
'   Returns HTML for Leaflet  header
'
'   d3  JES     05OCT1999      Created

    Dim sHTML As String
    sHTML = m_fwHTML.Bold("information leaflets") & " provide information for readers carrying out research in a particular area. Click the title to see the information leaflet.<P>"
    AddLeafHeader = sHTML
    
    
    
End Function

Private Function AddLeafLink(t_sTitle As String, t_sLeafID As String) As String
'   Description
'   Generates hypertext link to the leaflet loader page from a leaflet title
'   The loader page is passed the id of the row in the leaflet db table
'
'   d12  JES     19OCT1999      Created
'

On Error GoTo ErrorHandler
    Dim sSource As String
    sSource = "VWPSBrowser.AddLeafLink"
    
    Dim sLink As String

    sLink = gk_sPAGERdLeaflet & "?" & gk_sFORMLeafletID & "=" & t_sLeafID
    
    AddLeafLink = m_fwHTML.Hyperlink(sLink, m_fwHTML.Font(t_sTitle, t_iSize:=2), gk_sFRAMELeafletframe)

    Exit Function
ErrorHandler:
    ' Only use SetAbort if this method did an update,
    ' (and then there should have been no transient storage).
    ' Otherwise use SetComplete
    m_objContext.SetComplete
    
    ShowError sSource
End Function

Private Sub AddLeafTableHeader()
'   Adds the leaflet results table headings
'
'   d12  JES     19OCT1999      Created

On Error GoTo ErrorHandler
    
    Dim sSource As String
    sSource = "VWReaderSearch.AddLeafTableHeader"
    Dim typHeaderRow As RowStyle
    
    typHeaderRow = m_fwHTML.GetRowStyle(RowBackgroundRed)
    
    m_fwLeafTable.StartRow typHeaderRow
    
    '+ remove hardcoded column names
    m_fwLeafTable.AddHCellAsHTML m_fwHTML.Font("Matching Leaflets", COLOURWhite, 2)
    
    m_fwLeafTable.EndRow
    
Exit Sub
ErrorHandler:
    ' Only use SetAbort if this method did an update,
    ' (and then there should have been no transient storage).
    ' Otherwise use SetComplete
    m_objContext.SetComplete
    
    ShowError sSource
End Sub

Public Function RefineAdvancedSearch() As String
'   Call back from ASP to refine an advanced search
'   Converts the bosearch into a partial search and saves it
'   then returns the search number.
'
'   d18 JES     3-Nov-1999  Created
On Error GoTo ErrorHandler
    Dim sSource As String
    sSource = "VWReaderSearch.RefineAdvancedSearch"
    
    'Get the search instance number and build the search from it
    m_sSearchNumber = m_aspRequest(gk_sFORMSearchNumber)
    Set m_boSearch = New BOSearch
    m_boSearch.PROCATTransientInterface_GetState (m_sSearchNumber)
    
    'The new partial search we are creating
    Dim oboPartialSearch As New BOPartialSearch
    
    'Convert the search to a partial search object
    oboPartialSearch.DestructSearch m_boSearch
    
    Dim sSearchNumber As String
    sSearchNumber = GetInstanceNumber()
    'Store the EDID
    oboPartialSearch.EditSetId = m_aspRequest(gk_sParamEditSetID)
    
    'Store the partial search
    oboPartialSearch.PROCATTransientInterface_SetState sSearchNumber
    
    'Remove the search object as no longer needed
    'm_boSearch.PROCATTransientInterface_DeleteState m_sSearchNumber
    
    RefineAdvancedSearch = sSearchNumber
    
Exit Function
ErrorHandler:
    m_objContext.SetComplete
    
    ShowError sSource
End Function


Public Function BuildAdvancedSearch() As Boolean
'   Call back from ASP to build a search from a partial advanced
'   search object.
'
'   d14 JES     27-Oct-1999     Created
'   d16 JES     1-Nov-1999      Added restrictions
'   d18 JES     3-Nov-1999      Added sorttype
On Error GoTo ErrorHandler
    Dim sSource As String
    sSource = "VWReaderSearch.BuildAdvancedSearch"
    
    'The partial search to work with
    Dim oboPartialSearch As New BOPartialSearch
    
    'Get the search number
    Dim sSearchNumber As String
    sSearchNumber = m_aspRequest(gk_sFORMAdvSNumber)
  
    'Load the partial search
    oboPartialSearch.PROCATTransientInterface_GetState sSearchNumber
    
    Set m_boSearch = New BOSearch
    
    'Convert the partial search to the root term
    m_boSearch.RootSearchTerm = oboPartialSearch.PartialToRoot()
    
    'Get the sort type
    m_boSearch.SortType = oboPartialSearch.SortType
  
    'Add the search type
    m_boSearch.SearchType = oboPartialSearch.SearchType
    
    '+Set the context - use resource file text
    If m_boSearch.SearchType = Advanced Then
        m_boSearch.Context = " an advanced search"
    ElseIf m_boSearch.SearchType = Field Then
        m_boSearch.Context = " a field search"
    End If
    
    'Set the operator precedence
    m_boSearch.OperatorPrecedence = oboPartialSearch.Precedence
        
    ' Parse the restriction string into search breadth restrictions
    If Not oboPartialSearch.WidthRestriction = "" Then
        ParseRestrictions (oboPartialSearch.WidthRestriction)
    End If
    
    'Set the catalogue levels to search over
    If oboPartialSearch.LettercodeRestriction Then m_boSearch.AddDepthRestriction ctLetterCode
    If oboPartialSearch.DivisionRestriction Then m_boSearch.AddDepthRestriction ctDivision
    If oboPartialSearch.ClassRestriction Then m_boSearch.AddDepthRestriction ctClass
    If oboPartialSearch.HeaderRestriction Then m_boSearch.AddDepthRestriction ctHeaders
    If oboPartialSearch.SubheaderRestriction Then m_boSearch.AddDepthRestriction ctSubHeaders
    If oboPartialSearch.PieceRestriction Then m_boSearch.AddDepthRestriction ctPiece
    If oboPartialSearch.ItemRestriction Then m_boSearch.AddDepthRestriction ctItem
        
    'Create an instance reference for this search
    m_sSearchNumber = GetInstanceNumber()
    
    'Instruct us to send a stats message for this search
    If m_eUsage = ESearchReader Then
        Set m_boMessage = New BOMessage
        m_boSearch.SendStats = True
    End If
    
    ' Store the search transiently - we only need to do this now,
    ' after searches are created or altered.
    m_boSearch.PROCATTransientInterface_SetState m_sSearchNumber
    
    BuildAdvancedSearch = True
    
Exit Function
ErrorHandler:
    ' Only use SetAbort if this method did an update,
    ' (and then there should have been no transient storage).
    ' Otherwise use SetComplete
    m_objContext.SetComplete
    BuildAdvancedSearch = False
    ShowError sSource
End Function

Public Function BuildBasicSearch() As Boolean
'   Build the search from the ASP request
'
'   d1  WDP     8-Jul-1999   Created
'   d5  WDP     29-Jul-1999  Now doesnt search over arrangement zone
'   d11 JES     04-OCT-1999  Set search type
'   d12 JES     19-OCT-1999  Added VQL parser for search term
'   d13 WDP     19-Oct-1999  Moved VQL parsing to BOVerityResults
'                            This means the parsed term is not
'                            kept for refinement, context etc.
'   d26 JES     03-Dec-1999  Set stats message object
'   d29 WDP     20-Jan-2000  Search over map designation
'   d30 WDP     04-Feb-2000  Remove depth restricitons to improve performance
'   d35 WDP     17-Mar-2000  Fixed date defects in BuildBasicSearch
'   d37 DAJ     08-May-2000  Modified Basic Search to use 'yyyy' format for dates
'   d41 DAJ     06-Jun-2000  Fixed date conversion problem in Basic search

On Error GoTo ErrorHandler
    
    'Number of fields in the form that translate into search terms
    'Currently txtSearchTerm, txtFirstDate, txtLastDate
    Const klNumberOfFields As Long = 3
    'Number of zones that the free text search is over
    'Currently Title, Scope-Content, Admin history
    Const klFreeTextSearchZones As Long = 4
    
    Dim sSource As String
    sSource = "VWReaderSearch.BuildBasicSearch"
    
    Dim sSearchTerm As String
    Dim sFirstDate As String
    Dim sLastDate As String
    Dim sRestriction As String
    Dim sSortType As String
    Dim ofwAppData As New FWApplicationData
    
    ' Get fields from Request
    sSearchTerm = m_aspRequest(gk_sFORMBASICSearchTerm)
    sFirstDate = m_aspRequest(gk_sFORMBASICFirstDate)
    sLastDate = m_aspRequest(gk_sFORMBASICLastDate)
    sRestriction = m_aspRequest(gk_sFORMBASICRestriction)
    sSortType = m_aspRequest(gk_sFORMBASICSortType)
    
    'Array to hold search terms
    Dim oiTerms(1 To klNumberOfFields) As ISearchTerm
    'Count of the number of top level terms we create
    Dim lTermsCount As Long
    lTermsCount = 0
    
    ' Free text search term
    If Not (sSearchTerm = "") Then
        
        ' We build a combination term which is a disjunction of
        ' free text terms on various zones
        Dim boCombTerm As New BOCombinationTerm
        Dim boChildTerms(1 To klFreeTextSearchZones) As New BOCriteriaTerm
        
        'Set the comb term type
        boCombTerm.TermType = CTTOR
        
        ' Set the data items (which will map to zones) for
        ' each child term
        boChildTerms(1).Item = gk_sDIScopeContent
        boChildTerms(2).Item = gk_sDITitle
        boChildTerms(3).Item = gk_sDIAdminHistory
        boChildTerms(4).Item = gk_sDIMapDesignation
        
        'Set other child term properties and add to the comb term
        Dim l As Long
        For l = 1 To klFreeTextSearchZones
            With boChildTerms(l)
                .Operator = CTOVerityIN
                .TermType = CTTFreeText
                .Value = sSearchTerm
            End With
            'Add the child term
            boCombTerm.AddTerm boChildTerms(l)
        Next
        
        'Set the free text term and increment the terms count
        Set oiTerms(1) = boCombTerm
        lTermsCount = lTermsCount + 1
        
    'Raise an error if the search term is empty
    Else
        sSource = sSource & ": No search term specified."
        
            RaiseError sSource, PEUSEREmptySearchTermField, _
                ofwAppData.GetDisplayInformation(rrfErrorInvalidFieldSearchTerm)
        
    End If
    
    ' Period restriction
    If Not (sFirstDate = "") Then
    
        ' d37 - Use 'yyyy' format
        If Not (IsYear(sFirstDate)) Then
            sSource = sSource & ": invalid start date."
            
                RaiseError sSource, PEUSERInvalidDateField, _
                ofwAppData.GetDisplayInformation(rrfErrorInvalidFieldFirstDate), """" & sFirstDate & """"
            
            
        End If
            
        Dim boFirstDateTerm As New BOCriteriaTerm
        'Dim lFirstDate As Long
        'lFirstDate = CLng(VariantDateAsYYYYMMDD(sFirstDate))
        
        boFirstDateTerm.TermType = CTTValue
        boFirstDateTerm.Item = gk_sDIFirstDate
        boFirstDateTerm.Value = Format("01/01/" & sFirstDate, gk_sReaderSearchDateFormat)     ' d37 - Add the day/month
        boFirstDateTerm.Operator = CTOGreaterThan
        
        'Set the first date term and increment the terms count
        Set oiTerms(2) = boFirstDateTerm
        lTermsCount = lTermsCount + 1
        
    End If

    If Not (sLastDate = "") Then
    
        ' d37 - Use 'yyyy' format
        If Not (IsYear(sLastDate)) Then
            sSource = sSource & ": invalid end date."
            
                RaiseError sSource, PEUSERInvalidDateField, _
                ofwAppData.GetDisplayInformation(rrfErrorInvalidFieldLastDate), sLastDate
            
        End If
        
        Dim boLastDateTerm As New BOCriteriaTerm
        'Dim lLastDate As Long
        
        'lLastDate = CLng(VariantDateAsYYYYMMDD(sLastDate))
        boLastDateTerm.TermType = CTTValue
        boLastDateTerm.Item = gk_sDILastDate
        boLastDateTerm.Value = Format("31/12/" & sLastDate, gk_sReaderSearchDateFormat)     ' d37 - Add the day/month
        boLastDateTerm.Operator = CTOLessThan
        
        'Set the last date term and increment the terms count
        Set oiTerms(3) = boLastDateTerm
        lTermsCount = lTermsCount + 1
        
    End If
    
    'If we have a first and last date then check first < last
    If Not ((sLastDate = "") Or (sFirstDate = "")) Then
        If CDate(sLastDate) < CDate(sFirstDate) Then
            
                sSource = sSource & ": First date > last date."
                RaiseError sSource, PEUSERInvalidDateRange, _
                ofwAppData.GetDisplayInformation(rrfErrorInvalidDateRange)
            
        End If
    End If
    
    'Set up the search
    Set m_boSearch = New BOSearch
    '+Set the context - use resource file text
    m_boSearch.Context = " a basic search on """ & sSearchTerm & """"
    
    'Set the sort type
    Dim eSort As ESearchSortType
    Select Case sSortType
        Case gk_sFORMSORTCoveringDates:
            eSort = SORTDate
        Case gk_sFORMSORTFormerRef:
            eSort = SORTFormerReference
        Case gk_sFORMSORTReference:
            eSort = SORTReference
        Case gk_sFORMSORTRelevance:
            eSort = SORTRelevance
        Case Else
            sSource = sSource & ": Invalid sort type from request."
            GoTo ErrorHandler
    End Select
    
    m_boSearch.SortType = eSort
    
    'The number of terms dictates what action to take
    Select Case lTermsCount
        Case 0:
            sSource = sSource & ": No search terms set"
            GoTo ErrorHandler
        Case 1:
            Dim lCount As Long
            For lCount = 1 To klNumberOfFields
                'Set as the root term the term that exists
                If Not oiTerms(lCount) Is Nothing Then
                    m_boSearch.RootSearchTerm = oiTerms(lCount)
                End If
            Next
        Case Else
            'Create a combination term to be a conjunct of all
            'the search terms
            Dim boRootTerm As New BOCombinationTerm
            boRootTerm.TermType = CTTAND
            
            Dim lCounter As Long
            For lCounter = 1 To klNumberOfFields
                'Only add terms if they have been set
                If Not oiTerms(lCounter) Is Nothing Then
                    boRootTerm.AddTerm oiTerms(lCounter)
                End If
            Next
            m_boSearch.RootSearchTerm = boRootTerm
    
        End Select
    
    ' Add restrictions
    ' Parse the restriction string into search breadth restrictions
    If Not sRestriction = "" Then
        ParseRestrictions (sRestriction)
    End If
    
    'Set the catalogue levels to search over
    '+ for basic search we should only search class downwards
    '+ but for demo purposes we are searching the whole catalogue
    '+ remove thenext two lines (lettercode and division)
    'm_boSearch.AddDepthRestriction ctLetterCode
    'm_boSearch.AddDepthRestriction ctDivision
    'm_boSearch.AddDepthRestriction ctClass
    'm_boSearch.AddDepthRestriction ctHeaders
    'm_boSearch.AddDepthRestriction ctSubHeaders
    'm_boSearch.AddDepthRestriction ctPiece
    'm_boSearch.AddDepthRestriction ctItem
    
    'Add the search type
    m_boSearch.SearchType = Basic
    
    'Init the message object
    If m_eUsage = ESearchReader Then
        Set m_boMessage = New BOMessage
        m_boSearch.SendStats = True
    End If
    
    'Create an instance reference for this search
    m_sSearchNumber = GetInstanceNumber()
    
    ' Store the search transiently - we only need to do this now,
    ' after searches are created or altered.
    m_boSearch.PROCATTransientInterface_SetState m_sSearchNumber
    
    BuildBasicSearch = True
    
Exit Function
ErrorHandler:
    ' Only use SetAbort if this method did an update,
    ' (and then there should have been no transient storage).
    ' Otherwise use SetComplete
    m_objContext.SetComplete
    BuildBasicSearch = False
    ShowError sSource
End Function

Public Sub GoToCatEntry()
'   Perform a "go to" search. This causes the search to
'   be built and executed using the BORefChecker object
'   If there is > 1 hit then results are displayed listed as
'   for a normal search. Otherwise we redirect to the
'   display cat entry page
'
'   d19  WDP    3-Nov-1999      Created
'   d27  JES    6-Nov-1999      Stats message sent when more than 1 result is returned (results page is displayed)

On Error GoTo ErrorHandler
    
    Dim sSource As String
    sSource = "VWReaderSearch.GoToCatEntry"
    
    Dim sRef As String
    'Get the ref from the request
    sRef = Trim(m_aspRequest.Item(gk_sParamGoToReference))
    
    If sRef = "" Then
        RaiseError sSource, peUSEREmptyReference, "No reference specified."
    End If
    
    Dim oboRefChecker As New BORefChecker
    
    'Check the ref against current PRO and then
    'former departmental refs and PRO refs
    If oboRefChecker.GoToCatEntry(sRef) = True Then
        If oboRefChecker.Results.ResultCount > 1 Or oboRefChecker.Results.ResultType = RESULTTooMany Then
            'If we have > 1 hit display the results as for a normal search
            Set m_oiResults = oboRefChecker.Results
            Set m_boSearch = oboRefChecker.Search
            m_sSearchNumber = GetInstanceNumber()
            m_boSearch.PROCATTransientInterface_SetState m_sSearchNumber
            
            'Send the stats message
            If m_eUsage = ESearchReader Then
                Set m_boMessage = New BOMessage
                LogStats
            End If
            
            'm_oiResults.MoveFirst
            DisplayResults
        Else
            'If we have 1 hit reditect to the cat entry page
            Set m_oiResults = oboRefChecker.Results
            Set m_boSearch = oboRefChecker.Search
            'Send the stats message
            If m_eUsage = ESearchReader Then
                Set m_boMessage = New BOMessage
                LogStats
            End If
            Dim sRedirectURL As String
            'Set the access method as this will not be set by the GetTargetURL method.
            'Since m_boSearch has not been initialised.
            sRedirectURL = GetTargetURL
            m_aspResponse.Redirect sRedirectURL
        End If
    Else
        RaiseError sSource, PEUSERInvalidGoToReference, _
            "Could not find that reference"
    End If
    
    m_objContext.SetComplete
Exit Sub
ErrorHandler:
    ' Only use SetAbort if this method did an update,
    ' (and then there should have been no transient storage).
    ' Otherwise use SetComplete
    m_objContext.SetComplete
    
    ShowError sSource
End Sub

Public Sub BuildPopularSearch(ByVal t_sName As String)
'   Description
'   Builds a popular search
'
'   d10 JES 30SEP1999   Created
'   d11 JES 04OCT1999   Set search type
'   d26 JES 03Dec1999   Added popular search name to stats message

On Error GoTo ErrorHandler
    Dim sSource As String
    sSource = "VWReaderSearch.BuildPopularSearch"
    
    'Load the search into a BOPopularSearch object
    Dim oboPopularSearch As New BOPopSearchLoader
    oboPopularSearch.LoadData t_sName
    
    'Reference the BOPopularSearch search from our BOSearch object
    Set m_boSearch = New BOSearch
    m_boSearch.Deserialize oboPopularSearch.SearchString
    'm_boSearch.SearchType = Popular
    
    If m_eUsage = ESearchReader Then
        Set m_boMessage = New BOMessage
        'Add the search name to stats message
        m_boMessage.AddField EMFPopularSearchID, t_sName
        m_boSearch.SendStats = True
    End If
    
    'Create an instance reference for this search
    m_sSearchNumber = GetInstanceNumber()
    
    ' Store the search transiently - we only need to do this now,
    ' after searches are created or altered.
    m_boSearch.PROCATTransientInterface_SetState m_sSearchNumber
    
    
    
Exit Sub
ErrorHandler:
    ' Only use SetAbort if this method did an update,
    ' (and then there should have been no transient storage).
    ' Otherwise use SetComplete
    m_objContext.SetComplete
    
    ShowError sSource
    
End Sub

Public Sub BuildSavedSearch()
'   Description
'   Builds a saved search
'
'   d22 WDP    11-Nov-1999  Created
'   d26 JES    03-Dec-1999  Added saved search field to the message

On Error GoTo ErrorHandler
    Dim sSource As String
    sSource = "VWReaderSearch.BuildSavedSearch"
    
    Dim lSearchID As Long
    lSearchID = CLng(m_aspRequest.Item(gk_sFORMSavedSearchID))
    
    Set m_boSearch = New BOSearch
    m_boSearch.GetFromSavedSearch lSearchID
    
    'Create an instance reference for this search
    m_sSearchNumber = GetInstanceNumber()
    
    'Add the saved search field to the message
    If m_eUsage = ESearchReader Then
        Set m_boMessage = New BOMessage
        m_boMessage.AddField EMFSavedSearch, True
        m_boSearch.SendStats = True
    End If
    
    ' Store the search transiently - we only need to do this now,
    ' after searches are created or altered.
    m_boSearch.PROCATTransientInterface_SetState m_sSearchNumber
    
    
    
    
Exit Sub
ErrorHandler:
    ' Only use SetAbort if this method did an update,
    ' (and then there should have been no transient storage).
    ' Otherwise use SetComplete
    m_objContext.SetComplete
    
    ShowError sSource
    
End Sub


Public Sub BuildIndexSearch()
'   Description
'   Builds an index search
'
'   d20 WDP     4-Nov-1999      Created

On Error GoTo ErrorHandler
    Dim sSource As String
    sSource = "VWReaderSearch.BuildIndexSearch"
    
    Dim lAFTermID As Long
    Dim eTermType As EAFTermType
    Dim sTermName As String
    
    lAFTermID = m_aspRequest.Item(gk_sFORMAFTermID)
    eTermType = m_aspRequest.Item(gk_sFORMAFGroup)
    sTermName = m_aspRequest(gk_sFORMAdvSIndexName)
        
    'Run the search over the catalogue
    Dim sDataItem As String, sPrefix As String
    
    Select Case eTermType
        Case AFPerson
            sDataItem = gk_sDIPerson
            sPrefix = gk_sPersonTermPrefix
        Case AFPlace
            sDataItem = gk_sDIPlace
            sPrefix = gk_sPlaceTermPrefix
        Case AFSubject
            sDataItem = gk_sDISubject
            sPrefix = gk_sSubjectTermPrefix
        Case AFCorporateBody
            sDataItem = gk_sDICorporateName
            sPrefix = gk_sCorpBodyTermPrefix
    End Select
    
    Dim oboTerm As New BOCriteriaTerm
    oboTerm.TermType = CTTIndex
    oboTerm.Operator = CTOSubstring
    oboTerm.Item = sDataItem
    oboTerm.Value = sPrefix & lAFTermID
    
    Set m_boSearch = New BOSearch
    
    m_boSearch.RootSearchTerm = oboTerm
    
    '+Set the context - use resource file text
    m_boSearch.Context = " an index search on """ & sTermName & """"
    
    'Set the sort type
    m_boSearch.SortType = SORTReference
    
    m_boSearch.SearchType = Index
    
    'Add the index type and id to the message
    If m_eUsage = ESearchReader Then
        Set m_boMessage = New BOMessage
        m_boMessage.AddField EMFIndexType, CStr(eTermType)
        m_boMessage.AddField EMFIndexID, CStr(lAFTermID)
        m_boSearch.SendStats = True
    End If
    
    'Create an instance reference for this search
    m_sSearchNumber = GetInstanceNumber()
    
    ' Store the search transiently - we only need to do this now,
    ' after searches are created or altered.
    m_boSearch.PROCATTransientInterface_SetState m_sSearchNumber
    
    
    
Exit Sub
ErrorHandler:
    ' Only use SetAbort if this method did an update,
    ' (and then there should have been no transient storage).
    ' Otherwise use SetComplete
    m_objContext.SetComplete
    
    ShowError sSource
    
End Sub


''''''''''''''''' Implementation ''''''''''''''''''''''

'
' Private functions
'


Private Function PageStartHTML() As String
'
'
'   d1  WDP     8-Jul-1999      Created (stub)
'   d32 JCC     10-Feb-2000  Added white background

On Error GoTo ErrorHandler
    
    Dim sSource As String
    sSource = "VWReaderSearch.PageStartHTML"
    Dim sHTML As String
    '! temporary for testing
    sHTML = "<BODY BGCOLOR=#FFFFFF>"
    
    PageStartHTML = sHTML
    
Exit Function

ErrorHandler:
    m_objContext.SetComplete
    ShowError sSource
End Function


Private Function HeaderHTML() As String
'   Header info for the results page
'
'   d1  WDP     8-Jul-1999      Created (stub)
'   d6  WDP     12-Aug-1999     Added help button

On Error GoTo ErrorHandler
    
    Dim sSource As String
    sSource = "VWReaderSearch.HeaderHTML"
    
    Dim sHTML As String
    Dim sText As String
    
    'temporary - use text from resource file
    sHTML = "<P><B>Search Results</B></P>"
    
    ' Check for reader usage first
    If m_eUsage = ESearchReader Then
        sHTML = sHTML & m_fwHTML.JavascriptForHelpWindow
        sHTML = sHTML & m_fwHTML.Hyperlink(m_fwHTML.HelpWindowLink(gk_sHelpPagesPath & gk_sPAGESearchResultsHelp), m_fwHTML.Image("", gk_sIMAGEURLPageHelp), "")
    End If

    'Add the help button
    'sHTML = sHTML & m_fwHTML.ImageButton(gk_sIMAGEURLPageHelp, _
    '       t_sURL:=gk_sPAGESearchResultsHelp, t_sTarget:=gk_sFRAMEHelpFrame)
    
    sText = "You ran" & m_boSearch.Context

    'Add restrictions
    Dim lCount As Long
    lCount = m_boSearch.NumberOfBRestrictions
    If Not lCount = 0 Then
        sText = sText & " restricted to reference(s): "
        Dim l As Long
        Dim sRestrictions As String
        For l = 1 To lCount
            If Not sRestrictions = "" Then
                sRestrictions = sRestrictions & ", "
            End If
            sRestrictions = sRestrictions & m_boSearch.GetBreadthRestriction(l)
        Next
        sText = sText & sRestrictions
    End If
    sText = sText & "."
    'd28
    sText = m_fwHTML.EncodeAsHTML(sText)
    HeaderHTML = sHTML & "<P>" & m_fwHTML.Font(sText, t_iSize:=2) & "</P>"
    
Exit Function

ErrorHandler:
    m_objContext.SetComplete
    ShowError sSource
End Function


Private Function ContextHTML() As String
'
'
'   d1  WDP     8-Jul-1999      Created (stub)
'   d3  WDP     13-Jul-1999     Fleshed out
'   d5  WDP     29-Jul-1999     Cosmetic changes
'   d9  WDP     17-Aug-1999     Altered to cope with both List and
'                               summary results and added more detail
'   d40 WDP     31-May-2000     Read results behaviour boundary constants from registry

On Error GoTo ErrorHandler
    
    Dim sSource As String
    sSource = "VWReaderSearch.ContextHTML"
                       
    Dim sHTML As String
    Dim ofwAppData As New FWApplicationData
    
    '+Temporary - use text from resource file
    
    
    If m_oiResults.ResultType = RESULTList Or m_oiResults.ResultType = RESULTTooMany Then
        Dim sSortString As String
        sSortString = " sorted by " & m_boSearch.GetSortString & "."
        Dim lHits As Long
        lHits = m_oiResults.Hits
        If lHits = 0 Then
            sHTML = sHTML & " There were <B>0</B> hits within catalogue entry details."
        End If
        If lHits = 1 Then
            sHTML = sHTML & " There was <B>1</B> hit on the Catalogue.</P>"
        End If
        If m_oiResults.ResultType = RESULTTooMany Then
            Dim lMaxResults As Long
            If m_oiResults.RestrictType = RSTRSingleClass Then
                lMaxResults = ofwAppData.GetApplicationConfigurationInfomation(pSearchSummaryMidBound)
            Else
                lMaxResults = ofwAppData.GetApplicationConfigurationInfomation(pSearchMaxDocs)
            End If
            sHTML = sHTML & " There were at least " & lMaxResults & " hits within catalogue entry details."
        End If
        If (Not m_oiResults.ResultType = RESULTTooMany) And lHits > 1 And lHits < ofwAppData.GetApplicationConfigurationInfomation(pSearchMaxDocs) Then
            sHTML = sHTML & _
                " There were <B>" & m_oiResults.Hits & _
                "</B> hits within catalogue entry details." & _
                " Hits <B>" & m_oiResults.RangeMin & _
                "</B> to <B>" & m_oiResults.RangeMax & _
                "</B> are shown below" & sSortString & ".</P>"
                
        End If
        
    Else
        Dim sRestrictLevel As String
        If m_oiResults.ResultType = RESULTLCodeSummary Then
            sRestrictLevel = "department"
        Else
            sRestrictLevel = "division (where applicable) and series"
        End If
        
        sHTML = sHTML & " There were <B>" & m_oiResults.Hits & _
                "</B> hits within catalogue entry details." & _
                " These hits are summarised below by " & _
                sRestrictLevel & ". " & _
                "Click on a 'hits' link to confine your search to that section of the catalogue."
    End If
    
    sHTML = m_fwHTML.Font(sHTML, t_iSize:=2)
    
    ContextHTML = sHTML & vbCrLf
    
Exit Function

ErrorHandler:
    m_objContext.SetComplete
    ShowError sSource
End Function

Private Sub LogStats()
'd26    JES     2-Dec-1999  Created
On Error GoTo ErrorHandler
    Dim sSource As String
    sSource = "VWReaderSearch.LogStats"
    Dim msmq As New MSMQHelper
    Dim sTime As String
    Dim sUserID As String
    Dim sUserType As String
    'Set the message type
    m_boMessage.MessageType = EMTSearchPerformed
    
    'Get user ID and type from cookie
    sUserID = m_aspRequest.Cookies(gk_sPROCATCookieName).Item(gk_sPROCATCookieUserID)
    sUserType = m_aspRequest.Cookies(gk_sPROCATCookieName).Item(gk_sPROCATCookieUserType)
    
    'Add the required fields to the message
    m_boMessage.AddField EMFTimeStamp, GetTimeStamp()
    m_boMessage.AddField EMFNumberOfResults, m_oiResults.Hits
    m_boMessage.AddField EMFSearchType, m_boSearch.SearchType
    m_boMessage.AddField EMFUserID, sUserID
    m_boMessage.AddField EMFUserType, sUserType
    
    'Send the message
    msmq.SendMessage m_boMessage
    
    'Destroy objects
    Set m_boMessage = Nothing
    Set msmq = Nothing
Exit Sub
ErrorHandler:
    m_objContext.SetComplete
    ShowError sSource
End Sub

Private Function ResultsHTML() As String
'
'
'   d1  WDP     8-Jul-1999      Created (stub)
'   d9  WDP     17-Aug-1999     Altered to cope with both List and
'                               summary results
'   d24 JES     25-Nov-1999     Altered to display all the returned results in the
'                               case of editorial
'   d25 WDP     2-Dec-1999      Added call to AddSaveButton
'   d29 WDP     20-Jan-1999     Altered behaviour if no or too many results
'   d42 DAJ     07-Jun-2000  Added extra help text for when a reference search returns too many results

On Error GoTo ErrorHandler
    
    Dim sSource As String
    sSource = "VWReaderSearch.ResultsHTML"
    
    'HTML for editorial form
    Dim sFormStart As String
    Dim sBadResultsMsg As String
    
    'Results table
    Dim ofwTable As New FWTable
    
    
    
    'Set the table properties
    Dim typDefaultRowStyle As RowStyle
    Dim typDefaultCellStyle As CellStyle
    Dim typTableStyle As TableStyle
    typDefaultRowStyle = m_fwHTML.GetRowStyle(RowBackgroundPink)
    typDefaultCellStyle = m_fwHTML.GetCellStyle(CELLPlainStyle)
    ofwTable.DefaultRowStyle = typDefaultRowStyle
    ofwTable.DefaultCellStyle = typDefaultCellStyle
    
    typTableStyle = m_fwHTML.GetTableStyle(TABPlainStyle)
    typTableStyle.lCELLPADDING = mk_lTablePadding
    typTableStyle.lCELLSPACING = mk_lTableSpacing
    typTableStyle.sWidth = "100%"
    typTableStyle.sWidth = mk_sTableWidth
    
    '+debugging
    Dim lHits As Long
    Dim EResultType As EResultType
    
    lHits = m_oiResults.Hits
    EResultType = m_oiResults.ResultType
    
    
    ofwTable.TableStyle = typTableStyle
    
    'If this is editorial start a form to collect checkbox details
    If m_eUsage = ESearchEditorial Then
        Dim sEditsetParam As String
        Dim sCancelButton As String
        sEditsetParam = "?" & gk_sParamEditSetID & "=" & m_aspRequest(gk_sParamEditSetID)
        sFormStart = m_fwHTML.FormStart(gk_sLinkAddCataloguesToEditset & sEditsetParam _
                                                , gk_sMainFrame, "Test", MPost)

        
        
    End If
    
    If Not (m_oiResults.ResultCount < 1 Or m_oiResults.ResultType = RESULTTooMany) Then
        'Add the column headings
        AddTableHeader ofwTable
        
        'Add the results row by row
        'If this is reader then
        If m_eUsage = ESearchReader Then
            Dim l As Long
            l = 0
            While ((Not m_oiResults.IsAtEnd) And (l < mk_lResultsToDisplay))
                If m_oiResults.ResultType = RESULTList Then
                    AddTableRow ofwTable
                Else
                    AddSummaryTableRow ofwTable
                End If
                m_oiResults.MoveNext
                l = l + 1
            Wend
            
        'If editorial then add all the results
        ElseIf m_eUsage = ESearchEditorial Then
            m_lRowCount = 1
            While Not m_oiResults.IsAtEnd
                AddTableRow ofwTable
                m_lRowCount = m_lRowCount + 1
                m_oiResults.MoveNext
            Wend
            
        End If
        
        'Add refine and prev/next buttons if they are needed
        AddResultButtons ofwTable
       
        'Add the change sort type HTML if needed
        If m_oiResults.ResultType = RESULTList And Not (m_eUsage = ESearchEditorial) Then
            AddChangeSort ofwTable
        End If
        
        'If the search has been restricted from its original state
        'by summary links then we add a button to move back to the
        'previous summary page
        If m_boSearch.HasOldRestrictions Then
            AddUpLevelButton ofwTable
        End If
        
        If m_eUsage = ESearchReader Then
            'Add save search button
            Dim sReaderType As String
            sReaderType = GetDataFromReaderCookie(gk_sPROCATCookieUserType)
        
            If Not sReaderType = "" Then
                If Not (CInt(sReaderType) = EUTAnonymousKew Or _
                        CInt(sReaderType) = EUTAnonymousInternet) Then
                    AddSaveButton ofwTable
                End If
                
            End If
        End If
    
    Else
        If Not (m_boSearch.SearchType = Popular Or _
            m_boSearch.SearchType = Index Or _
            m_boSearch.SearchType = Reference) Then
            sBadResultsMsg = "<P>Please refine your search by clicking the button below.</P>"
            sBadResultsMsg = m_fwHTML.Font(sBadResultsMsg, t_iSize:=2)
        End If
        
        ' d42 - Added extra bit of help if this is a reference search
        If m_boSearch.SearchType = Reference Then
            sBadResultsMsg = "<P>To reduce the number of hits you can refine the search using the " & m_fwHTML.Hyperlink(gk_sPAGEFieldSearchFrameset, "field search", gk_sFRAMEBottomFrame) & ".</P>"
            sBadResultsMsg = m_fwHTML.Font(sBadResultsMsg, t_iSize:=2)
        End If
        
        AddResultButtons ofwTable, False
    End If
    
    ResultsHTML = sBadResultsMsg & ofwTable.GetHTML()
    
    'For editorial add a hidden field with the number of checkboxes stored in it
    'Add the form start and end
    'Also add a cancel button
    If m_eUsage = ESearchEditorial Then
        'Add a cancel button
        sCancelButton = m_fwHTML.ImageButton(gk_sIMAGEURLCancelButton, t_sURL:=gk_sLinkAddCataloguesToEditset & sEditsetParam)

        ResultsHTML = sFormStart & _
                      sCancelButton & _
                      ResultsHTML & _
                      m_fwHTML.HiddenField(mk_sNameCheckBoxCount, CStr(m_lRowCount)) & _
                      "</FORM>"
    End If
        
Exit Function

ErrorHandler:
    m_objContext.SetComplete
    ShowError sSource
End Function




Private Sub AddResultButtons(ByRef t_fwTable As FWTable, Optional t_bShowPrevNext As Boolean = True)
'   Add refine and prev/next buttons to the table passed if necessary
'
'   d2  WDP     13-Jul-1999     Created
'   d9  WDP     17-Aug-1999     Altered to cope with both List and summary results
'   d11 JES     04-Oct-1999     Don't add refine button to popular search
'   d20 WDP     4-Nov-1999      Dont add refine for index search
'   d23 JES     22-Nov-1999     Added usage field
'   d24 JES     25-Nov-1999     Removed prev/next buttons for editorial search
'   d29 WDP     29-Jan-2000     Added t_bShowPrevNext parameter that can
'                               be set to false to not show prev/next for 0 or too many results
'   d31 DTM     09-Feb-2000     Created gk_sImageSelectAll
'   d34 DAJ     09-Mar-2000     Added deselect all button
'   d36 DAJ     05-May-2000  Removed the Refine button for GoTo search results
On Error GoTo ErrorHandler
    
    Dim sSource As String
    sSource = "VWReaderSearch.AddResultButtons"
                       
    
    Dim typStyle As RowStyle
    
    typStyle = m_fwHTML.GetRowStyle(ROWPlainStyle)
    t_fwTable.StartRow typStyle
    
    ' d11/d36 - Don't add refine button for certain types of search
    If Not (m_boSearch.SearchType = Popular Or _
            m_boSearch.SearchType = Index Or _
            m_boSearch.SearchType = Reference) Then
        'Add refine button
        Dim sRefineHtml As String
        Dim sRefineURL As String
        'The URL depends on the search context (basic, field etc.)
        sRefineURL = GetRefineURL(m_boSearch.Context)
        
            sRefineURL = sRefineURL & "?" & _
                    gk_sFORMSearchNumber & "=" & m_sSearchNumber & "&" & _
                    gk_sParamEditSetID & "=" & m_aspRequest(gk_sParamEditSetID)
        
        
        sRefineHtml = m_fwHTML.ImageButton(gk_sIMAGEURLRefineButton, _
                    t_sURL:=sRefineURL)
        t_fwTable.AddCellAsHTML sRefineHtml
    Else
        t_fwTable.AddEmptyCell
    End If
    
    If Not (m_eUsage = ESearchEditorial) And t_bShowPrevNext = True Then
    'If this is not an editorial search then add prev next buttons
        'Add previous and next buttons if necessary
        Dim sPrevHtml As String
        Dim sNextHtml As String
        
        If m_oiResults.RangeMin > 1 Then
            '+ Add prev button
            Dim sPrevURL As String
            sPrevURL = gk_sPAGEMoreResults & "?" & _
                    gk_sFORMResultMarker & "=" & (m_oiResults.RangeMin - mk_lResultsToDisplay) & "&" & _
                    gk_sFORMSearchNumber & "=" & m_sSearchNumber
            sPrevHtml = m_fwHTML.ImageButton(gk_sIMAGEURLPrevButton, _
                        t_sURL:=sPrevURL)
            
        End If
        
        If ((m_oiResults.ResultType = RESULTLCodeSummary Or _
                m_oiResults.ResultType = RESULTClassSummary) And _
            (m_oiResults.RangeMin + mk_lResultsToDisplay - 1) <= m_oiResults.ResultCount) Or _
            (m_oiResults.ResultType = RESULTList And _
            m_oiResults.RangeMax < m_oiResults.Hits) Then
            
            '+ Add next button
            Dim sNextURL As String
            sNextURL = gk_sPAGEMoreResults & "?" & _
                    gk_sFORMResultMarker & "=" & (m_oiResults.RangeMin + mk_lResultsToDisplay) & "&" & _
                    gk_sFORMSearchNumber & "=" & m_sSearchNumber
            sNextHtml = m_fwHTML.ImageButton(gk_sIMAGEURLNextButton, _
                        t_sURL:=sNextURL)
            
        End If
        'Add them together as one cell
        Dim typCellStyle As CellStyle
        typCellStyle = m_fwHTML.GetCellStyle(CELLPlainStyle)
        typCellStyle.eALIGN = HALRight
        t_fwTable.AddCellAsHTML (sPrevHtml & sNextHtml), typCellStyle
        
    'if this is editorial add an OK button and a select/deselect all button
    Else
        If m_eUsage = ESearchEditorial Then
            Dim sSubmit As String
            sSubmit = m_fwHTML.SubmitImage(gk_sImagesPath & gk_sImageOK, "OK", "OK")
            
            'Select all button is script that checks all the checkboxes
            ' d34 - Deselect all button is similar thing
            Dim sSelectScript As String, sDeselectScript As String
            Dim sSelectAll As String, sDeselectAll As String
            Dim iCount As Integer
            
            For iCount = 1 To m_lRowCount - 1
                sSelectScript = sSelectScript & mk_sCheckboxName & iCount & ".checked=true;"
                sDeselectScript = sDeselectScript & mk_sCheckboxName & iCount & ".checked=false;"
            Next
            
            sSelectAll = m_fwHTML.ImageClick(gk_sImagesPath & gk_sImageSelectAll, sSelectScript)
            sDeselectAll = m_fwHTML.ImageClick(gk_sImagesPath & gk_sImageDeselectAll, sDeselectScript)
                                
            Dim typSubCellStyle As CellStyle
            typSubCellStyle = m_fwHTML.GetCellStyle(CELLPlainStyle)
            
            ' d34 - Split up the submit and select/deselect buttons into separate columns
            t_fwTable.AddCellAsHTML (sSubmit), typSubCellStyle
            typSubCellStyle.eALIGN = HALRight
            typSubCellStyle.lCOLSPAN = 3  ' 5 cols - refine button - OK button = 3
            t_fwTable.AddCellAsHTML (sSelectAll & sDeselectAll), typSubCellStyle
        End If
    End If
    
    t_fwTable.EndRow
    
Exit Sub

ErrorHandler:
    m_objContext.SetComplete
    ShowError sSource
End Sub


Private Sub AddUpLevelButton(ByRef t_fwTable As FWTable)
'   Add button to return to summary screen
'
'   d9  WDP     13-Jul-1999     Created
'   d23 JES     22-Nov-1999     Added usage
On Error GoTo ErrorHandler
    
    Dim sSource As String
    sSource = "VWReaderSearch.AddUpLevelButton"
    Dim sHTML As String
                       
    Dim typStyle As RowStyle
    typStyle = m_fwHTML.GetRowStyle(ROWPlainStyle)
    t_fwTable.StartRow typStyle
    
    Dim sURL As String
    'The URL depends on the search context (basic, field etc.)
    sURL = gk_sPAGEBackToSummary & "?" & _
            gk_sFORMSearchNumber & "=" & m_boSearch.OldSearchNumber
    sHTML = m_fwHTML.ImageButton(gk_sIMAGEURLUpLevelButton, _
                t_sURL:=sURL)
    t_fwTable.AddCellAsHTML sHTML
    
Exit Sub

ErrorHandler:
    m_objContext.SetComplete
    ShowError sSource
End Sub

Private Sub AddSaveButton(ByRef t_fwTable As FWTable)
'   Add button to save search
'
'   d25     WDP     2-Dec-1999      Created
On Error GoTo ErrorHandler
    
    Dim sSource As String
    sSource = "VWReaderSearch.AddSaveButton"
    Dim sHTML As String
                       
    Dim typStyle As RowStyle
    typStyle = m_fwHTML.GetRowStyle(ROWPlainStyle)
    t_fwTable.StartRow typStyle
    
    Dim sURL As String
    
    sURL = gk_sPAGESaveSearch & "?" & _
            gk_sFORMSearchNumber & "=" & m_sSearchNumber
    sHTML = m_fwHTML.ImageButton(gk_sIMAGEURLSaveSearch, _
                t_sURL:=sURL)
    t_fwTable.AddCellAsHTML sHTML
    
Exit Sub

ErrorHandler:
    m_objContext.SetComplete
    ShowError sSource
End Sub

Private Sub AddChangeSort(ByRef t_fwTable As FWTable)
'   Add the  re-sort options to the results
'
'   d4  WDP     13-Jul-1999     Created
'   d23 JES     22-Nov-1999     Added usage to url
On Error GoTo ErrorHandler
    
    Dim sSource As String
    sSource = "VWReaderSearch.AddChangeSort"
                       
    
    Dim sHTML As String
    Dim colReSortText As New Collection
    Dim sReSortString As String
    Dim sResortURL As String
    Dim typCellStyle As CellStyle
    
    'Part of the url to go to with the sort type missing
    sResortURL = gk_sPAGESortResults & "?" & _
            gk_sFORMSearchNumber & "=" & m_sSearchNumber & "&" & _
            gk_sFORMBASICSortType & "="
    
    sReSortString = "Sort results "
    
    'Add all the possible sort types as hyperlink strings
    colReSortText.Add m_fwHTML.Hyperlink( _
        (sResortURL & gk_sFORMSORTCoveringDates), _
        (sReSortString & "by covering dates."), gk_sFRAMEBottomFrame)
    colReSortText.Add m_fwHTML.Hyperlink( _
        (sResortURL & gk_sFORMSORTFormerRef), _
        (sReSortString & "by former reference."), gk_sFRAMEBottomFrame)
    colReSortText.Add m_fwHTML.Hyperlink( _
        (sResortURL & gk_sFORMSORTReference), _
        (sReSortString & "by catalogue reference."), gk_sFRAMEBottomFrame)
    colReSortText.Add m_fwHTML.Hyperlink( _
        (sResortURL & gk_sFORMSORTRelevance), _
        (sReSortString & "using relevance ranking."), gk_sFRAMEBottomFrame)
    
    'Remove the link for the sort type we have just done
    Select Case m_boSearch.SortType
        Case SORTDate:
            colReSortText.Remove (1)
        Case SORTFormerReference:
            colReSortText.Remove (2)
        Case SORTReference:
            colReSortText.Remove (3)
        Case SORTRelevance:
            colReSortText.Remove (4)
        Case Else
            sSource = sSource & ": Invalid sort type."
            GoTo ErrorHandler
    End Select
    
    typCellStyle = m_fwHTML.GetCellStyle(CELLWhiteStyle)
    typCellStyle.eVALIGN = VALMiddle
    typCellStyle.lCOLSPAN = 3
    
    Dim sURL
    'Add the links to the table
    For Each sURL In colReSortText
        t_fwTable.StartRow
        t_fwTable.AddCellAsHTML m_fwHTML.Font(CStr(sURL), t_iSize:=2), typCellStyle
        t_fwTable.EndRow
    Next
    
    'Removed old display method - too clumsy looking
    
    'Dim sHTML As String
    'Dim sSortButtonHTML As String
    'Dim sFormHTML As String
    'Dim sSortResultsURL As String
    'Dim typCellStyle As CellStyle
    
    'sSortResultsURL = gk_sPAGESortResults
    
    'sSortButtonHTML = m_fwHTML.FormStart(sSortResultsURL, "main", t_eMethod:=MGet)
    'sSortButtonHTML = sSortButtonHTML & _
        m_fwHTML.SubmitImage(gk_sIMAGEURLSearchButton, "AForm", t_sValue:="GO")
               
    '+ move to resource file
    'sHTML = "You can select an alternative sorting method for your results."
    'sHTML = sHTML & "<BR>" & Radio(gk_sFORMBASICSortType, gk_sFORMSORTRelevance, "Relevance ranking", True)
    'sHTML = sHTML & "<BR>" & Radio(gk_sFORMBASICSortType, gk_sFORMSORTReference, "Sort by catalogue reference")
    'sHTML = sHTML & "<BR>" & Radio(gk_sFORMBASICSortType, gk_sFORMSORTFormerRef, "Sort by former reference")
    'sHTML = sHTML & "<BR>" & Radio(gk_sFORMBASICSortType, gk_sFORMSORTCoveringDates, "Sort by date")
    'sHTML = sHTML & m_fwHTML.HiddenField(gk_sFORMSearchNumber, m_sSearchNumber)
    'sHTML = sHTML & m_fwHTML.HiddenField("feck", m_sSearchNumber)
    'sHTML = m_fwHTML.Font(sHTML, t_iSize:=2) & "</FORM>"
    
    'typCellStyle = m_fwHTML.GetCellStyle(CELLWhiteStyle)
    'typCellStyle.eVALIGN = VALMiddle
    
    't_fwTable.StartRow
    't_fwTable.AddCellAsHTML sSortButtonHTML, typCellStyle
    't_fwTable.AddCellAsHTML sHTML
    't_fwTable.EndRow
    
Exit Sub

ErrorHandler:
    m_objContext.SetComplete
    ShowError sSource
End Sub


Private Function NoResultsHTML() As String
'
'
'   d1  WDP     8-Jul-1999      Created (stub)
'   d18 JES     3-Nov-1999      Added refine button when this is an advanced search.

On Error GoTo ErrorHandler
    
    Dim sSource As String
    sSource = "VWReaderSearch.NoResultsHTML"
    '+temporary
    NoResultsHTML = "No Results - try a different search"
Exit Function

ErrorHandler:
    m_objContext.SetComplete
    ShowError sSource
End Function

Private Function TooManyResultsHTML() As String
'   HTML to display if the search brouht back too many results
'
'   d9  WDP     13-Jul-1999      Created

On Error GoTo ErrorHandler
    
    Dim sSource As String
    sSource = "VWReaderSearch.TooManyResultsHTML"
    '+temporary
    TooManyResultsHTML = "Too many results"
Exit Function

ErrorHandler:
    m_objContext.SetComplete
    ShowError sSource
End Function

Private Function PageEndHTML() As String
'
'
'   d1  WDP     8-Jul-1999      Created (stub)

On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim sHTML As String
    
    sSource = "VWReaderSearch.PageEndHTML"

    If m_eUsage = ESearchEditorial Then
        sHTML = m_fwHTML.JavascriptForHelpWindow
        sHTML = sHTML & m_fwHTML.Hyperlink(m_fwHTML.HelpWindowLink(gk_sHelpPagesPath & gk_sPAGESearchResultsHelp), m_fwHTML.Image("", gk_sIMAGEURLPageHelp), "")
    End If

    
Exit Function

ErrorHandler:
    m_objContext.SetComplete
    ShowError sSource
End Function


Private Sub AddTableHeader(ByRef t_fwTable As FWTable)
'
'
'   d1  WDP     8-Jul-1999      Created (stub)
'   d9  WDP     13-Aug-1999     Added code to deal with summary results
'   d24 JES     25-Nov-1999     Added switch for reader or editorial headers
'   d33 WDP     22-Feb-2000  Added dates in search results for editorial

On Error GoTo ErrorHandler
    
    Dim sSource As String
    sSource = "VWReaderSearch.PageEndHTML"
    
    Dim typHeaderStyle As RowStyle
    
    'Set the header row style
    typHeaderStyle = m_fwHTML.GetRowStyle(RowBackgroundRed)
    
    'Override the table default rowstyle with the
    'header style for this row
    t_fwTable.StartRow typHeaderStyle
    'We force this col to be wide enough for all
    'references to fit on one line
    Dim typCellStyle As CellStyle
    typCellStyle = m_fwHTML.GetCellStyle(CELLPlainStyle)
    typCellStyle.sWidth = "20%"
    'Add the correct headings for this table
    If m_eUsage = ESearchReader Then
    'For reader add these headings
        t_fwTable.AddCellAsHTML "<B>" & m_fwHTML.Font _
                (m_fwHTML.EncodeAsHTML(gk_sDISPLAYRefHeader), COLOURWhite, 2), typCellStyle
        t_fwTable.AddCellAsHTML "<B>" & m_fwHTML.Font _
                (m_fwHTML.EncodeAsHTML(gk_sDISPLAYTitleHeader), COLOURWhite, 2)
        t_fwTable.AddCellAsHTML "<B>" & m_fwHTML.Font _
                (m_fwHTML.EncodeAsHTML(gk_sDISPLAYDatesHeader), COLOURWhite, 2)
        'If we have a results summary we display a column for hits
        If m_oiResults.ResultType = RESULTLCodeSummary Or _
            m_oiResults.ResultType = RESULTClassSummary Then
            t_fwTable.AddCellAsHTML "<B>" & m_fwHTML.Font _
                    (m_fwHTML.EncodeAsHTML(gk_sDISPLAYHitsHeader), COLOURWhite, 2)
            
        Else
            'Add a column for former ref if thats what we sorted by
            '(not needed for summary results)
            If m_boSearch.SortType = SORTFormerReference Then
                t_fwTable.AddCellAsHTML "<B>" & m_fwHTML.Font _
                    (m_fwHTML.EncodeAsHTML(gk_sDISPLAYFormerRefHeader), COLOURWhite, 2)
            End If
        
        End If
        
    ElseIf m_eUsage = ESearchEditorial Then
    'For editorial add these headings
    
        'Reference heading
        t_fwTable.AddCellAsHTML "<B>" & m_fwHTML.Font _
                (m_fwHTML.EncodeAsHTML(gk_sDISPLAYRefHeader), COLOURWhite, 2), typCellStyle
        'Levels heading
        t_fwTable.AddCellAsHTML "<B>" & m_fwHTML.Font _
                (m_fwHTML.EncodeAsHTML(gk_sDISPLAYLevels), COLOURWhite, 2)
        'Title heading
        t_fwTable.AddCellAsHTML "<B>" & m_fwHTML.Font _
                (m_fwHTML.EncodeAsHTML(gk_sDISPLAYTitleHeader), COLOURWhite, 2)
        'Dates
        t_fwTable.AddCellAsHTML "<B>" & m_fwHTML.Font _
                (m_fwHTML.EncodeAsHTML(gk_sDISPLAYDatesHeader), COLOURWhite, 2)
        'Checkbox header
        t_fwTable.AddCellAsHTML "<B>" & m_fwHTML.Font _
            (m_fwHTML.EncodeAsHTML(gk_sDISPLAYAddHeader), COLOURWhite, 2)
    
    End If
    
    t_fwTable.EndRow
    
Exit Sub

ErrorHandler:
    m_objContext.SetComplete
    ShowError sSource
End Sub


Private Sub AddTableRow(t_fwTable As FWTable)
'
'
'   d1  WDP     8-Jul-1999      Created
'   d3  WDP     14-Jul-1999     Made treatment of lettercode refs different
'   d7  SPB     12-Aug-1999     Change target of link in AddTableRow
'                               to be DisplayCatDetailsSearch.asp?<ref parts>
'   d9  WDP     13-Aug-1999     Altered to add hits column for summary results
'   d24 JES     25-Nov-1999     Adds a checkbox for the editorial version
'   d33 WDP     22-Feb-2000  Added dates in search results for editorial


On Error GoTo ErrorHandler
    
    Dim sSource As String
    sSource = "VWReaderSearch.AddTableRow"
    
    Dim sTargetURL As String
    'The data to display
    Dim sReference As String
    Dim sTitleScope As String
    Dim sDates As String
    Dim sFormerRef As String
    
    'Construct the reference to display
    sReference = GetReference()
    
    'If we have a Lettercode or class ref then we want it bold and bigger
    Select Case m_oiResults.GetData(gk_sDICatalogueLevel)
        Case ctLetterCode:
            sReference = m_fwHTML.Font(m_fwHTML.Strong(sReference), t_iSize:=3)
        Case ctClass, ctDivision:
            sReference = m_fwHTML.Font(m_fwHTML.Strong(sReference), t_iSize:=2)
        Case Else
            sReference = m_fwHTML.Font(sReference, t_iSize:=2)
    End Select
    
    t_fwTable.StartRow
    
    If m_eUsage = ESearchReader Then
    'Add the reader row
    
        'Add the reference
        t_fwTable.AddCellAsHTML sReference
        
        'Build target URL containing reference information
        sTargetURL = GetTargetURL()
        
        'Encode the title as HTML display text (deal with special characters like < and >)
        sTitleScope = m_fwHTML.EncodeAsHTML(m_oiResults.GetData(gk_sDITitleOrScope))
        
        'If we have a lettercode ref we want it bold
        If m_oiResults.GetData(gk_sDICatalogueLevel) = ctLetterCode Then
            sTitleScope = m_fwHTML.Strong(sTitleScope)
        End If
        
        'Make the title a hyperlink to the reference
        'Target frame = bottom
        sTitleScope = m_fwHTML.Hyperlink(sTargetURL, t_sText:=m_fwHTML.Font _
            (sTitleScope, t_iSize:=2), t_sTarget:=gk_sFRAMEBottomFrame)
            
        'Add the Title/Scope and content
        t_fwTable.AddCellAsHTML sTitleScope
            
        'Encode the dats as HTML to allow for any special chars in them (< > etc)
        sDates = m_fwHTML.Font(m_fwHTML.EncodeAsHTML _
                (m_oiResults.GetData(gk_sDICoveringDates)), t_iSize:=2)
        
        'Add the covering dates
        t_fwTable.AddCellAsHTML sDates
        
        If m_oiResults.ResultType = RESULTLCodeSummary Or _
            m_oiResults.ResultType = RESULTClassSummary Then
            'Add the hits info.
            Dim sHits As String
            sHits = m_fwHTML.Font(m_fwHTML.EncodeAsHTML _
                (m_oiResults.GetData(gk_sDIHits)), t_iSize:=2)
            '+TO DO Add this as a hyperlink to refine the search
            t_fwTable.AddCellAsHTML sHits
            
        Else
            'Add a col with the former ref only if we sort by it
            'We dont want this in the summary
            If m_boSearch.SortType = SORTFormerReference Then
                sFormerRef = m_fwHTML.Font(m_fwHTML.EncodeAsHTML _
                    (m_oiResults.GetData(gk_sDIFormerRef)), t_iSize:=2)
                t_fwTable.AddHCellAsHTML sFormerRef
                
            End If
        End If
    
    'Add the editorial rows
    ElseIf m_eUsage = ESearchEditorial Then
        'Add the reference
        t_fwTable.AddCellAsHTML sReference
        'Add the levels
        t_fwTable.AddCellAsHTML m_fwHTML.EncodeAsHTML(GetCatalogueLevelDescription(m_oiResults.GetData(gk_sDICatalogueLevel)))
        'Add the title/scope
        t_fwTable.AddCellAsHTML m_fwHTML.EncodeAsHTML(m_oiResults.GetData(gk_sDITitleOrScope))
        'Encode the dats as HTML to allow for any special chars in them (< > etc)
        sDates = m_fwHTML.Font(m_fwHTML.EncodeAsHTML _
                (m_oiResults.GetData(gk_sDICoveringDates)), t_iSize:=2)
        
        'Add the covering dates
        t_fwTable.AddCellAsHTML sDates
        
        'Add the checkbox field
        'Store in the value field the catalogue ID and Catalogue level values
        Dim sValue As String
        sValue = CStr(DBIDFromDocKey(m_oiResults.GetData(gk_sDIVerityDocKey), m_oiResults.GetData(gk_sDICatalogueLevel))) & gk_sDataDelimiter & _
                CStr(m_oiResults.GetData(gk_sDICatalogueLevel))
                
        sValue = m_fwHTML.Checkbox("", mk_sCheckboxName & Trim(CStr(m_lRowCount)), False, sValue)
        
        t_fwTable.AddCellAsHTML sValue
    End If
    
    t_fwTable.EndRow
    
Exit Sub

ErrorHandler:
    m_objContext.SetComplete
    ShowError sSource
End Sub


Private Sub AddSummaryTableRow(t_fwTable As FWTable)
'   Add a row of summary results to the table
'
'   d9  WDP     13-Aug-1999     Created


On Error GoTo ErrorHandler
    
    Dim sSource As String
    sSource = "VWReaderSearch.AddSummaryTableRow"
    'The URL to link to if there is a hit at the summary level
    Dim sTargetURL As String
    'The URL to link to search within a summary level
    Dim sHitsURL As String
    'The data to display
    Dim sReference As String
    Dim sTitleScope As String
    Dim sDates As String
    Dim sLetterCode As String
    Dim sDivision As String
    Dim sClass As String
    Dim iLevel As Integer
    
    sReference = GetReference()
    
    sLetterCode = m_oiResults.GetData(gk_sDILetterCodeRef)
    sDivision = m_oiResults.GetData(gk_sDIDivisionRef)
    sClass = m_oiResults.GetData(gk_sDIClassRef)
    iLevel = CInt(m_oiResults.GetData(gk_sDICatalogueLevel))
    
    sHitsURL = gk_sPAGESearchWithinHits & "?" & _
                gk_sFORMLettercodeRef & "=" & sLetterCode & "&" & _
                gk_sFORMDivisionRef & "=" & sDivision & "&" & _
                gk_sFORMClassRef & "=" & sClass & "&" & _
                gk_sFORMSearchNumber & "=" & m_sSearchNumber
    
    'Construct the reference to display
    sReference = m_fwHTML.Font(m_fwHTML.Strong(sReference), t_iSize:=2)
    
    If m_oiResults.ResultType = RESULTClassSummary Then
    'tab in classes and divisions for viewing clarity
        sReference = gk_sHTMLNonBreakSpace & sReference
        If Not sClass = "" Then
            sReference = gk_sHTMLNonBreakSpace & sReference
        Else
        
        End If
    End If
    
    t_fwTable.StartRow
    'Add the reference
    'We force this column to be a certain size
    'so the reference fits in one line
    Dim typCellStyle As CellStyle
    typCellStyle = m_fwHTML.GetCellStyle(CELLPlainStyle)
    typCellStyle.sWidth = "20%"
    t_fwTable.AddCellAsHTML sReference, typCellStyle
    
    'If the summary entry was actually a search hit itself
    'then we display the title as a link to the entry and
    'dates etc.
    If Not SummaryEntryIsHit() Then
        m_oiResults.GetTitleFromDB
    End If
    
    'Build target URL containing reference information
    sTargetURL = GetTargetURL()
    
    'Encode the title as HTML display text (deal with special characters like < and >)
    sTitleScope = m_fwHTML.Font(m_fwHTML.EncodeAsHTML(m_oiResults.GetData(gk_sDITitleOrScope)), _
                        t_iSize:=2)
    
    'Make the title a hyperlink to the reference
    'Target frame = bottom. We only do this if we
    ' have a hit at this level. otherwose the title is plain text
    If SummaryEntryIsHit() Then
        sTitleScope = m_fwHTML.Hyperlink(sTargetURL, t_sText:=sTitleScope, t_sTarget:=gk_sFRAMEBottomFrame)
    End If
    
    'Add the Title/Scope and content
    t_fwTable.AddCellAsHTML sTitleScope
        
    'Encode the dats as HTML to allow for any special chars in them (< > etc)
    sDates = m_fwHTML.Font(m_fwHTML.EncodeAsHTML _
            (m_oiResults.GetData(gk_sDICoveringDates)), t_iSize:=2)
    
    'Add the covering dates
    t_fwTable.AddCellAsHTML sDates
                
    'Add the hits info. The total hits
    ' for divisions are not shown as a link
    ' as they are a sum of the hits for classes
    ' listed below them
    Dim sHits As String
    Dim lHits As Long
    lHits = CLng(m_oiResults.GetData(gk_sDIHits))
    
    sHits = m_fwHTML.Font(m_fwHTML.EncodeAsHTML(CStr(lHits)), t_iSize:=2)
    
    If (Not iLevel = CInt(ECATALOGUELEVEL.ctDivision)) And _
        Not (iLevel = CInt(ECATALOGUELEVEL.ctLetterCode) And m_oiResults.ResultType = RESULTClassSummary) And _
        Not (SummaryEntryIsHit() And lHits = 1) Then
        sHits = m_fwHTML.Hyperlink(sHitsURL, sHits, gk_sFRAMEBottomFrame)
    End If
    
    t_fwTable.AddCellAsHTML sHits
        
    t_fwTable.EndRow
    
Exit Sub

ErrorHandler:
    m_objContext.SetComplete
    ShowError sSource
End Sub


Private Function GetReference() As String
'   Creates the PRO reference from the individual parts:
'   Lettercode, Division, Class, Header, Subheader, Piece and Item
'   using the CreateReference glocal method
'
'   d1  WDP     9-Jul-1999      Created
'   d2  WDP     12-Jul-1999     Now use CreateReference method
'   d9  WDP     17-Aug-1999     Cope with whether we are displaying
'                               a summary or not

On Error GoTo ErrorHandler
    
    Dim sSource As String
    sSource = "VWReaderSearch.GetReference"
    Select Case m_oiResults.ResultType
        Case RESULTList
            GetReference = CreateReference(VariantStringAsString(m_oiResults.GetData(gk_sDILetterCodeRef)), _
                                    VariantStringAsString(m_oiResults.GetData(gk_sDIDivisionRef)), _
                                    VariantStringAsString(m_oiResults.GetData(gk_sDIClassRef)), _
                                    VariantStringAsString(m_oiResults.GetData(gk_sDIHeaderRef)), _
                                    VariantStringAsString(m_oiResults.GetData(gk_sDISubheaderRef)), _
                                    VariantStringAsString(m_oiResults.GetData(gk_sDIPieceRef)), _
                                    VariantStringAsString(m_oiResults.GetData(gk_sDIItemRef)), _
                                    CInt(m_oiResults.GetData(gk_sDICatalogueLevel)), _
                                    VariantStringAsString(m_oiResults.GetData(gk_sDISubclassRef)))
        Case RESULTLCodeSummary
            GetReference = m_oiResults.GetData(gk_sDILetterCodeRef)
        Case RESULTClassSummary
            GetReference = CreateReference(m_oiResults.GetData(gk_sDILetterCodeRef), _
                                m_oiResults.GetData(gk_sDIDivisionRef), _
                                m_oiResults.GetData(gk_sDIClassRef), _
                                t_eCatLevel:=m_oiResults.GetData(gk_sDICatalogueLevel), _
                                t_sSubClass:=VariantStringAsString(m_oiResults.GetData(gk_sDISubclassRef)))
        Case Else
            RaiseError sSource, PEGeneralInternalError, "Invalid result type."
    End Select
    

Exit Function

ErrorHandler:
    m_objContext.SetComplete
    ShowError sSource
End Function


Private Function GetTargetURL(Optional t_eSearchType As ESearchType) As String
'   Creates the URL to display cat entry details from the individual parts:
'   Lettercode, Division, Class, Header, Subheader, Piece and Item
'
'   d7  SPB     12-Aug-1999     Created
'   d19 WDP     3-Nov-1999      Altered to use the DB ID,cat
'                               level and highlight info returned
'                               by Verity

On Error GoTo ErrorHandler
    
    Dim sSource As String
    sSource = "VWReaderSearch.GetTargetURL"
    
    Dim lID As Long, eLevel As ECATALOGUELEVEL, sHighlightInfo As String
    Dim sDocKey As String
    sDocKey = m_oiResults.GetData(gk_sDIVerityDocKey)
    lID = DBIDFromDocKey(sDocKey, eLevel)
    sHighlightInfo = URLEncode(m_oiResults.GetData(gk_sDIHighlightInfo))
    
    GetTargetURL = gk_sLinkDisplayCatalogueDetails & "?" & _
        gk_sParamCatalogueID & "=" & lID & _
        "&" & gk_sParamCatalogueLevel & "=" & eLevel & _
        "&" & gk_sParamHighlightInfo & "=" & sHighlightInfo & _
        "&" & gk_sParamDetailsTabNumber & "=" & ETabType.etTabSummary
    
    'Add the Access method for the details page.
    'This allows a stats message to be sent to work out most popular
    'search method, etc
    If m_boSearch.SearchType = Basic Then
        GetTargetURL = GetTargetURL & "&" & gk_sParamAccessMethod & "=" & EAccessMethod.EAMBasicSearch
    ElseIf m_boSearch.SearchType = Index Then
        GetTargetURL = GetTargetURL & "&" & gk_sParamAccessMethod & "=" & EAccessMethod.EAMIndexSearch
    ElseIf m_boSearch.SearchType = Field Then
        GetTargetURL = GetTargetURL & "&" & gk_sParamAccessMethod & "=" & EAccessMethod.EAMFieldSearch
    ElseIf m_boSearch.SearchType = Advanced Then
        GetTargetURL = GetTargetURL & "&" & gk_sParamAccessMethod & "=" & EAccessMethod.EAMAdvancedSearch
    ElseIf m_boSearch.SearchType = Reference Then
        GetTargetURL = GetTargetURL & "&" & gk_sParamAccessMethod & "=" & EAccessMethod.EAMReferenceSearch
    End If
    
    'Add the search levels parameter string
    Dim typLevels As CatReferenceParts
    typLevels.sLettercodeRef = VariantStringAsString(m_oiResults.GetData(gk_sDILetterCodeRef))
    typLevels.sDivisionRef = VariantStringAsString(m_oiResults.GetData(gk_sDIDivisionRef))
    typLevels.sClassRef = VariantStringAsString(m_oiResults.GetData(gk_sDIClassRef))
    typLevels.sSubclassRef = VariantStringAsString(m_oiResults.GetData(gk_sDISubclassRef))
    typLevels.sHeaderRef = VariantStringAsString(m_oiResults.GetData(gk_sDIHeaderRef))
    typLevels.sSubheaderRef = VariantStringAsString(m_oiResults.GetData(gk_sDISubheaderRef))
    typLevels.sPieceRef = VariantStringAsString(m_oiResults.GetData(gk_sDIPieceRef))
    typLevels.sItemRef = VariantStringAsString(m_oiResults.GetData(gk_sDIItemRef))
    
    GetTargetURL = GetTargetURL & "&" & _
                gk_sParamEncodedRef & "=" & URLEncode(GetEncodedStringFromReference(typLevels))
Exit Function

ErrorHandler:
    m_objContext.SetComplete
    ShowError sSource
End Function


Private Sub ParseRestrictions(t_sRestrictions)
'   Parses the restrictions string (from the basic search form) and
'   adds breadth restrictions to the search. Note this operates under
'   the assumption that the restrictions are a comma separated list
'   of strings which themselves do not contain any commas.
'
'   d2  WDP     12-Jul-1999     Created
'   d19 WDP     3-Nov-1999      Altered to do proper restriction
'                               checking using BORefChecker
'   d29 WDP     20-Jan-2000     Change all restrictions to upper case before checking

On Error GoTo ErrorHandler
    
    Dim sSource As String
    sSource = "VWReaderSearch.ParseRestrictions"
        
    Dim sRestrictions As String
    sRestrictions = t_sRestrictions
    Dim colRestrs As New Collection
        
    'Loop round removing the next restriction from the string
    While Not sRestrictions = ""
        
        Dim lPos As Long
        Dim sRestr As String
        
        'Find the next comma
        lPos = InStr(sRestrictions, ",")
        If lPos > 0 Then
            'remove the next restriction from the string
            sRestr = UCase(Trim(Left$(sRestrictions, lPos - 1)))
            sRestrictions = Mid$(sRestrictions, lPos + 1)
        Else
            'If there are no commas left just use what remains
            sRestr = UCase(Trim(sRestrictions))
            sRestrictions = ""
        End If
        
        colRestrs.Add sRestr
    
    Wend
    
    Dim oboRefChecker As New BORefChecker
    If oboRefChecker.VerifyReference(colRestrs) Then
        'Add the restricitons to the search if they all
        'check out
        Dim l As Long
        For l = 1 To colRestrs.Count
            m_boSearch.AddBreadthRestriction colRestrs(l)
        Next
    Else
        Dim sInvalidRestrs As String
        Dim m As Long
        For m = 1 To colRestrs.Count
            'The invalid restrs are returned in colRestrs
            If Not sInvalidRestrs = "" Then
                sInvalidRestrs = sInvalidRestrs & ", "
            End If
            sInvalidRestrs = sInvalidRestrs & colRestrs(m)
        Next
        RaiseError sSource, PEUSERInvalidBreadthRestrictions, "Invalid breadth restricitons: " & sInvalidRestrs, sInvalidRestrs
    End If
    
Exit Sub

ErrorHandler:
    m_objContext.SetComplete
    ConvertError sSource
End Sub


Private Function GetRefineURL(t_sSearchContext As String) As String
'   Gets the page name for refining searches from the context
'
'   d3  WDP     13-Jul-1999     Created (stub)
'   d17 JES     2-Nov-1999      Added URL for advanced search
On Error GoTo ErrorHandler
    
    Dim sSource As String
    sSource = "VWReaderSearch.GetRefineURL"
    
  
        Select Case m_boSearch.SearchType
        Case Basic
            GetRefineURL = gk_sPAGERefineBasicSearchFrameset
        Case Advanced, Field
            GetRefineURL = gk_sPAGERefineAdvancedSearch
        End Select
    
Exit Function
ErrorHandler:
    m_objContext.SetComplete
    ShowError sSource
End Function


Private Function FindInTerm(t_sFieldName As String, t_oiSearchTerm As ISearchTerm) As String
'   This method recursively searches through a search term looking
'   for terms relating to the field name passed. It returns
'   the value associated with that fieldname. If none is found
'   the method returns ""
'
'   d3  WDP     13-Jul-1999     Created
On Error GoTo ErrorHandler
    
    Dim sSource As String
    sSource = "VWReaderSearch.FindInTerm"

    If t_oiSearchTerm.TermType = STCriteriaTerm Then
        'If we have a criteria term, we simply look in
        'its item and value members
        Dim oboTerm As New BOCriteriaTerm
        Set oboTerm = t_oiSearchTerm
        If oboTerm.Item = t_sFieldName Then
            FindInTerm = oboTerm.Value
        Else
            FindInTerm = ""
        End If
    Else
        Dim oboCombTerm As New BOCombinationTerm
        Set oboCombTerm = t_oiSearchTerm
        Dim l As Integer
        'If we have a combination term we must recursively look at
        'all the child terms until we find a match
        For l = 1 To oboCombTerm.GetNumberOfTerms
            Dim sRes As String
            'Recursive call
            sRes = FindInTerm(t_sFieldName, oboCombTerm.GetTerm(l))
            If Not sRes = "" Then
                FindInTerm = sRes
                Exit Function
            End If
        Next
        'If no match is found return ""
        FindInTerm = ""
    End If
                
Exit Function
ErrorHandler:
    m_objContext.SetComplete
    ShowError sSource
End Function

Private Function SummaryEntryIsHit() As Boolean
'   Indicates if the current record in the results (for summary results)
'   is a search hit itself
'
'   d9  WDP     17-Aug-1999     Created

On Error GoTo ErrorHandler
    
    Dim sSource As String
    sSource = "VWReaderSearch.SummaryEntryIsHit"
        
    If Not (m_oiResults.ResultType = RESULTLCodeSummary Or _
                m_oiResults.ResultType = RESULTClassSummary) Then
        RaiseError sSource, PEGeneralInternalError, "Results are not of summary type."
    Else
        Dim sHit As String
        sHit = m_oiResults.GetData(gk_sDIHitAtSummaryLevel)
        If Not sHit = "" Then
            SummaryEntryIsHit = True
        Else
            SummaryEntryIsHit = False
        End If
        
    End If
Exit Function

ErrorHandler:
    m_objContext.SetComplete
    ShowError sSource
End Function



'
' Class initialisation and MTS event handlers
'

Private Sub Initialise()
    Set m_objContext = GetObjectContext()
    
    ' Set asp objects
    Set m_aspResponse = m_objContext.Item("Response")
    Set m_aspRequest = m_objContext.Item("Request")
    Set m_aspServer = m_objContext.Item("Server")
    
    ' Specific initialisation
    ' <Specific initialisation here>
    Set m_fwLeafTable = New FWTable
    m_bDisplayLeaflets = False
    
    

End Sub

Private Sub ObjectControl_Activate()
    Initialise
End Sub

Private Sub ObjectControl_Deactivate()
    Set m_objContext = Nothing
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = True
End Function


'Temporary testing methods

Public Sub DoSearch()
' run a test search to try display methods


    
    Dim sSource As String
    sSource = "VWReaderSearch.DoSearch"
                   
    'Dim boTerm As New BOCombinationTerm
    'Dim boRootTerm As New BOCombinationTerm
    Dim boTerm1 As New BOCriteriaTerm
    'Dim boTerm2 As New BOCriteriaTerm
    'Dim boTerm3 As New BOCriteriaTerm
    'Dim sResults() As String
    
    boTerm1.Item = gk_sDIScopeContent
    boTerm1.Value = "of"
    boTerm1.Operator = CTOVerityIN
    boTerm1.Negation = False
    boTerm1.TermType = CTTFreeText
    
    'boTerm2.Item = "Scope and Content"
    'boTerm2.Value = "chicks"
    'boTerm2.Operator = CTOVerityIN
    'boTerm2.Negation = False
    'boTerm2.TermType = CTTFreeText
    
    'boTerm3.Item = "Scope and Content"
    'boTerm3.Value = "freezing"
    'boTerm3.Operator = CTOVerityIN
    'boTerm3.Negation = False
   ' boTerm3.TermType = CTTFreeText
    
    'boTerm.AddTerm boTerm1
    'boTerm.AddTerm boTerm3
    'boTerm.TermType = CTTAND
    
    'boRootTerm.AddTerm boTerm
    'boRootTerm.AddTerm boTerm2
    'boRootTerm.TermType = CTTOR
    
    Set m_boSearch = New BOSearch
    m_boSearch.RootSearchTerm = boTerm1
    Dim oboExec As New BOSearchExecutive
    
    'Execute the search and get the results
    oboExec.Search = m_boSearch
    oboExec.RunSearch
    Set m_oiResults = New ISearchResults
    Set m_oiResults = oboExec.Results
    DisplayResults
    
Exit Sub

ErrorHandler:
    m_objContext.SetComplete
    ShowError sSource
End Sub








