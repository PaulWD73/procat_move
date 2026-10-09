Attribute VB_Name = "TempMod"
Option Explicit

'The number of matches to display on each page
Public Const gk_lPSBrowserTableSize As Integer = 3
'Just get all the matches (for guided search results)
Public Const gk_lPSSearchGetAllMatches As Integer = -1

Public Const gk_sPAGEPSList As String = "PSBrowser.asp"

Public Const gk_sFORMPSListQuery As String = "sQuery"

Public Const gk_sFORMPSListStartIndex As String = "lStartIndex"

Public Const gk_sFORMPSListUsage As String = "eUsage"

Public Const gk_sFORMPSSearchResultsID As String = "iID"
Public Const gk_sFORMPSSearchResultsName As String = "sName"

'Search page (goes here to run the search)
Public Const gk_sPagePSRunSearch As String = "PSRunSearch.asp"
Public Const gk_sPagePSEditSearch As String = "PSEditSearch.asp"

