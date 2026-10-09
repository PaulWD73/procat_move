Attribute VB_Name = "BrowserConstants"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    BrowserConstants.bas
' System:       PROCAT, PRO
' Copyright:    (C) Quidnunc Limited
'
' Description:  Holds global methods and declaration needed for the Special
'               records used in boCatalogueBrowser and will be required by the
'               View class that will use it
'
' Amendment history:
'   d1   ksk      23Jun99      Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Public Const gk_sValueSeperator As String = ","

'These enumerations are used for the Specail record returned from the BOCatalgoueBrowser
Public Enum EBROWSERSPECIALRECORD
    bsrPageUP = 1
    bsrPageDown = 2
End Enum
Public Enum EBROWSERSPECIALRECORDTYPE
    bsrtNoMorePreviousPages = 1
    bsrtPreviousPage = 2
    bsrtNoMoreNextPage = 3
    bsrtNextPage = 4
End Enum

Public Enum ECOMPARISONTYPE
    ctEqual
    ctGreater
    ctLess
End Enum



Public Function GetBrowserSpecialRecordString(t_eBrowserSpecialRecord As EBROWSERSPECIALRECORD)
'Function to return a unique string value corresponding to the special record
    Select Case t_eBrowserSpecialRecord
    Case EBROWSERSPECIALRECORD.bsrPageUP
        GetBrowserSpecialRecordString = "PU"
    Case EBROWSERSPECIALRECORD.bsrPageDown
        GetBrowserSpecialRecordString = "PD"
    End Select
End Function


Public Function GetBrowserSpecialRecordTypeString(t_eBrowserSpecialRecord As EBROWSERSPECIALRECORDTYPE)
'Function to return a unique string to signify the type of the special record.
    Select Case t_eBrowserSpecialRecord
    Case EBROWSERSPECIALRECORDTYPE.bsrtNoMorePreviousPages
        GetBrowserSpecialRecordTypeString = "No&Previous&Page"
    Case EBROWSERSPECIALRECORDTYPE.bsrtPreviousPage
        GetBrowserSpecialRecordTypeString = "Previous&Page"
    Case EBROWSERSPECIALRECORDTYPE.bsrtNoMoreNextPage
        GetBrowserSpecialRecordTypeString = "No&Next&Page"
    Case EBROWSERSPECIALRECORDTYPE.bsrtNextPage
        GetBrowserSpecialRecordTypeString = "Next&Page"
    End Select
End Function



