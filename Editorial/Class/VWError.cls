VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VWError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    VWError.cls
' System:       PROCAT, PRO
' Copyright:    (C) Quidnunc Limited
'
' Description:  View catalogue browser class with all display items for catalogues
'
' Amendment history:
'   d1   ksk     08Jun99      Created.
'   d1   ksk     08Jun99      Added DisplayInvalidPage method
'   d3   ndk     14Jul99      Added DisplayUnhandledErrorPage method
'   d4   ksk     29Jul99      Added methods DisplayIncorrectDataLength and
'                               DisplayIncorrectDataType
'   d5   ksk     08Oct99      Added method to display the access denied error
'   d6   ndk     16Oct99      Added method DisplayErrorDescriptionOnly to display errors
'                             which show only the description
'   d7   ndk     12Nov1999    Removed unused errorhandlers (which now call dislay description only handler)
'   d8   wdp     23Feb2000    Added calls to BackButtonHTML from error display methods
'   d9   alm     14Nov2001    Added support for back button span multiple pages
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private m_ofwRefData As FWReferenceData
Private m_ofwHtmlRef As FWHTML

Private m_objContext As Object
Private m_aspResponse As Response
Private m_aspRequest As Request
Private m_aspServer As Server
Private m_aspSession As Session

' Constants


Public Sub DisplayInvalidEditsetPage(Optional ByVal t_sLinkbackPage As String = "")
'
'   d2  ksk     08Jul1999   Created
'   d9  alm     14Nov2001    Added support for back button span multiple pages
On Error GoTo ErrorHandler
    Dim sHTML As String
    sHTML = m_ofwHtmlRef.Bold(m_ofwHtmlRef.Font(m_ofwRefData.Search( _
                gError, EErrorResource.erInvalidEditSet), COLOURRed))
    m_aspResponse.Write sHTML & BackButtonHTML(t_sLinkbackPage)
Exit Sub
ErrorHandler:
End Sub

Public Sub DisplayInvalidNextPage(Optional ByVal t_sLinkbackPage As String = "")
'
'   d2  ksk     08Jul1999   Created
'   d9  alm     14Nov2001    Added support for back button span multiple pages
On Error GoTo ErrorHandler
    Dim sHTML As String
    sHTML = m_ofwHtmlRef.Bold(m_ofwHtmlRef.Font(m_ofwRefData.Search( _
                        gError, EErrorResource.erInvalidNextPage), COLOURRed))
    m_aspResponse.Write sHTML & BackButtonHTML(t_sLinkbackPage)
Exit Sub
ErrorHandler:
End Sub

Public Sub DisplayInvalidCatalogueReferencePage(Optional ByVal t_sLinkbackPage As String = "")
'
'   d2  ksk     08Jul1999   Created
'   d9  alm     14Nov2001    Added support for back button span multiple pages
On Error GoTo ErrorHandler
    Dim sHTML As String
    sHTML = m_ofwHtmlRef.Bold(m_ofwHtmlRef.Font(m_ofwRefData.Search( _
                        gError, EErrorResource.erInvalidCatalogueReference), COLOURRed))
    m_aspResponse.Write sHTML & BackButtonHTML(t_sLinkbackPage)
Exit Sub
ErrorHandler:
End Sub

Public Sub DisplayUnhandledErrorPage(ByVal vErrorMessage As Variant, _
                                     Optional ByVal t_sLinkbackPage As String = "")
'
'   d3   ndk     14Jul99      Created
'   d9   alm     14Nov2001    Added support for back button span multiple pages
On Error GoTo ErrorHandler
    Dim sHTML As String
    sHTML = m_ofwHtmlRef.Bold(m_ofwHtmlRef.Font("Unhandled Error Encountered : " & "<p>" & vErrorMessage, COLOURRed))
    m_aspResponse.Write sHTML & BackButtonHTML(t_sLinkbackPage)
Exit Sub
ErrorHandler:

End Sub


Public Sub DisplayAccessDenied(t_sDescription As String, _
                               Optional ByVal t_sLinkbackPage As String = "")
'   d5   ksk     08Oct99      Added method to display the access denied error
'   d9   alm     14Nov2001    Added support for back button span multiple pages
On Error GoTo ErrorHandler
    Dim sHTML As String
    sHTML = m_ofwHtmlRef.Bold(m_ofwHtmlRef.Font(m_ofwRefData.Search( _
                        gError, EErrorResource.erAccessDenied) & t_sDescription, COLOURRed))
    m_aspResponse.Write sHTML & BackButtonHTML(t_sLinkbackPage)
Exit Sub
ErrorHandler:

End Sub

Public Sub DisplayErrorDescriptionOnly(t_sDescription As String, _
                                       Optional ByVal t_sLinkbackPage As String = "")
'   d6   ndk     16Oct99      Method to display errors which show only the description (without source
'   d9   alm     14Nov2001    Added support for back button span multiple pages
On Error GoTo ErrorHandler
    Dim sHTML As String
    sHTML = m_ofwHtmlRef.Bold(m_ofwHtmlRef.Font(t_sDescription, COLOURRed))
    m_aspResponse.Write sHTML & BackButtonHTML(t_sLinkbackPage)
Exit Sub
ErrorHandler:
End Sub




''''''''''''''''' Implementation ''''''''''''''''''''''

'
' Private functions
'


'
' Class initialisation and MTS event handlers
'

Private Sub Initialise()
    Set m_objContext = GetObjectContext()
    
    ' Set asp objects
    Set m_aspResponse = m_objContext.Item("Response")
    Set m_aspRequest = m_objContext.Item("Request")
    Set m_aspServer = m_objContext.Item("Server")
    Set m_aspSession = m_objContext.Item("Session")
    
    Set m_ofwRefData = New FWReferenceData
    Set m_ofwHtmlRef = New FWHTML
End Sub

Private Sub Terminate()
    Set m_objContext = Nothing
    
    ' Set asp objects
    Set m_aspResponse = Nothing
    Set m_aspRequest = Nothing
    Set m_aspServer = Nothing
    Set m_aspSession = Nothing
    Set m_ofwRefData = Nothing
    Set m_ofwHtmlRef = Nothing
End Sub

Private Sub Class_Initialize()
    Initialise
End Sub

Private Sub Class_Terminate()
    Terminate
End Sub






