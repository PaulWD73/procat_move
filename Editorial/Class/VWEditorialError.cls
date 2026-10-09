VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "VWEditorialError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    VWEditorialError.cls
' System:       PROCAT, PRO
' Copyright:    (C) Quidnunc Limited
'
' Description:  View class for editorial errors
'
' Amendment history:
'   d1  WDP 23Feb2000   Created
'   d2  ALM 14Nov2001   Added linkback to jump back multiple pages
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Option Explicit

Implements ObjectControl

Private m_objContext As Object
Private m_aspResponse As Response
Private m_aspRequest As Request


Public Sub DisplayError()
'Displays the error on in the browser
'No error handling (this is still part of the errorhandler)
'd1 WDP 23Feb2000   Created
'd2 ALM 14Nov2001   Added linkback to jump back multiple pages

    Dim lErrorNumber As Long, sErrorSource As String
    Dim sErrorDesc As String, sExtraText As String
    Dim sLinkbackPage As String
    
    'Get the error number
    lErrorNumber = m_aspRequest.Item(gk_sPARAMErrorNumber)
    sErrorSource = m_aspRequest.Item(gk_sPARAMErrorSource)
    sErrorDesc = m_aspRequest.Item(gk_sPARAMErrorDesc)
    sExtraText = m_aspRequest.Item(gk_sPARAMErrorExtraText)
                    
    'Get the linkback number
    If Not IsEmpty(m_aspRequest.Item(gk_sPARAMErrorLinkBack)) Then
        sLinkbackPage = m_aspRequest.Item(gk_sPARAMErrorLinkBack)
    Else
        sLinkbackPage = ""
    End If
        
    PROCATEditorialErrorHandling lErrorNumber, sErrorSource, sErrorDesc, sLinkbackPage
    
End Sub



'
' Class initialisation and MTS event handlers
'

Private Sub Initialise()
    Set m_objContext = GetObjectContext()
    
    ' Set asp objects
    Set m_aspResponse = m_objContext.Item("Response")
    Set m_aspRequest = m_objContext.Item("Request")
    

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
