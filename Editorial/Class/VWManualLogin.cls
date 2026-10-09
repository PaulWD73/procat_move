VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "VWManualLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    VWManualLogin.cls
' System:       PROCAT, PRO
' Copyright:    (C) Quidnunc Limited
'
' Description:
'
' Used when the user wants to login by typing the login name
'
' Amendment history:
'
' d1    ndk     23Nov1999       Created.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit


'Interfaces
Implements ObjectControl

'private constants

'Private Variables
Private m_objObjectContext As ObjectContext
Private m_aspResponse  As Response
Private m_aspRequest As Request

Private Const ksNameLoginName As String = "LoginName"
Private m_lLoggedUserID As Long

'Private Constants
Private Sub ObjectControl_Activate()
    Initialise
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    Terminate
End Sub
Public Property Get Name()
    Name = "VWManualLogin"
End Property

Private Sub Initialise()
    
    On Error GoTo ErrorHandler
    
    Dim sSource As String
    sSource = Me.Name & ".Initialise"
    
    'Get the object context
    Set m_objObjectContext = GetObjectContext()
    If m_objObjectContext Is Nothing Then
        RaiseError sSource, peOCUnableToGetObjectContext, "Objectcontext is nothing"
    End If
    
    'get references to ASP objects
    Set m_aspResponse = m_objObjectContext.Item(gk_sResponseObject)
    Set m_aspRequest = m_objObjectContext.Item(gk_sRequestObject)
    
    'First time when a user logs, UserID will be zero. In Login method, the same line code should exist
    m_lLoggedUserID = GetDataFromSession(dsUserID, m_objObjectContext.Item(gk_sSessionObject))
    
    Exit Sub
    
ErrorHandler:
    m_objObjectContext.SetAbort
    ShowError sSource
End Sub

Private Sub Terminate()
    Set m_objObjectContext = Nothing 'release
    Set m_aspResponse = Nothing 'release
    Set m_aspRequest = Nothing 'release
End Sub
' Methods
'
Public Sub ManualLogin()
'
' d1    ndk     23Nov1999       Created.
'
' To validate login info. and permit user to proceed further
' Gets login name as parameter
'
On Error GoTo ErrorHandler
    Dim sSource As String, sLoginName As String
    
    sSource = Me.Name & ".ManualLogin"
    
    Dim oBOUser As BOUser
    
    Set oBOUser = New BOUser
   
    oBOUser.Initialize
  
    sLoginName = m_aspRequest.Form(ksNameLoginName)
    
    'Call the business object to validate the nt logon name of the user
    oBOUser.Login sLoginName
    
    'Store the users details in the session object
    StoreDataInSession DATASTORE.dsUserID, oBOUser.UserID
    StoreDataInSession DATASTORE.dsUserInitials, oBOUser.UserInitials
    StoreDataInSession DATASTORE.dsUserName, oBOUser.UserName
    StoreDataInSession DATASTORE.dsUserNTLogon, oBOUser.UserNTLogon
    
    m_lLoggedUserID = GetDataFromSession(dsUserID, m_objObjectContext.Item(gk_sSessionObject))
    
    m_objObjectContext.SetComplete
    
    Exit Sub
ErrorHandler:
    m_objObjectContext.SetAbort
    ShowError sSource
End Sub


