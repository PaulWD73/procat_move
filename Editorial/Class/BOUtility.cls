VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "BOUtility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    BOUtility.cls
' System:       PROCAT, PRO
' Copyright:    (C) Quidnunc Limited
'
' Description:  Utility methods that are run when the session ends and other methods
'
' Instancing : Multiuse
'
' Amendment history:
'
' d1    NDK     08Jun1999       Created.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit


Public Sub CleanUpSessionData()
' Description:
'
' Called in the session_onend event (in global.asa)
' Removes all the data pertaining to the session from the database
' Most of these will be stored by the transient manager
'
' Amendment history:
'
' d1    NDK     08Jun1999       Created.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    'Private Variables
    On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim sSql As String
    Dim sSessionId As String, sTableName As String
    
    Dim oFWDBHelper As FWDBHelper
    Dim oBOEditsetEntryLockManager  As BOEditSetEntryLockManager

    sSource = "BOUtility.CleanUpSessionData()"
    
    Set oFWDBHelper = New FWDBHelper
    Set oBOEditsetEntryLockManager = New BOEditSetEntryLockManager
    
    'Get the session id
    sSessionId = CStr(GetObjectContext().Item(gk_sSessionObject).SessionID)
    
    'Get the table name
    sTableName = gk_sDefaultTransientStorageTable
    
    sSql = "DELETE " & _
            "FROM " & sTableName & " " & _
            "WHERE SessionID = " & sSessionId
            
    'Execute the sql
    oFWDBHelper.Execute sSql
        
    'To unlock all the catalogues held by the user of this session
    oBOEditsetEntryLockManager.UnlockEditSetCatalogueEntryLocksForSession sSessionId
    
    'Release the objects
    Set oFWDBHelper = Nothing
    Set oBOEditsetEntryLockManager = Nothing
    
    Exit Sub
    
ErrorHandler:
    ShowError sSource
End Sub



