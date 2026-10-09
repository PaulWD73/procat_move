VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BOEditSetEntryLockManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    BOEditSetEntryLockManager.cls
' System:       PROCAT, PRO
' Copyright:    (C) Quidnunc Limited
' Instancing:   Private
' Description:
'
' Business object to manage the locking and unlocking of catalogue entries
' while editing
'
' Amendment history:
'
' d1    ndk     16Oct1999       Created.
' d2    ndk     02Nov1999       Locking is now done tab wise
' d3    ndk     22Nov1999       Locking and unlocking always return true
'                               They dont execute any other code
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

'private framework objects
Private m_oFWDBHelper As FWDBHelper

'private property variables
Private m_lUserID As Long
Private m_vUserName As Variant
Private m_vLockedDateTime As Variant
Private m_vLockedTabID As Variant
Private m_vLockedTabName As Variant
Private m_vSessionID As Variant

Private Sub Class_Initialize()
    Initialise
End Sub
Private Sub Class_Terminate()
    Terminate
End Sub
Private Sub Initialise()
    Set m_oFWDBHelper = New FWDBHelper
End Sub
Private Sub Terminate()
    Set m_oFWDBHelper = Nothing
End Sub

Public Property Get Name()
    Name = "BOCatalogueEntryLockManager"
End Property

Public Property Get LockedUserID() As Long
    'Returns the id of the user who has locked the catalogue
    'This property is set when LockEditSetCatalogueEntry is called and the attempt to lock a cataogue fails
    
    LockedUserID = m_lUserID
    
End Property
Public Property Get LockedUserName() As Variant
    'Returns the name of the user who has locked the catalogue
    'This property is set when LockEditSetCatalogueEntry is called and the attempt to lock a cataogue fails
    
    LockedUserName = m_vUserName
    
End Property
Public Property Get LockedTabID() As Variant
    'Returns the id of the locked tab for the catalogue
    'This property is set when LockEditSetCatalogueEntry is called and the attempt to lock a cataogue fails
    
    LockedTabID = m_vLockedTabID
    
End Property
Public Property Get LockedTabName() As Variant
    'Returns the name of the locked tab for the catalogue
    'This property is set when LockEditSetCatalogueEntry is called and the attempt to lock a cataogue fails
    
    LockedTabName = m_vLockedTabName
    
End Property

Public Property Get LockedDateTime() As Variant
    'Returns the date-time string when the catalogue was locked
    'This property is set when LockEditSetCatalogueEntry is called and the attempt to lock a cataogue fails
    
    LockedDateTime = m_vLockedDateTime

End Property

Public Property Get LockedSessionID() As Variant
    'Returns the session-id of the user who has locked the catalogue
    'This property is set when LockEditSetCatalogueEntry is called and the attempt to lock a cataogue fails
    
    LockedSessionID = m_vSessionID
    
End Property
Public Function LockEditSetCatalogueEntry( _
    ByVal t_eTabToLock As ETabType, _
    ByVal t_lCatalogueID As Long, _
    ByVal t_eCatalogueLevel As ECATALOGUELEVEL, _
    ByVal t_lEditSetID As Long, _
    ByVal t_lUserID As Long, _
    ByVal t_sSessionId As String) As Boolean
    
' Locks a catalogue entry in the editset
'
' Amendment history:
'
' d1    ndk     16Oct1999       Created.
' d2    ndk     02Nov1999       Locking is now done tab wise, edit set id is removed from condition while checking
'                               for locks because the same catalogue can not be locked in two editsets
' d3    ndk     22Nov1999       Locking and unlocking always return true
'                               They dont execute any other code


    On Error GoTo ErrorHandler
    
    Dim sSource As String, sSql As String
    Dim oFWRecordlist As FWRecordList
    
    sSource = Me.Name & ".LockEditSetCatalogueEntry"
    
    'd3
    LockEditSetCatalogueEntry = True
    
    Exit Function
    
    'default
    LockEditSetCatalogueEntry = False
    
    'To check whether the catalogue is already locked
    'construct sql
    sSql = _
        " SELECT " & _
            gk_sDBCOLTabToLock & "," & _
            gk_sDBCOLCatalogueID & "," & _
            gk_sDBCOLCatalogueLevel & "," & _
            gk_sDBCOLEditSetID & "," & _
            gk_sDBCOLUserID & "," & _
            gk_sDBCOLSessionID & "," & _
            gk_sDBCOLDateTimeOfLocking & _
        " FROM " & _
            gk_sDBTABLELockedCatalogues & _
        " WHERE " & _
            gk_sDBCOLTabToLock & "=" & t_eTabToLock & " AND " & _
            gk_sDBCOLCatalogueID & "=" & t_lCatalogueID & " AND " & _
            gk_sDBCOLCatalogueLevel & "=" & GetCatalogueLevelCode(t_eCatalogueLevel)
    
    'Get the list of records
    Set oFWRecordlist = m_oFWDBHelper.FindBySql(sSql)
    
    If oFWRecordlist.RecordCount = 0 Then
        'The catalogue is not locked
        'Insert a record in the table to signal that the catalogue is locked
        sSql = _
            " INSERT " & gk_sDBTABLELockedCatalogues & "(" & _
                gk_sDBCOLTabToLock & "," & _
                gk_sDBCOLCatalogueID & "," & _
                gk_sDBCOLCatalogueLevel & "," & _
                gk_sDBCOLEditSetID & "," & _
                gk_sDBCOLUserID & "," & _
                gk_sDBCOLSessionID & "," & _
                gk_sDBCOLDateTimeOfLocking & ") " & _
            " VALUES (" & _
                t_eTabToLock & "," & _
                t_lCatalogueID & "," & _
                GetCatalogueLevelCode(t_eCatalogueLevel) & "," & _
                t_lEditSetID & "," & _
                t_lUserID & "," & _
                t_sSessionId & "," & _
                "GetDate()" & ")"
        
        'insert the record
        m_oFWDBHelper.Execute sSql
        
        LockEditSetCatalogueEntry = True
        
    Else    'There is/are some entries for the catalogue in the lock information table
        'The query is expected to give only one record as the primary key consists
        'of a combination of editsetid, catalogueid and catalogue level
                
        'set the properties (LockedUserID,LockedDateTime, LockedSessionID
        m_lUserID = oFWRecordlist.GetFieldValue(gk_sDBCOLUserID)
        m_vLockedDateTime = oFWRecordlist.GetFieldValue(gk_sDBCOLDateTimeOfLocking)
        m_vSessionID = oFWRecordlist.GetFieldValue(gk_sDBCOLSessionID)
        m_vLockedTabID = oFWRecordlist.GetFieldValue(gk_sDBCOLTabToLock)
        
        'Check whether it was locked by the same person
        If m_lUserID = t_lUserID And Trim(m_vSessionID) = Trim(t_sSessionId) And _
            m_vLockedTabID = t_eTabToLock Then
        
            'The same catalogue and same tab was locked by the same person earlier in the same session
            'Let him to edit the catalogues tab , he might have moved away from the screen accidentally
            'This option is given to prevent having to give other options to release locks
            
            LockEditSetCatalogueEntry = True
        Else
            'to get the name of the user
            sSql = _
                " SELECT " & gk_sDBCOLUserName & _
                " FROM " & gk_sDBTABLEUsers & _
                " WHERE " & gk_sDBCOLUserID & "=" & m_lUserID
            
            Set oFWRecordlist = m_oFWDBHelper.FindBySql(sSql)
            
            If oFWRecordlist.RecordCount = 0 Then
                m_vUserName = "User Not Found !"
            Else
                m_vUserName = oFWRecordlist.GetFieldValue(gk_sDBCOLUserName)
            End If
            
            'to get the id of the tab locked
            Select Case m_vLockedTabID
                Case ETabType.etTabContext
                    m_vLockedTabName = "Context"
                Case ETabType.etTabSummary
                    m_vLockedTabName = "Summary"
                Case ETabType.etTabAccess
                    m_vLockedTabName = "Access"
                Case ETabType.etTabContent
                    m_vLockedTabName = "Content"
                Case ETabType.etTabAdminHistory
                    m_vLockedTabName = "Admin History"
                Case ETabType.etTabIndexTerms
                    m_vLockedTabName = "Index Terms"
            End Select
                    
            'Locking was not successful
            LockEditSetCatalogueEntry = False
        End If
    End If
    
    'Release objects
    
    Set oFWRecordlist = Nothing
    
    Exit Function
    
ErrorHandler:
    ConvertError sSource
End Function

Public Sub UnlockLockEditSetCatalogueEntry( _
    ByVal t_eTabToUnLock As ETabType, _
    ByVal t_lCatalogueID As Long, _
    ByVal t_eCatalogueLevel As ECATALOGUELEVEL, _
    ByVal t_lEditSetID As Long, _
    ByVal t_lUserID As Long, _
    ByVal t_sSessionId As String)
    
' The lock is released only if it is held by that user in that session
'
' Amendment history:
'
' d1    ndk     16Oct1999       Created.
' d3    ndk     22Nov1999       Locking and unlocking always return true
'                               They dont execute any other code
    
    On Error GoTo ErrorHandler
    
    Dim sSource As String, sSql As String
        
    'd3
    Exit Sub
    
    sSource = Me.Name & ".UnlockLockEditSetCatalogueEntry"
    
    'To unlock a catalogue entry, delete all its entries from the lock info table
    'construct sql
    sSql = _
        " DELETE " & _
            gk_sDBTABLELockedCatalogues & _
        " WHERE " & _
            gk_sDBCOLTabToLock & "=" & t_eTabToUnLock & " AND " & _
            gk_sDBCOLCatalogueID & "=" & t_lCatalogueID & " AND " & _
            gk_sDBCOLCatalogueLevel & "=" & GetCatalogueLevelCode(t_eCatalogueLevel) & " AND " & _
            gk_sDBCOLEditSetID & "=" & t_lEditSetID & " AND " & _
            gk_sDBCOLUserID & "=" & t_lUserID & " AND " & _
            gk_sDBCOLSessionID & "=" & Quotes(t_sSessionId)
    
    'Delete the records
    m_oFWDBHelper.Execute sSql
    
    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub


Public Sub UnlockEditSetCatalogueEntryLocksForSession( _
    ByVal t_sSessionId As String)
    
'   All the locks held by the user of that session are released
'   This method is called generally when the session times out
'
' Amendment history:
'
' d1    ndk     16Oct1999       Created.
' d3    ndk     22Nov1999       Locking and unlocking always return true
'                               They dont execute any other code
    
    On Error GoTo ErrorHandler
    
    Dim sSource As String, sSql As String
        
    sSource = Me.Name & ".UnlockEditSetCatalogueEntryLocksForSession"
    
    'd3
    Exit Sub
    
    'To unlock a catalogue entry, delete all its entries from the lock info table
    'construct sql
    sSql = _
        " DELETE " & _
            gk_sDBTABLELockedCatalogues & _
        " WHERE " & _
            gk_sDBCOLSessionID & "=" & Quotes(t_sSessionId)
    
    'Delete the records
    m_oFWDBHelper.Execute sSql
    
    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub




