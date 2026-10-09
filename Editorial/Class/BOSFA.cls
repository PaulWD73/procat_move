VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BOSFA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    BOSFA.cls
' System:       PRO, PROCAT
' Copyright:    (C) Quidnunc Limited
'
' Description:  This business object used in editorial and its main function is to
'               create and update sfas
'
' Amendment history:
'   d1  NDK     09NOV1999   Created
'   d2  NDK     18Nov1999   Removed all references to release status
'   d3  KSK     25Nov1999   Replace the database constant
'   d7  TMJ     21Mar2000   Added Error Checking to Delete
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'private objects and types
Private m_typSFADetails As SFADetails
Private m_oFWDBHelper As New FWDBHelper

'private enums

Private Enum ESFAReleaseStatus
    ESRSApproved = 0
    ESRSCandidate = 1
End Enum
Public Property Get Name()
    Name = "BOSFA"
End Property
Public Property Let Data(t_eSFADetails As SFADetails)
    m_typSFADetails = t_eSFADetails
End Property
Public Property Get Data() As SFADetails
    Data = m_typSFADetails
End Property

Public Sub Create()
'   Description
'   Creates a new sfa by using the data stored in m_typSFADetails
'
'   d1  NDK     09NOV1999   Created
'

On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim sSQL As String

    sSource = Me.Name & ".Create"
    
    'Build the query
    'table
    sSQL = " INSERT " & gk_sDBTABLESFA & "("
    
    'columns
    sSQL = sSQL & gk_sDBCOLSFATitleText & ","
    sSQL = sSQL & gk_sDBCOLPublisher & ","
    sSQL = sSQL & gk_sDBCOLAuthor & ","
    sSQL = sSQL & gk_sDBCOLPublicationDate & ","
    sSQL = sSQL & gk_sDBCOLInPROLibrary & ","
    sSQL = sSQL & gk_sDBCOLInPROReadingRoom & ","
    sSQL = sSQL & gk_sDBCOLSFASubtitleText & ","
    sSQL = sSQL & gk_sDBCOLExternalSFAReference & ","
    sSQL = sSQL & gk_sDBCOLEdition & ","
    sSQL = sSQL & gk_sDBCOLPublicationNote & ","
    sSQL = sSQL & gk_sDBCOLElectronicSFAFilename & ","
    sSQL = sSQL & gk_sDBCOLCreatedOn & ","
    sSQL = sSQL & gk_sDBCOLSFATypeID
    
    'values
    sSQL = sSQL & " ) VALUES ("
    sSQL = sSQL & QuoteValueWithQuotes(m_typSFADetails.sTitle) & ","
    sSQL = sSQL & QuoteValueWithQuotes(m_typSFADetails.sPublisher) & ","
    sSQL = sSQL & QuoteValueWithQuotes(m_typSFADetails.sAuthor) & ","
    sSQL = sSQL & QuoteValueWithQuotes(m_typSFADetails.sPublicationDate) & ","
    sSQL = sSQL & IIf(m_typSFADetails.bInProLibaray, 1, 0) & ","
    sSQL = sSQL & IIf(m_typSFADetails.bInProReference, 1, 0) & ","
    sSQL = sSQL & QuoteValueWithQuotes(m_typSFADetails.sSubTitle) & ","
    sSQL = sSQL & QuoteValueWithQuotes(m_typSFADetails.sExternalReference) & ","
    sSQL = sSQL & QuoteValueWithQuotes(m_typSFADetails.sEdition) & ","
    sSQL = sSQL & QuoteValueWithQuotes(m_typSFADetails.sPublicationNote) & ","
    sSQL = sSQL & QuoteValueWithQuotes(m_typSFADetails.sElectronicSFAFilename) & ","
    sSQL = sSQL & "Getdate()" & ","
    sSQL = sSQL & m_typSFADetails.lTypeID & ")"
    
    'Insert the row
    
    m_oFWDBHelper.Execute sSQL
   
    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub
Public Sub Update()
'   Description
'
'   Updates an existing sfa with the details
'
'   d1  NDK     09NOV1999   Created
'

On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim sSQL As String
    
    sSource = Me.Name & ".Update"
    
    'Build the query

    'table
    sSQL = " UPDATE " & gk_sDBTABLESFA & " SET "
    
    'fields
    sSQL = sSQL & gk_sDBCOLSFATitleText & "=" & QuoteValueWithQuotes(m_typSFADetails.sTitle) & ","
    sSQL = sSQL & gk_sDBCOLPublisher & "=" & QuoteValueWithQuotes(m_typSFADetails.sPublisher) & ","
    sSQL = sSQL & gk_sDBCOLAuthor & "=" & QuoteValueWithQuotes(m_typSFADetails.sAuthor) & ","
    sSQL = sSQL & gk_sDBCOLPublicationDate & "=" & QuoteValueWithQuotes(m_typSFADetails.sPublicationDate) & ","
    sSQL = sSQL & gk_sDBCOLInPROLibrary & "=" & IIf(m_typSFADetails.bInProLibaray, 1, 0) & ","
    sSQL = sSQL & gk_sDBCOLInPROReadingRoom & "=" & IIf(m_typSFADetails.bInProReference, 1, 0) & ","
    sSQL = sSQL & gk_sDBCOLSFASubtitleText & "=" & QuoteValueWithQuotes(m_typSFADetails.sSubTitle) & ","
    sSQL = sSQL & gk_sDBCOLExternalSFAReference & "=" & QuoteValueWithQuotes(m_typSFADetails.sExternalReference) & ","
    sSQL = sSQL & gk_sDBCOLEdition & "=" & QuoteValueWithQuotes(m_typSFADetails.sEdition) & ","
    sSQL = sSQL & gk_sDBCOLPublicationNote & "=" & QuoteValueWithQuotes(m_typSFADetails.sPublicationNote) & ","
    sSQL = sSQL & gk_sDBCOLElectronicSFAFilename & "=" & QuoteValueWithQuotes(m_typSFADetails.sElectronicSFAFilename)
    
    'condition
    sSQL = sSQL & " WHERE " & gk_sDBCOLSFAReferenceID & "=" & m_typSFADetails.lID
    
    m_oFWDBHelper.Execute sSQL
    
    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub
    
Public Sub Delete(ByVal t_lSFAReferenceID As Long)
'   Description
'
'   Deletes the SFA from the SFA table, no tables are checked before deleting
'   If any foreign key violatins occur then an error is raised
'
'   d1  NDK     09NOV1999   Created
'   d6  KSK     25Nov1999   Replaced the Database error constant with
'                           with a general constant
'   d7  TMJ     21Mar2000   Added checking to ensure that SFA not referenced
'

On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim sSQL As String
    Dim List As FWRecordList
       
    sSource = Me.Name & ".Delete"
    
    sSQL = "SELECT " & gk_sSFAReferenceID & " FROM " & gk_sDBTABLEPCSFA & _
           " WHERE " & gk_sDBCOLSFAReferenceID & " = " & t_lSFAReferenceID
           
    Set List = m_oFWDBHelper.FindBySql(sSQL)
    
    If List.RecordCount > 0 Then
        RaiseError sSource, peCanNotDeleteSFAWithReferences, _
            "This SFA has is linked to Catalogue entries, so cannot be deleted"
    End If
           
    sSQL = "SELECT " & gk_sSFAReferenceID & " FROM " & _
           gk_sDBTABLEAuditTrailPCSFA & " WHERE " & gk_sDBCOLSFAReferenceID & _
           " = " & t_lSFAReferenceID
    
    Set List = m_oFWDBHelper.FindBySql(sSQL)
    
    If List.RecordCount > 0 Then
        RaiseError sSource, peCanNotDeleteSFAWithReferences, _
            "This SFA has references in the Audit Trail, so cannot be deleted"
    End If
    
    sSQL = "SELECT " & gk_sSFAReferenceID & " FROM " & _
           gk_sDBTABLEEditsetPCSFA & " WHERE " & gk_sDBCOLSFAReferenceID & _
           " = " & t_lSFAReferenceID
    
    Set List = m_oFWDBHelper.FindBySql(sSQL)
    
    If List.RecordCount > 0 Then
        RaiseError sSource, peCanNotDeleteSFAWithReferences, _
            "This SFA has references in an Editset, so cannot be deleted"
    End If
    
    
    'Build the query
    sSQL = " DELETE " & gk_sDBTABLESFA & " WHERE "
    sSQL = sSQL & gk_sDBCOLSFAReferenceID & "=" & t_lSFAReferenceID
    
    'Delete the row
    m_oFWDBHelper.Execute sSQL
   
    Exit Sub
ErrorHandler:
    If Err.Number = gk_lErrorNumberDatabase Then
        'This could occur here only because there are records refering to this sfa being deleted
        RaiseError sSource, peCanNotDeleteSFAWithReferences, _
            "This SFA has references in other tables, can not delete this"
    Else
        ConvertError sSource
    End If
End Sub
Public Function SFAHasReference(ByVal t_lSFAReferenceID As Long) As Boolean
'   Description
'
'   Checks if the SFA has any reference to ES_PC_SFA table, the Delete function was not to perform this
'   check and handle it in the error handler.Now the need to show a message if there
'   are references necessitates this function
'
'
'   d4  CRM     10FEB2000   Created
'
'

On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim sSQL As String
    Dim ofwDBHelper As New FWDBHelper
    Dim ofwRecList As FWRecordList
                
    sSource = "BOSFA.SFAHasReference"
    
    sSQL = "SELECT * FROM " & gk_sDBTABLEEditsetPCSFA & _
           " WHERE " & gk_sDBCOLSFAReferenceID & " = " & t_lSFAReferenceID
    Set ofwRecList = ofwDBHelper.FindBySql(sSQL)
    
    SFAHasReference = IIf(ofwRecList.RecordCount > 0, True, False)
    
Exit Function
ErrorHandler:
    ConvertError sSource
End Function


    
'
' Class initialisation
'

Private Sub Initialise()
   ' Specific initialisation
End Sub

Private Sub Class_Initialize()
    Initialise
End Sub

Private Sub Class_Terminate()
    ' specific clear up
End Sub



