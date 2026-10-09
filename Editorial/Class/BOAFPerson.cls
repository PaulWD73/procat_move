VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BOAFPerson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    BOAFPerson.cls
' System:       PRO, PROCAT
' Copyright:    (C) Quidnunc Limited
'
' Description:  This class encapsultes person AF term.
'               Written in the same lines as BOLoader class send by WDP.
'
' Amendment history:
'   d1  Geedha Sivanadian     6-Aug-1999      Created
'   d2  Geedha Sivanadian     12-Aug-1999     Handled all strings of length more than 128 differently.
'                                             If you convert all variant to strings using variantasstring
'                                             Then it would allow you to pass only max of 128 chars.
'   d3  Sree Krishna          20-Aug-1999     Changed the using of SP to updated
'                                             statement
'   d4  Aravinda A            08-Oct-1999     Modified save method to take care of Nulls
'                                             Added methods and properties for See/See also links
'                                             Replaced hardcoded table and column constants
'   d5  Sree Krishna          04-Nov-1999     Call to reference data for refreshing after change
'   d6  Sree Krishna          25-Nov-1999     Replaced the Database error number
'                                             with a constant
'   d7  Aravinda A            08-Dec-1999     Limited see links to one, added rejection criteria
'   d8  Aravinda A            21-Dec-1999     Defect fixes (minor)
'   d9  Aravinda A            20-Jan-2000      Added function IsTermHasLinks
'   d10 Mahesh CR             16-Feb-2000     Added function NonPrefTermsRelPrefTermIsIncomOrApproved
'   d10 WDP                   23-Feb-2000     Added function IsLinkedtoOtherTerm
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Interfaces
' Implements the IAFSaver interface for all AF Save classes
Implements IAFSaver

'SInce this object has to be transiently saved.
Implements PROCATTransientInterface

' Instance variables
Private m_ofwSQLExecuteHelper As FWDBHelper
Private m_objContext As Object


Private m_typData As PersonAFTerm

'd4
Private m_colPersonSeeLinks As Collection
Private m_colPersonSeeAlsoLinks As Collection

Public Property Let Data(t_typPersonAFTerm As PersonAFTerm)
    IAFSaver_Data = t_typPersonAFTerm
End Property

Public Property Get Data() As PersonAFTerm
    Data = m_typData
End Property

Public Property Get SeeLinks() As Collection
'd4
    Set SeeLinks = m_colPersonSeeLinks
End Property

Public Property Get SeeAlsoLinks() As Collection
'd4
    Set SeeAlsoLinks = m_colPersonSeeAlsoLinks
End Property



Private Sub Initialise()
    ' Specific initialisation
    Set m_ofwSQLExecuteHelper = New FWDBHelper
    Set m_objContext = GetObjectContext()
    Set m_colPersonSeeLinks = New Collection
    Set m_colPersonSeeAlsoLinks = New Collection
    
End Sub

Private Sub Class_Initialize()
    Initialise
End Sub

Private Sub Class_Terminate()
    Set m_ofwSQLExecuteHelper = Nothing
    Set m_colPersonSeeLinks = Nothing
    Set m_colPersonSeeAlsoLinks = Nothing
    ' specific clear up
End Sub

'IAFSaver implementation

Private Property Let IAFSaver_Data(t_vPersonAFTerm As Variant)
        m_typData = t_vPersonAFTerm
End Property


Public Function IsLinkedToOtherTerm() As Boolean
'   Check if terms is linked to/from other AF terms
'
'   d10 WDP 23-Feb-2000     Created
On Error GoTo ErrorHandler
    Dim sSource As String
    sSource = "BOAFPerson.IAFSaver_IsLinkedToOtherTerm"
    
    Dim sSql As String
    Dim ofwResults As FWRecordList
    Dim ofwHelper As New FWDBHelper
    
    sSql = " SELECT * FROM " & gk_sDBTABLEPersonLink & " WHERE " & _
            gk_sDBCOLPersonLinkRefID1 & "=" & m_typData.lID & " OR " & _
            gk_sDBCOLPersonLinkRefID2 & "=" & m_typData.lID
    
    Set ofwResults = ofwHelper.FindBySql(sSql)
    
    If ofwResults.RecordCount > 0 Then
        IsLinkedToOtherTerm = True
    Else
        IsLinkedToOtherTerm = False
    End If
    
    Exit Function
ErrorHandler:
    ConvertError sSource

End Function

Private Function IAFSaver_Save(t_lID As Long) As Long
'   This method Saves the data passed from the frontend.
'   Frontend will call the let property to initialise the collection. Load will pass the values from
'   the collection to backend.
'
'   Built in the same lines as IAFLoader send by WDP.
'   d1  Geedha Sivanadian     6Aug1999
'   d3  Sree Krishna          20Aug1999     Changed the using of SP to updated
'                                           statement
'   d4  Aravinda A            08Oct1999     Modified to take care of Nulls
'                                           Modified to add a new term
'   d5  Sree Krishna          04Nov1999     Call to reference data for refreshing after change
'   d6  Sree Krishna          25Nov1999     Replaced the Database error number
'                                           with a constant
'   d7  Aravinda A            08Dec1999     Modified rejection logic. Rejection now calls stored procedure SP_Edit_RejectAFTerm
'                                           Added Created_By field
'   d8  Aravinda A            21Dec1999     changing the datatype of birth date & death date to integer
'   d10 Mahesh CR             16-Feb-2000   Added check function NonPrefTermsRelPrefTermIsIncomOrApproved
    
On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim sSql As String
    Dim lCreatorID As Long 'Incase of adding new term
     Dim lReturnCode As Long 'd7

    sSource = "BOAFPerson.save"
    'd2 Handled biog histroy alone differently since it gave ODBC errors.
    

    If t_lID > 0 Then
        If m_typData.lAuthorityStatus = gk_lRejected Then
            'Rejection should call the SP_Edit_RejectAFTerm
            sSql = "SP_Edit_RejectAFTerm @tiAFTermType = " & EAFTermType.AFPerson _
                    & ", @tiAFTermID = " & t_lID
            lReturnCode = m_ofwSQLExecuteHelper.FindBySql(sSql).GetFieldValue(0)
            LogMessage sSource & ": Saved the given person details", pComplete
        Else
            'If the term is not rejected, update it
            sSql = "UPDATE " & gk_sDBTABLEPerson & " SET " & _
                    gk_sDBCOLSurname_Text & " = " & QuoteValueWithQuotes(m_typData.sSurname) & " ,"
            If m_typData.lAuthorityStatus <> 0 Then
                sSql = sSql & gk_sDBCOLAuthority_Status & " = " & StringToNullOrLong(m_typData.lAuthorityStatus) & " ,"
                'if the term's status is changed to Approved/ApprovedIncomplete
                'See desc of function NonPrefTermsRelPrefTermIsIncomOrApproved
                If m_typData.lAuthorityStatus = gk_lApprovedComplete Or m_typData.lAuthorityStatus = gk_lApprovedInComplete Then
                    NonPrefTermsRelPrefTermIsIncomOrApproved
                End If
            End If
            sSql = sSql & gk_sDBCOLGender_Indicator & " = " & QuoteValueWithQuotes(m_typData.sGender) & " ," & _
                    gk_sDBCOLPerson_Title & " = " & QuoteValueWithQuotes(m_typData.sTitle) & " ," & _
                    gk_sDBCOLadditional_elements_of_name & "= " & QuoteValueWithQuotes(m_typData.sAdditionalNameElements) & " ," & _
                    gk_sDBCOLPretitle_Text & " = " & QuoteValueWithQuotes(m_typData.sPretitle) & " ," & _
                    gk_sDBCOLForename_Text & " = " & QuoteValueWithQuotes(m_typData.sForenames) & " ," & _
                    gk_sDBCOLAlternative_Status_Text & " = " & QuoteValueWithQuotes(m_typData.sAlternativeStatus) & " ," & _
                    gk_sDBCOLBirth_Date & " = " & VariantToNullOrLong(m_typData.sBirthDate) & " ," & _
                    gk_sDBCOLUncertain_Birth_Date_Code & " = " & QuoteValueWithQuotes(m_typData.sUncertainBirthDateCode) & " ,"
            sSql = sSql & gk_sDBCOLBiog_History_Text & "= " & QuoteValueWithQuotes(m_typData.sBiogHistory) & " ," & _
                    gk_sDBCOLValidation_text & " = " & QuoteValueWithQuotes(m_typData.sValidation) & " ," & _
                    gk_sDBCOLPersonDeathDate & " = " & VariantToNullOrLong(m_typData.sDeathDate) & " ," & _
                    gk_sDBCOLUncertain_Death_Date_Code & " = " & QuoteValueWithQuotes(m_typData.sUncertainDeathDateCode)
            
            sSql = sSql & " WHERE " & gk_sDBCOLPersonReferenceID & "  =  " & t_lID
            m_ofwSQLExecuteHelper.Execute sSql
            lReturnCode = -1 'Updation
        End If
        
        LogMessage sSource & ": Saved the given person details", pComplete
    Else
        'add a new term: In our case t_lID is -1
        
        'Get the ID of the creator from the Session
        lCreatorID = CLng(GetDataFromSession(dsUserID))
        
        sSql = "INSERT INTO " & gk_sDBTABLEPerson & " (" & _
                    gk_sDBCOLSurname_Text & "," _
                    & gk_sDBCOLAuthority_Status & ", " _
                    & gk_sDBCOLGender_Indicator & ", " _
                    & gk_sDBCOLPerson_Title & ", " _
                    & gk_sDBCOLadditional_elements_of_name & ", " _
                    & gk_sDBCOLPretitle_Text & " , " _
                    & gk_sDBCOLForename_Text & ", " _
                    & gk_sDBCOLAlternative_Status_Text & ", " _
                    & gk_sDBCOLPersonBirthDate & ", " _
                    & gk_sDBCOLUncertain_Birth_Date_Code & ", "
                    
        sSql = sSql & gk_sDBCOLBiog_History_Text & ", " _
                    & gk_sDBCOLPersonValidation & ", " & gk_sDBCOLPersonDeathDate & ", " _
                    & gk_sDBCOLUncertain_Death_Date_Code & ", " & gk_sDBCOLAFCreated_By & ") " _
                    & "VALUES(" & QuoteValueWithQuotes(m_typData.sSurname) & " ," _
                    & StringToNullOrLong(m_typData.lAuthorityStatus) & " ," _
                    & QuoteValueWithQuotes(m_typData.sGender) & " ," _
                    & QuoteValueWithQuotes(m_typData.sTitle) & " ," _
                    & QuoteValueWithQuotes(m_typData.sAdditionalNameElements) & " ," _
                    & QuoteValueWithQuotes(m_typData.sPretitle) & " ," _
                    & QuoteValueWithQuotes(m_typData.sForenames) & " ," _
                    & QuoteValueWithQuotes(m_typData.sAlternativeStatus) & " ," _
                    & VariantToNullOrLong(m_typData.sBirthDate) & " ," _
                    & QuoteValueWithQuotes(m_typData.sUncertainBirthDateCode) & " ," _
                    & QuoteValueWithQuotes(m_typData.sBiogHistory) & " ," _
                    & QuoteValueWithQuotes(m_typData.sValidation) & " ," _
                    & VariantToNullOrLong(m_typData.sDeathDate) & " ," _
                    & QuoteValueWithQuotes(m_typData.sUncertainDeathDateCode) & ", " _
                    & lCreatorID & ")"
                    
           lReturnCode = m_ofwSQLExecuteHelper.InsertAndReturnIdentity(sSql)
           LogMessage sSource & ": Added new Person AF term", pComplete
    End If
    
    'Return the code
    IAFSaver_Save = lReturnCode
    Dim ofwReferenceData As FWReferenceData
    Set ofwReferenceData = New FWReferenceData
    ofwReferenceData.Refresh (EGroup.gPersonOnly)
    ofwReferenceData.Refresh (EGroup.gPerson)
    ofwReferenceData.Refresh (EGroup.gAllIndexTerms)
    Set ofwReferenceData = Nothing

    m_objContext.SetComplete
    
Exit Function
ErrorHandler:

    ConvertError sSource
    m_objContext.SetAbort

End Function

Private Sub IAFSaver_SaveLinks(t_lID As Long)
'Saves the links of the term t_lID to the database
'
'PARAMETERS: t_lID -> ID of the term whose links to be saved

'd4     aga     14-Oct-1999     Created
    On Error GoTo ErrorHandler
    Dim sSource As String
    Dim sSql As String
    Dim iCount As Integer
    
    sSource = Me.Name & ".IAFSaver_SaveLinks"
    
    '   Now delete all the items from Person_Link table for this person
    sSql = "DELETE " & gk_sDBTABLEPersonLink & " WHERE " & gk_sDBCOLPersonLinkRefID1 & "=" & t_lID
    m_ofwSQLExecuteHelper.Execute sSql
    
    'Insert the values into the database from the collection
    
    'Insert see links
    For iCount = 1 To m_colPersonSeeLinks.Count
    
        sSql = "INSERT INTO " & gk_sDBTABLEPersonLink & " (" _
            & gk_sDBCOLPersonLinkTypeID & ", " _
            & gk_sDBCOLPersonLinkRefID1 & ", " _
            & gk_sDBCOLPersonLinkRefID2 & ") " _
            & " VALUES ( " & EAFSeeLinkType.EAFSeeLink & "," & t_lID & "," & m_colPersonSeeLinks.Item(iCount).ID & ")"
            
        m_ofwSQLExecuteHelper.Execute sSql
        
    Next
    
    'Insert see also links
    For iCount = 1 To m_colPersonSeeAlsoLinks.Count
    
        sSql = "INSERT INTO " & gk_sDBTABLEPersonLink & " (" _
            & gk_sDBCOLPersonLinkTypeID & ", " _
            & gk_sDBCOLPersonLinkRefID1 & ", " _
            & gk_sDBCOLPersonLinkRefID2 & ") " _
            & " VALUES ( " & EAFSeeLinkType.EAFSeealsoLink & "," & t_lID & "," & m_colPersonSeeAlsoLinks.Item(iCount).ID & ")"
            
        m_ofwSQLExecuteHelper.Execute sSql
        
    Next
    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub
Public Property Get Name()
    Name = "BOAFPerson"
End Property

Private Property Let IAFSaver_SeeAlsoLinksData(t_vSeeAlsoLinks As Collection)
'd4
    Set m_colPersonSeeAlsoLinks = t_vSeeAlsoLinks
End Property

Private Property Let IAFSaver_SeeLinksData(t_vSeeLinks As Collection)
'd4
    Set m_colPersonSeeLinks = t_vSeeLinks
End Property

'
' PROCATTransientInterface methods (since this object needs to be transiently stored)
'
Public Sub PROCATTransientInterface_DeleteState(Optional t_sInstanceReference As String)
'   Create an instance reference from SessionID, Class name and optional name.

    Dim sInstanceReference As String
    Dim ofwTransientManager As FWTransientManager

    sInstanceReference = CreateDelimitedString(Me.Name, t_sInstanceReference)

    Set ofwTransientManager = New FWTransientManager
    
    ofwTransientManager.DeleteData GetObjectContext.Item(gk_sSessionObject).SessionID, _
                                    sInstanceReference
    
    LogMessage "Deleted BOPerson state" & sInstanceReference, pComplete
    
End Sub

Public Sub PROCATTransientInterface_GetState(Optional t_sInstanceReference As String)
'   Create an instance reference from SessionID, Class name and optional name.
'   Retrieve serialised string from FWTransientManager, and use to populate state.

'd4     aga     14-Oct-1999     Added collections
    Dim sInstanceReference As String
    Dim sSerialisedString As String
    Dim ofwTransientManager As FWTransientManager

    sInstanceReference = CreateDelimitedString(Me.Name, t_sInstanceReference)

    Set ofwTransientManager = New FWTransientManager
    sSerialisedString = _
        ofwTransientManager.GetData(GetObjectContext.Item(gk_sSessionObject).SessionID, sInstanceReference)
    If Trim(sSerialisedString) = "" Then
        LogMessage "BOPerson state not found " & sInstanceReference, EProfile.pProfileInformation
    Else
        m_typData.lID = CLng(GetDataFromSerialisedString(sSerialisedString))
        m_typData.lAuthorityStatus = CLng(GetDataFromSerialisedString(sSerialisedString))
        m_typData.sSurname = GetDataFromSerialisedString(sSerialisedString)
        m_typData.sGender = GetDataFromSerialisedString(sSerialisedString)
        m_typData.sTitle = GetDataFromSerialisedString(sSerialisedString)
        m_typData.sAdditionalNameElements = GetDataFromSerialisedString(sSerialisedString)
        m_typData.sPretitle = GetDataFromSerialisedString(sSerialisedString)
        m_typData.sForenames = GetDataFromSerialisedString(sSerialisedString)
        m_typData.sAlternativeStatus = GetDataFromSerialisedString(sSerialisedString)
        m_typData.sBirthDate = GetDataFromSerialisedString(sSerialisedString)
        m_typData.sUncertainBirthDateCode = GetDataFromSerialisedString(sSerialisedString)
        m_typData.sBiogHistory = GetDataFromSerialisedString(sSerialisedString)
        m_typData.sValidation = GetDataFromSerialisedString(sSerialisedString)
        m_typData.sDeathDate = GetDataFromSerialisedString(sSerialisedString)
        m_typData.sUncertainDeathDateCode = GetDataFromSerialisedString(sSerialisedString)
        
        'd4 Collections
        Set m_colPersonSeeLinks = ExtractLinkItemsFromSerialIsedString(sSerialisedString)
        Set m_colPersonSeeAlsoLinks = ExtractLinkItemsFromSerialIsedString(sSerialisedString)
        
        LogMessage "Read BOPerson state" & sInstanceReference, EProfile.pComplete
    End If

End Sub

Public Sub PROCATTransientInterface_SetState(Optional t_sInstanceReference As String)
'   Create an instance reference from SessionID, Class name and optional name.
'   Construct serialised string from state, and use FWTransientManager to store.
'd4     aga     13-Oct-1999     Added collections
    Dim sInstanceReference As String
    Dim sSerialisedString As String
    Dim ofwTransientManager As FWTransientManager

    sInstanceReference = CreateDelimitedString(Me.Name, t_sInstanceReference)

    sSerialisedString = ""
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, CStr(m_typData.lID))
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, CStr(m_typData.lAuthorityStatus))
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, m_typData.sSurname)
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, m_typData.sGender)
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, m_typData.sTitle)
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, m_typData.sAdditionalNameElements)
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, m_typData.sPretitle)
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, m_typData.sForenames)
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, m_typData.sAlternativeStatus)
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, m_typData.sBirthDate)
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, m_typData.sUncertainBirthDateCode)
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, m_typData.sBiogHistory)
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, m_typData.sValidation)
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, m_typData.sDeathDate)
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, m_typData.sUncertainDeathDateCode)
    
    'd4 collections
    sSerialisedString = SerialiseLinkItem(sSerialisedString, m_colPersonSeeLinks)
    sSerialisedString = SerialiseLinkItem(sSerialisedString, m_colPersonSeeAlsoLinks)
    
    Set ofwTransientManager = New FWTransientManager
    ofwTransientManager.AddData GetObjectContext.Item(gk_sSessionObject).SessionID, _
                                sInstanceReference, sSerialisedString

    LogMessage "Stored BOPerson state" & sInstanceReference, EProfile.pComplete

End Sub



Public Sub Load( _
    ByVal t_eLoadType As EAFTermLoadType, _
    Optional ByVal t_lID As Long, Optional ByVal t_lLinkTermID As Long, _
    Optional ByVal t_sLinkTermValue As String, Optional t_eAFSeelinktype As EAFSeeLinkType)

'
'Description: Loads the details of the bo,depending on the value of t_eLoadType either the details are
'loaded from the loader class or from the stored stage
'
'Amendment history:
'd4     aga     14-Oct-1999     Created
'd7     aga     10-Dec-1999     Check for links before adding a see link
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String
     'd7
    Dim iHasLinks As Integer
    Dim sSql As String
    Dim eSeelinkType As EAFSeeLinkType
    
    
    sSource = "BOAFPerson.Load"
    
    
    'The user clicked on Edit button and control has come here
    If t_eLoadType = EAFTermLoadType.EAFLoadFromDatabase Then
        
        'Delete the state information (if any)
        PROCATTransientInterface_DeleteState
        
        'Load from the loader class
        LoadFromDatabase t_lID
    Else
    
        Select Case t_eLoadType
        
            Case EAFTermLoadType.EAFLoadFromStateAddTerm
            'The user has selected a pick term to add and the control has come here
                
                'Read the stored state info
                PROCATTransientInterface_GetState
                'Add the new link term to the collection
                
                
                If IsMissing(t_eAFSeelinktype) Then
                    eSeelinkType = EAFSeeLink
                Else
                    eSeelinkType = t_eAFSeelinktype
                End If
'
                If eSeelinkType = EAFSeeLink Then
'                    'Check if the term has links to the catelogues
                    sSql = "SP_Edit_IsAFTermLinkedToCatalogueEntry @tiAFTermType=" & EAFTermType.AFPerson _
                        & ", " & "@tiAFTermID=" & m_typData.lID
                    iHasLinks = m_ofwSQLExecuteHelper.FindBySql(sSql).GetFieldValue(0)
                    If iHasLinks = 0 Then
                        AddNewLinkTerm t_lLinkTermID, t_sLinkTermValue, EAFSeeLink
                    End If
                Else
                    AddNewLinkTerm t_lLinkTermID, t_sLinkTermValue, EAFSeealsoLink
                End If
                PROCATTransientInterface_SetState
            
            Case EAFTermLoadType.EAFLoadFromStateRemoveTerm
                'The user has selected a pick term to remove and the control has come here
                
                'Read the stored state info
                PROCATTransientInterface_GetState
                            
                'Remove the link term from the collection
                If Not IsMissing(t_eAFSeelinktype) Then
                    RemoveLinkTerm t_lLinkTermID, t_eAFSeelinktype
                Else
                    RemoveLinkTerm t_lLinkTermID
                End If
                
            Case EAFTermLoadType.EAFLoadFromStateTermNoChange
                'The user has selected cancel operation while adding/removing a link term and the control has come here
                
                'Just Read the stored stage info
                PROCATTransientInterface_GetState
            Case Else
                PROCATTransientInterface_GetState
                
        End Select
        
    End If

    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub

Private Sub LoadFromDatabase(ByVal t_lID As Long)
'
'Description: Loads the object details from the loader class
'
' Amendment history:
'd4     aga     14-Oct-1999     Created
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim oBOAFPersonLoader As BOAFPersonLoader
    
    sSource = "BOAFPerson.LoadFromDatabase"
    
    Set oBOAFPersonLoader = New BOAFPersonLoader
        
    'Set the properties
    oBOAFPersonLoader.IAFLoader_Load t_lID
    m_typData = oBOAFPersonLoader.Data
    
    'Now load the see links
    oBOAFPersonLoader.LoadLinks t_lID
    Set m_colPersonSeeLinks = oBOAFPersonLoader.SeeLinks
    Set m_colPersonSeeAlsoLinks = oBOAFPersonLoader.SeeAlsoLinks
    
    'Release the objects
    Set oBOAFPersonLoader = Nothing
    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub

Private Sub AddNewLinkTerm( _
    t_lLinkTermID As Long, t_sLinkTermValue As String, _
    Optional t_eAFSeelinktype As EAFSeeLinkType = EAFSeeLinkType.EAFSeeLink)
'
'Description: Adds a link item to the collection
'
'Amendment history:
'   d4  Aravinda A            14-Oct-1999     Created
'   d7  Aravinda A            08-Dec-1999     Limited see links to one
'
    Dim sSource As String
    Dim recLinkItemRecord  As GenericAFLinkTerm
    Dim iTemp As Integer
    
    
    sSource = "BOAFPerson.AddNewLinkTerm"
    On Error GoTo ErrorHandler
    recLinkItemRecord.ID = t_lLinkTermID
    recLinkItemRecord.Value = t_sLinkTermValue
    If t_eAFSeelinktype = EAFSeealsoLink Then
        m_colPersonSeeAlsoLinks.Add recLinkItemRecord, CStr(recLinkItemRecord.ID)
    Else
        'd7
        'Remove items from existing collection and add the new one
        For iTemp = 1 To m_colPersonSeeLinks.Count
            m_colPersonSeeLinks.Remove 1    'Always use index as 1 as the col is reindexed
        Next
        m_colPersonSeeLinks.Add recLinkItemRecord, CStr(recLinkItemRecord.ID)
    End If
    
    Exit Sub
ErrorHandler:
    If Err.Number = 457 Then
        'If the item already exists in the collection, don't do any thing
        Exit Sub
    End If
    ConvertError sSource
End Sub


Private Sub RemoveLinkTerm( _
    t_lLinkTermID As Long, Optional t_eAFSeelinktype As EAFSeeLinkType = EAFSeeLinkType.EAFSeeLink)
'
'Description: Removes the link item from the collection
'
'Amendment history:
'd4     aga     14-Oct-1999     Created
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim recLinkItemRecord  As GenericAFLinkTerm
    sSource = "BOAFPerson.RemoveLinkTerm"
    On Error GoTo ErrorHandler
    If t_eAFSeelinktype = EAFSeeLinkType.EAFSeealsoLink Then
        m_colPersonSeeAlsoLinks.Remove CStr(t_lLinkTermID)
    Else
        m_colPersonSeeLinks.Remove CStr(t_lLinkTermID)
    End If

    Exit Sub

ErrorHandler:
    If Err.Number = 5 Then
        'If no more links exits, the remove gives this error.
        'Do nothing then.
        Exit Sub
    End If
    ConvertError sSource
End Sub

Public Function SerialiseLinkItem( _
    ByVal t_sSerialIsedLinkItem As String, _
    ByVal t_vColLinkItem As Collection) As String
'
'Description: Convert the link items (id and description) stored in the collection into a string
'             to be stored in persistence manager (except related and separate material)
'
' Amendment history:
'd4     aga     14-Oct-1999     Created
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim recLinkItemRecord  As Variant
                
    sSource = "BOAFPerson.SerialiseLinkItem"
    
    SerialiseLinkItem = t_sSerialIsedLinkItem
    
    'If there are no items in the collection store 0 in the serialised string
    If t_vColLinkItem.Count = 0 Then
        SerialiseLinkItem = AddDataToSerialisedString(SerialiseLinkItem, CStr(0))
    Else
        'Store the number of items in the collection in the string
        SerialiseLinkItem = AddDataToSerialisedString(SerialiseLinkItem, CStr(t_vColLinkItem.Count))
        For Each recLinkItemRecord In t_vColLinkItem
            'Serialise the id
            SerialiseLinkItem = _
                AddDataToSerialisedString(SerialiseLinkItem, CStr(recLinkItemRecord.ID))
            'Serialise the description
            SerialiseLinkItem = _
                AddDataToSerialisedString(SerialiseLinkItem, VariantStringAsString(recLinkItemRecord.Value))
        Next
    End If
    
    Exit Function
    
ErrorHandler:
    ConvertError sSource
End Function
Public Function ExtractLinkItemsFromSerialIsedString( _
        ByRef t_sSerialisedString As String) As Collection
'
'Description: Convert the link items (id and description) stored in the string into
'             a collection (except related and separate material)
'
' Amendment history:
'd4     aga     14-Oct-1999     Created
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim recLinkItemRecord  As GenericAFLinkTerm, iTotalRecords As Integer, iCounter As Integer
    
    sSource = "BOAFPerson.ExtractLinkItemsFromSerialIsedString"
    
    Set ExtractLinkItemsFromSerialIsedString = New Collection
    
    'Get the total items to be put in the collection,  if there are 0 items then the collection will remain empty
    iTotalRecords = CInt(GetDataFromSerialisedString(t_sSerialisedString))
    
    For iCounter = 1 To iTotalRecords
        'Get the id
        recLinkItemRecord.ID = CLng(GetDataFromSerialisedString(t_sSerialisedString))
        'Get the description
        recLinkItemRecord.Value = GetDataFromSerialisedString(t_sSerialisedString)
        
        'Store it on the collection
        ExtractLinkItemsFromSerialIsedString.Add recLinkItemRecord, CStr(recLinkItemRecord.ID)
    Next
    
    Exit Function
    
ErrorHandler:
    ConvertError sSource
End Function

Public Function IsTermHasLinks() As Boolean

'   Checks whether the term has links to the catelogues
'
'd4     aga     20-Jan-2000     Created
    Dim iHasLinks As Integer
    Dim sSql As String
    
    IsTermHasLinks = True
    'Check if the term has links to the catelogues
    sSql = "SP_Edit_IsAFTermLinkedToCatalogueEntry @tiAFTermType=" & EAFTermType.AFPerson _
        & ", " & "@tiAFTermID=" & m_typData.lID
    iHasLinks = m_ofwSQLExecuteHelper.FindBySql(sSql).GetFieldValue(0)
    If iHasLinks = 0 Then
        IsTermHasLinks = False
    End If
End Function

Private Sub NonPrefTermsRelPrefTermIsIncomOrApproved()
'Description    :   Check to see that when a Non_prefered term's status
'                   is being changed to InComplete or Approved the realted
'                   prefereed terms linked to it should be Incomplete or
'                   approved .
'   d10 Mahesh CR             16-Feb-2000    Added function
On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim oboPersonLd As New BOAFPersonLoader
    Dim sDesc As String
    sSource = Me.Name & ".NonPrefTermsRelPrefTermIsIncomOrApproved"
    
    
    sDesc = "You cannot change the status of this term because the preferred term associated with it will not be visible to the Reader"
    'if this is a non preferred term
    If m_colPersonSeeLinks.Count > 0 Then
        oboPersonLd.IAFLoader_Load m_colPersonSeeLinks.Item(1).ID
        'and if the pref term has a status of ApprovedComplete or ApprovedIncomplete
        If Not (oboPersonLd.Data.lAuthorityStatus = gk_lApprovedComplete Or oboPersonLd.Data.lAuthorityStatus = gk_lApprovedInComplete) Then
              'Raise an error
              RaiseError sSource, PECanNotMakeNonPrefTermIcompApprovWithPrefTermIcompApprov, sDesc
        End If
    End If
    Set oboPersonLd = Nothing
Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub

