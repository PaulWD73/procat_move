VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BOAFCorporateBody"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    BOAFCorporateBody.cls
' System:       PRO, PROCAT
' Copyright:    (C) Quidnunc Limited
'
' Description:  This class encapsultes CorporateBody AF term.
'               Written in the same lines as BOLoader class send by WDP.
'
' Amendment history:
'   d1  Geedha Sivanadian     11-Aug-1999      Created
'   d2  Geedha Sivanadian     12-Aug-1999      handled text area values differently
'                                              since ODBC was giving errors.
'   d3  Sree Krishna          20-Aug-1999      Changed the using of SP to updated
'                                              statement
'   d4  AGA                   08-Oct-1999      Modifed Save method to take care of nulls
'                                              Modified Save to add a new term
'                                              Added methods and properties for See/See also links
'                                              Replaced hardcoded table and column constants
'   d5  Sree Krishna          04-Nov-1999      Call to reference data for refreshing after change
'   d6  Sree Krishna          25-Nov-1999      Replaced the Database error number
'                                              with a constant
'   d7  Aravinda A            08-Dec-1999      Limited see links to one
'   d8  Aravinda A            22-Dec-1999      Defect fixes (minor)
'   d9  Aravinda A            20-Jan-2000      Added function IsTermHasLinks
'   d10 Mahesh CR             16-Feb-2000      Added function NonPrefTermsRelPrefTermIsIncomOrApproved
'   d11 WDP                   23-Feb-2000      Added function IsLinkedToOtherTerm
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


Private m_typData As CorporateBodyAFTerm
'd4
Private m_colCorporateBodySeeLinks As Collection
Private m_colCorporateBodySeeAlsoLinks As Collection

Public Property Let Data(t_typCorporateBodyAFTerm As CorporateBodyAFTerm)
    IAFSaver_Data = t_typCorporateBodyAFTerm
End Property

Public Property Get Data() As CorporateBodyAFTerm
    Data = m_typData
End Property

Public Property Get SeeLinks() As Collection
'd4
    Set SeeLinks = m_colCorporateBodySeeLinks
End Property

Public Property Get SeeAlsoLinks() As Collection
'd4
    Set SeeAlsoLinks = m_colCorporateBodySeeAlsoLinks
End Property


Private Sub Initialise()
    ' Specific initialisation
    Set m_ofwSQLExecuteHelper = New FWDBHelper
    Set m_objContext = GetObjectContext()
    'aga
    Set m_colCorporateBodySeeLinks = New Collection
    Set m_colCorporateBodySeeAlsoLinks = New Collection
    
End Sub

Private Sub Class_Initialize()
    Initialise
End Sub

Private Sub Class_Terminate()
    Terminate
End Sub
Private Sub Terminate()
'd4
    Set m_ofwSQLExecuteHelper = Nothing
    ' specific clear up
    Set m_colCorporateBodySeeLinks = Nothing
    Set m_colCorporateBodySeeAlsoLinks = Nothing
End Sub
'IAFSaver implementation

Private Property Let IAFSaver_Data(t_vCorporateBodyAFTerm As Variant)
        m_typData = t_vCorporateBodyAFTerm
End Property


Private Function IAFSaver_IsLinkedToOtherTerm() As Boolean

End Function

Private Function IAFSaver_Save(t_lID As Long) As Long
'   This method Saves the data passed from the frontend.
'   Frontend will call the let property to initialise the collection. Load will pass the values from
'   the collection to backend.
'
'   RETURNS: ID of the term inserted in case of insertion Or -1 in case of updation, other values in case of rejection

'   Built in the same lines as IAFLoader send by WDP.
'   d1  Geedha Sivanadian     6Aug1999
'   d3  Sree Krishna          20Aug1999
'   d4  AGA                   08Oct1999       Modified the SQL query string to take care of nulls
'                                             Modified to add a new term
'   d6  Sree Krishna          25-Nov-1999     Replaced the Database error number
'                                             with a constant
'   d7  Aravinda A            09-Dec-1999     Added Created_By field, added rejection criteria
'   d8  Aravinda A            21-Dec-1999     changing the datatype of start date & end date to integer
'   d10 Mahesh CR             16-Feb-2000   Added check function NonPrefTermsRelPrefTermIsIncomOrApproved
On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim sSql As String
    Dim lCreatorID As Long 'd7
    Dim lReturnCode As Long

    sSource = "BOAFCorporateBody.save"
    If t_lID > 0 Then
        'd4
        
        If m_typData.lAuthorityStatus = gk_lRejected Then
            
            'Rejection should call the SP_Edit_RejectAFTerm
            sSql = "SP_Edit_RejectAFTerm @tiAFTermType = " & EAFTermType.AFCorporateBody _
                    & ", @tiAFTermID = " & t_lID
            lReturnCode = m_ofwSQLExecuteHelper.FindBySql(sSql).GetFieldValue(0)
        Else
            'If the term is not rejected, update it
            
            sSql = "Update " & gk_sDBTABLECorporateBody & " SET "
            
            If m_typData.lAuthorityStatus <> 0 Then
                sSql = sSql & gk_sDBCOLAuthority_Status & " = " & StringToNullOrLong(m_typData.lAuthorityStatus) & " ,"
                'if the term's status is changed to Approved/ApprovedIncomplete
                'See desc of function NonPrefTermsRelPrefTermIsIncomOrApproved
                If m_typData.lAuthorityStatus = gk_lApprovedComplete Or m_typData.lAuthorityStatus = gk_lApprovedInComplete Then
                    NonPrefTermsRelPrefTermIsIncomOrApproved
                End If
            End If
            
            sSql = sSql & gk_sDBCOLCorporateBodyName & " = " & QuoteValueWithQuotes(m_typData.sTermText) & " ," & _
                    gk_sDBCOLCorporateBodyDateEnd & "  = " & VariantToNullOrLong(m_typData.sEndDate) & " ," & _
                    gk_sDBCOLRemitFunction & "  = " & QuoteValueWithQuotes(m_typData.sRemitFunction) & " ," & _
                    gk_sDBCOLCorporateBodyDateStart & " = " & VariantToNullOrLong(m_typData.sStartDate) & " ," & _
                    gk_sDBCOLCorporateBodyUncertainStartDate & " = " & QuoteValueWithQuotes(m_typData.sUncertainStartDateCode) & " ," & _
                    gk_sDBCOLVariantText & " = " & QuoteValueWithQuotes(m_typData.sVariantText) & " ," & _
                    gk_sDBCOLCorporateBodyValidation & " = " & QuoteValueWithQuotes(m_typData.sValidation) & " ," & _
                    gk_sDBCOLCorporateBodyUncertainEndDate & " = " & QuoteValueWithQuotes(m_typData.sUncertainEndDateCode) & " ," & _
                    gk_sDBCOLJurisdiction & "  = " & QuoteValueWithQuotes(m_typData.sJurisdiction) & " ," & _
                    gk_sDBCOLCorporateBodyIsPlaceOfDep & " = " & IIf(m_typData.bIsPlaceOfDeposit, 1, 0) & " ," & _
                    gk_sDBCOLNationalPlaceOfDepositCode & " = " & QuoteValueWithQuotes(m_typData.sPlaceOfDepositCode)
        
            sSql = sSql & " WHERE " & gk_sDBCOLCorporateBodyReferenceID & " =  " & t_lID
            
            m_ofwSQLExecuteHelper.Execute (sSql)
            IAFSaver_Save = -1 'Updation
        End If
        
        
        
        LogMessage sSource & ": Saved the given CorporateBody details", pComplete
    Else
        'Add new term if the t_lID <= 0. In our case t_lID = -1 for this purpose
        'Get the ID of the creator from the Session
        lCreatorID = CLng(GetDataFromSession(dsUserID))
        
        sSql = "INSERT INTO " & gk_sDBTABLECorporateBody & " (" _
                    & gk_sDBCOLCorporateBodyName & ", " _
                    & gk_sDBCOLAuthority_Status & ", " _
                    & gk_sDBCOLCorporateBodyDateEnd & ", " _
                    & gk_sDBCOLRemitFunction & ", " _
                    & gk_sDBCOLCorporateBodyDateStart & ", " _
                    & gk_sDBCOLCorporateBodyUncertainStartDate & ", " _
                    & gk_sDBCOLVariantText & ", " _
                    & gk_sDBCOLCorporateBodyValidation & ", " _
                    & gk_sDBCOLCorporateBodyUncertainEndDate & ", " _
                    & gk_sDBCOLJurisdiction & ", " _
                    & gk_sDBCOLCorporateBodyIsPlaceOfDep & ", " _
                    & gk_sDBCOLNationalPlaceOfDepositCode & ", " _
                    & gk_sDBCOLAFCreated_By & ") "
                    
        sSql = sSql & "VALUES(" & QuoteValueWithQuotes(m_typData.sTermText) & " ," _
                    & StringToNullOrLong(m_typData.lAuthorityStatus) & " ," _
                    & VariantToNullOrLong(m_typData.sEndDate) & " ," _
                    & QuoteValueWithQuotes(m_typData.sRemitFunction) & " ," _
                    & VariantToNullOrLong(m_typData.sStartDate) & " ," _
                    & QuoteValueWithQuotes(m_typData.sUncertainStartDateCode) & " ," _
                    & QuoteValueWithQuotes(m_typData.sVariantText) & " ," _
                    & QuoteValueWithQuotes(m_typData.sValidation) & " ," _
                    & QuoteValueWithQuotes(m_typData.sUncertainEndDateCode) & " ," _
                    & QuoteValueWithQuotes(m_typData.sJurisdiction) & " ," _
                    & IIf(m_typData.bIsPlaceOfDeposit, 1, 0) & " ," _
                    & QuoteValueWithQuotes(m_typData.sPlaceOfDepositCode) & "," _
                    & lCreatorID & ")"
                    
           'IAFSaver_Save = m_ofwSQLExecuteHelper.InsertAndReturnIdentity(sSql)
           lReturnCode = m_ofwSQLExecuteHelper.InsertAndReturnIdentity(sSql)
           
           LogMessage sSource & ": Added new Corporate Body AF term", pComplete
    End If
    'RETURN THE code
    IAFSaver_Save = lReturnCode
    Dim ofwReferenceData As FWReferenceData
    Set ofwReferenceData = New FWReferenceData
    ofwReferenceData.Refresh (EGroup.gCorporateBodyOnly)
    ofwReferenceData.Refresh (EGroup.gCorporateBody)
    ofwReferenceData.Refresh (EGroup.gAllIndexTerms)
    Set ofwReferenceData = Nothing
    
    m_objContext.SetComplete
    
Exit Function
ErrorHandler:

    ConvertError sSource
    m_objContext.SetAbort

End Function

Private Sub IAFSaver_SaveLinks(t_lID As Long)
'   Saves the links of the term
'PARAMETERS: t_lID -> ID of the term whose links to be saved

'd4     AGA    14 Oct 1999      Created
    On Error GoTo ErrorHandler
    Dim sSource As String
    Dim sSql As String
    Dim iCount As Integer
    
    sSource = Me.Name & ".IAFSaver_SaveLinks"
    
    ' delete all the items from Person_Link table for this person
    sSql = "DELETE " & gk_sDBTABLECorpBodyLink & " WHERE " & gk_sDBCOLCorpBodyLinkRefID1 & "=" & t_lID
    m_ofwSQLExecuteHelper.Execute sSql
    
    '   Insert the see link values into the database from the collection
    For iCount = 1 To m_colCorporateBodySeeLinks.Count
    
        sSql = "INSERT INTO " & gk_sDBTABLECorpBodyLink _
            & " (" & gk_sDBCOLCorpBodyLinkTypeID & ", " _
            & gk_sDBCOLCorpBodyLinkRefID1 & ", " _
            & gk_sDBCOLCorpBodyLinkRefID2 & ") " _
            & " VALUES ( " & EAFSeeLinkType.EAFSeeLink & "," & t_lID & "," & m_colCorporateBodySeeLinks.Item(iCount).ID & ")"
            
        m_ofwSQLExecuteHelper.Execute sSql
    Next
    
    '   Insert the see also link values into the database from the collection
    For iCount = 1 To m_colCorporateBodySeeAlsoLinks.Count
    
        sSql = "INSERT INTO " & gk_sDBTABLECorpBodyLink _
            & " (" & gk_sDBCOLCorpBodyLinkTypeID & ", " _
            & gk_sDBCOLCorpBodyLinkRefID1 & ", " _
            & gk_sDBCOLCorpBodyLinkRefID2 & ") " _
            & " VALUES ( " & EAFSeeLinkType.EAFSeealsoLink & "," & t_lID & "," & m_colCorporateBodySeeAlsoLinks.Item(iCount).ID & ")"
            
        m_ofwSQLExecuteHelper.Execute sSql
    Next
    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub
Public Property Get Name()
    Name = "BOAFCorporateBody"
End Property

Private Property Let IAFSaver_SeeAlsoLinksData(t_vSeeAlsoLinks As Collection)
    Set m_colCorporateBodySeeAlsoLinks = t_vSeeAlsoLinks
End Property

Private Property Let IAFSaver_SeeLinksData(t_vSeeLinks As Collection)
    Set m_colCorporateBodySeeLinks = t_vSeeLinks
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
    
    LogMessage "Deleted BOCorporateBody state" & sInstanceReference, pComplete
    
End Sub

Public Sub PROCATTransientInterface_GetState(Optional t_sInstanceReference As String)
'   Create an instance reference from SessionID, Class name and optional name.
'   Retrieve serialised string from FWTransientManager, and use to populate state.

'd4     aga     13-Oct-1999     Added collections
    Dim sInstanceReference As String
    Dim sSerialisedString As String
    Dim ofwTransientManager As FWTransientManager

    sInstanceReference = CreateDelimitedString(Me.Name, t_sInstanceReference)

    Set ofwTransientManager = New FWTransientManager
    sSerialisedString = _
        ofwTransientManager.GetData(GetObjectContext.Item(gk_sSessionObject).SessionID, sInstanceReference)
    If Trim(sSerialisedString) = "" Then
        LogMessage "BOCorporateBody state not found " & sInstanceReference, EProfile.pProfileInformation
    Else
        m_typData.lID = CLng(GetDataFromSerialisedString(sSerialisedString))
        m_typData.lAuthorityStatus = CLng(GetDataFromSerialisedString(sSerialisedString))
        m_typData.bIsPlaceOfDeposit = GetDataFromSerialisedString(sSerialisedString)
        m_typData.sEndDate = GetDataFromSerialisedString(sSerialisedString)
        m_typData.sJurisdiction = GetDataFromSerialisedString(sSerialisedString)
        m_typData.sPlaceOfDepositCode = GetDataFromSerialisedString(sSerialisedString)
        m_typData.sRemitFunction = GetDataFromSerialisedString(sSerialisedString)
        m_typData.sStartDate = GetDataFromSerialisedString(sSerialisedString)
        m_typData.sTermText = GetDataFromSerialisedString(sSerialisedString)
        m_typData.sUncertainEndDateCode = GetDataFromSerialisedString(sSerialisedString)
        m_typData.sUncertainStartDateCode = GetDataFromSerialisedString(sSerialisedString)
        m_typData.sValidation = GetDataFromSerialisedString(sSerialisedString)
        m_typData.sVariantText = GetDataFromSerialisedString(sSerialisedString)
       
        'd4 Collections
        Set m_colCorporateBodySeeLinks = ExtractLinkItemsFromSerialIsedString(sSerialisedString)
        Set m_colCorporateBodySeeAlsoLinks = ExtractLinkItemsFromSerialIsedString(sSerialisedString)
        
        LogMessage "Read BOCorporateBody state" & sInstanceReference, EProfile.pComplete
    End If

End Sub

Public Sub PROCATTransientInterface_SetState(Optional t_sInstanceReference As String)
'   Create an instance reference from SessionID, Class name and optional name.
'   Construct serialised string from state, and use FWTransientManager to store.
'
'd4     aga     13-Oct-1999     Added collections
    Dim sInstanceReference As String
    Dim sSerialisedString As String
    Dim ofwTransientManager As FWTransientManager

    sInstanceReference = CreateDelimitedString(Me.Name, t_sInstanceReference)

    sSerialisedString = ""
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, CStr(m_typData.lID))
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, CStr(m_typData.lAuthorityStatus))
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, m_typData.bIsPlaceOfDeposit)
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, m_typData.sEndDate)
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, m_typData.sJurisdiction)
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, m_typData.sPlaceOfDepositCode)
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, m_typData.sRemitFunction)
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, m_typData.sStartDate)
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, m_typData.sTermText)
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, m_typData.sUncertainEndDateCode)
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, m_typData.sUncertainStartDateCode)
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, m_typData.sValidation)
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, m_typData.sVariantText)
    
    'd4  see/Seealso links collection
    sSerialisedString = SerialiseLinkItem(sSerialisedString, m_colCorporateBodySeeLinks)
    sSerialisedString = SerialiseLinkItem(sSerialisedString, m_colCorporateBodySeeAlsoLinks)
    
    Set ofwTransientManager = New FWTransientManager
    ofwTransientManager.AddData GetObjectContext.Item(gk_sSessionObject).SessionID, _
                                sInstanceReference, sSerialisedString

    LogMessage "Stored BOCorporateBody state" & sInstanceReference, EProfile.pComplete

End Sub

Public Sub Load( _
    ByVal t_eLoadType As EAFTermLoadType, _
    Optional ByVal t_lID As Long, Optional ByVal t_lLinkTermID As Long, _
    Optional ByVal t_sLinkTermValue As String, Optional t_eAFSeelinktype As EAFSeeLinkType)
'AGA : TEST
'
'Description: Loads the details of the bo,depending on the value of t_eLoadType either the details are
'loaded from the loader class or from the stored stage
'
'Amendment history:

'd4     aga     13-Oct-1999     Created
'd7     aga     10-Dec-1999     Check for links before adding a see link

    On Error GoTo ErrorHandler
    
    Dim sSource As String
    'd7
    Dim iHasLinks As Integer
    Dim sSql As String
    Dim eSeelinkType As EAFSeeLinkType
    sSource = Me.Name & ".Load"
    
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
                'd7
                If IsMissing(t_eAFSeelinktype) Then
                    eSeelinkType = EAFSeeLink
                Else
                    eSeelinkType = t_eAFSeelinktype
                End If
                'd7
                If eSeelinkType = EAFSeeLink Then
                    'Check if the term has links to the catelogues
                    sSql = "SP_Edit_IsAFTermLinkedToCatalogueEntry @tiAFTermType=" & EAFTermType.AFCorporateBody _
                        & ", " & "@tiAFTermID=" & m_typData.lID
                    iHasLinks = m_ofwSQLExecuteHelper.FindBySql(sSql).GetFieldValue(0)
                    If iHasLinks = 0 Then
                        AddNewLinkTerm t_lLinkTermID, t_sLinkTermValue, EAFSeeLink
                    End If
                Else
                    AddNewLinkTerm t_lLinkTermID, t_sLinkTermValue, EAFSeealsoLink
                End If
                
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
'd4     aga     13-Oct-1999     Created
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim oBOAFCorporateBodyLoader As BOAFCorpBodyLoader
    
    sSource = Me.Name & ".LoadFromDatabase"
    
    Set oBOAFCorporateBodyLoader = New BOAFCorpBodyLoader
        
    'Set the properties
    oBOAFCorporateBodyLoader.IAFLoader_Load t_lID
    m_typData = oBOAFCorporateBodyLoader.Data
    
    'Now load the see links
    oBOAFCorporateBodyLoader.LoadLinks t_lID
    Set m_colCorporateBodySeeLinks = oBOAFCorporateBodyLoader.SeeLinks
    Set m_colCorporateBodySeeAlsoLinks = oBOAFCorporateBodyLoader.SeeAlsoLinks
    
    'Release the objects
    Set oBOAFCorporateBodyLoader = Nothing
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
'   d4  Aravinda A            13-Oct-1999     Created
'   d7  Aravinda A            08-Dec-1999     Limited see links to one
'
    Dim sSource As String
    Dim recLinkItemRecord  As GenericAFLinkTerm
    Dim iTemp As Integer
    
    sSource = Me.Name & ".AddNewLinkTerm"
    On Error GoTo ErrorHandler
    recLinkItemRecord.ID = t_lLinkTermID
    recLinkItemRecord.Value = t_sLinkTermValue
    If t_eAFSeelinktype = EAFSeealsoLink Then
        m_colCorporateBodySeeAlsoLinks.Add recLinkItemRecord, CStr(recLinkItemRecord.ID)
    Else
        'd7
        'Remove items from existing collection and add the new one
        For iTemp = 1 To m_colCorporateBodySeeLinks.Count
            m_colCorporateBodySeeLinks.Remove 1    'Always use index as 1 as the col is reindexed
        Next
        m_colCorporateBodySeeLinks.Add recLinkItemRecord, CStr(recLinkItemRecord.ID)
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
'd4     aga     13-Oct-1999     Created
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim recLinkItemRecord  As GenericAFLinkTerm
    sSource = Me.Name & ".RemoveLinkTerm"
    On Error GoTo ErrorHandler
    If t_eAFSeelinktype = EAFSeeLinkType.EAFSeealsoLink Then
        m_colCorporateBodySeeAlsoLinks.Remove CStr(t_lLinkTermID)
    Else
        m_colCorporateBodySeeLinks.Remove CStr(t_lLinkTermID)
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
'd4     aga     13-Oct-1999     Created
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim recLinkItemRecord  As Variant
                
    sSource = Me.Name & ".SerialiseLinkItem"
    
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
'd4     aga     13-Oct-1999     Created
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim recLinkItemRecord  As GenericAFLinkTerm, iTotalRecords As Integer, iCounter As Integer
    
    sSource = Me.Name & ".ExtractLinkItemsFromSerialIsedString"
    
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
    sSql = "SP_Edit_IsAFTermLinkedToCatalogueEntry @tiAFTermType=" & EAFTermType.AFCorporateBody _
        & ", " & "@tiAFTermID=" & m_typData.lID
    iHasLinks = m_ofwSQLExecuteHelper.FindBySql(sSql).GetFieldValue(0)
    If iHasLinks = 0 Then
        IsTermHasLinks = False
    End If
End Function


Public Function IsLinkedToOtherTerm() As Boolean
'   Check if terms is linked to/from other AF terms
'
'   d11 WDP 23-Feb-2000     Created
On Error GoTo ErrorHandler
    Dim sSource As String
    sSource = "BOAFCorporateBody.IsLinkedToOtherTerm"
    
    Dim sSql As String
    Dim ofwResults As FWRecordList
    Dim ofwHelper As New FWDBHelper
    
    sSql = " SELECT * FROM " & gk_sDBTABLECorpBodyLink & " WHERE " & _
            gk_sDBCOLCorpBodyLinkRefID1 & "=" & m_typData.lID & " OR " & _
            gk_sDBCOLCorpBodyLinkRefID2 & "=" & m_typData.lID
    
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


Private Sub NonPrefTermsRelPrefTermIsIncomOrApproved()
'Description    :   Check to see that when a Non_prefered terms status
'                   is being changed to InComplete or Approved the realted
'                   prefereed terms linked to it should be Incomplete or
'                   approved .
'   d10 Mahesh CR             16-Feb-2000    Added function
On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim oboCorpBodyLd As New BOAFCorpBodyLoader
    Dim sDesc As String
    sSource = Me.Name & ".NonPrefTermsRelPrefTermIsIncomOrApproved"
    
    sDesc = "You cannot change the status of this term because the preferred term associated with it will not be visible to the Reader"
    'if this is a non preferred term
    If m_colCorporateBodySeeLinks.Count > 0 Then
        oboCorpBodyLd.IAFLoader_Load m_colCorporateBodySeeLinks.Item(1).ID
        'and if the pref term has a status of ApprovedComplete or ApprovedIncomplete
        If Not (oboCorpBodyLd.Data.lAuthorityStatus = gk_lApprovedComplete Or oboCorpBodyLd.Data.lAuthorityStatus = gk_lApprovedInComplete) Then
              'Raise an error
              RaiseError sSource, PECanNotMakeNonPrefTermIcompApprovWithPrefTermIcompApprov, sDesc
        End If
    End If
Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub

