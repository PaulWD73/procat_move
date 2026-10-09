VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BOCatalogueDefaultValues"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    BOCatalogueDefaultValues.cls
' System:       PROCAT, PRO
' Copyright:    (C) Quidnunc Limited
' Instancing : Private
'
' Description:  Used to set default values when catalogue entries are create
'               from scratch and by accessioning
'
' Amendment history:
'
' d1    NDK     22Nov1999       Created.
' d2    NDK     10Dec1999       Change in updation of default values to division
' d3    WDP     25May2000       Get default PoD, Legal status, language and access cond from registry
' d4    DAJ     12Feb2001       Fixed problem with getting class children with no division parent,
'                               and general error handling
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'private business objects and framwork objects
Private m_oFWDBHelper As FWDBHelper
Private m_oBOCatalogueDBHelperLd As BOCatalogueDBHelperLd
Private m_oBOCatalogueDBHelper As BOCatalogueDBHelper
Private m_oBOCatalogueElementMatrix As BOCatalogueElementMatrix

'private varaibles
Private m_lEditsetID As Long
Private m_vCurrentCatalogueEntryTypeID As ECATALOGUEENTRYSTATUS

'default values for the lettercode
Private m_colDefaultCorpBodyIdForPlaceOfDeposit As Collection
Private m_vDefaultLegalStatusID As Variant
Private m_vDefaultLanguageID As Variant
Private m_vDefaultAccessConditionsID As Variant

'values to be assigned for empty coloumns
Private m_colNewCorpBodyIdForPlaceOfDeposit As Collection
Private m_vNewLegalStatusID As Variant
Private m_vNewLanguageID As Variant
Private m_vNewAccessConditionsID As Variant

'current values in the coloumns
Private m_colCurrentCorpBodyIdForPlaceOfDeposit As Collection
Private m_vCurrentLegalStatusID As Variant
Private m_vCurrentLanguageID As Variant
Private m_vCurrentAccessConditionsID As Variant

'private consts
Private Const ksDefaultPlaceOfDepositText As String = "'Public Record Office, Kew'"
Private Const ksDefaultLegalStatusText As String = "'Public Records'"
Private Const ksDefaultLanguageText As String = "'English'"
Private Const ksDefaultAccessConditionText As String = "'Subject to 30 year closure'"


Private Sub Class_Initialize()
    Initialise
End Sub

Private Sub Class_Terminate()
    Terminate
End Sub

Private Sub Initialise()
    Set m_oFWDBHelper = New FWDBHelper
    Set m_oBOCatalogueDBHelperLd = New BOCatalogueDBHelperLd
    Set m_oBOCatalogueDBHelper = New BOCatalogueDBHelper
    Set m_oBOCatalogueElementMatrix = New BOCatalogueElementMatrix
    
    Set m_colCurrentCorpBodyIdForPlaceOfDeposit = New Collection
    Set m_colNewCorpBodyIdForPlaceOfDeposit = New Collection
    Set m_colDefaultCorpBodyIdForPlaceOfDeposit = New Collection
End Sub
Private Sub Terminate()
    Set m_oFWDBHelper = Nothing
    Set m_oBOCatalogueDBHelperLd = Nothing
    Set m_oBOCatalogueDBHelper = Nothing
    Set m_oBOCatalogueElementMatrix = Nothing
    
    Set m_colCurrentCorpBodyIdForPlaceOfDeposit = Nothing
    Set m_colNewCorpBodyIdForPlaceOfDeposit = Nothing
    Set m_colDefaultCorpBodyIdForPlaceOfDeposit = Nothing
End Sub

Public Property Get Name()
    Name = "BOEditorialDBUtility"
End Property

Public Function DefaultValuesForCreatedCatalogueEntry( _
    ByVal t_lEditSetID As Long, _
    ByVal t_bCreatingLettercode As Boolean, _
    Optional ByVal t_bParentPresentInEditorial As Boolean, _
    Optional ByVal t_lParentCatalogueID As Long, _
    Optional ByVal t_eParentCatalogueLevelID As ECATALOGUELEVEL) As FWRecord
'
'
' Returns the default values to be set for a catalogue created from scratch
' The data is returned in a FWRecord wrapper
' When a lettercode is being created it is enough to give the editsetid and the catalogue level
' For all other levels the parent catalogue details should be passed.
' If the parent is not already present in editorial then the value is taken from reader
'
'
' d1    NDK     22Nov1999       Created.

    On Error GoTo ErrorHandler

    Dim sSql As String, sSource As String
    Dim sSerialisedPlaceOfDeposit As String
    
    sSource = Me.Name & ".DefaultValuesForCreatedCatalogueEntry"
    
    'store value of editset id in private variable
    m_lEditsetID = t_lEditSetID
    
    'initialise
    Set DefaultValuesForCreatedCatalogueEntry = New FWRecord
    
    If t_bCreatingLettercode Then
        'get the default values for lettercode
        GetDefaultValuesForLettercode
        
        'transfer the default values to new value
        TransferDefaultValuesToNewValues
    Else
        'get the values of the parent catalogue
        RetrieveCurrentValuesOfCatalogue _
            t_lParentCatalogueID, t_eParentCatalogueLevelID, _
            IIf(t_bParentPresentInEditorial, ESectionType.SECeditorial, ESectionType.SECReader)
            
        'transfer from current to new variables
        
        TransferCurrentValuesToNewValues
        
    End If
    
    'Add the properties
    DefaultValuesForCreatedCatalogueEntry.AddAttribute gk_sDBCOLLegalStatusID, m_vNewLegalStatusID
    DefaultValuesForCreatedCatalogueEntry.AddAttribute gk_sDBCOLLanguageID, m_vNewLanguageID
    DefaultValuesForCreatedCatalogueEntry.AddAttribute gk_sDBCOLAccessConditionID, m_vNewAccessConditionsID
    DefaultValuesForCreatedCatalogueEntry.AddAttribute _
        gk_sNamePlaceOfDeposit, _
        m_oBOCatalogueDBHelper.SerialiseLinkItem( _
        sSerialisedPlaceOfDeposit, m_colNewCorpBodyIdForPlaceOfDeposit)
    
    Exit Function
ErrorHandler:
    ConvertError sSource
End Function
Public Sub AssignDefaultValuesToAccessionedEntries( _
    ByVal t_lEditSetID As Long)
'
'
' Accepts an editset and assigns default values to all accessioned entries.
' The values are set only if there is not value present for that coloumn
' For lettercodes the default values are set if any relevant fields are empty.
' First the empty values for lettercodes are set and then recursively the empty values in each levels
' are set after taking their parents values.
' The updation is done only for accessioned entries
'
' d1    NDK     22Nov1999       Created.

    On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim oFWRecordlist As fwRecordList

    sSource = Me.Name & ".AssignDefaultValuesToAccessionedEntries"
    
    'set the value of editsetid
    m_lEditsetID = t_lEditSetID
    
    'retrieve the default value for lettercode and stores in variable
    GetDefaultValuesForLettercode
    
    Set oFWRecordlist = m_oBOCatalogueDBHelperLd.RecordListFromMainExtnTable( _
        gk_sDBCOLEditSetID & "=" & t_lEditSetID, _
        ESectionType.SECeditorial, _
        ECATALOGUELEVEL.ctLetterCode, _
        cttExtensionTable)
        
    'loop through all the lettercodes in the editset
    Do While Not oFWRecordlist.IsAtEnd
    
        'transfer the default values to new value
        TransferDefaultValuesToNewValues
        
        'update the relevant empty fields in this lettercode with default values
        ReplaceEmptyValuesWithDefaultOrParentValues _
            oFWRecordlist.GetFieldValue(gk_sDBCOLLetterCodeID), _
            ECATALOGUELEVEL.ctLetterCode
            
        oFWRecordlist.MoveNext
    Loop
    
    'release
    Set oFWRecordlist = Nothing
    
    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub
Private Sub ReplaceEmptyValuesWithDefaultOrParentValues( _
    ByVal t_lCatalogueID As Long, _
    ByVal t_eCatalogueLevelID As ECATALOGUELEVEL)
'
'
' Prepares data and calls the methods to update the values in the catalogues
' This method calls itself recursively till all the levels are updated
'
' d1    NDK     22Nov1999       Created.
' d2    NDK     10Dec1999       Division's language defaulted to 'english' if its and its parents language is empty/null
'
    On Error GoTo ErrorHandler
    
    Dim oFWChildrenList1 As fwRecordList, oFWChildrenList2 As fwRecordList
    Dim sIDColumn1 As String, sIDColumn2 As String
    Dim eCatalogueLevelID1 As ECATALOGUELEVEL, eCatalogueLevelID2 As ECATALOGUELEVEL
    
    Dim sSource As String
    
    sSource = Me.Name & ".ReplaceEmptyValuesWithDefaultOrParentValues"
    
    'get the current values for the catalogue entry
    'stores the current values of the catalogue in m_vCurrent... variables
        
    RetrieveCurrentValuesOfCatalogue t_lCatalogueID, t_eCatalogueLevelID, ESectionType.SECeditorial
    
    
    'The m_vNew.. variables contain the parents value initially.
    'If current value of any of the relevant fields is not null/empty then
    'the these current values are stored in the m_vnew.. variables
    'The tables will be updated with the values in m_vNew.. variables
    
    'legal status
    If Not IsNull(m_vCurrentLegalStatusID) Then
        m_vNewLegalStatusID = m_vCurrentLegalStatusID
    End If
    
    'Language
    If Not IsNull(m_vCurrentLanguageID) Then
        m_vNewLanguageID = m_vCurrentLanguageID
    End If
    
    'for division, set value to 'english' if its and its parents language are empty
    If t_eCatalogueLevelID = ECATALOGUELEVEL.ctDivision Then
        If IsNull(m_vNewLanguageID) Then
            m_vNewLanguageID = m_vDefaultLanguageID
        End If
    End If
    
    'Access condition
    If Not IsNull(m_vCurrentAccessConditionsID) Then
        m_vNewAccessConditionsID = m_vCurrentAccessConditionsID
    End If
    
    'Place of deposit
    If Not m_colCurrentCorpBodyIdForPlaceOfDeposit.Count = 0 Then
        Set m_colNewCorpBodyIdForPlaceOfDeposit = _
            m_colCurrentCorpBodyIdForPlaceOfDeposit
    End If
        
    'check if the catalogue is Imported and update it
    If m_vCurrentCatalogueEntryTypeID = ECATALOGUEENTRYSTATUS.cesImported Then
        
        'actually run the sqls to update the values
        UpdateValuesInTable t_lCatalogueID, t_eCatalogueLevelID
    
    End If
    
    'initialise
    
    'create the recordsets with chidlren details
    CreateRecordlistsWithChildrenDetails _
        t_lCatalogueID, t_eCatalogueLevelID, _
        oFWChildrenList1, oFWChildrenList2, _
        sIDColumn1, sIDColumn2, _
        eCatalogueLevelID1, eCatalogueLevelID2
        
    'First recordlist
    If Not oFWChildrenList1 Is Nothing Then
        Do While Not oFWChildrenList1.IsAtEnd
            
            'calls itself recursively to update children
            ReplaceEmptyValuesWithDefaultOrParentValues _
                oFWChildrenList1.GetFieldValue(sIDColumn1), _
                eCatalogueLevelID1
            
            'next child
            oFWChildrenList1.MoveNext
        Loop
    End If
    
    'Second recordlist
    If Not oFWChildrenList2 Is Nothing Then
        Do While Not oFWChildrenList2.IsAtEnd
                    
            'calls itself recursively to update children
            ReplaceEmptyValuesWithDefaultOrParentValues _
                oFWChildrenList2.GetFieldValue(sIDColumn2), _
                eCatalogueLevelID2
            
            'next child
            oFWChildrenList2.MoveNext
        Loop
    End If
    
    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub

Private Sub UpdateValuesInTable( _
    ByVal t_lCatalogueID As Long, _
    ByVal t_eCatalogueLevelID As ECATALOGUELEVEL)
'
' Actually runs the sqls to update the values in the tables
' The values that update are prepared in the ReplaceEmptyValuesWithDefaultOrParentValues method
'
' d1    NDK     22Nov1999       Created.
' d4    DAJ     12Feb2001       Fixed problem with the ',' characters

    On Error GoTo ErrorHandler
    
    Dim sSource As String, sSql As String
    Dim sValueToAssign As String
    
    sSource = Me.Name & ".UpdateValuesInTable"

    'basic sql
    sSql = " UPDATE " & _
        GetTableName(cttExtensionTable, SECeditorial, t_eCatalogueLevelID) & " SET "
    
    'legal status
    
    'value
    sValueToAssign = IIf(IsNull(m_vNewLegalStatusID), "Null", m_vNewLegalStatusID)
    'sql
    sSql = sSql & gk_sDBCOLLegalStatusID & "=" & sValueToAssign
    
    'language (Only for class and above)
    If t_eCatalogueLevelID = ECATALOGUELEVEL.ctClass Or _
        t_eCatalogueLevelID = ECATALOGUELEVEL.ctDivision Or _
        t_eCatalogueLevelID = ECATALOGUELEVEL.ctLetterCode Then
            
            'value
            sValueToAssign = IIf(IsNull(m_vNewLanguageID), "Null", m_vNewLanguageID)
    
            'sql
            sSql = sSql & ", " & gk_sDBCOLLanguageID & "=" & sValueToAssign
    End If
    
    'Access condition id (only for subheader and above)
    If t_eCatalogueLevelID = ECATALOGUELEVEL.ctSubHeaders Or _
        t_eCatalogueLevelID = ECATALOGUELEVEL.ctHeaders Or _
        t_eCatalogueLevelID = ECATALOGUELEVEL.ctClass Or _
        t_eCatalogueLevelID = ECATALOGUELEVEL.ctDivision Or _
        t_eCatalogueLevelID = ECATALOGUELEVEL.ctLetterCode Then
            'value
            sValueToAssign = IIf(IsNull(m_vNewAccessConditionsID), "Null", m_vNewAccessConditionsID)
            
            'sql
            sSql = sSql & ", " & gk_sDBCOLAccessConditionID & "=" & sValueToAssign

    End If
    
    'where clause
    sSql = sSql & " WHERE " & _
        gk_sDBCOLEditSetID & "=" & m_lEditsetID & " AND " & _
        GetColumnName(cctIDColumn, t_eCatalogueLevelID) & "=" & t_lCatalogueID
    
    'update the coloumns in the extension table
    m_oFWDBHelper.Execute sSql
    
    'now for the place of deposit (if there are already places of deposit, they are left alone)
    If m_colCurrentCorpBodyIdForPlaceOfDeposit.Count = 0 Then
        'there were no place of deposits, add them
        m_oBOCatalogueDBHelper.UpdateLinkTable _
            m_lEditsetID, t_lCatalogueID, t_eCatalogueLevelID, elPlaceOfDeposit, _
            m_colNewCorpBodyIdForPlaceOfDeposit, False
    End If
        
    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub


Private Sub RetrieveCurrentValuesOfCatalogue( _
    ByVal t_lCatalogueID As Long, _
    ByVal t_eCatalogueLevelID As ECATALOGUELEVEL, _
    ByVal t_eSectionType As ESectionType)
'
' Retrieves the default values from the catalogue passed as parameter
' The retrieving is done from the section mentioned in the parameter t_eSectionType
' The retrieved data are stored in private varaibles which are accessible in the module
'
' d1    NDK     22Nov1999       Created.

    On Error GoTo ErrorHandler
    
    Dim sSql As String, sSource As String
    Dim sConditionForMainAndExtnTables As String
    Dim oFWExtnRecordList As fwRecordList
    
    sSource = Me.Name & ".RetrieveCurrentValuesOfCatalogue"
    
    'Conditions
    If t_eSectionType = ESectionType.SECReader Then
        sConditionForMainAndExtnTables = _
            GetColumnName(ECATALOGUECOLUMNTYPE.cctIDColumn, GetCatalogueLevelCode(t_eCatalogueLevelID)) & "=" & t_lCatalogueID
    Else
        'for editorial
        sConditionForMainAndExtnTables = _
            gk_sDBCOLEditSetID & "=" & m_lEditsetID & " AND " & _
            GetColumnName(ECATALOGUECOLUMNTYPE.cctIDColumn, GetCatalogueLevelCode(t_eCatalogueLevelID)) & "=" & t_lCatalogueID
    End If
    
    'get the data in the extension dtable
    Set oFWExtnRecordList = _
        m_oBOCatalogueDBHelperLd.RecordListFromMainExtnTable( _
            sConditionForMainAndExtnTables, _
            t_eSectionType, t_eCatalogueLevelID, ECATALOGUETABLETYPE.cttExtensionTable)
    
    'In live this field is not available
    If t_eSectionType = ESectionType.SECeditorial Then
        'entry type of the current catalogue
        m_vCurrentCatalogueEntryTypeID = oFWExtnRecordList.GetFieldValue(gk_sDBCOLEditSetEntryTypeID)
    End If
    
    'to get place of deposit
    Set m_colCurrentCorpBodyIdForPlaceOfDeposit = _
        m_oBOCatalogueDBHelperLd.GenerateDetailsUsingIdsFromDatabase( _
            t_lCatalogueID, t_eCatalogueLevelID, _
            t_eSectionType, elPlaceOfDeposit, m_lEditsetID)
            
    'Legal status code
    m_vCurrentLegalStatusID = oFWExtnRecordList.GetFieldValue(gk_sDBCOLLegalStatusID)
 
    'Language (Only for class and above)
    If t_eCatalogueLevelID = ECATALOGUELEVEL.ctClass Or _
        t_eCatalogueLevelID = ECATALOGUELEVEL.ctDivision Or _
        t_eCatalogueLevelID = ECATALOGUELEVEL.ctLetterCode Then
        
            m_vCurrentLanguageID = oFWExtnRecordList.GetFieldValue(gk_sDBCOLLanguageID)
    End If
    
    'Access condition id (only for subheader and above)
    If t_eCatalogueLevelID = ECATALOGUELEVEL.ctSubHeaders Or _
        t_eCatalogueLevelID = ECATALOGUELEVEL.ctHeaders Or _
        t_eCatalogueLevelID = ECATALOGUELEVEL.ctClass Or _
        t_eCatalogueLevelID = ECATALOGUELEVEL.ctDivision Or _
        t_eCatalogueLevelID = ECATALOGUELEVEL.ctLetterCode Then
        
            m_vCurrentAccessConditionsID = oFWExtnRecordList.GetFieldValue(gk_sDBCOLAccessConditionID)

    End If
    
    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub




Private Sub GetDefaultValuesForLettercode()
'
' Populates the variables storing the default values for a lettercode
' These default values are found in the master tables
'
' d1    NDK     22Nov1999       Created.
' d3    WDP     25May2000       Get default PoD, Legal status, language and access cond from registry

    On Error GoTo ErrorHandler
    
    Dim sSql As String, sSource As String
    Dim ofwAppData As New FWApplicationData
    
    sSource = Me.Name & ".GetDefaultValuesForLettercode"
    
    'place of deposit
    Dim vPlaceOfDepositID As Variant, recLinkItem As GenericLinkRecord
    
    vPlaceOfDepositID = ColumnContent( _
        gk_sDBTABLECorporateBody, _
        gk_sDBCOLCorpBodyReferenceID, _
        gk_sDBCOLCorpBodyNameText & " = " & _
        QuoteValueWithQuotes(ofwAppData.GetApplicationConfigurationInfomation(pDefaultPlaceOfDeposit)), _
        ksDefaultPlaceOfDepositText)
    
    If Not IsNull(vPlaceOfDepositID) Then
        recLinkItem.lID = vPlaceOfDepositID
        m_colDefaultCorpBodyIdForPlaceOfDeposit.ADD recLinkItem
    End If
    
    'legal status
    m_vDefaultLegalStatusID = ColumnContent( _
        gk_sDBTABLELegalStatus, _
        gk_sDBCOLLegalStatusID, _
        gk_sDBCOLLegalStatusDescription & " = " & _
        QuoteValueWithQuotes(ofwAppData.GetApplicationConfigurationInfomation(pDefaultLegalStatus)) & _
            " AND " & _
        gk_sDBCOLCurrencyIndicator & "= 1", _
        ksDefaultLegalStatusText)
            
    'language
    m_vDefaultLanguageID = ColumnContent( _
        gk_sDBTABLELanguage, _
        gk_sDBCOLLanguageID, _
        gk_sDBCOLLanguageDescription & " = " & _
        QuoteValueWithQuotes(ofwAppData.GetApplicationConfigurationInfomation(pDefaultLanguage)) & _
            " AND " & _
        gk_sDBCOLCurrencyIndicator & "= 1", _
        ksDefaultLanguageText)
    
    'access conditions
    m_vDefaultAccessConditionsID = ColumnContent( _
        gk_sDBTABLEAccessCondition, _
        gk_sDBCOLAccessConditionID, _
        gk_sDBCOLAccessConditionDescription & " = " & _
        QuoteValueWithQuotes(ofwAppData.GetApplicationConfigurationInfomation(pDefaultAccessCondition)) & _
            " AND " & _
        gk_sDBCOLCurrencyIndicator & "= 1", _
        ksDefaultAccessConditionText)
            
   Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub


Private Function ColumnContent( _
    ByVal t_sTableName As String, _
    ByVal t_sColumToReturnName As String, _
    ByVal t_sConditionString As String, _
    ByVal t_sValueBeingSearched As String) As Variant
'
' Returns the content of the coloumn. This is used by method GetDefaultValuesForLettercode
' to get the values for the default coloumns.
' The method raises an error if exactly one row is not found
'
' d1    NDK     22Nov1999       Created.

    On Error GoTo ErrorHandler
    
    Dim sSql As String, sSource As String
    Dim oFWRecordlist As fwRecordList

    sSource = Me.Name & ".ColumnContent"
    
    'get the recordset
    sSql = _
        " SELECT * FROM " & t_sTableName & _
        " WHERE " & t_sConditionString
        
    'get the recordlist
    Set oFWRecordlist = m_oFWDBHelper.FindBySql(sSql)
    
    'there should be only one record
    If oFWRecordlist.RecordCount = 0 Then
        RaiseError sSource, peDefaultValuesForCataloguesRowNotPresent, _
            "Row containing value " & t_sValueBeingSearched & " not found "
    End If
    
    If oFWRecordlist.RecordCount > 1 Then
        RaiseError sSource, peDefaultValuesForCataloguesMultipleRowsPresent, _
            "More than one row containing value " & t_sValueBeingSearched & " were found "
    End If
    
    'return the value
    ColumnContent = oFWRecordlist.GetFieldValue(t_sColumToReturnName)
    
    'release
    Set oFWRecordlist = Nothing
    
    Exit Function
ErrorHandler:
    ConvertError sSource
End Function




Private Sub CreateRecordlistsWithChildrenDetails( _
    ByVal t_lCatalogueID As Long, _
    ByVal t_eCatalogueLevel As ECATALOGUELEVEL, _
    ByRef t_oFWChildrenList1 As fwRecordList, _
    ByRef t_oFWChildrenList2 As fwRecordList, _
    ByRef t_sIDColumn1 As String, _
    ByRef t_sIDColumn2 As String, _
    ByRef t_eCatalogueLevelID1 As ECATALOGUELEVEL, _
    ByRef t_eCatalogueLevelID2 As ECATALOGUELEVEL)
'
' Creates a record list of children for the catalogue passed as parameter
' For lettercodes, classes,headers two recordlists are created
' The created recordlists and other details are stored in variables
' passed by reference
'
' d1    NDK     22Nov1999       Created.
' d4    DAJ     12Feb2001       Fixed problem with getting class children with no division parent

    On Error GoTo ErrorHandler
    
    Dim sSource As String
    
    sSource = Me.Name & ".CreateRecordlistsWithChildrenDetails"

    Select Case t_eCatalogueLevel
    
        Case ECATALOGUELEVEL.ctLetterCode
        'lettercode could have children in division and class tables
         
            'list of divisions under the letterocode
            Set t_oFWChildrenList1 = GetRecordListFromTable _
                (GetTableName(cttMainTable, ESectionType.SECeditorial, ECATALOGUELEVEL.ctDivision), _
                m_lEditsetID, gk_sDBCOLLetterCodeID, t_lCatalogueID)
            t_sIDColumn1 = gk_sDBCOLDivisionID
            t_eCatalogueLevelID1 = ECATALOGUELEVEL.ctDivision
            
            'list of classes directly under the letterocode (d4 - which have division id as zero, not null)
            Set t_oFWChildrenList2 = GetRecordListFromTable _
                (GetTableName(cttMainTable, ESectionType.SECeditorial, ECATALOGUELEVEL.ctClass), _
                m_lEditsetID, gk_sDBCOLLetterCodeID, t_lCatalogueID, gk_sDBCOLDivisionID, , True)
            t_sIDColumn2 = gk_sDBCOLClassID
            t_eCatalogueLevelID2 = ECATALOGUELEVEL.ctClass
            
        Case ECATALOGUELEVEL.ctDivision
        'division could have children in class table

            'list of classes under the division
            Set t_oFWChildrenList1 = GetRecordListFromTable _
                (GetTableName(cttMainTable, ESectionType.SECeditorial, ECATALOGUELEVEL.ctClass), _
                m_lEditsetID, gk_sDBCOLDivisionID, t_lCatalogueID)
            t_sIDColumn1 = gk_sDBCOLClassID
            t_eCatalogueLevelID1 = ECATALOGUELEVEL.ctClass

        Case ECATALOGUELEVEL.ctClass
        'class could have children in header and piece tables

            'list of headers under class
            Set t_oFWChildrenList1 = GetRecordListFromTable _
                (GetTableName(cttMainTable, ESectionType.SECeditorial, ECATALOGUELEVEL.ctHeaders), _
                m_lEditsetID, gk_sDBCOLClassID, t_lCatalogueID)
            t_sIDColumn1 = gk_sDBCOLHeaderID
            t_eCatalogueLevelID1 = ECATALOGUELEVEL.ctHeaders
            
            'list of pieces directly under the class (which have header and subheader as null)
            Set t_oFWChildrenList2 = GetRecordListFromTable _
                (GetTableName(cttMainTable, ESectionType.SECeditorial, ECATALOGUELEVEL.ctPiece), _
                m_lEditsetID, gk_sDBCOLClassID, t_lCatalogueID, gk_sDBCOLHeaderID, gk_sDBCOLSubHeaderID)
            t_sIDColumn2 = gk_sDBCOLPieceID
            t_eCatalogueLevelID2 = ECATALOGUELEVEL.ctPiece

        Case ECATALOGUELEVEL.ctHeaders
        'header could have children in sub header and piece

            'list of sub headers under header
            Set t_oFWChildrenList1 = GetRecordListFromTable _
                (GetTableName(cttMainTable, ESectionType.SECeditorial, ECATALOGUELEVEL.ctSubHeaders), _
                m_lEditsetID, gk_sDBCOLHeaderID, t_lCatalogueID)
            t_sIDColumn1 = gk_sDBCOLSubHeaderID
            t_eCatalogueLevelID1 = ECATALOGUELEVEL.ctSubHeaders

            'list of pieces directly under header (which have subheader id as null)
            Set t_oFWChildrenList2 = GetRecordListFromTable _
                (GetTableName(cttMainTable, ESectionType.SECeditorial, ECATALOGUELEVEL.ctPiece), _
                m_lEditsetID, gk_sDBCOLHeaderID, t_lCatalogueID, gk_sDBCOLSubHeaderID)
            t_sIDColumn2 = gk_sDBCOLPieceID
            t_eCatalogueLevelID2 = ECATALOGUELEVEL.ctPiece

        Case ECATALOGUELEVEL.ctSubHeaders
        'subheader could have children in piece

            'list of pieces under sub header
            Set t_oFWChildrenList1 = GetRecordListFromTable _
                (GetTableName(cttMainTable, ESectionType.SECeditorial, ECATALOGUELEVEL.ctPiece), _
                m_lEditsetID, gk_sDBCOLSubHeaderID, t_lCatalogueID)
            t_sIDColumn1 = gk_sDBCOLPieceID
            t_eCatalogueLevelID1 = ECATALOGUELEVEL.ctPiece
        
        Case ECATALOGUELEVEL.ctPiece
        'piece could have children in item

            'list of items under piece
            Set t_oFWChildrenList1 = GetRecordListFromTable _
                (GetTableName(cttMainTable, ESectionType.SECeditorial, ECATALOGUELEVEL.ctItem), _
                m_lEditsetID, gk_sDBCOLPieceID, t_lCatalogueID)
            t_sIDColumn1 = gk_sDBCOLItemID
            t_eCatalogueLevelID1 = ECATALOGUELEVEL.ctItem
            
        Case ECATALOGUELEVEL.ctItem
            'there would be no children for item
    End Select

    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub

Private Function GetRecordListFromTable( _
    ByVal t_sTableName As String, _
    ByVal t_lEditSetID As Long, _
    ByVal t_sIDColumnName As String, _
    ByVal t_lIDColumnValue As Long, _
    Optional ByVal t_sColumn1ToBeNull As String = "", _
    Optional ByVal t_sColumn2ToBeNull As String = "", _
    Optional ByVal t_bColumn1IsZeroNotNull As Boolean = False, _
    Optional ByVal t_bColumn2IsZeroNotNull As Boolean = False) As fwRecordList
   
'
'Returns a record list from the editorial table where t_sIDColumnName  = t_vIDColumnValue
'If the t_sColumn1ToBeNull and t_sColumn2ToBeNull are not empty then the the records returned
'should have these columns set as null
'
' Amendment history:
' d1    NDK     22Nov1999       Created.
' d4    DAJ     12Feb2001       Added ability to specify if column was a zero or null

    On Error GoTo ErrorHandler

    Dim sSource As String, sSql As String
    
    sSource = Me.Name & ".GetRecordListFromTable"
    
    'basic query
    sSql = _
        " SELECT * FROM " & t_sTableName & _
        " WHERE " & _
        gk_sDBCOLEditSetID & "=" & t_lEditSetID & " AND " & _
        t_sIDColumnName & "=" & t_lIDColumnValue
    
    'This t_sColumn1ToBeNull has to be null (d4 - or zero)
    If Not t_sColumn1ToBeNull = "" Then
        sSql = sSql & " AND " & t_sColumn1ToBeNull & IIf(t_bColumn1IsZeroNotNull, " = 0", " IS  Null")
    End If
    
    'This t_sColumn2ToBeNull has to be null (d4 - or zero)
    If Not t_sColumn2ToBeNull = "" Then
        sSql = sSql & " AND " & t_sColumn2ToBeNull & IIf(t_bColumn2IsZeroNotNull, " = 0", " IS  Null")
    End If
    
    'return the recordlist
    Set GetRecordListFromTable = m_oFWDBHelper.FindBySql(sSql)
    
    Exit Function
ErrorHandler:
    ConvertError sSource
End Function



Private Sub TransferDefaultValuesToNewValues()
'
'Transfers the default values of a lettercode to new values
'
' Amendment history:
' d1    NDK     22Nov1999       Created.
    
    On Error GoTo ErrorHandler

    Dim sSource As String
    
    sSource = Me.Name & ".TransferDefaultValuesToNewValues"
    
    m_vNewLegalStatusID = m_vDefaultLegalStatusID
    m_vNewLanguageID = m_vDefaultLanguageID
    m_vNewAccessConditionsID = m_vDefaultAccessConditionsID
    Set m_colNewCorpBodyIdForPlaceOfDeposit = m_colDefaultCorpBodyIdForPlaceOfDeposit

Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub

Private Sub TransferCurrentValuesToNewValues()
'
'Transfers the current values to new values for all levels
'
' Amendment history:
' d1    NDK     22Nov1999       Created.
    
    On Error GoTo ErrorHandler

    Dim sSource As String
    
    sSource = Me.Name & ".TransferCurrentValuesToNewValues"
    
    m_vNewLegalStatusID = m_vCurrentLegalStatusID
    m_vNewLanguageID = m_vCurrentLanguageID
    m_vNewAccessConditionsID = m_vCurrentAccessConditionsID
    Set m_colNewCorpBodyIdForPlaceOfDeposit = m_colCurrentCorpBodyIdForPlaceOfDeposit

Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub

