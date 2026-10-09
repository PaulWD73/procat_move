VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BOCatalogueDBHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    BOCatalogueDBHelper.cls
' System:       PROCAT, PRO
' Copyright:    (C) Quidnunc Limited
'
' Description:  Contains the database interaction functions used for editing and
'               displaying catalogue details tab. The loader class BOCatalogueDBHelperLd is used
'               to perform loading functions
' Amendment history:
'   d1      ndk     25Aug1999       Created
'   d2      ndk     21Oct1999       Added method UpdateEditSetEntryTypeStatus
'   d3      ndk     02Nov1999       Changed the name of the method UpdateEditSetEntryTypeStatus to
'                                   SetCatalogueStatusAsModified
'   d4      ndk     18Nov1999       updating sSubjectTermText value for subject link terms
'   d5      ndk     24Nov1999       Removed method RecordListFromMainExtnTable
'   d6      wdp     03Dec1999       Added UpdateRecordStatus to update the record status flag
'                                   in edit set and live
'   d7      crm     02Feb2000       Updated code to provide links for Leaflets to catalogues
'   d8      crm     11Feb2000       changed GenerateDetailsUsingSavedIDs (the catalogue to leaflet link portion)
'                                   to correctly persist optional descriptions between
'                                   navigations
'   d9      alm     02-Oct-2001     Created BulkUpdateRecordStatus
'   d10     fjd     05Feb2003       Renamed LCI flag to AR
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private m_oFWDBHelper As FWDBHelper
Private m_oBOCatalogueDBHelperLd As BOCatalogueDBHelperLd

Public Property Get Name()
    Name = "BOCatalogueDBHelper"
End Property
Private Sub Class_Initialize()
    Initialise
End Sub
Private Sub Class_Terminate()
    Terminate
End Sub
Private Sub Initialise()
    Set m_oFWDBHelper = New FWDBHelper
    Set m_oBOCatalogueDBHelperLd = New BOCatalogueDBHelperLd
End Sub
Private Sub Terminate()
    Set m_oFWDBHelper = Nothing
    Set m_oBOCatalogueDBHelperLd = Nothing
End Sub
    
Public Function GenerateDetailsUsingSavedIDs( _
    ByVal t_lEditSetID As Long, _
    ByVal t_lCatalogueID As Long, _
    ByVal t_eCatalogueLevelID As ECATALOGUELEVEL, _
    ByVal t_eLinkTermType As ELinkTermType, _
    ByVal t_colIDsAndDescriptions As Collection) As Collection
'
'Description: Generates details of link items by converting the information in the collections
'             into to complete details
' Amendment history:
'   d1      ndk     25Aug1999   Created
'   d8      crm     11Feb2000       changed GenerateDetailsUsingSavedIDs (the catalogue to leaflet link portion)
'                                   to correctly persist optional descriptions between
'                                   navigations

    On Error GoTo ErrorHandler

    Dim sSource As String
    Dim sIDString As String, recLinkItemWithIDAndDescription As Variant, colLinkItemsWithoutDescription As Collection
    Dim sKey As String
    Dim recPersonDetails As PersonDetails, recCorporateBodyDetails As CorporateBodyDetails, recPlaceDetails As PlaceDetails, recSubjectDetails As SubjectDetails, recSFADetails As SFADetails
    Dim recLeafletDetail As LeafletDetails
    
    sSource = Me.Name & ".GenerateDetailsUsingSavedIDs"
    
    Set GenerateDetailsUsingSavedIDs = New Collection
            
    If t_colIDsAndDescriptions.Count = 0 Then       'The collection is empty
        Exit Function
    End If
            
    'scan the collection and construct a string which contains all the ids
    For Each recLinkItemWithIDAndDescription In t_colIDsAndDescriptions
        If sIDString = "" Then
            sIDString = recLinkItemWithIDAndDescription.lID
        Else
            sIDString = sIDString & "," & recLinkItemWithIDAndDescription.lID
        End If
    Next
            
    'Get the list of ids (without description) from the table using the loader class
    Set colLinkItemsWithoutDescription = m_oBOCatalogueDBHelperLd.DetailsFromDatabase( _
        t_lCatalogueID, _
        t_eCatalogueLevelID, _
        ESectionType.SECeditorial, _
        t_eLinkTermType, _
        True, _
        t_lEditSetID, _
        sIDString)
        
    'Add the description and other details to the collection
    For Each recLinkItemWithIDAndDescription In t_colIDsAndDescriptions
    
        'Index value
        sKey = gk_sPrefixForLinkTerm & recLinkItemWithIDAndDescription.lID
        
        Select Case t_eLinkTermType
        
            'Person type records
            Case ELinkTermType.elImmedieteSourceOfAcquisitionPerson, _
                ELinkTermType.elCreatorNamesPerson, _
                ELinkTermType.elIndexTermPerson
            
                recPersonDetails = colLinkItemsWithoutDescription.Item(sKey)
                
                'Get the description
                recPersonDetails.sOptionalDescription = recLinkItemWithIDAndDescription.sOptionalDescription
                
                'Other details
                recPersonDetails.lID = colLinkItemsWithoutDescription.Item(sKey).lID
                recPersonDetails.sForeName = colLinkItemsWithoutDescription.Item(sKey).sForeName
                
                GenerateDetailsUsingSavedIDs.Add recPersonDetails
            
            'Corporate body type records
            Case ELinkTermType.elImmedieteSourceOfAcquisitionCorporateBody, _
                ELinkTermType.elCopiesInformation, _
                ELinkTermType.elLocationOfOriginals, _
                ELinkTermType.elCreatorNamesCorporateBody, _
                ELinkTermType.elPlaceOfDeposit, _
                ELinkTermType.elIndexTermCorporateBody
            
                recCorporateBodyDetails = colLinkItemsWithoutDescription.Item(sKey)

                'Get the description
                recCorporateBodyDetails.sOptionalDescription = recLinkItemWithIDAndDescription.sOptionalDescription

                'Other details
                recCorporateBodyDetails.lID = colLinkItemsWithoutDescription.Item(sKey).lID
                recCorporateBodyDetails.sCorporateBodyName = colLinkItemsWithoutDescription.Item(sKey).sCorporateBodyName
                recCorporateBodyDetails.sUncertainStartDateCode = colLinkItemsWithoutDescription.Item(sKey).sUncertainStartDateCode
                recCorporateBodyDetails.sStartDate = colLinkItemsWithoutDescription.Item(sKey).sStartDate
                recCorporateBodyDetails.sUncertainEndDateCode = colLinkItemsWithoutDescription.Item(sKey).sUncertainEndDateCode
                recCorporateBodyDetails.sEndDate = colLinkItemsWithoutDescription.Item(sKey).sEndDate
                
                GenerateDetailsUsingSavedIDs.Add recCorporateBodyDetails
                
            'Place type records
            Case ELinkTermType.elIndexTermPlace
                recPlaceDetails = colLinkItemsWithoutDescription.Item(sKey)

                'Get the description
                recPlaceDetails.sOptionalDescription = recLinkItemWithIDAndDescription.sOptionalDescription

                'Other details
                recPlaceDetails.lID = colLinkItemsWithoutDescription.Item(sKey).lID
                recPlaceDetails.sPlaceName = colLinkItemsWithoutDescription.Item(sKey).sPlaceName
                recPlaceDetails.sParishName = colLinkItemsWithoutDescription.Item(sKey).sParishName
                recPlaceDetails.sTownName = colLinkItemsWithoutDescription.Item(sKey).sTownName
                recPlaceDetails.sCountyName = colLinkItemsWithoutDescription.Item(sKey).sCountyName
                recPlaceDetails.sCountryName = colLinkItemsWithoutDescription.Item(sKey).sCountryName
                recPlaceDetails.sUncertainStartDate = colLinkItemsWithoutDescription.Item(sKey).sUncertainStartDate
                recPlaceDetails.sStartDate = colLinkItemsWithoutDescription.Item(sKey).sStartDate
                recPlaceDetails.sUncertainEndDate = colLinkItemsWithoutDescription.Item(sKey).sUncertainEndDate
                recPlaceDetails.sEndDate = colLinkItemsWithoutDescription.Item(sKey).sEndDate
                
                GenerateDetailsUsingSavedIDs.Add recPlaceDetails
            
            'Subject type records
            Case ELinkTermType.elIndexTermSubject
                recSubjectDetails = colLinkItemsWithoutDescription.Item(sKey)

                'Get the description
                recSubjectDetails.sOptionalDescription = recLinkItemWithIDAndDescription.sOptionalDescription

                'Other details
                recSubjectDetails.lID = colLinkItemsWithoutDescription.Item(sKey).lID
                recSubjectDetails.sSubjectDefinition = colLinkItemsWithoutDescription.Item(sKey).sSubjectDefinition
                recSubjectDetails.sSubjectTermText = colLinkItemsWithoutDescription.Item(sKey).sSubjectTermText
                
                
                GenerateDetailsUsingSavedIDs.Add recSubjectDetails
            
            'SFA type records
            Case ELinkTermType.elPublicationNote, ELinkTermType.elUnPublishedFindingAids
            
                recSFADetails = colLinkItemsWithoutDescription.Item(sKey)

                'Get the description
                recSFADetails.sOptionalDescription = recLinkItemWithIDAndDescription.sOptionalDescription

                'Other details
                recSFADetails.lID = colLinkItemsWithoutDescription.Item(sKey).lID
                recSFADetails.sTitle = colLinkItemsWithoutDescription.Item(sKey).sTitle
                recSFADetails.sSubTitle = colLinkItemsWithoutDescription.Item(sKey).sSubTitle
                recSFADetails.sEdition = colLinkItemsWithoutDescription.Item(sKey).sEdition
                recSFADetails.sAuthor = colLinkItemsWithoutDescription.Item(sKey).sAuthor
                recSFADetails.sPublisher = colLinkItemsWithoutDescription.Item(sKey).sPublisher
                recSFADetails.sPublicationDate = colLinkItemsWithoutDescription.Item(sKey).sPublicationDate
                recSFADetails.bInProLibaray = colLinkItemsWithoutDescription.Item(sKey).bInProLibaray
                recSFADetails.bInProReference = colLinkItemsWithoutDescription.Item(sKey).bInProReference

                GenerateDetailsUsingSavedIDs.Add recSFADetails
            '   d7      crm     02Feb2000
            Case ELinkTermType.elPublicationNoteLeaflet, ELinkTermType.elUnPublishedFindingAidLeaflet
                recLeafletDetail.lID = colLinkItemsWithoutDescription.Item(sKey).lID
                recLeafletDetail.sTitle = colLinkItemsWithoutDescription.Item(sKey).sTitle
                '   d8      crm     11Feb2000
                recLeafletDetail.sOptionalDescription = recLinkItemWithIDAndDescription.sOptionalDescription
                
                GenerateDetailsUsingSavedIDs.Add recLeafletDetail
        End Select
    Next            'Move to the next record
    
    Exit Function

ErrorHandler:
    ConvertError sSource
End Function


Public Function GenerateDetailsUsingIdsFromDatabase( _
    ByVal t_lCatalogueID As Long, _
    ByVal t_eCatalogueLevelID As ECATALOGUELEVEL, _
    ByVal t_eSectionType As ESectionType, _
    ByVal t_eLinkTermType As ELinkTermType, _
    Optional ByVal t_lEditSetID As Long) As Collection
'
'Description: Generates details of link items by reading the database
'             The loader object is used to perform this task
'
' Amendment history:
'   d1      ndk     25Aug1999   Created
'
    On Error GoTo ErrorHandler

    Dim sSource As String
    
    sSource = Me.Name & ".GenerateDetailsUsingIdsFromDatabase"

    'Call the load objects method to get the details
    Set GenerateDetailsUsingIdsFromDatabase = _
        m_oBOCatalogueDBHelperLd.GenerateDetailsUsingIdsFromDatabase( _
            t_lCatalogueID, _
            t_eCatalogueLevelID, _
            t_eSectionType, _
            t_eLinkTermType, _
            t_lEditSetID)
    Exit Function
    
ErrorHandler:
    ConvertError sSource
End Function



Private Function RelatedSepareteMaterialsDetailsFromDatabase( _
    ByVal t_lCatalogueID As Long, _
    ByVal t_eCatalogueLevelID As ECATALOGUELEVEL, _
    ByVal t_eSectionType As ESectionType, _
    ByVal t_eLinkTermType As ELinkTermType, _
    Optional ByVal t_lEditSetID As Long) As Collection
    
'
'Description: Gets the related and separate material details , the loader object
'             is used to perform the query
' Amendment history:
'   d1   ndk      25Aug1999      Created
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String, sSql As String
    
    sSource = Me.Name & ".RelatedSepareteMaterialsDetailsFromDatabase"

    Set RelatedSepareteMaterialsDetailsFromDatabase = _
        m_oBOCatalogueDBHelperLd.RelatedSepareteMaterialsDetailsFromDatabase( _
        t_lCatalogueID, _
        t_eCatalogueLevelID, _
        t_eSectionType, _
        t_eLinkTermType, _
        t_lEditSetID)
    
    Exit Function

ErrorHandler:
    ConvertError sSource
End Function

Public Function UnknownDescription( _
    ByVal t_lCatalogueID As Long, _
    ByVal t_eCatalogueLevelID As ECATALOGUELEVEL, _
    ByVal t_eSectionType As ESectionType, _
    ByVal t_eProperty As ECATALOGUEPROPERTIES, _
    Optional ByVal t_lEditSetID As Long) As Variant
'
'Description: Returns the unknwon description for link items, the details are extracted
'             from the loader class
' Amendment history:
'   d1      ndk     25Aug1999   Created
'
    On Error GoTo ErrorHandler

    Dim sSource As String

    sSource = Me.Name & ".UnknownDescription"
    
    UnknownDescription = _
        m_oBOCatalogueDBHelperLd.UnknownDescription( _
            t_lCatalogueID, _
            t_eCatalogueLevelID, _
            t_eSectionType, _
            t_eProperty, _
            t_lEditSetID)
        
    Exit Function

ErrorHandler:
    ConvertError sSource
End Function


Public Function UpdateMainExtnTable( _
            ByVal t_sConditionString As String, _
            ByVal t_sTableName As String, _
            ByVal t_sFieldName As String, _
            ByVal t_vFieldContent As Variant)
'
'Description: Updates the specfied field and specified table for the main and extension table
'
' Amendment history:
'   d1   ndk      25Aug1999      Created
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String, sSql As String
    sSource = Me.Name & ".UpdateMainExtnTable"
    
    'Construct the sql
    
    'Update specfic field in the table
    sSql = _
        " UPDATE " & t_sTableName & _
        " SET " & _
            t_sFieldName & "=" & _
            IIf(IsNull(t_vFieldContent) Or IsEmpty(t_vFieldContent) Or LTrim(RTrim(t_vFieldContent)) = "", "Null", t_vFieldContent) & _
        " WHERE " & t_sConditionString
            
    'Execute the sql
    m_oFWDBHelper.Execute sSql
    
    Exit Function
    
ErrorHandler:
    ConvertError sSource
End Function


Public Sub UpdateLinkTable( _
    ByVal t_lEditSetID As Long, _
    ByVal t_lCatalogueID As Long, _
    ByVal t_eCatalogueLevelID As ECATALOGUELEVEL, _
    ByVal t_eLinkTermType As ELinkTermType, _
    ByVal t_colLinkItems As Collection, _
    ByVal t_bHasOptionalDescription As Boolean)

'
'
'Description: Update the link items table (except related and separate material)
'
' Amendment history:
'   d1     ndk      25Aug1999      Created
'   d7     crm     02Feb2000       Updated code to provide links for Leaflets to catalogues
    On Error GoTo ErrorHandler
    
    Dim sSource As String, sSql As String
    Dim sMasterTable As String, sLinksTable As String
    Dim sIDColumnName As String, eProperty As ECATALOGUEPROPERTIES
    Dim recLinkItemRecord  As Variant
    
    sSource = Me.Name & ".UpdateLinkTable"

    'case structure to get the following details
    'Links table - table to extract link item details from
    'Master table - master table for the link items (af files)
    'ID column name - id column for the link and master tables
    'Property - property id of the link item
    
    Select Case t_eLinkTermType
        Case ELinkTermType.elImmedieteSourceOfAcquisitionPerson
            sLinksTable = GetTableName(cttPerson, ESectionType.SECeditorial)
            sMasterTable = gk_sDBTABLEPerson
            sIDColumnName = gk_sDBCOLPersonReferenceID
            eProperty = ECATALOGUEPROPERTIES.cpImmediateSourceOfAcquisition
        Case ELinkTermType.elImmedieteSourceOfAcquisitionCorporateBody
            sLinksTable = GetTableName(cttCorporateBody, ESectionType.SECeditorial)
            sMasterTable = gk_sDBTABLECorporateBody
            sIDColumnName = gk_sDBCOLCorporateBodyReferenceID
            eProperty = ECATALOGUEPROPERTIES.cpImmediateSourceOfAcquisition
        Case ELinkTermType.elCopiesInformation
            sLinksTable = GetTableName(cttCorporateBody, ESectionType.SECeditorial)
            sMasterTable = gk_sDBTABLECorporateBody
            sIDColumnName = gk_sDBCOLCorporateBodyReferenceID
            eProperty = ECATALOGUEPROPERTIES.cpCopiesInformation
        Case ELinkTermType.elLocationOfOriginals
            sLinksTable = GetTableName(cttCorporateBody, ESectionType.SECeditorial)
            sMasterTable = gk_sDBTABLECorporateBody
            sIDColumnName = gk_sDBCOLCorporateBodyReferenceID
            eProperty = ECATALOGUEPROPERTIES.cpLocationOfOriginals
        Case ELinkTermType.elCreatorNamesCorporateBody
            sLinksTable = GetTableName(cttCorporateBody, ESectionType.SECeditorial)
            sMasterTable = gk_sDBTABLECorporateBody
            sIDColumnName = gk_sDBCOLCorporateBodyReferenceID
            eProperty = ECATALOGUEPROPERTIES.cpCreatorNames
        Case ELinkTermType.elCreatorNamesPerson
            sLinksTable = GetTableName(cttPerson, ESectionType.SECeditorial)
            sMasterTable = gk_sDBTABLEPerson
            sIDColumnName = gk_sDBCOLPersonReferenceID
            eProperty = ECATALOGUEPROPERTIES.cpCreatorNames
        Case ELinkTermType.elPlaceOfDeposit
            sLinksTable = GetTableName(cttCorporateBody, ESectionType.SECeditorial)
            sMasterTable = gk_sDBTABLECorporateBody
            sIDColumnName = gk_sDBCOLCorporateBodyReferenceID
            eProperty = ECATALOGUEPROPERTIES.cpPlaceOfDeposit
        Case ELinkTermType.elPublicationNote
            sLinksTable = GetTableName(cttSFA, ESectionType.SECeditorial)
            sMasterTable = gk_sDBTABLESFA
            sIDColumnName = gk_sDBCOLSFAReferenceID
            eProperty = ECATALOGUEPROPERTIES.cpPublicationNote
        Case ELinkTermType.elUnPublishedFindingAids
            sLinksTable = GetTableName(cttSFA, ESectionType.SECeditorial)
            sMasterTable = gk_sDBTABLESFA
            sIDColumnName = gk_sDBCOLSFAReferenceID
            eProperty = ECATALOGUEPROPERTIES.cpUnpublishedFindingAids
        Case ELinkTermType.elIndexTermCorporateBody
            sLinksTable = GetTableName(cttCorporateBody, ESectionType.SECeditorial)
            sMasterTable = gk_sDBTABLECorporateBody
            sIDColumnName = gk_sDBCOLCorporateBodyReferenceID
            eProperty = ECATALOGUEPROPERTIES.cpIndexTermCorporateName
        Case ELinkTermType.elIndexTermPerson
            sLinksTable = GetTableName(cttPerson, ESectionType.SECeditorial)
            sMasterTable = gk_sDBTABLEPerson
            sIDColumnName = gk_sDBCOLPersonReferenceID
            eProperty = ECATALOGUEPROPERTIES.cpIndexTermPersonName
        Case ELinkTermType.elIndexTermPlace
            sLinksTable = GetTableName(cttPlace, ESectionType.SECeditorial)
            sMasterTable = gk_sDBTABLEPlace
            sIDColumnName = gk_sDBCOLPlaceReferenceID
            eProperty = ECATALOGUEPROPERTIES.cpIndexTermPlaceName
        Case ELinkTermType.elIndexTermSubject
            sLinksTable = GetTableName(cttSubject, ESectionType.SECeditorial)
            sMasterTable = gk_sDBTABLESubject
            sIDColumnName = gk_sDBCOLSubjectReferenceID
            eProperty = ECATALOGUEPROPERTIES.cpIndexTermSubject
            '   d7      crm     02Feb2000
        Case ELinkTermType.elPublicationNoteLeaflet
            sLinksTable = GetTableName(cttLeaflet, ESectionType.SECeditorial)
            sMasterTable = gk_sDBTABLELeaflet
            sIDColumnName = gk_sDBCOLLeafletReferenceID
            eProperty = ECATALOGUEPROPERTIES.cpPublicationNote
        Case ELinkTermType.elUnPublishedFindingAidLeaflet
            sLinksTable = GetTableName(cttLeaflet, ESectionType.SECeditorial)
            sMasterTable = gk_sDBTABLELeaflet
            sIDColumnName = gk_sDBCOLLeafletReferenceID
            eProperty = ECATALOGUEPROPERTIES.cpUnpublishedFindingAids
    End Select
        
    'Delete the link entries from the table
    sSql = _
        " DELETE  " & sLinksTable & _
        " WHERE " & _
            gk_sDBCOLEditSetID & "=" & t_lEditSetID & " AND " & _
            gk_sDBCOLCatalogueID & "=" & t_lCatalogueID & " AND " & _
            gk_sDBCOLCatalogueLevel & "=" & GetCatalogueLevelCode(t_eCatalogueLevelID) & " AND " & _
            gk_sDBCOLPropertyID & "=" & eProperty

    'Execute the sql
    m_oFWDBHelper.Execute sSql
    
    For Each recLinkItemRecord In t_colLinkItems
        'Columns
        sSql = _
            " INSERT " & sLinksTable & "(" & _
            sIDColumnName & "," & _
            gk_sDBCOLEditSetID & "," & _
            gk_sDBCOLCatalogueID & "," & _
            gk_sDBCOLCatalogueLevel & "," & _
            gk_sDBCOLPropertyID & "," & _
            gk_sDBCOLDescription & ")"
    
        'Values
        sSql = _
            sSql & " VALUES ( " & _
            recLinkItemRecord.lID & "," & _
            t_lEditSetID & "," & _
            t_lCatalogueID & "," & _
            GetCatalogueLevelCode(t_eCatalogueLevelID) & "," & _
            eProperty
        
        'to insert null if the string is empty
        If Not t_bHasOptionalDescription Or VariantStringAsString(recLinkItemRecord.sOptionalDescription) = "" Then
            sSql = sSql & "," & "Null"
        Else
            sSql = sSql & "," & QuoteValueWithQuotes(VariantStringAsString(recLinkItemRecord.sOptionalDescription))
        End If
        
        sSql = sSql & ")"
         
        'Execute the sql
        m_oFWDBHelper.Execute sSql
    
    Next
    
    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub
      
Public Sub UpdateRelatedSeparateMaterialLinks( _
    ByVal t_lEditSetID As Long, _
    ByVal t_lCatalogueID As Long, _
    ByVal t_eCatalogueLevelID As ECATALOGUELEVEL, _
    ByVal t_eLinkTermType As ELinkTermType, _
    ByVal t_colLinkItems As Collection)

'Description: Update the related and separate material link tables
'
' Amendment history:
'   d1     ndk      25Aug1999      Created
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String, sSql As String
    Dim vRelatedSeparateMaterial As Variant, vCatalogueLink As Variant
    
    sSource = Me.Name & ".UpdateRelatedSeparateMaterialLinks"
        
    Dim sMasterTable As String, sLinksTable As String
    Dim sIDColumnName As String, eProperty As ECATALOGUEPROPERTIES
    Dim recLinkItemRecord  As Variant
    
    If t_eLinkTermType = ELinkTermType.elRelatedMaterial Then
        eProperty = ECATALOGUEPROPERTIES.cpRelatedMaterials
    ElseIf t_eLinkTermType = ELinkTermType.elSeparateMaterial Then
        eProperty = ECATALOGUEPROPERTIES.cpSeperatedMaterials
    End If
    
    'Delete the current links
    sSql = _
        " DELETE " & GetTableName(cttLink, ESectionType.SECeditorial) & _
        " WHERE " & _
        gk_sDBCOLCatalogueID1 & "=" & t_lCatalogueID & " AND " & _
        gk_sDBCOLCatalogueLevel1 & "=" & GetCatalogueLevelCode(t_eCatalogueLevelID) & " AND " & _
        gk_sDBCOLPropertyID & "=" & eProperty & " AND " & _
        gk_sDBCOLEditSetID & "=" & t_lEditSetID
    
    'Execute the sql
    m_oFWDBHelper.Execute sSql
    
    'Insert every link item
    For Each vRelatedSeparateMaterial In t_colLinkItems
        sSql = _
            " INSERT " & GetTableName(cttLink, ESectionType.SECeditorial) & "(" & _
            gk_sDBCOLEditSetID & "," & _
            gk_sDBCOLPropertyID & "," & _
            gk_sDBCOLCatalogueID1 & "," & _
            gk_sDBCOLCatalogueLevel1 & "," & _
            gk_sDBCOLCatalogueID2 & "," & _
            gk_sDBCOLCatalogueLevel2 & "," & _
            gk_sDBCOLDescription & ")" & _
        " VALUES (" & _
            t_lEditSetID & "," & _
            eProperty & "," & _
            t_lCatalogueID & "," & _
            GetCatalogueLevelCode(t_eCatalogueLevelID) & "," & _
            vRelatedSeparateMaterial.lCatalogueID & "," & _
            vRelatedSeparateMaterial.iCatalogueLevelID & "," & _
            QuoteValueWithQuotes(vRelatedSeparateMaterial.sOptionalDescription) & ")"
            'Execute the aql
            m_oFWDBHelper.Execute sSql
    Next        'Move to the next link
    
    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub
      
Public Sub UpdateUnknownDescription( _
    ByVal t_lEditSetID As Long, _
    ByVal t_lCatalogueID As Long, _
    ByVal t_eCatalogueLevelID As ECATALOGUELEVEL, _
    ByVal t_eProperty As ECATALOGUEPROPERTIES, _
    ByVal t_vUnknownDescription As Variant)

'
'Description: Update unknown description for the link items
'
' Amendment history:
'   d1     ndk      25Aug1999      Created
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String, sSql As String
    
    sSource = Me.Name & ".UpdateUnknownDescription"
        
    'Delete the existing description from the table
    sSql = _
        " DELETE  " & gk_sDBTABLEEditsetPCLinkDataElement & _
        " WHERE " & _
            gk_sDBCOLEditSetID & "=" & t_lEditSetID & " AND " & _
            gk_sDBCOLCatalogueID & "=" & t_lCatalogueID & " AND " & _
            gk_sDBCOLCatalogueLevel & "=" & GetCatalogueLevelCode(t_eCatalogueLevelID) & " AND " & _
            gk_sDBCOLPropertyID & "=" & t_eProperty
            
    'Execute the sql
    m_oFWDBHelper.Execute sSql
    
    If Not VariantStringAsString(t_vUnknownDescription) = "" Then
        sSql = _
            " INSERT " & gk_sDBTABLEEditsetPCLinkDataElement & "(" & _
                gk_sDBCOLEditSetID & "," & _
                gk_sDBCOLCatalogueID & "," & _
                gk_sDBCOLCatalogueLevel & "," & _
                gk_sDBCOLPropertyID & "," & _
                gk_sDBCOLDescription & ")" & _
            " VALUES (" & _
                t_lEditSetID & "," & _
                t_lCatalogueID & "," & _
                GetCatalogueLevelCode(t_eCatalogueLevelID) & "," & _
                t_eProperty & "," & _
                QuoteValueWithQuotes(t_vUnknownDescription) & ")"
            
        'Execute the sql
        m_oFWDBHelper.Execute sSql
    End If
    
    Exit Sub
    
ErrorHandler:
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
'   d1     ndk      25Aug1999      Created
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
                AddDataToSerialisedString(SerialiseLinkItem, CStr(recLinkItemRecord.lID))
            'Serialise the description
            SerialiseLinkItem = _
                AddDataToSerialisedString(SerialiseLinkItem, VariantStringAsString(recLinkItemRecord.sOptionalDescription))
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
'   d1     ndk      25Aug1999      Created
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim recLinkItemRecord  As GenericLinkRecord, iTotalRecords As Integer, iCounter As Integer
    
    sSource = Me.Name & ".ExtractLinkItemsFromSerialIsedString"
    
    Set ExtractLinkItemsFromSerialIsedString = New Collection
    
    'Get the total items to be put in the collection,  if there are 0 items then the collection will remain empty
    iTotalRecords = CInt(GetDataFromSerialisedString(t_sSerialisedString))
    
    For iCounter = 1 To iTotalRecords
        'Get the id
        recLinkItemRecord.lID = CLng(GetDataFromSerialisedString(t_sSerialisedString))
        'Get the description
        recLinkItemRecord.sOptionalDescription = GetDataFromSerialisedString(t_sSerialisedString)
        
        'Store it on the collection
        ExtractLinkItemsFromSerialIsedString.Add recLinkItemRecord, gk_sPrefixForLinkTerm & recLinkItemRecord.lID
    Next
    
    Exit Function
    
ErrorHandler:
    ConvertError sSource
End Function


Public Function SerialiseRelatedSeparatedMaterial( _
    ByVal t_sSerialIsedRelatedSeparatedMaterial As String, _
    ByVal t_vColLinkItem As Collection) As String
'
'Description: Convert the link items of related and separate material (id and description) stored in the
'             collection into a string to be stored in persistence manager
'
' Amendment history:
'   d1     ndk      25Aug1999      Created
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim recLinkItemRecord  As Variant, recCatalogueLink As Variant
                
    sSource = Me.Name & ".SerialiseRelatedSeparatedMaterial"
        
    SerialiseRelatedSeparatedMaterial = t_sSerialIsedRelatedSeparatedMaterial
    
    'If there are no items in the collection store 0 in the serialised string
    If t_vColLinkItem.Count = 0 Then
        SerialiseRelatedSeparatedMaterial = AddDataToSerialisedString(SerialiseRelatedSeparatedMaterial, CStr(0))
    Else
        'Add the number of rows in the main collection
        SerialiseRelatedSeparatedMaterial = AddDataToSerialisedString(SerialiseRelatedSeparatedMaterial, CStr(t_vColLinkItem.Count))
                
        'For each record
        For Each recLinkItemRecord In t_vColLinkItem
            'Add the description
            SerialiseRelatedSeparatedMaterial = _
                AddDataToSerialisedString(SerialiseRelatedSeparatedMaterial, VariantStringAsString(recLinkItemRecord.sOptionalDescription))    'Description
            
                SerialiseRelatedSeparatedMaterial = _
                    AddDataToSerialisedString(SerialiseRelatedSeparatedMaterial, recLinkItemRecord.lCatalogueID)   'CatalogueID
                SerialiseRelatedSeparatedMaterial = _
                    AddDataToSerialisedString(SerialiseRelatedSeparatedMaterial, recLinkItemRecord.iCatalogueLevelID)   'Catalogue Level
            Next    'Next link for the description
    End If
    
    Exit Function
    
ErrorHandler:
    ConvertError sSource
End Function


Public Function ExtractRelatedSeparateMaterialFromSerialIsedString( _
    ByRef t_sSerialIsedRelatedSeparatedMaterial As String) As Collection
'
'Description: Convert the link items of related and separate material (id and description) stored in the string into
'             a collection
'
' Amendment history:
'   d1     ndk      25Aug1999      Created
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim recRelatedSeparateMaterial As RelatedSeparateMaterial
    Dim iTotalRecords  As Integer, iCounter As Integer
        
    sSource = Me.Name & ".ExtractRelatedSeparateMaterialFromSerialIsedString"
    
    Set ExtractRelatedSeparateMaterialFromSerialIsedString = New Collection
    
    'Get the number of rows in the collection
    iTotalRecords = CInt(GetDataFromSerialisedString(t_sSerialIsedRelatedSeparatedMaterial))
    
    For iCounter = 1 To iTotalRecords
        'Get the description, catalogue id and catalogue level
        recRelatedSeparateMaterial.sOptionalDescription = GetDataFromSerialisedString(t_sSerialIsedRelatedSeparatedMaterial)
        recRelatedSeparateMaterial.lCatalogueID = CLng(GetDataFromSerialisedString(t_sSerialIsedRelatedSeparatedMaterial))
        recRelatedSeparateMaterial.iCatalogueLevelID = CLng(GetDataFromSerialisedString(t_sSerialIsedRelatedSeparatedMaterial))
        
        'Add the record to the main collection
        ExtractRelatedSeparateMaterialFromSerialIsedString.Add recRelatedSeparateMaterial
        
    Next    'Next record
    
    Exit Function
    
ErrorHandler:
    ConvertError sSource
End Function


Public Sub SetCatalogueStatusAsModified( _
    ByVal sExtensionTableName As String, _
    ByVal t_sConditionString As String)
'
'Description: Updates the edit_set_entry_type_id field
'             on updation of a catalogue entry
'
' Amendment history:
'   d2      ndk     21Oct1999      Created
'   d3      ndk     02Nov1999      Changed the name of this method from UpdateEditSetEntryTypeStatus
'                                  to SetCatalogueStatusAsModified

    On Error GoTo ErrorHandler
    
    Dim sSource As String, sSql As String
    
    sSource = Me.Name & ".SetCatalogueStatusAsModified"
    
    'The followin sql does the following
    'If the status is Created it sets it as CreatedAndModified
    'if the status is AddedFromLive it sets it as AddedFromLiveAndModified
    'if the status is Imported it sets it as ImportedAndModified
    
    'The updation is done only if the current value is any of these (Created,AddedFromLive,Imported)
    
    sSql = _
        " UPDATE " & sExtensionTableName & _
        " SET " & gk_sDBCOLEntryTypeID & " = " & _
            " CASE " & gk_sDBCOLEntryTypeID & " " & _
            " WHEN " & ECATALOGUEENTRYSTATUS.cesPROCATCreated & _
                " THEN " & ECATALOGUEENTRYSTATUS.cesPROCATCreatedModified & _
            " WHEN " & ECATALOGUEENTRYSTATUS.cesAddedFromLive & _
                " THEN " & ECATALOGUEENTRYSTATUS.cesAddedFromLiveModified & _
            " WHEN " & ECATALOGUEENTRYSTATUS.cesImported & _
                " THEN " & ECATALOGUEENTRYSTATUS.cesImportedModified & _
            " ELSE " & gk_sDBCOLEntryTypeID & " END " & _
        " WHERE " & t_sConditionString & " AND " & _
            gk_sDBCOLEntryTypeID & " in (" & _
                ECATALOGUEENTRYSTATUS.cesPROCATCreated & "," & _
                ECATALOGUEENTRYSTATUS.cesAddedFromLive & "," & _
                ECATALOGUEENTRYSTATUS.cesImported & ")"
    'Execute the sql
    m_oFWDBHelper.Execute sSql
    
    Exit Sub
    
ErrorHandler:
    ConvertError sSource
End Sub

Public Sub UpdateRecordStatus(t_lEditSetID As Long, _
                              t_lCatID As Long, _
                              t_eLevel As ECATALOGUELEVEL, _
                              t_lFlagValue As Long)
'
'Description: Uses SP to update the record_status flag in the edit set
'             and in live (if it has been changed). For pieces and
'             items only. t_lFlag value contains the combined values
'             of the AccessionedAsCopy, ExtractOnTransfer, AR and
'             Orderable flags to update the record status field with
'
' Amendment history:
'   d6  WDP     3-Dec-1999      Created

    On Error GoTo ErrorHandler
    
    Dim sSource As String, sSql As String
    
    sSource = Me.Name & ".UpdateRecordStatus"
    
    sSql = " sp_edit_update_record_status " & t_lEditSetID & ", " & _
            t_lCatID & ", " & t_eLevel & ", " & t_lFlagValue
    
    'Execute the sql
    m_oFWDBHelper.Execute sSql
    
    Exit Sub
    
ErrorHandler:
    ConvertError sSource
End Sub

Public Sub BulkUpdateRecordStatus(t_lEditSetID As Long, _
                                  t_eLevel As ECATALOGUELEVEL, _
                                  t_lFlagValue As Long, _
                                  t_bTurnOn As Boolean)
'
'Description: Uses SP to bulk update the record_status flag in the edit set
'             and in live (if it has been changed). For pieces and
'             items only. t_lFlag value contains the combined values
'             of the AccessionedAsCopy, ExtractOnTransfer, AR and
'             Orderable flags to update the record status field with
'
' Amendment history:
'   d9  ALM     02-Oct-2001      Created based on UpdateRecordStatus
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String, sSql As String
    
    sSource = Me.Name & ".BulkUpdateRecordStatus"
    
    sSql = " sp_edit_bulk_update_record_status " & t_lEditSetID & ", " _
           & t_eLevel & ", " & t_lFlagValue & ", " & CStr(Abs(CInt(t_bTurnOn)))
    
    'Execute the sql
    m_oFWDBHelper.Execute sSql
    
    Exit Sub
    
ErrorHandler:
    ConvertError sSource
End Sub


