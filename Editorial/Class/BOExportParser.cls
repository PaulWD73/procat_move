VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BOExportParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    BOExportParser.cls
' System:       PRO, PROCAT
' Copyright:    (C) Quidnunc Limited
'
' Description:  Class to add all the parsed catalogue entries genertaed by BOEADExport to
'               an XML tree.
'
'
' Amendment history:
'   Vers    Auth        Date            Comment
'   d1      A.Farrow    01-Aug-2000     Created from prototype parser
'   d2      A.Farrow    11-Aug-2000     Added member variables and code to construct
'   d3      D Johnston  15-Aug-2000     Added checking for null numeric fields and added
'                                       the dates for the index terms
'   d4      D Johnston  18-Sep-2000     Added otherlevel code for headers in BuildFirstSection
'   d5      D Johnston  31-Oct-2000     Fixed problem in Title field as the title XSL pattern is
'                                       nested and StripUnwantedTags cannot handle it
'   d6      Tom Howden  25-Feb-2002     Fixed defect (0214)
'                                       problem when closure code = NULL.
'   d7      F Devereux  01-Nov-2002     Fixed multiple unittitle elements, incident 0239
'   d8      F Devereux  01-Nov-2002     Applied multiple unittitle fix to other elements
'
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Class constants

'+ Move
Private Const gk_sDBTABLEEADData As String = "EADData"
Private Const gk_sDBCOLEADDataID As String = "ID"
Private Const gk_sDBCOLEADDataValue As String = "Value"

Private Const ksElementDID As String = "./did"
Private Const ksElementDIDReplace As String = "."

' Switch for using AF IDs rather than text matching
' d10 set to true, use attribute to get id
Private Const m_kbUseAFIDs As Boolean = True

' Switch for doing hardcoded language map (eg from eng to English)
' d10 changed meaning, true = use ISO_code column , false = use language_desc column
Private Const m_kbDoLanguageMap As Boolean = True

' Level names
'+ Read from ref data
Private m_ksLevelLettercode As String
Private m_ksLevelDivision As String
Private m_ksLevelClass As String
Private m_ksLevelHeader As String
Private m_ksLevelSubHeader As String
Private m_ksLevelPiece As String
Private m_ksLevelItem As String

' Delimeters for parsing IDs
Private m_ksIDDelimeter As String
Private m_ksIDDivision As String
Private m_ksIDClass As String
Private m_ksIDSubclass As String
Private m_ksIDHeader As String
Private m_ksIDSubheader As String
Private m_ksIDPiece As String
Private m_ksIDItem As String

' Attribute values
Private m_ksAttrRefPROCAT As String
Private m_ksAttrRefFormerPRO As String
Private m_ksAttrRefFormerDept As String
Private m_ksAttrCoveringDates As String
Private m_ksAttrAccumulationDates As String
Private m_ksAttrScaleNumber As String
Private m_ksAttrCreatorNames As String
Private m_ksAttrPlaceDeposit As String
Private m_ksAttrLevel As String

' Attribute names
Private m_ksAttrNameLabel As String
Private m_ksAttrNameID As String
Private m_ksAttrNameLevel As String
Private m_ksAttrNameOtherLevel As String
Private m_ksAttrLegalStatus As String
Private m_ksAttrNameLegalStatus As String
Private m_ksAttrNameOtherLegalStatus As String
Private m_ksAttrNameLanguage As String
Private m_ksAttrNameTarget As String

' Patterns for extracting elements
Private m_ksEADTag As String
Private m_ksArchdescTag As String
Private m_ksElementTitle As String
Private m_ksElementReference As String
Private m_ksElementCreatorNamesCorpBody As String
Private m_ksElementCreatorNamesPerson As String
Private m_ksElementCoveringDates As String
Private m_ksElementPhysDescExtent As String
Private m_ksElementPhysDescForm As String
Private m_ksElementPhysDescDimensions As String
Private m_ksElementPhysDescCondition As String
Private m_ksElementMapScale As String
Private m_ksElementMapDesignation As String
Private m_ksElementPlaceDeposit As String
Private m_ksElementAccessConditions As String
Private m_ksElementRestrictionsUse As String
Private m_ksElementSourceAcquisition As String
Private m_ksElementCustodialHistory As String
Private m_ksElementAccumulationDates As String
Private m_ksElementAppraisal As String
Private m_ksElementAccruals As String
Private m_ksElementAltFormAvail As String
Private m_ksElementScopeContent As String
Private m_ksElementArrangement As String
Private m_ksElementPublicationNote As String
Private m_ksElementUnpublishedFindingAids As String
Private m_ksElementRelatedMaterial As String
Private m_ksElementSeparatedMaterial As String
Private m_ksElementNote As String
Private m_ksElementAdminHistory As String
Private m_ksElementCorpName As String
Private m_ksElementPerson As String
Private m_ksElementPlace As String
Private m_ksElementSubject As String
Private m_ksElementNameCorporateBody As String
Private m_ksElementNamePerson As String
Private m_ksElementExtref As String
Private m_ksElementNameRef As String
Private m_ksElementCopiesInformation As String
Private m_ksElementLocationOriginals As String

' Lengths
Private Const m_klLengthRefLettercode As Long = 4
Private Const m_klLengthRefPiece As Long = 40
Private Const m_klLengthRefItem As Long = 40
Private Const m_klLengthRefPROCAT As Long = 100
Private Const m_klLengthRefFormerPRO As Long = 200
Private Const m_klLengthRefFormerDept As Long = 200
Private Const m_klLengthCoveringDates As Long = 255
Private Const m_klLengthPhysDescDimensions As Long = 20
Private Const m_klLengthPhysDescCondition As Long = 255
Private Const m_klLengthRestrictionsUse As Long = 60
Private Const m_klLengthAccumulationDates As Long = 255
Private Const m_klLengthMapDesignation As Long = 100
Private Const m_klLengthSingleLinkTerm As Long = 8000
Private Const m_klLengthLargeText As Long = 8000

Private Const mk_sPersonIndexText As String = "Person index ID attribute"
Private Const mk_sCorpNameIndexText As String = "Corporate body index ID attribute"
Private Const mk_sPlaceIndexText As String = "Place index ID attribute"
Private Const mk_sSubjectIndexText As String = "Subject index ID attribute"
Private Const mk_sPlaceOfDepIndexText As String = "Place of deposit ID attribute"
Private Const mk_sCreatorCBIndexText As String = "Creator names corporate body ID attribute"
Private Const mk_sCreatorPerIndexText As String = "Creator names person ID attribute"
Private Const mk_sLocationCBIndexText As String = "Location originals corp body ID attribute"
Private Const mk_sCopiesCBIndexText As String = "Copies corp body ID attribute"
Private Const mk_sSourceOfAcqCBIndexText As String = "Source of acquisition corp body ID attribute"
Private Const mk_sSourceOfAcqPerIndexText As String = "Source of acquisition person ID attribute"
Private Const mk_sAuthFileIDAttr As String = "authfilenumber"

Private Const m_ksDSCTagName As String = "dsc"
Private Const m_ksDSCTagAttribute As String = "type"
Private Const m_ksDSCTagValue As String = "combined"

'Export front end stuff
Private m_ksExportEADHeaderXML As String
Private m_ksExportFrontMatterXML As String
Private m_ksExportXMLVersion As String
Private m_ksExportDocType As String

' Patterns for getting child lists
Private m_ksElementArchdescChildList As String
Private m_ksElementCEChildList As String


' Enum
Private Enum ELinkDataItemState
    ENotInPair
    EReadPair
    EReadDescription
    EReadCorporateBody
    EReadperson
End Enum

' Instance variables

' EAD file
Private m_sEADFile As String

' The list of DOM nodes found by the hierarchy walk
Private m_colNodeList As Collection

' List of parent ParsedCatalogueEntry objects added in by parser
Private m_colParentList As Collection

' The current list index for reading out of the cat entry list
Private m_lListIndex As Long

' The current list index for reading out of the parent list
Private m_lParentListIndex As Long

' The current catalogue entry to be returned
Private m_catEntry As New ParsedCatalogueEntry

'! Current Level, ID and Reference
Private m_eLevel As ECATALOGUELEVEL
Private m_sID As String
Private m_sReference As String

' State list of current parent references (implied by hierarchy)
Private m_sLettercode As String
Private m_lDivisionNo As Long
Private m_lClassNo As Long
Private m_lSubClassNo As Long
Private m_lHeaderNo As Long
Private m_lSubheaderNo As Long
Private m_sPieceRef As String
Private m_sItemRef As String

' Database searcher object
Private m_ofwSearcher As New FWDBHelper

' Object for querying catalogue element matrix
Private m_oboCatalogueMatrix As New BOCatalogueElementMatrix

'The DOM object
Dim m_domXmlDoc As DOMDocument

'The current attribue name and value
Private m_sCurrentAttrName As String
Private m_sCurrentAttrValue As Variant
Private m_sCanBeMultiple As Boolean
Private m_bCheckingAttr As Boolean

'Flag to indicate we are using the test parsing from VWTestEADParse
'If this is true (default is false) then we use m_sLogfile
'to log to
Private m_bForTestParse As Boolean
Private m_sLogfile As String

Private m_lBatchID As Long

Private Sub SetReferenceData()
'   Get instance data from the database
'
'   Vers    Auth        Date            Comment
'   d1      A.Farrow    01-Aug-2000     Created from XML-Parser SetReferenceData()
'

On Error GoTo ErrorHandler:
Dim sSource As String
sSource = "SetReference.SetReferenceData"
    
    
    Dim sSql As String
    sSql = "SELECT " & gk_sDBCOLEADDataID & ", " & gk_sDBCOLEADDataValue & _
            " FROM " & gk_sDBTABLEEADData & " (nolock)"
    
    Dim ofwResults As FWRecordList
    Set ofwResults = m_ofwSearcher.FindBySql(sSql)
    
    If ofwResults.RecordCount = 0 Then
        RaiseError sSource, PEGeneralInternalError, "Could not read EAD reference data from the database."
    End If
        
    Dim sKey As String
    Dim sValue As String
    ofwResults.MoveFirst
    While Not ofwResults.IsAtEnd
        sKey = Trim(LCase(ofwResults.GetFieldValue(gk_sDBCOLEADDataID)))
        sValue = ofwResults.GetFieldValue(gk_sDBCOLEADDataValue)
        
    Select Case sKey
        Case "lettercodelevel"
            m_ksLevelLettercode = sValue
        Case "divisionlevel"
            m_ksLevelDivision = sValue
        Case "classlevel"
            m_ksLevelClass = sValue
        Case "headerlevel"
            m_ksLevelHeader = sValue
        Case "subheaderlevel"
            m_ksLevelSubHeader = sValue
        Case "filelevel"
            m_ksLevelPiece = sValue
        Case "itemlevel"
            m_ksLevelItem = sValue
        Case "iddelimeter"
             m_ksIDDelimeter = sValue
        Case "divisiondelimeter"
             m_ksIDDivision = sValue
        Case "classdelimeter"
            m_ksIDClass = sValue
        Case "subclassdelimeter"
            m_ksIDSubclass = sValue
        Case "headerdelimeter"
            m_ksIDHeader = sValue
        Case "subheaderdelimeter"
            m_ksIDSubheader = sValue
        Case "piecedelimeter"
            m_ksIDPiece = sValue
        Case "itemdelimeter"
            m_ksIDItem = sValue
        Case "attrrefprocat"
            m_ksAttrRefPROCAT = sValue
        Case "attrrefformerpro"
            m_ksAttrRefFormerPRO = sValue
        Case "attrrefformerdept"
            m_ksAttrRefFormerDept = sValue
        Case "attrcoveringdates"
            m_ksAttrCoveringDates = sValue
        Case "attraccumulationdates"
            m_ksAttrAccumulationDates = sValue
        Case "attrscalenumber"
            m_ksAttrScaleNumber = sValue
        Case "attrcreatornames"
            m_ksAttrCreatorNames = sValue
        Case "attrplacedeposit"
            m_ksAttrPlaceDeposit = sValue
        Case "attrlevel"
            m_ksAttrLevel = sValue
        Case "attrnamelabel"
            m_ksAttrNameLabel = sValue
        Case "attrnameid"
            m_ksAttrNameID = sValue
        Case "attrnamelevel"
            m_ksAttrNameLevel = sValue
        Case "attrnameotherlevel"
            m_ksAttrNameOtherLevel = sValue
        Case "attrlegalstatus"
            m_ksAttrLegalStatus = sValue
        Case "attrnamelegalstatus"
            m_ksAttrNameLegalStatus = sValue
        Case "attrnameotherlegalstatus"
            m_ksAttrNameOtherLegalStatus = sValue
        Case "attrnamelanguage"
            m_ksAttrNameLanguage = sValue
        Case "eadtag"
            m_ksEADTag = sValue
        Case "archdesctag"
            m_ksArchdescTag = sValue
        Case "elementtitle"
            m_ksElementTitle = sValue
        Case "elementreference"
            m_ksElementReference = sValue
        Case "elementcreatornamescorpbody"
            m_ksElementCreatorNamesCorpBody = sValue
        Case "elementcreatornamesperson"
            m_ksElementCreatorNamesPerson = sValue
        Case "elementcoveringdates"
            m_ksElementCoveringDates = sValue
        Case "elementphysdescextent"
            m_ksElementPhysDescExtent = sValue
        Case "elementphysdescform"
            m_ksElementPhysDescForm = sValue
        Case "elementphysdescdimensions"
            m_ksElementPhysDescDimensions = sValue
        Case "elementphysdesccondition"
            m_ksElementPhysDescCondition = sValue
        Case "elementmapscale"
            m_ksElementMapScale = sValue
        Case "elementplacedeposit"
            m_ksElementPlaceDeposit = sValue
        Case "elementaccessconditions"
            m_ksElementAccessConditions = sValue
        Case "elementrestrictionsuse"
            m_ksElementRestrictionsUse = sValue
        Case "elementsourceacquisition"
            m_ksElementSourceAcquisition = sValue
        Case "elementcustodialhistory"
            m_ksElementCustodialHistory = sValue
        Case "elementaccumulationdates"
            m_ksElementAccumulationDates = sValue
        Case "elementappraisal"
            m_ksElementAppraisal = sValue
        Case "elementaccruals"
            m_ksElementAccruals = sValue
        Case "elementaltformavail"
            m_ksElementAltFormAvail = sValue
        Case "elementscopecontent"
            m_ksElementScopeContent = sValue
        Case "elementarrangement"
            m_ksElementArrangement = sValue
        Case "elementpublicationnote"
            m_ksElementPublicationNote = sValue
        Case "elementunpublishedfindingaids"
            m_ksElementUnpublishedFindingAids = sValue
        Case "elementrelatedmaterial"
            m_ksElementRelatedMaterial = sValue
        Case "elementseparatedmaterial"
            m_ksElementSeparatedMaterial = sValue
        Case "elementnote"
            m_ksElementNote = sValue
        Case "elementadminhistory"
            m_ksElementAdminHistory = sValue
        Case "elementcorpname"
            m_ksElementCorpName = sValue
        Case "elementperson"
            m_ksElementPerson = sValue
        Case "elementplace"
            m_ksElementPlace = sValue
        Case "elementsubject"
            m_ksElementSubject = sValue
        Case "elementnamecorporatebody"
            m_ksElementNameCorporateBody = sValue
        Case "elementnameperson"
            m_ksElementNamePerson = sValue
        Case "elementextref"
            m_ksElementExtref = sValue
        Case "elementarchdescchildlist"
            m_ksElementArchdescChildList = sValue
        Case "elementcechildlist"
            m_ksElementCEChildList = sValue
        Case "elementnameref"
            m_ksElementNameRef = sValue
        Case "attrnametarget"
            m_ksAttrNameTarget = sValue
        Case "elementcopiesinformation"
            m_ksElementCopiesInformation = sValue
        Case "elementlocationoriginals"
            m_ksElementLocationOriginals = sValue
            
        ' d20 - Added
        Case "elementmapdesignation"
            m_ksElementMapDesignation = sValue
        
        ' Export specific values
        Case "exporteadheaderxml"
            m_ksExportEADHeaderXML = sValue
        Case "exportfrontmatterxml"
            m_ksExportFrontMatterXML = sValue
        Case "exportxmlversion"
            m_ksExportXMLVersion = sValue
        Case "exportdoctype"
            m_ksExportDocType = sValue
            
        Case Else
            Debug.Print sKey
        End Select

        
        'Move to next record
        ofwResults.MoveNext
    Wend
    
Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub
Public Sub Initialise()
'
'   Reset the instance variables.
'
'   d1  spb     10Aug99     Created
'   d12 WDP     10Feb2000   Set m_bForTestParse = False
On Error GoTo ErrorHandler:
Dim sSource As String
sSource = "XMLParser.Initialise"
    
    m_bForTestParse = False
    ' Get the EAD reference data from the database
    SetReferenceData
    
    ' Reset lists
    Set m_colParentList = New Collection
    m_lParentListIndex = 0
    Set m_colNodeList = New Collection
    m_lListIndex = 0
    
    ' Current values of cat entry
    'm_lLevel = 0
    m_sID = ""
    m_sReference = ""
    
    ' Reference parts state machine
    m_sLettercode = ""
    m_lDivisionNo = 0
    m_lClassNo = 0
    m_lSubClassNo = 0
    m_lHeaderNo = 0
    m_lSubheaderNo = 0
    m_sPieceRef = ""
    m_sItemRef = ""
    
    Set m_domXmlDoc = New DOMDocument
    
Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub


Public Function CreateExportNodes(t_oboParsedCatalogueEntry As ParsedCatalogueEntry, t_iLevel As Integer, t_bIsWanted As Boolean) As IXMLDOMNode
'   Description :  This creates a new DOM and then adds the passed catalogue entry values
'
'   Vers    Auth        Date            Comment
'   d1      A.Farrow    01-Aug-2000     Created
'   d3      D Johnston  15-Aug-2000     Added checking for null numeric fields and added
'                                       the dates for the index terms
'   d5      D Johnston  31-Oct-2000     Fixed problem in Title field as the title XSL pattern is
'                                       nested and StripUnwantedTags cannot handle it
'   d7      F Devereux  01-Nov-2002     Fixed another problem in unittitle field (multiple unittitle
'                                       elements) as the title XSL pattern includes attributes
'                                       and StripUnwantedTags cannot handle it
'   d8      F Devereux  01-Nov-2002     Applied multiple unittitle fix to other elements

Dim sSource As String
sSource = "XMLParser.Parse: "
    
    ' Create the new DOM
    Dim domTitleElement As IXMLDOMElement
    Set m_domXmlDoc = New DOMDocument
    Dim domText As IXMLDOMText
    Dim tempXML As IXMLDOMElement
    
    Dim oboPerson As PersonAFTerm
    Dim oboCorp As CorporateBodyAFTerm
    Dim oboPlace As PlaceAFTerm
    Dim oboSubject As SubjectAFTerm
    Dim oBOAFPersonLoader As BOAFPersonLoader
    Dim oboAfCorpLoader As BOAFCorpBodyLoader
    Dim oboAFPlaceLoader As BOAFPlaceLoader
    Dim oboAFSubjectLoader As BOAFSubjectLoader
    
    Dim lCreatorNames As Collection
    Dim clinks As Collection
    Dim lIdInCollection As Long
    Dim i As Integer
    Dim typair As GenericLinkRecord
    Dim sPattern As String
    Dim sRemaining As String
    Dim sData As String
    
    ' d5 - Added temporary pattern for title tags in the database
    Dim sUnnestedTitlePattern As String
    
    Set oBOAFPersonLoader = New BOAFPersonLoader
    Set oboAfCorpLoader = New BOAFCorpBodyLoader
    Set oboAFPlaceLoader = New BOAFPlaceLoader
    Set oboAFSubjectLoader = New BOAFSubjectLoader
    
    m_eLevel = GetCatalogueLevelEnum(t_iLevel)
    
    Set domTitleElement = BuildFirstSection(t_oboParsedCatalogueEntry)
    
    ' We must add a <did> tag regardless of whether we are going to put data in it or not
    Set tempXML = BuildXml(ksElementDID, domTitleElement)
    
    If t_bIsWanted = True Then
        
    
    
    ' Now we can keep adding all the elements to this titleelement
    'Title
    If m_oboCatalogueMatrix.LevelHasCatalogueProperty(m_eLevel, cpTitle) Then
        Set tempXML = BuildXml(m_ksElementTitle, domTitleElement)
        StripUnwantedTags tempXML, StripDidAndAttrsFromPattern(m_ksElementTitle), t_oboParsedCatalogueEntry.Title
        tempXML.setAttribute m_sCurrentAttrName, m_sCurrentAttrValue
    End If
    
    'These all use the same pattern so we set the boolean
    m_sCanBeMultiple = True
    
        'Reference
    If m_oboCatalogueMatrix.LevelHasCatalogueProperty(m_eLevel, cpReference) Then
        Set tempXML = BuildXml(m_ksElementReference, domTitleElement)
        tempXML.Text = GetRefFromDBID(m_eLevel, t_oboParsedCatalogueEntry.ID)
        tempXML.setAttribute m_ksAttrNameLabel, m_ksAttrRefPROCAT
    End If
    
    m_sCurrentAttrName = ""
    m_sCurrentAttrValue = ""
    
    'Former department reference
    If m_oboCatalogueMatrix.LevelHasCatalogueProperty(m_eLevel, cpFormerDepartmentReference) Then
        Set tempXML = BuildXml(m_ksElementReference, domTitleElement)
        tempXML.Text = t_oboParsedCatalogueEntry.ReferenceDepartment
        tempXML.setAttribute m_ksAttrNameLabel, m_ksAttrRefFormerDept
    End If
    
    m_sCurrentAttrName = ""
    m_sCurrentAttrValue = ""
    
    'Former PRO reference
    If m_oboCatalogueMatrix.LevelHasCatalogueProperty(m_eLevel, cpFormerPROReference) Then
        Set tempXML = BuildXml(m_ksElementReference, domTitleElement)
        tempXML.Text = t_oboParsedCatalogueEntry.ReferencePRO
        tempXML.setAttribute m_ksAttrNameLabel, m_ksAttrRefFormerPRO
    End If
    
    m_sCurrentAttrName = ""
    m_sCurrentAttrValue = ""
    
    ' Map designation
    If m_oboCatalogueMatrix.LevelHasCatalogueProperty(m_eLevel, cpMapDesignation) Then
        Set tempXML = BuildXml(m_ksElementMapDesignation, domTitleElement)
        StripUnwantedTags tempXML, StripDidAndAttrsFromPattern(m_ksElementMapDesignation), t_oboParsedCatalogueEntry.MapDesignation
        tempXML.setAttribute m_sCurrentAttrName, m_sCurrentAttrValue
    End If
    
    m_sCurrentAttrName = ""
    m_sCurrentAttrValue = ""
    
    'Here we iterate through the creator names
    If m_oboCatalogueMatrix.LevelHasCatalogueProperty(m_eLevel, cpCreatorNames) Then
        Set lCreatorNames = t_oboParsedCatalogueEntry.CreatorNamesPersons
        
        For i = 1 To lCreatorNames.Count
            lIdInCollection = lCreatorNames.Item(i)
            oBOAFPersonLoader.IAFLoader_Load lIdInCollection
            oboPerson = oBOAFPersonLoader.Data
        
            Set tempXML = BuildXml(m_ksElementCreatorNamesPerson, domTitleElement)
            tempXML.Text = oboPerson.sPretitle & " " & oboPerson.sForenames & " " & oboPerson.sSurname & "," & _
                oboPerson.sTitle & ", " & DateText(oboPerson.sBirthDate, oboPerson.sDeathDate, oboPerson.sUncertainBirthDateCode, oboPerson.sUncertainDeathDateCode)
            tempXML.setAttribute mk_sAuthFileIDAttr, CStr(oboPerson.lID)
            Set tempXML = tempXML.parentNode
            tempXML.setAttribute m_ksAttrNameLabel, m_ksAttrCreatorNames
        Next
        
    End If
    
    m_sCurrentAttrName = ""
    m_sCurrentAttrValue = ""
    
    'Here we iterate through the corp creators
    If m_oboCatalogueMatrix.LevelHasCatalogueProperty(m_eLevel, cpCreatorNames) Then
        
        Set lCreatorNames = t_oboParsedCatalogueEntry.CreatorNamesCorpBodies
        
        For i = 1 To lCreatorNames.Count
            lIdInCollection = lCreatorNames.Item(i)
            oboAfCorpLoader.IAFLoader_Load lIdInCollection
            oboCorp = oboAfCorpLoader.Data
          
            Set tempXML = BuildXml(m_ksElementCreatorNamesCorpBody, domTitleElement)
            tempXML.Text = oboCorp.sTermText & " " & DateText(oboCorp.sStartDate, oboCorp.sEndDate, oboCorp.sUncertainStartDateCode, oboCorp.sUncertainEndDateCode)
            tempXML.setAttribute mk_sAuthFileIDAttr, CStr(oboCorp.lID)
            Set tempXML = tempXML.parentNode
            tempXML.setAttribute m_ksAttrNameLabel, m_ksAttrCreatorNames
        Next
    End If
    
   
    m_sCurrentAttrName = ""
    m_sCurrentAttrValue = ""
    
    ' Get the Covering dates
    If m_oboCatalogueMatrix.LevelHasCatalogueProperty(m_eLevel, cpCoveringDates) Then
        Set tempXML = BuildXml(m_ksElementCoveringDates, domTitleElement)
        tempXML.Text = t_oboParsedCatalogueEntry.DateText
        tempXML.setAttribute m_ksAttrNameLabel, m_ksAttrCoveringDates
    End If
    
    m_sCurrentAttrName = ""
    m_sCurrentAttrValue = ""
    
    ' Get the Physical descritpion
        'Extent
    If m_oboCatalogueMatrix.LevelHasCatalogueProperty(m_eLevel, cpPhysicalDescription_RecordQuantity) Then
        Set tempXML = BuildXml(m_ksElementPhysDescExtent, domTitleElement)
        
        If t_oboParsedCatalogueEntry.PhysdescExtentNull = False Then
            tempXML.Text = t_oboParsedCatalogueEntry.PhysdescExtent
        End If
    End If
    
    m_sCurrentAttrName = ""
    m_sCurrentAttrValue = ""
    
        'Form
    If m_oboCatalogueMatrix.LevelHasCatalogueProperty(m_eLevel, cpPhysicalDescription_Form) Then
        Set tempXML = BuildXml(m_ksElementPhysDescForm, domTitleElement)
        tempXML.Text = GetPhysicalForm(t_oboParsedCatalogueEntry.PhysdescFormID)
    End If
    
    m_sCurrentAttrName = ""
    m_sCurrentAttrValue = ""
    
        'Dimension
    If m_oboCatalogueMatrix.LevelHasCatalogueProperty(m_eLevel, cpPhysicalDescription_Dimension) Then
        Set tempXML = BuildXml(m_ksElementPhysDescDimensions, domTitleElement)
        tempXML.Text = t_oboParsedCatalogueEntry.PhysdescDimensions
    End If
    
    m_sCurrentAttrName = ""
    m_sCurrentAttrValue = ""
    
        'Condition
    If m_oboCatalogueMatrix.LevelHasCatalogueProperty(m_eLevel, cpPhysicalCondition) Then
        Set tempXML = BuildXml(m_ksElementPhysDescCondition, domTitleElement)
        tempXML.Text = t_oboParsedCatalogueEntry.PhysdescCondition
    End If
    
    ' These can not be replicated more than once
    m_sCanBeMultiple = False
    
    m_sCurrentAttrName = ""
    m_sCurrentAttrValue = ""
    
    ' The map Scale - NOTE we have to manually append a <P> as the pattern does not have it
    If m_oboCatalogueMatrix.LevelHasCatalogueProperty(m_eLevel, cpMapscaleNumber) Then
        Set tempXML = BuildXml(m_ksElementMapScale, domTitleElement)
        tempXML.setAttribute m_ksAttrNameLabel, m_ksAttrScaleNumber
        
        Set tempXML = BuildXml(m_ksElementMapScale & "/p", domTitleElement)
        If t_oboParsedCatalogueEntry.ScaleNumberNull = False Then
            tempXML.Text = t_oboParsedCatalogueEntry.ScaleNumber
        End If
    End If
    
    m_sCurrentAttrName = ""
    m_sCurrentAttrValue = ""
    
    ' The place of deposit
    If m_oboCatalogueMatrix.LevelHasCatalogueProperty(m_eLevel, cpPlaceOfDeposit) Then
        
        ' Build the Place Of Deposit node whether there are entries or not.
        Set tempXML = BuildXml(m_ksElementPlaceDeposit, domTitleElement)
        
        Set lCreatorNames = t_oboParsedCatalogueEntry.PlaceDepositCorpBodies
        
        For i = 1 To lCreatorNames.Count
            lIdInCollection = lCreatorNames.Item(i)
            oboAfCorpLoader.IAFLoader_Load lIdInCollection
            oboCorp = oboAfCorpLoader.Data
          
            m_sCanBeMultiple = True
            Set tempXML = BuildXml(m_ksElementPlaceDeposit & "/" & m_ksElementNameCorporateBody, domTitleElement)
            tempXML.Text = oboCorp.sTermText
            tempXML.setAttribute mk_sAuthFileIDAttr, CStr(oboCorp.lID)
        Next
        
        m_sCanBeMultiple = False
        
    End If
    
    m_sCurrentAttrName = ""
    m_sCurrentAttrValue = ""
    
    'admininfo
        'accessrestriction - This valid for levels 1-5.  Piece and item elements
        '                   display text and 'Closure Code'  which is not infact valid EAD, but
        '                   had been specified by PRO.
    If m_oboCatalogueMatrix.LevelHasCatalogueProperty(m_eLevel, cpAccessCondition) Then
        Set tempXML = BuildXml(m_ksElementAccessConditions, domTitleElement)
        tempXML.Text = GetAccessCondition(t_oboParsedCatalogueEntry.AccessConditionID)
    Else
        Set tempXML = BuildXml(m_ksElementAccessConditions, domTitleElement)
        tempXML.Text = GetClosureInfo(t_oboParsedCatalogueEntry.ID, m_eLevel)
    End If

    m_sCurrentAttrName = ""
    m_sCurrentAttrValue = ""
    
        'User Restrictions
    If m_oboCatalogueMatrix.LevelHasCatalogueProperty(m_eLevel, cpRestrictionsOnUse) Then
        Set tempXML = BuildXml(m_ksElementRestrictionsUse, domTitleElement)
        tempXML.Text = t_oboParsedCatalogueEntry.RestrictionsOnUse
    End If
    
    m_sCurrentAttrName = ""
    m_sCurrentAttrValue = ""
    
    m_sCanBeMultiple = True
        'Aquisition info
    If m_oboCatalogueMatrix.LevelHasCatalogueProperty(m_eLevel, cpImmediateSourceOfAcquisition) Then
        
        Set lCreatorNames = t_oboParsedCatalogueEntry.SourceAcquisitionPersons
        
        For i = 1 To lCreatorNames.Count
            typair = lCreatorNames.Item(i)
            
            'Load the data about the person
            oBOAFPersonLoader.IAFLoader_Load typair.lID
            oboPerson = oBOAFPersonLoader.Data
            
            Set tempXML = BuildXml(m_ksElementSourceAcquisition, domTitleElement)
            Set domText = m_domXmlDoc.createTextNode(typair.sOptionalDescription)
            tempXML.appendChild domText
        
            Set tempXML = BuildXml(m_ksElementSourceAcquisition _
                & "/" & m_ksElementNamePerson, domTitleElement)
            tempXML.Text = oboPerson.sSurname & "," & oboPerson.sForenames
            tempXML.setAttribute mk_sAuthFileIDAttr, CStr(oboPerson.lID)
        Next
    End If
   
    m_sCurrentAttrName = ""
    m_sCurrentAttrValue = ""
    
    '  Get the source of aqusition  corporate bodies
    If m_oboCatalogueMatrix.LevelHasCatalogueProperty(m_eLevel, cpImmediateSourceOfAcquisition) Then
        
        Set lCreatorNames = t_oboParsedCatalogueEntry.SourceAcquisitionCorpBodies
        
        For i = 1 To lCreatorNames.Count
            typair = lCreatorNames.Item(i)
            oboAfCorpLoader.IAFLoader_Load typair.lID
            oboCorp = oboAfCorpLoader.Data
     
            ' Build the link text (optional)
            m_sCanBeMultiple = False
            
            ' We append the text node which represents the link
            Set tempXML = BuildXml(m_ksElementSourceAcquisition, domTitleElement)
            Set domText = m_domXmlDoc.createTextNode(typair.sOptionalDescription)
            tempXML.appendChild domText
            m_sCanBeMultiple = True
            
            Set tempXML = BuildXml(m_ksElementSourceAcquisition _
                & "/" & m_ksElementNameCorporateBody, domTitleElement)
            tempXML.Text = oboCorp.sTermText
            tempXML.setAttribute mk_sAuthFileIDAttr, CStr(oboCorp.lID)
        Next
        
        'Append the hanging text to the bottom of the aquisition text
            m_sCanBeMultiple = False
            Set tempXML = BuildXml(m_ksElementSourceAcquisition, domTitleElement)
            Set domText = m_domXmlDoc.createTextNode(t_oboParsedCatalogueEntry.SourceAcquisitionSingleText)
            tempXML.appendChild domText
            
    End If
    
    m_sCurrentAttrName = ""
    m_sCurrentAttrValue = ""
    
    ' cust history
    If m_oboCatalogueMatrix.LevelHasCatalogueProperty(m_eLevel, cpCustodialHistory) Then
        Set tempXML = BuildXml(m_ksElementCustodialHistory, domTitleElement)
        tempXML.Text = t_oboParsedCatalogueEntry.CustHist
    End If
    
    m_sCurrentAttrName = ""
    m_sCurrentAttrValue = ""
    
    ' accumulation dates
    If m_oboCatalogueMatrix.LevelHasCatalogueProperty(m_eLevel, cpAccumulationDates) Then
        Set tempXML = BuildXml(m_ksElementAccumulationDates, domTitleElement)
        tempXML.Text = t_oboParsedCatalogueEntry.AccumulationDates
    End If
    
    m_sCurrentAttrName = ""
    m_sCurrentAttrValue = ""
    
     ' accurals
    If m_oboCatalogueMatrix.LevelHasCatalogueProperty(m_eLevel, cpAccruals) Then
        Set tempXML = BuildXml(m_ksElementAccruals, domTitleElement)
        tempXML.Text = t_oboParsedCatalogueEntry.Accruals
    End If
    
     ' appraisal
    If m_oboCatalogueMatrix.LevelHasCatalogueProperty(m_eLevel, cpAppraisal_DestructionInformation) Then
        Set tempXML = BuildXml(m_ksElementAppraisal, domTitleElement)
        tempXML.Text = t_oboParsedCatalogueEntry.Appraisal
    End If
    
    m_sCurrentAttrName = ""
    m_sCurrentAttrValue = ""
    
    'Location of originals
    If m_oboCatalogueMatrix.LevelHasCatalogueProperty(m_eLevel, cpLocationOfOriginals) Then
        
        Set lCreatorNames = t_oboParsedCatalogueEntry.LocationOriginalsCorpBodies
        
        For i = 1 To lCreatorNames.Count
            typair = lCreatorNames.Item(i)
            oboAfCorpLoader.IAFLoader_Load typair.lID
            oboCorp = oboAfCorpLoader.Data
               
            Set tempXML = BuildXml(m_ksElementLocationOriginals, domTitleElement)
            ' We append the text node which represents the link
            Set tempXML = BuildXml(m_ksElementLocationOriginals _
                & "/" & m_ksElementNameCorporateBody, domTitleElement)
            tempXML.Text = oboCorp.sTermText
            tempXML.setAttribute mk_sAuthFileIDAttr, CStr(oboCorp.lID)
        Next
        
        'We set the hanging text for this, which in fact part of the GenericLink description
        m_sCanBeMultiple = False
        Set tempXML = BuildXml(m_ksElementLocationOriginals, domTitleElement)
        Set domText = m_domXmlDoc.createTextNode(typair.sOptionalDescription)
        tempXML.appendChild domText
        
        'Set the attribute value for the child node
        sPattern = GetParentNode(m_ksElementCopiesInformation)
        Set tempXML = domTitleElement.selectSingleNode(sPattern)
        Set tempXML = tempXML.lastChild
        tempXML.setAttribute m_sCurrentAttrName, m_sCurrentAttrValue
        
    End If
    
    
    ' NOTE :This has the same structure as location of originals.  Due to the 'eadata' layout
    ' we have to call a special build duplicate method
    m_sCanBeMultiple = True
    m_bCheckingAttr = True
    m_sCurrentAttrName = ""
    m_sCurrentAttrValue = ""
    
    ' Copy Information
    If m_oboCatalogueMatrix.LevelHasCatalogueProperty(m_eLevel, cpCopiesInformation) Then
        
        Set lCreatorNames = t_oboParsedCatalogueEntry.CopiesCorpBodies
        
       
            Dim lPos As Long
            Dim newXML As IXMLDOMElement
            
            sPattern = GetParentNode(m_ksElementCopiesInformation, sRemaining)
       
            'We strip out the node and pass that
            Set tempXML = domTitleElement.selectSingleNode(sPattern)
            Set tempXML = BuildXml(sRemaining, tempXML)
            m_bCheckingAttr = False
             
        For i = 1 To lCreatorNames.Count
           
            typair = lCreatorNames.Item(i)
            oboAfCorpLoader.IAFLoader_Load typair.lID
            oboCorp = oboAfCorpLoader.Data
           
            Set domText = m_domXmlDoc.createTextNode(typair.sOptionalDescription)
            tempXML.appendChild domText

            ' We have slotted into admmininfo now we add to that
           '
            Set newXML = m_domXmlDoc.createElement(m_ksElementNameCorporateBody)
            newXML.setAttribute mk_sAuthFileIDAttr, CStr(oboCorp.lID)
            newXML.Text = oboCorp.sTermText
            tempXML.appendChild newXML
            
        Next
        
        
        'Append the hanging text to the bottom of the aquisition text
        
        'Set tempXML = BuildXml(m_ksElementCopiesInformation, domTitleElement)
        Set domText = m_domXmlDoc.createTextNode(t_oboParsedCatalogueEntry.CopiesSingleText)
        tempXML.appendChild domText
            
        ' we set the attribute value
        Set tempXML = domTitleElement.selectSingleNode(sPattern)
        Set tempXML = tempXML.lastChild
        tempXML.setAttribute m_sCurrentAttrName, m_sCurrentAttrValue
        
    End If
    
    m_sCanBeMultiple = False
    
    'Scope and content
    If m_oboCatalogueMatrix.LevelHasCatalogueProperty(m_eLevel, cpScopeAndContent) Then
        If t_oboParsedCatalogueEntry.ScopeContent = "" Then
            Set tempXML = BuildXml(m_ksElementScopeContent & "/p", domTitleElement)
        Else
            Set tempXML = BuildXml(m_ksElementScopeContent, domTitleElement)
            StripUnwantedTags tempXML, StripDidAndAttrsFromPattern(m_ksElementScopeContent), t_oboParsedCatalogueEntry.ScopeContent
        End If
    End If
    
    'Arrangement
    If m_oboCatalogueMatrix.LevelHasCatalogueProperty(m_eLevel, cpArrangement) Then
        Set tempXML = BuildXml(m_ksElementArrangement, domTitleElement)
        StripUnwantedTags tempXML, StripDidAndAttrsFromPattern(m_ksElementArrangement), t_oboParsedCatalogueEntry.Arrangement
    End If
    
    'Publication Note
    If m_oboCatalogueMatrix.LevelHasCatalogueProperty(m_eLevel, cpPublicationNote) Then
        
        Set tempXML = BuildXml(m_ksElementPublicationNote, domTitleElement)
        tempXML.Text = t_oboParsedCatalogueEntry.PublicationNote
        
        Set clinks = t_oboParsedCatalogueEntry.PublicationNoteLeafletLinks
        'We just append all links as plain text
        For i = 1 To clinks.Count
            typair = clinks.Item(i)
            tempXML.Text = tempXML.Text & " " & typair.sOptionalDescription
        Next
        
        Set clinks = t_oboParsedCatalogueEntry.PublicationNoteSFALinks
        'We just append all links as plain text
        For i = 1 To clinks.Count
            typair = clinks.Item(i)
            tempXML.Text = tempXML.Text & " " & typair.sOptionalDescription
        Next
    End If
    
    'Unpublished finding aids
    If m_oboCatalogueMatrix.LevelHasCatalogueProperty(m_eLevel, cpUnpublishedFindingAids) Then
        
        Set tempXML = BuildXml(m_ksElementUnpublishedFindingAids, domTitleElement)
        tempXML.Text = t_oboParsedCatalogueEntry.UnpublishedFindingAids
        
        Set clinks = t_oboParsedCatalogueEntry.UnpublishedFindingAidLeafletLinks
        'We just append all links as plain text
        For i = 1 To clinks.Count
            typair = clinks.Item(i)
            tempXML.Text = tempXML.Text & " " & typair.sOptionalDescription
        Next
        
        Set clinks = t_oboParsedCatalogueEntry.UnpublishedFindingAidSFALinks
        'We just append all links as plain text
        For i = 1 To clinks.Count
            typair = clinks.Item(i)
            tempXML.Text = tempXML.Text & " " & typair.sOptionalDescription
        Next
    End If
    
    'Related materials
    If m_oboCatalogueMatrix.LevelHasCatalogueProperty(m_eLevel, cpRelatedMaterials) Then
        
        Dim catRef As CatReferenceParts
        Dim sRef As String
    
        Dim linkedMaterial As RelatedSeparateMaterial
        Set clinks = t_oboParsedCatalogueEntry.RelatedMaterialLinks
        
        For i = 1 To clinks.Count
            linkedMaterial = clinks.Item(i)
            
            Set tempXML = BuildXml(m_ksElementRelatedMaterial, domTitleElement)
            Set domText = m_domXmlDoc.createTextNode(linkedMaterial.sOptionalDescription)
            tempXML.appendChild domText
        
            m_sCanBeMultiple = True
            Set tempXML = BuildXml(m_ksElementRelatedMaterial _
                & "/" & m_ksElementNameRef, domTitleElement)
                
            catRef = GetEntireRefFromDBID(GetCatalogueLevelEnum(linkedMaterial.iCatalogueLevelID), linkedMaterial.lCatalogueID)
            sRef = GenerateExportReference(catRef, linkedMaterial.iCatalogueLevelID)
                
            tempXML.Text = GetRefFromDBID(GetCatalogueLevelEnum(linkedMaterial.iCatalogueLevelID), linkedMaterial.lCatalogueID)
            tempXML.setAttribute m_ksAttrNameTarget, sRef
            m_sCanBeMultiple = False
        Next
         'Append the hanging text
        Set tempXML = BuildXml(m_ksElementRelatedMaterial, domTitleElement)
        Set domText = m_domXmlDoc.createTextNode(t_oboParsedCatalogueEntry.RelatedMaterialSingleText)
        tempXML.appendChild domText
        
    End If
    
    
     'Sperated materials
    If m_oboCatalogueMatrix.LevelHasCatalogueProperty(m_eLevel, cpSeperatedMaterials) Then
        
        Set clinks = t_oboParsedCatalogueEntry.SeparatedMaterialLinks
        
        For i = 1 To clinks.Count
            linkedMaterial = clinks.Item(i)
            
            Set tempXML = BuildXml(m_ksElementSeparatedMaterial, domTitleElement)
            Set domText = m_domXmlDoc.createTextNode(linkedMaterial.sOptionalDescription)
            tempXML.appendChild domText
        
            m_sCanBeMultiple = True
            Set tempXML = BuildXml(m_ksElementSeparatedMaterial _
                & "/" & m_ksElementNameRef, domTitleElement)
                
            catRef = GetEntireRefFromDBID(GetCatalogueLevelEnum(linkedMaterial.iCatalogueLevelID), linkedMaterial.lCatalogueID)
            sRef = GenerateExportReference(catRef, linkedMaterial.iCatalogueLevelID)
                
            tempXML.Text = GetRefFromDBID(GetCatalogueLevelEnum(linkedMaterial.iCatalogueLevelID), linkedMaterial.lCatalogueID)
            tempXML.setAttribute m_ksAttrNameTarget, sRef
            m_sCanBeMultiple = False
        Next
        'Append the hanging text
        Set tempXML = BuildXml(m_ksElementSeparatedMaterial, domTitleElement)
        Set domText = m_domXmlDoc.createTextNode(t_oboParsedCatalogueEntry.SeparatedMaterialSingleText)
        tempXML.appendChild domText
        
    End If
    
    'Note element
     If m_oboCatalogueMatrix.LevelHasCatalogueProperty(m_eLevel, cpNote) Then
        Set tempXML = BuildXml(m_ksElementNote, domTitleElement)
        tempXML.Text = t_oboParsedCatalogueEntry.Note
    End If
    
    'bioghist element - must have a mandatory <p> tag
    If m_oboCatalogueMatrix.LevelHasCatalogueProperty(m_eLevel, cpAdministrative_BiographicalHistory) Then
        If t_oboParsedCatalogueEntry.History = "" Then
            ' Value is empty, so put in the <p> tag anyway
            Set tempXML = BuildXml(m_ksElementAdminHistory & "/p", domTitleElement)
        Else
            Set tempXML = BuildXml(m_ksElementAdminHistory, domTitleElement)
            StripUnwantedTags tempXML, StripDidAndAttrsFromPattern(m_ksElementAdminHistory), t_oboParsedCatalogueEntry.History
        End If
    End If
    
    'The index terms
    
    'The Corp name
    If m_oboCatalogueMatrix.LevelHasCatalogueProperty(m_eLevel, cpIndexTermCorporateName) Then
        
        Set clinks = t_oboParsedCatalogueEntry.CorpBodyIndex
        
        For i = 1 To clinks.Count
            lIdInCollection = clinks.Item(i)
            oboAfCorpLoader.IAFLoader_Load lIdInCollection
            oboCorp = oboAfCorpLoader.Data
          
            'Set multiple to true so we can have as many corp names as we want
            m_sCanBeMultiple = True
            
            Set tempXML = BuildXml(m_ksElementCorpName, domTitleElement)
            tempXML.Text = oboCorp.sTermText & " " & DateText(oboCorp.sStartDate, oboCorp.sEndDate, oboCorp.sUncertainStartDateCode, oboCorp.sUncertainEndDateCode)
            tempXML.setAttribute mk_sAuthFileIDAttr, CStr(oboCorp.lID)
        Next
        m_sCanBeMultiple = False
    End If
    
        'The person index
     If m_oboCatalogueMatrix.LevelHasCatalogueProperty(m_eLevel, cpIndexTermPersonName) Then
     
        Set clinks = t_oboParsedCatalogueEntry.CreatorNamesPersons
        
        For i = 1 To clinks.Count
            lIdInCollection = clinks.Item(i)
            oBOAFPersonLoader.IAFLoader_Load lIdInCollection
            oboPerson = oBOAFPersonLoader.Data
            
            m_sCanBeMultiple = True
        
            Set tempXML = BuildXml(m_ksElementPerson, domTitleElement)
            
            tempXML.Text = oboPerson.sPretitle & " " & oboPerson.sForenames & " " & oboPerson.sSurname & "," & _
                oboPerson.sTitle & ", " & DateText(oboPerson.sBirthDate, oboPerson.sDeathDate, oboPerson.sUncertainBirthDateCode, oboPerson.sUncertainDeathDateCode)
            tempXML.setAttribute mk_sAuthFileIDAttr, CStr(oboPerson.lID)
            
        Next
        m_sCanBeMultiple = False
        
    End If
    
        'Place index
    If m_oboCatalogueMatrix.LevelHasCatalogueProperty(m_eLevel, cpIndexTermPlaceName) Then
     
        Set clinks = t_oboParsedCatalogueEntry.PlaceIndex
        
        For i = 1 To clinks.Count
            lIdInCollection = clinks.Item(i)
            oboAFPlaceLoader.IAFLoader_Load lIdInCollection
            oboPlace = oboAFPlaceLoader.Data
        
            m_sCanBeMultiple = True
            
            Set tempXML = BuildXml(m_ksElementPlace, domTitleElement)
            tempXML.Text = oboPlace.sTermText & " " & DateText(oboPlace.sStartDate, oboPlace.sEndDate, oboPlace.sUncertainStartDateCode, oboPlace.sUncertainEndDateCode)
            tempXML.setAttribute mk_sAuthFileIDAttr, CStr(oboPlace.lID)
        Next
        m_sCanBeMultiple = False
        
    End If
    
        'Subject index
    If m_oboCatalogueMatrix.LevelHasCatalogueProperty(m_eLevel, cpIndexTermSubject) Then
     
        Set clinks = t_oboParsedCatalogueEntry.SubjectIndex
        
        For i = 1 To clinks.Count
            lIdInCollection = clinks.Item(i)
            oboAFSubjectLoader.IAFLoader_Load lIdInCollection
            oboSubject = oboAFSubjectLoader.Data
        
            m_sCanBeMultiple = True
            
            Set tempXML = BuildXml(m_ksElementSubject, domTitleElement)
            tempXML.Text = oboSubject.sTermText
            tempXML.setAttribute mk_sAuthFileIDAttr, CStr(oboSubject.lID)
        Next
        m_sCanBeMultiple = False
        
    End If
    
    '  end if for the giant IF statement deciding whether to produce everything
    End If
        Set CreateExportNodes = domTitleElement
   
       

End Function

Private Function DateText(t_sStart As Variant, t_sEnd As Variant, Optional t_sUncertainStart As Variant = "", Optional t_sUncertainEnd As Variant = "") As String
'   Function that takes two strings as start and end dates, then parses
'   them to display a sensible date range
'
'   Vers    Auth        Date            Comment
'   d3      D Johnston  15-Aug-2000     Added checking for null numeric fields and added
'                                       the dates for the index terms

    Dim sDateRange As Variant
    sDateRange = ""
    
    If t_sStart = "" And t_sEnd = "" Then
        DateText = sDateRange
        Exit Function
    End If
    
    sDateRange = t_sUncertainStart & t_sStart & "-" & t_sUncertainEnd & t_sEnd
    sDateRange = Trim(sDateRange)
    
    If sDateRange = "-" Then
        sDateRange = ""
    End If
    
    DateText = CStr(sDateRange)
    
End Function

Private Function BuildFirstSection(t_oboParsedCatEntry As ParsedCatalogueEntry) As IXMLDOMElement
'   Vers    Auth        Date            Comment
'   d1      A.Farrow    01-Aug-2000     Created
'   d4      D Johnston  18-Sep-2000     Added otherlevel code for headers

    'Dim elements
    Dim domElement As IXMLDOMElement
    Dim domTitleElement As IXMLDOMElement
    Dim domText As IXMLDOMText
    Dim domOrigination As IXMLDOMElement
    Dim tempXML As IXMLDOMElement
    Dim dscXML As IXMLDOMElement
    
    
    If m_eLevel = ECATALOGUELEVEL.ctLetterCode Then
        Set tempXML = m_domXmlDoc.createElement(m_ksArchdescTag)
    Else
        Set dscXML = m_domXmlDoc.createElement(GetParentNode(m_ksElementArchdescChildList))
        Set tempXML = BuildXml(m_ksElementCEChildList, dscXML)
    End If
    
    Dim catRef As CatReferenceParts
    Dim sRef As String
    
    
    catRef = GetEntireRefFromDBID(m_eLevel, t_oboParsedCatEntry.ID)
    sRef = GenerateExportReference(catRef, GetCatalogueLevelCode(m_eLevel))
    tempXML.setAttribute m_ksAttrNameID, sRef
    tempXML.setAttribute m_ksAttrNameLevel, m_ksAttrNameOtherLevel
    
    'Determine the level and correct other level to display
    Select Case m_eLevel
    
        Case ECATALOGUELEVEL.ctLetterCode
            tempXML.setAttribute m_ksAttrNameOtherLevel, m_ksLevelLettercode
    
        Case ECATALOGUELEVEL.ctDivision
            tempXML.setAttribute m_ksAttrNameOtherLevel, m_ksLevelDivision
            
        Case ECATALOGUELEVEL.ctClass
            tempXML.setAttribute m_ksAttrNameOtherLevel, m_ksLevelClass
        
        ' d4 - Added
        Case ECATALOGUELEVEL.ctHeaders
            tempXML.setAttribute m_ksAttrNameOtherLevel, m_ksLevelHeader
            
        Case ECATALOGUELEVEL.ctSubHeaders
            tempXML.setAttribute m_ksAttrNameOtherLevel, m_ksLevelSubHeader
            
        Case ECATALOGUELEVEL.ctPiece
            tempXML.setAttribute m_ksAttrNameOtherLevel, m_ksLevelPiece
            
        Case ECATALOGUELEVEL.ctItem
            tempXML.setAttribute m_ksAttrNameOtherLevel, m_ksLevelItem
    
    End Select
    
    tempXML.setAttribute m_ksAttrNameLanguage, GetLanguage(t_oboParsedCatEntry.LanguageID)
    tempXML.setAttribute m_ksAttrNameLegalStatus, m_ksAttrNameOtherLegalStatus
    tempXML.setAttribute m_ksAttrNameOtherLegalStatus, GetLegalStatus(t_oboParsedCatEntry.LegalStatusID)
    
  
    Set BuildFirstSection = tempXML



End Function


Private Function GetLanguage(t_lID As Long) As String
'   Description :  This creates either the 'c' or archdesc section
'
'   Vers    Auth        Date            Comment
'   d1      A.Farrow    01-Aug-2000     Created
'

    Dim sSource As String

    sSource = "BOExportparser.Getlanguage"
    
    Dim sSql As String
    Dim ofwHelper As New FWDBHelper
    Dim ofwResult As FWRecordList
    
    sSql = "SELECT " & gk_sDBCOLLanugageISOCode & " FROM " & gk_sDBTABLELanguage & " (nolock)" _
        & " WHERE " & gk_sDBCOLLanguageID & " = " & t_lID
    
    Set ofwResult = ofwHelper.FindBySql(sSql)
    
    If Not ofwResult Is Nothing Then
        If ofwResult.RecordCount > 0 Then
                GetLanguage = ofwResult.GetFieldValue(gk_sDBCOLLanugageISOCode)
        End If
    Else
        GetLanguage = ""
    End If
    
End Function

Private Function GetLegalStatus(t_lID As Long) As String
'   Description :  Get the appropriate legal status from the DB
'
'   Vers    Auth        Date            Comment
'   d1      A.Farrow    01-Aug-2000     Created
'

    Dim sSource As String

    sSource = "BOExportparser.GetLegalStatus"
    
    Dim sSql As String
    Dim ofwHelper As New FWDBHelper
    Dim ofwResult As FWRecordList
    
    sSql = "SELECT " & gk_sDBCOLLegalStatusDescription & " FROM " & gk_sDBTABLELegalStatus & " (nolock)" _
        & " WHERE " & gk_sDBCOLLegalStatusID & " = " & t_lID
    
    Set ofwResult = ofwHelper.FindBySql(sSql)
    
    If Not ofwResult Is Nothing Then
        If ofwResult.RecordCount > 0 Then
                GetLegalStatus = ofwResult.GetFieldValue(gk_sDBCOLLegalStatusDescription)
        End If
    Else
        GetLegalStatus = ""
    End If
    
End Function

Private Function BuildXml(t_sPattern As String, t_XMLElement As IXMLDOMElement) As IXMLDOMElement
'   Description :  using this function we create the XML pattern
'
'   Vers    Auth        Date            Comment
'   d1      A.Farrow    02-Aug-2000     Created

    Dim lPos As Long
    Dim lAtrrPos As Long
    Dim sPattern As String
    Dim sRemaining As String
    Dim domElement As IXMLDOMElement
    Dim domTitleElement As IXMLDOMElement
    Dim domText As IXMLDOMText
    Dim nodeList As IXMLDOMNodeList
    Dim existingNode As IXMLDOMElement

    
    'Find the next slash
    lPos = InStr(t_sPattern, "/")
    If lPos > 0 Then
        'remove the next restriction from the string
        sPattern = Trim(Left$(t_sPattern, lPos - 1))
        sRemaining = Mid$(t_sPattern, lPos + 1)
        If sPattern = "." Then
            lPos = InStr(sRemaining, "/")
            If lPos = 0 Then
                lPos = Len(sRemaining) + 1
            End If
            'remove the next restriction from the string
            sPattern = Trim(Left$(sRemaining, lPos - 1))
            sRemaining = Mid$(sRemaining, lPos + 1)
        End If
    Else
        'If there are no slashes left
        sPattern = Trim(t_sPattern)
        sRemaining = ""
    End If
    
    'We decide if the pattern contains any attribute values and strip them out
    lPos = InStr(sPattern, "@")
    If lPos > 0 Then
        Dim sAttr As String
        sAttr = sPattern
        sPattern = Trim(Left$(sPattern, lPos - 2))
        sAttr = Mid$(sAttr, lPos + 1)
        lPos = InStr(sAttr, "=")
        m_sCurrentAttrName = Trim(Left$(sAttr, lPos - 1))
        
        ' Remove the single quotes
        sAttr = Mid$(sAttr, lPos + 1)
        lPos = InStr(sAttr, "]")
        m_sCurrentAttrValue = sAttr
        m_sCurrentAttrValue = Replace(m_sCurrentAttrValue, "'", "")
        m_sCurrentAttrValue = Replace(m_sCurrentAttrValue, "]", "")
        m_sCurrentAttrValue = Trim(m_sCurrentAttrValue)
    End If
    
    ' If the pattern does not contain more than one element
    If Not sPattern = "." Then
    
        Set nodeList = t_XMLElement.selectNodes(sPattern)

        'If the pattern does'nt exist already create it
        If nodeList.Length = 0 Or (m_sCanBeMultiple = True And sRemaining = "") Or m_bCheckingAttr Then
            Set domElement = m_domXmlDoc.createElement(sPattern)
            t_XMLElement.appendChild domElement
        Else
            Set domElement = t_XMLElement.selectSingleNode(sPattern)
        End If
        
        'If this is the final section of the pattern then we return this node
        If sRemaining = "" Then
            Set BuildXml = domElement
        Else
            'Otherwise we continue building the pattern
            Set BuildXml = BuildXml(sRemaining, domElement)
        End If
    End If


End Function

Private Function GetAccessCondition(t_lID As Long) As String
'   Description :  Get the appropriate legal status from the DB
'
'   Vers    Auth        Date            Comment
'   d1      A.Farrow    01-Aug-2000     Created
'

    Dim sSource As String

    sSource = "BOExportparser.GetAccessCondition"
    
    Dim sSql As String
    Dim ofwHelper As New FWDBHelper
    Dim ofwResult As FWRecordList
    
    sSql = "SELECT " & gk_sDBCOLAccessConditionDescription & " FROM " & gk_sDBTABLEAccessCondition & " (nolock)" _
        & " WHERE " & gk_sDBCOLAccessConditionID & " = " & t_lID
    
    Set ofwResult = ofwHelper.FindBySql(sSql)
    
    If Not ofwResult Is Nothing Then
        If ofwResult.RecordCount > 0 Then
                GetAccessCondition = ofwResult.GetFieldValue(gk_sDBCOLAccessConditionDescription)
        End If
    Else
        GetAccessCondition = ""
    End If
    
End Function

Private Function GetParentNode(t_sString As String, Optional ByRef sRemaining) As String
'   Description :  Some middle nodes within the pattern actually require attributes to be set
'
'   Vers    Auth        Date            Comment
'   d1      A.Farrow    01-Aug-2000     Created
'
 
    Dim sPattern As String
    'Dim sRemaining As String
    Dim lPos As Long
 
    lPos = InStr(t_sString, "/")
    If lPos > 0 Then
        'remove the next restriction from the string
        sPattern = Trim(Left$(t_sString, lPos - 1))
        sRemaining = Mid$(t_sString, lPos + 1)
            If sPattern = "." Then
                lPos = InStr(sRemaining, "/")
                If lPos > 0 Then
                'remove the next restriction from the string
                    sPattern = Trim(Left$(sRemaining, lPos - 1))
                    sRemaining = Mid$(sRemaining, lPos + 1)
                End If
            End If
    End If
    
    GetParentNode = sPattern
    
End Function
Private Function GenerateExportReference(t_cCatRef As CatReferenceParts, t_iLevelID As Integer) As String
'   Description :  Generates the export reference.
'
'   Vers    Auth        Date            Comment
'   d1      A.Farrow    07-Aug-2000     Created
'

Dim sRef As String

    
    If Not t_cCatRef.sLettercodeRef = "" Then
        sRef = sRef & t_cCatRef.sLettercodeRef
    End If
    If Not t_cCatRef.sDivisionRef = "" Then
        sRef = sRef & m_ksIDDelimeter & m_ksIDDivision & t_cCatRef.sDivisionRef
    End If
    If Not t_cCatRef.sClassRef = "" Then
        sRef = sRef & m_ksIDDelimeter & m_ksIDClass & t_cCatRef.sClassRef
    End If
    If Not t_cCatRef.sSubclassRef = "" Then
        sRef = sRef & m_ksIDSubclass & t_cCatRef.sSubclassRef
    End If
    If Not t_cCatRef.sHeaderRef = "" Then
        sRef = sRef & m_ksIDDelimeter & m_ksIDHeader & t_cCatRef.sHeaderRef
    End If
    If Not t_cCatRef.sSubheaderRef = "" Then
        sRef = sRef & m_ksIDDelimeter & m_ksIDSubheader & t_cCatRef.sSubheaderRef
    End If
    If Not t_cCatRef.sPieceRef = "" Then
        sRef = sRef & m_ksIDDelimeter & m_ksIDPiece & t_cCatRef.sPieceRef
    End If
    If t_iLevelID = 7 Then
        sRef = sRef & m_ksIDDelimeter & m_ksIDItem & t_cCatRef.sItemRef
    End If

    
 
    
    GenerateExportReference = sRef

   

End Function


Private Function GetClosureInfo(t_lID As Long, t_eLevel As ECATALOGUELEVEL) As String
'   Description :  For pieces and items we get the closure information
'
'   Vers    Auth        Date            Comment
'   d1      A.Farrow    07-Aug-2000     Created
'   d6      TomH        25-Feb-2002     Have altered so check to see if
'                                       closure code is NULL before exporting.

    Dim sSource As String

    sSource = "BOExportparser.GetClosureInfo"
    
    Dim sSql As String
    Dim ofwHelper As New FWDBHelper
    Dim ofwResult As FWRecordList
    Dim sTable As String
    Dim sIDCol As String
    Dim sInfo As String
    Dim sClosureCode As String
    
    
    'This function should only be called by for PIECE and ITEM, here we decide what called this function
    
    If m_eLevel = ctPiece Then
        sTable = gk_sDBTABLEPiece
        sIDCol = gk_sDBCOLPieceID
    Else
        sTable = gk_sDBTABLEItem
        sIDCol = gk_sDBCOLItemID
    End If
    
    sSql = "SELECT " & gk_sDBCOLClosureCode & "," & gk_sDBCOLClosureType & " FROM " & sTable & " (nolock)" _
        & " WHERE " & sIDCol & " = " & t_lID
    
    Set ofwResult = ofwHelper.FindBySql(sSql)
    
    If Not ofwResult Is Nothing Then
        If ofwResult.RecordCount > 0 Then
                sInfo = ofwResult.GetFieldValue(gk_sDBCOLClosureType)
                
                ' d6 have put in check to see if NULL
                If Not IsNull(ofwResult.GetFieldValue(gk_sDBCOLClosureCode)) Then
                    sClosureCode = ofwResult.GetFieldValue(gk_sDBCOLClosureCode)
                End If
        End If
    End If
    
    'We use the ClosureType from tbl_piece/tbl_item to get the full closure type text
    sSql = "SELECT " & gk_sDBCOLClosureTypeDescription & " FROM " & gk_sDBTABLEClosureType & " (nolock)" _
        & " WHERE " & gk_sDBCOLClosureType & " = " & QuoteValueWithQuotes(sInfo)

    Set ofwResult = ofwHelper.FindBySql(sSql)
    
    If Not ofwResult Is Nothing Then
        If ofwResult.RecordCount > 0 Then
                sInfo = ofwResult.GetFieldValue(gk_sDBCOLClosureTypeDescription)
        End If
    End If
    
    GetClosureInfo = sInfo & " " & sClosureCode
End Function

Private Function GetPhysicalForm(t_lID As Long)
'   Description :  Get the physical form of the catalogue
'
'   Vers    Auth        Date            Comment
'   d1      A.Farrow    07-Aug-2000     Created
'
Dim sSource As String

    sSource = "BOExportparser.GetPhysicalForm"
    
    Dim sSql As String
    Dim ofwHelper As New FWDBHelper
    Dim ofwResult As FWRecordList
    
    sSql = "SELECT " & gk_sDBCOLPhysicalDescriptionFormDescription & " FROM " & gk_sDBTABLEPhysicalForm & " (nolock)" _
        & " WHERE " & gk_sDBCOLPhysicalDescriptionFormID & " = " & t_lID
    
    Set ofwResult = ofwHelper.FindBySql(sSql)
    
    If Not ofwResult Is Nothing Then
        If ofwResult.RecordCount > 0 Then
                GetPhysicalForm = ofwResult.GetFieldValue(gk_sDBCOLPhysicalDescriptionFormDescription)
        End If
    Else
        GetPhysicalForm = ""
    End If
End Function

Private Sub StripUnwantedTags(t_xmlParentNode As IXMLDOMElement, t_sXMLPattern As String, t_sXmlString As String)
'   Description :  Takes in an XMLDOMElement and searches down through the children
'   to see find the lowest occurence of the pattern.  The lowest level found will then
'   be promoted up to the top level
'
'   Vers    Auth        Date            Comment
'   d1      A.Farrow    07-Aug-2000     Created
'
    Dim xmlNodeList As IXMLDOMNodeList, xmlChildrenList As IXMLDOMNodeList
    Dim xmlNode As IXMLDOMNode, xmlParentNode As IXMLDOMNode, xmlChild As IXMLDOMNode
    Dim xmlCloneNode As IXMLDOMNode
    Dim xmlLowestNode As DOMDocument
    
    Set xmlLowestNode = New DOMDocument
    
    xmlLowestNode.loadXML (t_sXmlString)
    
    ' Get any children with the correct pattern
    Set xmlNodeList = xmlLowestNode.selectNodes(t_sXMLPattern)
    
    If xmlNodeList.Length = 0 Then
        If xmlLowestNode.hasChildNodes = True Then
            For Each xmlNode In xmlLowestNode.childNodes
                t_xmlParentNode.appendChild xmlNode
            Next
        Else
            t_xmlParentNode.Text = xmlLowestNode.Text
        End If
        Exit Sub
    End If
    
    While xmlNodeList.Length > 0
        For Each xmlNode In xmlNodeList
            Set xmlChildrenList = xmlNode.childNodes
            If xmlChildrenList.Length > 0 Then
                For Each xmlChild In xmlChildrenList
                    'Set xmlParentNode = xmlNode.parentNode
                    t_xmlParentNode.appendChild xmlChild.cloneNode(True)
                    xmlNode.removeChild xmlChild
                Next
            End If
            xmlLowestNode.removeChild xmlNode
        Next
        
        Set xmlNodeList = xmlLowestNode.selectNodes(t_sXMLPattern)
    Wend
    
End Sub

Private Function StripDidAndAttrsFromPattern(t_sXMLPattern As String)
'   Description :  Takes in an XSL pattern and removes DID from its path and
'   also removes any attributes.  It should be used on patterns passed to
'   StripUnwantedTags to make sure that all unwanted duplicate tags are removed.
'
'   Vers    Auth        Date            Comment
'   d8      F Devereux  01-Nov-2002     Created
'
    Dim sResult As String

    ' strip DID
    sResult = Replace(t_sXMLPattern, ksElementDID, ksElementDIDReplace)

    ' strip attributes
    Dim lPos As Long
    lPos = InStr(sResult, "@")
    If lPos > 0 Then
        sResult = Trim(Left$(sResult, lPos - 2))
    End If

    StripDidAndAttrsFromPattern = sResult
    
End Function

Public Function DocType() As String
' Returns the DOCTYPE element from the Parser
    DocType = m_ksExportDocType & vbNewLine
End Function

Public Function XMLVersionElement() As IXMLDOMProcessingInstruction
    Set XMLVersionElement = m_domXmlDoc.createProcessingInstruction("xml", "version=""" & m_ksExportXMLVersion & """")
End Function

Public Function DSCElement() As IXMLDOMElement
    Set DSCElement = m_domXmlDoc.createElement(m_ksDSCTagName)
    DSCElement.setAttribute m_ksDSCTagAttribute, m_ksDSCTagValue
End Function

Public Function EADHeaderElement() As IXMLDOMElement
On Error GoTo ErrorHandler
    Dim sSource As String
    sSource = "EADHeaderElement"
    
    If m_domXmlDoc.loadXML(m_ksExportEADHeaderXML) Then
        Set EADHeaderElement = m_domXmlDoc.childNodes.Item(0)
    Else
        RaiseError sSource, PEGeneralInternalError, _
            "Invalid EAD header in table " & gk_sDBTABLEEADData
    End If
    
    Exit Function
    
ErrorHandler:
    ConvertError sSource
End Function

Public Function FrontMatterElement() As IXMLDOMElement
On Error GoTo ErrorHandler
    Dim sSource As String
    sSource = "FrontMatterElement"
    
    If m_domXmlDoc.loadXML(m_ksExportFrontMatterXML) Then
        Set FrontMatterElement = m_domXmlDoc.childNodes.Item(0)
    Else
        RaiseError sSource, PEGeneralInternalError, _
            "Invalid EAD for FRONTMATTER in table " & gk_sDBTABLEEADData
    End If
    
    Exit Function
    
ErrorHandler:
    ConvertError sSource
End Function
