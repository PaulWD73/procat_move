VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "VWCatalogueEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    VWCatalogueEntry.cls
' System:       PROCAT, PRO
' Copyright:    (C) Quidnunc Limited
'
' Description:  View,Create,Edit catalogue entries
'
' Amendment history:
'   d1  ndk     25Aug1999       Created (Some functions were moved here from VWEditSet and VWCatalogue)
'   d2  ndk     16Oct1999       Added catalogue locking functionality, added closure type and closure code
'   d3  ndk     28Oct1999       Changed catalogue locking to lock only the tab being edited to be locked
'                               Summary tab permits modification or piece key order and item_key_order
'                               Create by copying now permits copying only at the same level (except pieece and item)
'                               Move catalogue entry functionality coded
'                               Added sub CancelAddLinkTerm
'                               Integrated with new catalogue browser
'                               Locking done at tab level
'                               Changed text fields to text areas for text columns
'   d4  ndk     04Nov1999       display and saving of access conditions handled differently, resources for item and piece
'                               key order were put
'   d5  ndk     11Nov1999       Create by copying drop down for levels changed, edit button disabled for users with no rights
'                               Changed the way deleteeditsetentry method of bodeletecatalogue was called
'   d6  ndk     15Nov1999       Displaying scope and content at the top for piece and item
'                               Integrated sfa browser
'                               Security check to enable move button
'   d7  ndk     17Nov1999       Slight change in way security is handled in dislplaying tabs
'   d8  ndk     18Nov1999       Change in the was error messages are constructed
'   d9  ndk     20Nov1999       For lookups using lookup list of bolookup instead of fwreference data and other minor changes
'   d10 ndk     22Nov1999       When displaying browser for place of deposit , appropriate filter is set
'   d11 ndk     22Nov1999       Changed the name of the project for catalogue browser
'                               Warning before deletion made different depending on types of entry
'   d13 ndk     24Nov1999       Access condition is now updated when a editset entry is created from scratch or lettercode
'                               using default value of parents value
'   d14 ndk     25Nov1999       Minor changes in DisplayEditSummaryDetailsScreen
'   d15 ndk     07Dec1999       Implemented EAD Editing and made related changes
'                               Added ReDisplayEditCatalogueDetailsWithUpdatedEAD
'   d16 ndk     08Dec1999       The catalogue browser is now opened at appropriate level when
'                               selecting parent while moving and createing by copying
'                               Catalogue browser is displayed using response.redirect.
'                               Added help buttons
'                               Removed methods that called cataloguebrowser as a method call
'   d17 ndk     09Dec1999       Added option for choosing the selection type for adding related and separate materials
'                               Added DisplayMenuForRelatedSeparatedMaterial
'   d18 ndk     09Dec1999       Changed images for pick term buttons and xml launcher buttons
'                               Added private methods GenerateEditOnclickCode, GenerateUpdateOnclickCode to handle xml editor launching and updation
'                               Put deletion messages into resource file
'   d19 ndk     10Dec1999       Code added to add a read only entry into editorial as added entry
'                               Added AddReadOnlyEntryIntoEditSet
'   d20 wdp     10Dec1999       Added checkboxes for record status
'   d21 ndk     15Dec1999       Replacing single and double quotes with equivalent html codes before rendering ead data
'                               Replacing html codes with single and double quotes before saving/storing ead data
'   d22 ndk     16Dec1999       Changed the display length for reference field when creating catalogues
'                               Changed edit format for record opening date
'   d23 ndk     21Dec1999       Added security check where required in the methods
'   d24 ndk     27Dec1999       Changed to VI code use customizable styles
'   d25 ndk     27Dec1999       Added tooltips
'   d26 ndk     29Dec1999       Rechecking releasing of variables, especially if an error is deliberately raised
'   d27 ndk     29Dec1999       Using resource file for audit trail related display constants
'   d28 ndk     30Dec1999       Rectfied error - not releasing the collection
'                               Persisting catalogue level while creating
'   d29 crm     02Feb2000       Updated code to provide links for Leaflets to catalogues
'   d30 dtm     02feb2000       corrected typo of 'Catalogye' to 'Catalogue'
'   d31 crm     08Feb2000       Added QueryString parameter for PlaceOfDeposit
'                               pickterm(CorpBody)
'   d32 dtm     09Feb2000       Fixed defect in DisplayHeaderTab : IF A AND B OR C
'   d33 dtm     09Feb2000       Fixed defect in DisplayMenuForRelatedSeparatedMaterial : gk_sIMAGESmallGo already includes full path
'   d34 jcc     10Feb2000       removed covering dates from catalogue entry heading
'   d35 dtm     10Feb2000       Modified DisplayEditSummaryDetailsScreen so that a * is displayed when title is empty to indicate mandatoriness
'   d36 jcc     18Feb2000       and remove the -from- bit too! Also, should be a * on title all the time, above piece
'   d37 daj     21Feb2000       Changed the Add Leaflet button to point to the AddLeaflet button instead of AddTerm
'   d38 tmj     23Feb2000       Modified so could copy departments
'   d39 tmj     25Feb2000       Added additional parameter to be passed in querystring when finding the parent of a copied
'                               object
'   d40 tmj     29Feb2000       Modified so EditSetID is passed through to certain pages
'   d41 spb     29Feb2000       Changed ObjectControl_CanBePooled = True rather than false
'   d42 tmj     02Mar2000       fixed links when going through Audit Trail
'   d43 tmj     04Mar2000       changed cosmetic output of DisplayOneLineEditAndDisplayForContent
'   d44 wdp     06Mar2000       Fixed DisplayEditSummaryDetailsScreen to cope with
'                               title OR scope-content for create new entry and
'                               be able to switch between them.
'   d45 tmj     06Mar2000       Raised an error when user selects an item as a parent of a new entry
'   d46 tmj     21Mar2000       Added error checking to StoreSelectedParentForCreateByCopy
'   d47 tmj     24Mar2000       Passed Editset parameter to moving checking
'   d48 daj     27Mar2000       Tidied up the audit trail and fixed defect 766 - Audit trail 'back' links not working
'                               once you move off the summary tab
'   d49 tmj     30Mar2000       Fixed so displays EAD fields. This is achieved by Encoding the EAD fields as HTML when they
'                               are hidden so encoding the special characters twice
'   d50 daj     05Apr2000       Changed GenerateUpdateOnclickCode to call VB rather than the control directly
'                               (see FWHTML.EADEditorLauncher)
'   d51 daj     05Apr2000       Added Cancel buttons to the Related/Separated Material link forms
'   d52 daj     11Apr2000       Added extra situations where the summary edit screen should toggle the DTDtype
'   d53 daj     23May2000       Added checking of closure code to prevent invalid use of null errors
'   d54 daj     20Jun2000       Removed formatting of 'View this version' link
'   d55 map     07Nov2000       Closure code now set to Null if form returns an empty stringStringToNullOrString
'   d56 daj     04Jan2001       Added Covering Dates to the header details, and made Covering Dates non mandatory when editing
'                               headers and subheaders
'   d57 bcj     19Jan2001       Created StartAddFromExisting for enhancement 2 (enhancement set 1)
'   d58 bcj     14Jan2001       AddCataloguesToEditSet
'                               Changed lCatalogueID(1) > 0 with lCatalogueID(1) <> 0
'   d59 ???     ?????????       ?
'   d60 alm     13Feb2001       Allows the entry and editing of start / end dates
'   d61 daj     13Feb2001       Fixed bugs with enhancement done in d56
'   d62 bcj     13Feb2001       Added LinkToBrowser, as it cannot be accessed in VWSharedCatalogueEntry
'   d63 bcj     14Feb2001       Added CASE to LinkToBrowser for searching for related material (NOTE: Not done in reader)
'   d64 pnp     16Feb2001       Added functionality to get last reference
'   d67 gdb     20Feb2001       Added more arguments when invoking inputFiled() for closure code in DisplayEditAccessDetailScreen
'   d68 alm     27Feb2001       Modified to work with the new browser
'   d69 map     05Mar2001       Corrected functionality for getting last reference
'   d70 map     15May2001       Amended DisplayEditSummaryDetailsScreen
'   d71 map     11Jul2001       Added regnal years check box to bypass check against covering_dates_expressions
'   d72 alm     25Sep2001       Moved the 'Notes' field from the Content view to the Summary View
'   d73 alm     19Oct2001       Added function MoveMultiplePiecesToNewParent
'   d74 alm     25Oct2001       Check that a parent of a new / copied / moved entry does not exist in a
'                               delete editset and modified functions to support a delete edit set.
'                               Also added DisplayConfirmationForCatalogueUnDeletion
'   d75 fjd     05Feb2003       Renamed LCI flag to AR
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'Interfaces
Implements ObjectControl

Public Enum ECatalogueEntryEditType
'How is the catalogue entry being created
    ecEditCatalogueEntry = 0        'Keep this as 0
    ecCreateCatalogueFromScatch
    ecCreateCatalogueByCopying
    ecCreateLetterCodeFromScratch
End Enum

Public Enum ECatalogueEditFormLoadType
    'How the edit set entry details (edit) should be loaded
    etLoadFromDatabaseEditCatalogueEntry = 1
    etLoadFromSavedStateAndAddTermEditCatalogueEntry
    etLoadFromSavedStateAndRemoveTermEditCatalogueEntry
    etLoadFromSavedStateNoChangeEditCatalogueEntry    'Used when user clicks 'cancel' when he has to choose a link item from the list shown to him
    etLoadFromNoneCreateCatalogueEntry      'Used when creating a new catalogue entry
End Enum

'Private framework objects
Private m_oFWHTML As FWHTML
Private m_oFWAppData As FWApplicationData
Private m_ofwRefData  As FWReferenceData
Private m_oFWTable As FWTable

'Private asp/mts objects
Private m_objObjectContext As ObjectContext
Private m_aspResponse  As Response
Private m_aspRequest As Request

'Private framework styles used across all methods
Private m_typCellStyleLeftHeading As CellStyle, m_typCellStyleRightDetails As CellStyle
Private m_typCellStyleForOKCancel As CellStyle
Private m_typTableStyleContents As TableStyle
Private m_typFontStyleLeftHeading As FontStyle, m_typFontStyleRightDetails As FontStyle

'Private business objects
Private m_oBOEditorialPermissions  As BOEditorialPermissions
Private m_oBOCatalogueElementMatrix As BOCatalogueElementMatrix
Private m_oBOLookUp As BOLookup
Private m_oVWSharedCatalogueEntry As VWSharedCatalogueEntry

'Private Variables
Private m_lLoggedUserID As Long                     'ID of the logged user
Private m_sHtml As String
Private m_eBrowserUsage As ECATALOGUEBROWSERUSAGE   ' browser usage d58
'Private Constants

'For Add Catalogue Entry
Private Const ksNameCatalogueSearch As String = "FCS"
Private Const ksNameCheckBoxCount As String = "CBN"
Private Const ksNameCheckBoxes As String = "CHKBX"

'Constants for Parameters for values passes to and from asp
Private Const ksParamCatalogueID As String = "CID"
Private Const ksParamCatalogueToBeCopiedID As String = "CCPDID"
Private Const ksParamCatalogueToBeCopiedLevel As String = "CCPDLVL"
Private Const ksParamCatalogueLevel As String = "CL"
Private Const ksParamSearchDirection As String = "SD"
Private Const ksParamReference As String = "REF"
Private Const ksParamCatalogueTitle As String = "CT"
Private Const ksParamSortOrder As String = "SOD"
Private Const ksNextPageCode As String = "NPC"
Private Const ksNameOpenImage As String = "OCAT"
Private Const ksNameButtonMore As String = "btnMore"

Private Const ksParamParentCatalogueID As String = "PCID"
Private Const ksParamParentCatalogueLevel As String = "PCL"
Private Const ksNameCatalogueLevel As String = "CATLEVEL"

'field lengths and sizes
Private Const kiColumnsForLetterCodeReference As Integer = 6
Private Const kiColumnsForOtherCatalogueReference As Integer = 50
Private Const kiSizeOtherCatalogueReferences As Integer = 40
Private Const kiColumnsForTextFields As Integer = 80
Private Const kiRowsForTextFields As Integer = 4
Private Const kiSizeLinkItemDescription As Integer = 50

Private Const ksParamCatalogueEntryEditType As String = "CatalogueEntryEditTypeType"
Private Const ksParamLoadType As String = "LoadType"
Private Const ksParamLinkTermID As String = "LinkTermID"
Private Const ksParamLinkTermType As String = "LinkTermType"
Private Const ksParamLinkTermLevelID As String = "LinkTermLevelID"
Private Const ksParamMainCollectionKey As String = "MainCollectionKey"
Private Const ksParamEditingCancelled As String = "EditingCancelled"


'Summary Details
Private Const ksNameReference As String = "REFERENCE"
Private Const ksNameTitle As String = "TITLE"
Private Const ksNameLegalStatusID As String = "LEGALSTATUS"
Private Const ksNameLanguageID As String = "LANGUAGEID"
Private Const ksNameFormerReferenceDepartment As String = "FORMERREFERENCEDEPT"
Private Const ksNameFormerReferencePRO As String = "FORMERREFERENCEPRO"
Private Const ksNameMapDesignation As String = "MAPDESIGNATION"
Private Const ksNameCoveringDate As String = "COVERINGDATE"
Private Const ksNameStartDate As String = "STARTDATE"
Private Const ksNameEndDate As String = "ENDDATE"
Private Const ksNameMapScaleNumber As String = "MAPSCALENUMBER"
Private Const ksNamePhysicalDescriptionFormCode As String = "PHYSICALDESCRIPTIONFORMCODE"
Private Const ksNamePhysicalDescriptionRecordQty As String = "PHYSICALDESCRIPTIONQTY"
Private Const ksNamePhysicalDescriptionDimensions As String = "PHYSICALDESCRIPTIONDIMENSIONS"
Private Const ksNamePhysicalCondition As String = "PHYSICALCONDITION"
Private Const ksNamePieceKeyOrder As String = "PIECEKEYORDER"
Private Const ksNameItemKeyOrder As String = "ITEMKEYORDER"
Private Const ksNameDTDType As String = "DTDType"
Private Const ksNameIsRegnalDate As String = "REGNALDATE"

'Access Conditions
Private Const ksNameAccessConditionID As String = "ACCESSCONDITIONID"
Private Const ksNameClosureType As String = "CLOSURETYPE"
Private Const ksNameClosureCode As String = "CLOSURECODE"
Private Const ksNameRecordOpeningDate As String = "RECORDOPENINGDATE"
Private Const ksNameClosureStatus As String = "CLOSURESTATUS"
Private Const ksNameRestrictionsOnUse As String = "RESTRICTIONSONUSE"
Private Const ksNameCustodialHistory As String = "CUSTODIALHISTORY"
Private Const ksNameAccumulationDates As String = "ACCUMULATIONDATE"
Private Const ksNameAppDestructionInformation As String = "APPDESTRUCTIONINFO"
Private Const ksNameAccruals As String = "ACCRUALS"

'Admin History
Private Const ksNameAdminHistory As String = "ADMINHISTORY"

'Content
Private Const ksNameScopeAndContent As String = "SCOPEANDCONTENT"
Private Const ksNameArrangement As String = "ARRANGEMENT"
Private Const ksNameNote As String = "NOTE"

'Ids for link terms
Private Const ksNameIDImmediateSourceOfAcquisitionPerson As String = "IDIMMSOURCEPERSON"
Private Const ksNameIDImmediateSourceOfAcquisitionCorporateBody As String = "IDIMMSOURCECORPBODY"
Private Const ksNameIDCopiesInformation As String = "IDCOPIESINFO"
Private Const ksNameIDLocationOfOriginals As String = "IDLOCATIONOFORIGINALS"
Private Const ksNameIDCreatorNamesPerson As String = "IDCREATORNAMESPERSON"
Private Const ksNameIDCreatorNamesCorporateBody As String = "IDCREATORNAMESCB"
Private Const ksNameIDPlaceOfDeposit As String = "IDPLACEOFDEPOSIT"

Private Const ksNameIDPublicationNote As String = "IDPUBLICATIONNOTE"
Private Const ksNameIDUnpublishedFindingAids As String = "IDUNPUBLISHEDFINDINGAIDS"
Private Const ksNameIDRelatedMaterial As String = "IDRELATEDMATERIAL"
Private Const ksNameIDSeparateMaterial As String = "IDSEPARATEMATERIAL"

Private Const ksNameIDIndexTermCorporateBody As String = "IDINDEXTERMCORPORATEBODY"
Private Const ksNameIDIndexTermPerson As String = "IDINDEXTERMPERSON"
Private Const ksNameIDIndexTermPlace As String = "IDINDEXTERMPLACE"
Private Const ksNameIDIndexTermSubject As String = "IDINDEXTERMSUBJECT"

'IDs for levels
Private Const ksNameLevelRelatedMaterial As String = "LEVELRELATEDMATERIAL"
Private Const ksNameLevelSeparateMaterial As String = "LEVELSEPARATEMATERIAL"

'Descriptions for link terms
Private Const ksNameDescriptionImmediateSourceOfAcquisitionPerson As String = "DESCIMMSOURCEPERSON"
Private Const ksNameDescriptionImmediateSourceOfAcquisitionCorporateBody As String = "DESCIMMSOURCECORPBODY"
Private Const ksNameDescriptionCopiesInformation As String = "DESCCOPIESINFO"
Private Const ksNameDescriptionLocationOfOriginals As String = "DESCLOCATIONOFORIGINALS"
Private Const ksNameDescriptionPublicationNote As String = "DESCPUBLICATIONNOTE"
Private Const ksNameDescriptionUnpublishedFindingAids As String = "DESCUNPUBLISHEDFINDINGAIDS"
Private Const ksNameDescriptionRelatedMaterial As String = "DESCRELATEDMATERIAL"
Private Const ksNameDescriptionSeparateMaterial As String = "DESCSEPARATEMATERIAL"

'Unknown descriptions (unlinked) for link terms
Private Const ksNameUnknownImmediateSourceOfAcquisition As String = "UNKNOWNIMMSOURCE"
Private Const ksNameUnknownCopiesInformation As String = "UNKNOWNCOPIESINFO"
Private Const ksNameUnknownPublicationNote As String = "UNKNOWNPUBLICATIONNOTE"
Private Const ksNameUnknownUnPublishedFindingAids As String = "UNKNOWNUNPUBLISHEDFINDINGAIDS"
Private Const ksNameUnknownRelatedMaterial As String = "UNKNOWNRELATEDMATERIAL"
Private Const ksNameUnknownSeparateMaterial As String = "UNKNOWNSEPARATEMATERIAL"

'Button Name Constants
Private Const ksNameButtonOk As String = "cmdOK"
Private Const ksNameButtonCancel As String = "cmdCancel"

'Hidden field for ead editing
Private Const ksNameHiddenEADUpdatedScopeAndContent As String = "HidddenEADUpdatedScopeAndContent"
Private Const ksNameHiddenEADUpdatedArrangement As String = "HidddenEADUpdatedArrangement"
Private Const ksNameHiddenEADUpdatedAdminHistory As String = "HidddenEADUpdatedAdminHistory"
Private Const ksNameHiddenEADUpdatedTitle As String = "HidddenEADUpdatedTitle"
Private Const ksNameHiddenEADUpdatedMapDesignation As String = "HidddenEADUpdatedDesignation"

'Add Link Terms
Private Const ksNameButtonAddTermImmSrcPerson As String = "cmdAddTermImmSrcPerson"
Private Const ksNameButtonAddTermImmSrcCorpBody As String = "cmdAddTermImmSrcCorpBody"
Private Const ksNameButtonAddTermCopiesInformation As String = "cmdAddTermCopiesInformation"
Private Const ksNameButtonAddTermLocationOfOriginals As String = "cmdAddTermLocationOfOriginals"
Private Const ksNameButtonAddTermCreatorNamesPerson As String = "cmdAddTermCNPerson"
Private Const ksNameButtonAddTermCreatorNamesCorpBody As String = "cmdAddTermCNCorpBody"
Private Const ksNameButtonAddTermPlaceOfDeposit As String = "cmdAddTermPlaceOfDeposit"

Private Const ksNameButtonAddTermPublicationNote As String = "cmdAddTermPublicationNote"
Private Const ksNameButtonAddTermUnPublishedFindingAids = "cmdAddTermUnPublishedFindingAids"
Private Const ksNameButtonAddTermRelatedMaterial = "cmdAddTermRelatedMaterial"
Private Const ksNameButtonAddTermSeparateMaterial = "cmdAddTermSeparateMaterial"

Private Const ksNameButtonAddTermIndexTermCorporateBody As String = "cmdAddTermIndexTermCorporateBody"
Private Const ksNameButtonAddTermIndexTermPerson As String = "cmdAddTermIndexTermPerson"
Private Const ksNameButtonAddTermIndexTermPlace As String = "cmdAddTermIndexTermPlace"
Private Const ksNameButtonAddTermIndexTermSubject As String = "cmdAddTermIndexTermSubject"

'Remove Link Terms
Private Const ksNameButtonRemoveTermImmSrcPerson  As String = "cmdRemoveTermImmSrcPerson"
Private Const ksNameButtonRemoveTermImmSrcCorpBody  As String = "cmdRemoveTermImmSrcCorpBody"
Private Const ksNameButtonRemoveTermCopiesInformation  As String = "cmdRemoveTermCopiesInformation"
Private Const ksNameButtonRemoveTermLocationOfOriginals  As String = "cmdRemoveTermLocationOfOriginals"
Private Const ksNameButtonRemoveTermCreatorNamesPerson  As String = "cmdRemoveTermCNPerson"
Private Const ksNameButtonRemoveTermCreatorNamesCorpBody  As String = "cmdRemoveTermCNCorpBody"
Private Const ksNameButtonRemoveTermPlaceOfDeposit  As String = "cmdRemoveTermPlaceOfDeposit"

Private Const ksNameButtonRemoveTermPublicationNote  As String = "cmdRemoveTermPublicationNote"
Private Const ksNameButtonRemoveTermUnpublishedFindingAids As String = "cmdRemoveTermUnpublishedFindingAids"
Private Const ksNameButtonRemoveTermRelatedMaterial As String = "cmdRemoveTermRelatedMaterial"
Private Const ksNameButtonRemoveTermSeparateMaterial As String = "cmdRemoveTermSeparateMaterial"

Private Const ksNameButtonRemoveTermIndexTermCorporateBody As String = "cmdRemoveTermIndexTermCorporateBody"
Private Const ksNameButtonRemoveTermIndexTermPerson As String = "cmdRemoveTermIndexTermPerson"
Private Const ksNameButtonRemoveTermIndexTermPlace As String = "cmdRemoveTermIndexTermPlace"
Private Const ksNameButtonRemoveTermIndexTermSubject As String = "cmdRemoveTermIndexTermSubject"

Private Const mk_sValue As String = "checked"

Private Const ksNameButtonAddTermPNoteLeaflet As String = "cmdAddTermPNoteLeaflet"
Private Const ksNameButtonAddTermUnPFindingAidLeaflet As String = "cmdAddTermUnPFindingAidLeaflet"
Private Const ksNameButtonRemoveTermPNoteLeaflet As String = "cmdRemoveTermPNoteLeaflet"
Private Const ksNameButtonRemoveTermUnPFindingAidLeaflet As String = "cmdRemoveTermUnPFindingAidLeaflet"

'   d29 crm     02Feb2000
'ID for Leaflets
Private Const ksNameIDPublicationNoteLeaflet As String = "IDPUBLICATIONNOTELEAFLET"
Private Const ksNameIDUnpublishedFindingAidsLeaflet As String = "IDUNPUBLISHEDFINDINGAIDSLEAFLET"

'Desc for Leaflets
Private Const ksNameDescPublicationNoteLeaflet As String = "DESCPUBLICATIONNOTELEAFLET"
Private Const ksNameDescUnpublishedFindingAidsLeaflet As String = "DESCUNPUBLISHEDFINDINGAIDSLEAFLET"

'TODO
Private Const gk_sDisablePlaceOfDeposit As String = "DisablePlaceOfDeposit"

Public Property Get Name()
    Name = "VWCatalogueEntry"
End Property

Private Sub ObjectControl_Activate()
    Initialise
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = True
End Function

Private Sub ObjectControl_Deactivate()
    Terminate
End Sub
Private Sub Initialise()
'   Initialises the objects
'
'   d1  ndk     25Aug1999       Created
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String
    
    sSource = Me.Name & ".Initialise"
    
    'Get the object context
    Set m_objObjectContext = GetObjectContext()
    
    If m_objObjectContext Is Nothing Then
        RaiseError sSource, peOCUnableToGetObjectContext, "Objectcontext is nothing"
    End If
    
    'Initialise the asp objects
    Set m_aspResponse = m_objObjectContext.Item(gk_sResponseObject)
    Set m_aspRequest = m_objObjectContext.Item(gk_sRequestObject)
    
    'Initialise the framework obects
    Set m_oFWHTML = New FWHTML
    Set m_oFWAppData = New FWApplicationData
    Set m_ofwRefData = New FWReferenceData
    Set m_oFWTable = New FWTable
    
    'Initialise the business objects
    Set m_oBOEditorialPermissions = New BOEditorialPermissions
    Set m_oBOCatalogueElementMatrix = New BOCatalogueElementMatrix
    Set m_oBOLookUp = New BOLookup
    Set m_oVWSharedCatalogueEntry = New VWSharedCatalogueEntry
    
    'Get the users details
    m_lLoggedUserID = GetDataFromSession(dsUserID, m_objObjectContext.Item(gk_sSessionObject))
    
    'Initialise the table, row, font styles etc
    InitialiseStyles
    
    Exit Sub
    
ErrorHandler:
    ConvertError sSource
End Sub

Private Sub Terminate()
'   Releases the objects
'
'   d1  ndk     25Aug1999       Created

    On Error GoTo ErrorHandler
    
    Dim sSource As String
    
    sSource = Me.Name & ".Terminate"
    
    'Release the asp objects
    Set m_objObjectContext = Nothing
    Set m_aspResponse = Nothing
    Set m_aspRequest = Nothing
    
    'Release the framework objects
    Set m_oFWHTML = Nothing
    Set m_oFWAppData = Nothing
    Set m_ofwRefData = Nothing
    Set m_oFWTable = Nothing
    
    'Release the business objects
    Set m_oBOEditorialPermissions = Nothing
    Set m_oBOCatalogueElementMatrix = Nothing
    Set m_oBOLookUp = Nothing
    Set m_oVWSharedCatalogueEntry = Nothing
    
    Exit Sub
    
ErrorHandler:
    ConvertError sSource
End Sub

Public Sub DisplayCreateNewCatalogueMenu()
'
'Description :Displays the menu to create a new catalogue entry
'
'Amendment history
'   d01  ndk     25Aug1999       Integrated from VWEditset
'   d03  ndk     28Oct1999       asp page Links changed
'   d16  ndk     08Dec1999       Added help button
'   d23  ndk     21Dec1999       Added security check
'   d24  ndk     27Dec1999       Changed to VI code use customizable styles
'   d25  ndk     27Dec1999       Added tooltips
'

On Error GoTo ErrorHandler

    Dim sSource As String
    Dim sHTML As String
    Dim eSecurityCheckStatus As ESecurityCheckResult
    
    Dim lEditsetID As Long
    
    Dim typRowStyle As RowStyle
    Dim typCellStyle As CellStyle
    Dim typTableStyle As TableStyle
    
    Dim ofwTable As FWTable
    
    sSource = Me.Name & ".DisplayCreateNewCatalogueMenu"
    
    'Get the editset id
    lEditsetID = m_aspRequest.QueryString(gk_sParamEditSetID)
    
    'perform security check
    m_oBOEditorialPermissions.EditsetID = lEditsetID
    m_oBOEditorialPermissions.UserID = m_lLoggedUserID
    
    eSecurityCheckStatus = m_oBOEditorialPermissions.Check(EUserCanCreateCatalogueEntryOnEditSet)
    
    If eSecurityCheckStatus <> ESecurityCheckResult.ESRSecurityCheckSuccessful Then
        RaiseError sSource, _
               GetErrorEnumForSecurityCheckResult(eSecurityCheckStatus), _
               GetDescriptionForSecurityCheckError( _
                        GetErrorEnumForSecurityCheckResult(eSecurityCheckStatus))
    End If
    
    Set ofwTable = New FWTable
    
    
    typTableStyle = m_oFWHTML.GetTableStyle(TABContentStyle)
    typTableStyle.sWidth = "100%"
    typTableStyle.lCELLPADDING = 3
    
    ofwTable.TableStyle = typTableStyle

    typCellStyle = m_oFWHTML.GetCellStyle(CELLNormalStyle)
    typCellStyle.sWidth = "100%"
    typCellStyle.lCOLSPAN = 2
    
    'Heading
    ofwTable.StartRow
    ofwTable.AddCellAsHTML m_oFWHTML.Bold( _
                m_oFWAppData.GetDisplayInformation(erfDisplayCreateCatalogueMenuHeader1) _
                ), typCellStyle
    ofwTable.EndRow
        
    ofwTable.AddBlankRow
    
    'Heading
    ofwTable.StartRow
    ofwTable.AddCellAsHTML (m_oFWAppData.GetDisplayInformation(erfDisplayCreateCatalogueMenuHeader2)), typCellStyle
    ofwTable.EndRow
    
    ofwTable.AddBlankRow
    
    'Create from scratch
    ofwTable.StartRow
    ofwTable.AddCellAsHTML gk_sHTMLNonBreakSpace
    ofwTable.AddCellAsHTML _
        m_oFWHTML.Hyperlink(gk_sLinkStartCreateNewCatalogueFromScratch & _
                            "?" & gk_sParamEditSetID & "=" & lEditsetID, _
                            m_oFWAppData.GetDisplayInformation(erfDisplayCreateCatalogueMenuHyperLinkFromScratch), _
                            gk_sMainFrame, , GetToolTip(ETOOLTIP.ttCreateEntryFromScratch))
    ofwTable.EndRow
    
    ofwTable.AddBlankRow

    'Create by copying another entry
    ofwTable.StartRow
    ofwTable.AddCellAsHTML gk_sHTMLNonBreakSpace
    ofwTable.AddCellAsHTML _
        m_oFWHTML.Hyperlink(gk_sLinkStartCreateNewCatalogueByCopying & _
                            "?" & gk_sParamEditSetID & "=" & lEditsetID, _
                            m_oFWAppData.GetDisplayInformation(erfDisplayCreateCatalogueMenuHyperLinkByCopying), _
                            gk_sMainFrame, , GetToolTip(ETOOLTIP.ttCreateEntryByCopying))
    ofwTable.EndRow
    
    ofwTable.AddBlankRow

    'Create a new letter code
    ofwTable.StartRow
    ofwTable.AddCellAsHTML gk_sHTMLNonBreakSpace
    ofwTable.AddCellAsHTML _
        m_oFWHTML.Hyperlink(gk_sLinkDisplayCreateNewLetterCodeScreen & _
                        "?" & gk_sParamEditSetID & "=" & lEditsetID, _
                        m_oFWAppData.GetDisplayInformation(erfDisplayCreateCatalogueMenuHyperLinkNewLetterCode), _
                        gk_sMainFrame, , GetToolTip(ETOOLTIP.ttCreateLettercode))

    ofwTable.EndRow
    
    sHTML = sHTML & ofwTable.GetHTML
    
    'Add the help button
    sHTML = sHTML & HTMLToAddHelpButton(gk_sHelpLinkToCreateCatalogueMenu)
    
    m_aspResponse.Write sHTML

    Set ofwTable = Nothing
    
    m_objObjectContext.SetComplete
    
    Exit Sub
ErrorHandler:
    m_objObjectContext.SetAbort
    ShowError sSource
End Sub
Public Sub DisplayMenuForRelatedSeparatedMaterial()
'
'Description :Displays the menu choose catalogue selction method for related/separate material
'
'Amendment history
'
'   d17 ndk     09Dec1999       Created
'   d24 ndk     27Dec1999       Changed to VI code use customizable styles
'   d33 dtm     09Feb2000       Fixed defect: gk_sIMAGESmallGo already includes full path so adding it on again was erroneous
'   d51 daj     05Apr2000       Added Cancel buttons to the Related/Separated Material link forms
'
On Error GoTo ErrorHandler

    Dim sSource As String, sHTML As String
    
    Dim lCatalogueID As Long, eCatalogueLevelID As ECATALOGUELEVEL
    Dim sQueryString As String, sParamString As String
    
    Dim typRowStyle As RowStyle
    Dim typCellStyle As CellStyle
    Dim typTableStyle As TableStyle
    
    Dim ofwTable As FWTable
    
    sSource = Me.Name & ".DisplayMenuForRelatedSeparatedMaterial"
    
    'Get the editset id
    lCatalogueID = m_aspRequest.QueryString(gk_sParamCatalogueID)
    eCatalogueLevelID = m_aspRequest.QueryString(gk_sParamCatalogueLevel)
    
    Set ofwTable = New FWTable
    
    typTableStyle = m_oFWHTML.GetTableStyle(TABContentStyle)
    typTableStyle.sWidth = "100%"
    typTableStyle.lCELLPADDING = 3
    
    ofwTable.TableStyle = typTableStyle

    typCellStyle = m_oFWHTML.GetCellStyle(CELLNormalStyle)
    typCellStyle.sWidth = "100%"
    typCellStyle.lCOLSPAN = 2
    
    'Start the form
    sHTML = sHTML & m_oFWHTML.FormStart(gk_sLinkFindCatalogueAndStartCatalogueBrowser, gk_sMainFrame, gk_sHTMLDefaultFormName, MPost, True)
    
    'Heading
    ofwTable.StartRow
    ofwTable.AddCellAsHTML m_oFWHTML.Bold("Find catalogue entry to link to")
    
    ofwTable.EndRow
        
    ofwTable.AddBlankRow
    
    'generate the query string to open the browser at the required level
    sQueryString = GenerateQueryStringsForCatalogueBrowser(lCatalogueID, eCatalogueLevelID)
    
    'paramstring
    sParamString = gk_sLinkDisplayCatalogueBrowserForRelatedSeparateMaterial & "?" & sQueryString
    
    'Browse from this entry
    ofwTable.StartRow
    ofwTable.AddCellAsHTML _
        m_oFWHTML.Hyperlink(sParamString, _
                            "Browse from this entry", _
                            gk_sMainFrame)
    ofwTable.EndRow
    
    ofwTable.AddBlankRow

    'Browse from this reference
    ofwTable.StartRow
    ofwTable.AddCellAsHTML _
        "Browse from this reference: " & _
        m_oFWHTML.InputField(ITText, gk_sNameCatalogueReference, 10, 50, "", VCNone, True) & _
        m_oFWHTML.SubmitImage(gk_sIMAGESmallGo, "", "")
    ofwTable.EndRow

    ' d51 - Add a Cancel button
    ofwTable.AddBlankRow
    ofwTable.StartRow
    ofwTable.AddCellAsHTML m_oFWHTML.ImageButton(gk_sImagesPath & gk_sImageCancel, "cmdCancel", gk_sPAGEAFCancelPickTerm)
    ofwTable.EndRow
    
    ofwTable.AddBlankRow

    sHTML = sHTML & ofwTable.GetHTML
    
    'Add the help button
    sHTML = sHTML & HTMLToAddHelpButton(gk_sHelpLinkToMenuForRelatedSeparatedMaterial)
    
    m_aspResponse.Write sHTML

    Set ofwTable = Nothing
    
    m_objObjectContext.SetComplete
    
    Exit Sub
ErrorHandler:
    m_objObjectContext.SetAbort
    ShowError sSource
End Sub

Public Sub FindCatalogueAndStartCatalogueBrowser()
'
'Description : Finds the catalogue keyed by the user and starts the
'              catalogue browser from that point
'
'Amendment history
'
'   d17 ndk     09Dec1999       Created
'   d39 tmj     25Feb2000       Added (hard coded) extra parameter in querystring
'                               this tells the page if this is the first time it has been
'                               accessed
'
On Error GoTo ErrorHandler

    Dim sSource As String, lCatalogueID As Long, eCatalogueLevelID As ECATALOGUELEVEL
    Dim sQueryString As String, sCatalogueReference As String, lMatchesFound As Long
    
    sSource = Me.Name & ".FindCatalogueAndStartCatalogueBrowser"

    'Get catalogue reference from the form
    sCatalogueReference = Trim(m_aspRequest.Form(gk_sNameCatalogueReference))
    
    'Search for the entry
    lMatchesFound = SearchCatalogue(sCatalogueReference, lCatalogueID, eCatalogueLevelID)
    
    '!TODO
    'Search for the referennce
    'This part will be done by Joe Stapleton
    'The search function will put appropriate values in variables lCatalogueToBeCopiedId and lCatalogueToBeCopiedId
    'These are passed by reference to the function
    '
    'The search stub
    'Private Function SearchCatalogue( _
    '    ByVal t_sCatalogueReference As String, _
    '    ByRef t_lCatalogueToBeCopiedId As Long, _
    '    ByRef t_eCatalogueToBeCopiedLevel As ECATALOGUELEVEL)
    '
    ' Note : ( t_lCatalogueToBeCopiedId, t_eCatalogueToBeCopiedLevel are recieved by ref)
    '
    'Search for catalogue t_sCatalogueReference
    'if matches found is one then
    '   SearchCatalogue = 1
    '   t_lCatalogueToBeCopiedId = the catalogue id
    '   t_eCatalogueToBeCopiedLevel = the catalogue level
    'else
    '   SearchCatalogue = no. of matches found
    'Endif
    'End Function

    '!TODO Uncomment this line later
    'lMatchesFound = SearchCatalogue(sCatalogueReference, lCatalogueToBeCopiedId, eCatalogueToBeCopiedLevel)
    

    
    Select Case lMatchesFound
        Case 1  'this is ok
            'generate the query string to open the browser at the required level
            sQueryString = GenerateQueryStringsForCatalogueBrowser(lCatalogueID, eCatalogueLevelID)
            
            'Show page to display catalogue browser for selecting catalogue
            m_aspResponse.Redirect _
               gk_sLinkDisplayCatalogueBrowserForRelatedSeparateMaterial & "?" & sQueryString
            
        Case Is > 1 'this is not ok
            RaiseError sSource, peMoreThanOneCatalogueFoundForRelatedSeparateMaterial, "More than one catalogue found matching the reference you entered"
        Case Is < 1 ' this not ok
            RaiseError sSource, peNoCatalogueFoundForRelatedSeparateMaterial, "No catalogue was found matching the reference you entered"
    End Select
    
    m_objObjectContext.SetComplete
    
    Exit Sub
ErrorHandler:
    m_objObjectContext.SetAbort
    ShowError sSource
End Sub

Private Function SearchCatalogue( _
        ByVal t_sCatalogueReference As String, _
        ByRef t_lCatalogueToBeCopiedId As Long, _
        ByRef t_eCatalogueToBeCopiedLevel As ECATALOGUELEVEL) As Long
        
'
On Error GoTo ErrorHandler
    Dim sSource As String
    Dim oboRefChecker As New BORefChecker
    Dim sDocKey As String
    sSource = "VWCatalogueEntry.SearchCatalogue"
    
    'Check the ref against current PRO and then
    'former departmental refs and PRO refs
    If oboRefChecker.GoToCatEntry(t_sCatalogueReference) = True Then
        If oboRefChecker.Results.ResultCount > 0 Then
            oboRefChecker.Results.MoveFirst
            sDocKey = oboRefChecker.Results.GetData(gk_sDIVerityDocKey)
            t_lCatalogueToBeCopiedId = DBIDFromDocKey(sDocKey, t_eCatalogueToBeCopiedLevel)
        End If
    Else
        RaiseError sSource, PEUSERInvalidGoToReference, _
            "Could not find that reference"
    End If
    
    SearchCatalogue = oboRefChecker.Results.ResultCount

Exit Function
ErrorHandler:
    ConvertError sSource
End Function
Public Sub StoreSelectedCatalogueForCreateByCopy()
'
'Description
'
'Stores the selected copy when creating catalogue by copying (1st this happens)
'and redirects to the page displaying the catalogue browser for parent selection
'
'Amendment history
'
'   d3  ndk     28Oct1999       Created
'   d16 ndk     08Dec1999       Passing the query string when redirecting to the catalogue browser
'   d38 tmj     23Feb2000       Added branch so could copy a department
'   d68 alm     27Feb2001       Modified to work with the new browser
'

On Error GoTo ErrorHandler

    Dim sSource As String
    Dim lCatalogueToBeCopiedId As Long
    Dim eCatalogueToBeCopiedLevel As ECATALOGUELEVEL
    Dim lParentCatalogueID As Long
    Dim eParentCatalogueLevel As ECATALOGUELEVEL
    Dim sQueryString As String
    
    sSource = Me.Name & ".StoreSelectedCatalogueForCreateByCopy"

    'Get the data from the query string
    lCatalogueToBeCopiedId = m_aspRequest.QueryString(gk_sParamCatalogueID)
    eCatalogueToBeCopiedLevel = m_aspRequest.QueryString(gk_sParamCatalogueLevel)
    lParentCatalogueID = m_aspRequest.QueryString(gk_sParamParentCatalogueID)
    eParentCatalogueLevel = m_aspRequest.QueryString(gk_sParamparentCatalogueLevel)
        
    'Store these details in the session store
    StoreDataInSession dlCreateCatalogueCatalogueToBeCopiedID, lCatalogueToBeCopiedId
    StoreDataInSession diCreateCatalogueCatalogueToBeCopiedLevel, eCatalogueToBeCopiedLevel

    'If not trying to copy a department then find a parent ot copy under
    If eCatalogueToBeCopiedLevel = ctLetterCode Then
        'Show page to display copy summary details page for the department
        m_aspResponse.Redirect gk_sLinkDisplayCreateNewCatalogueByCopyingDepartmentScreen
    Else
        'generate the query string to open the browser at the required level
        sQueryString = GenerateQueryStringsForCatalogueBrowser(lParentCatalogueID, eParentCatalogueLevel)
        'Show page to display catalogue browser for selecting parent
        m_aspResponse.Redirect _
            gk_sLinkDisplayCatalogueBrowserForInitialCreateByCopyParentSelection & "?" & sQueryString
    End If
    
    m_objObjectContext.SetComplete
    
    Exit Sub
ErrorHandler:
    m_objObjectContext.SetAbort
    ShowError sSource
End Sub

Public Sub StoreSelectedParentForCreateByCopy()
'
'Description
'
'Stores the selected parent when creating catalogue by copying (2nd this happens)
'And then calls the asp page that displays the form to create a new catalogue entry by copying
'
'Amendment history
'
'   d3  ndk     28Oct1999       Created
'   d46 tmj     21Mar2000       Added Error Checking for Moved Parents
'   d47 tmj     24Mar2000       Passed EditsetID to Move Checks
'   d74 alm     25Oct2001       Checks that the parent does not exist is
'                               not marked for deletion
On Error GoTo ErrorHandler

    Dim sSource As String
    Dim oBOBasicCatalogueEntry As BOBasicCatalogueEntry
    Dim lCatID As Long
    Dim eCATLN As ECATALOGUELEVEL
    Dim lEditsetID As Long
    
    
    sSource = Me.Name & ".StoreSelectedParentForCreateByCopy"

    Set oBOBasicCatalogueEntry = New BOBasicCatalogueEntry

    If m_aspRequest(gk_sNameCatalogueLevel) = ECATALOGUELEVEL.ctItem Then
        RaiseError sSource, PEItemCannotBeParent, _
                   "An Item cannot be a parent"
        Exit Sub
    End If

    lCatID = m_aspRequest(gk_sNameCatalogueID)
    eCATLN = m_aspRequest(gk_sNameCatalogueLevel)
    lEditsetID = m_aspRequest(gk_sNameEditsetID)

    If Not oBOBasicCatalogueEntry.CheckNotMoved(lCatID, eCATLN, lEditsetID) Then
        RaiseError sSource, PEItemCannotBeParent, _
            "Cannot set Catalogue (" & _
            GenerateCatalogueReferenceFromID(lCatID, eCATLN) & ")" & _
            " to be a parent as it has been moved"

        Exit Sub
    End If
    
    ' d74 Determine if the parent has been marked for deletion
    If oBOBasicCatalogueEntry.CheckCatalogueParentMarkedForDeletion(lCatID, eCATLN) Then
        RaiseError sSource, EPROCATINTERNALERROR.PECatalogue, _
            gk_sErrorCatalogueParentMarkedForDeletion & _
            " (Catalogue Ref - " & GenerateCatalogueReferenceFromID(lCatID, eCATLN) & _
            ", Catalogue Level - " & GetCatalogueLevelDescription( _
                                GetCatalogueLevelCode(eCATLN)) & ")<br>"
        Exit Sub
    End If

    m_aspResponse.Redirect _
        gk_sLinkDisplayCreateNewCatalogueByCopyingScreen & "?" & m_aspRequest.QueryString
    
    m_objObjectContext.SetComplete
    
    Exit Sub
ErrorHandler:
    m_objObjectContext.SetAbort
    ShowError sSource
End Sub




Public Sub StoreSelectedParentForAddExisting()
'
'Description
'
'Stores the selected parent when creating catalogue from scratch
'
'Amendment history
'
'   d3  bcj     19Jan2001       Created

On Error GoTo ErrorHandler

    Dim oBOBasicCatalogueEntry As BOBasicCatalogueEntry
    Dim lCatID As Long
    Dim eCATLN As ECATALOGUELEVEL
    Dim lEditsetID As Long
    Dim sSource As String
    
    Set oBOBasicCatalogueEntry = New BOBasicCatalogueEntry
    
    sSource = Me.Name & ".StoreSelectedParentForAddExisting"
    
    If m_aspRequest(gk_sNameCatalogueLevel) = ECATALOGUELEVEL.ctItem Then
        RaiseError sSource, PEItemCannotBeParent, _
                    "An Item cannot be a parent"
        Exit Sub
        
    End If
    
    lCatID = m_aspRequest(gk_sNameCatalogueID)
    eCATLN = m_aspRequest(gk_sNameCatalogueLevel)
    lEditsetID = m_aspRequest(gk_sNameEditsetID)

    If Not oBOBasicCatalogueEntry.CheckNotMoved(lCatID, eCATLN, lEditsetID) Then
        RaiseError sSource, PEItemCannotBeParent, _
            "Cannot set Catalogue (" & _
            GenerateCatalogueReferenceFromID(lCatID, eCATLN) & ")" & _
            " to be a parent as it has been moved"

        Exit Sub
    End If

   

    m_objObjectContext.SetComplete
    
    Exit Sub
ErrorHandler:
    m_objObjectContext.SetAbort
    ShowError sSource
End Sub




Public Sub StoreSelectedParentForCreateFromScratch()
'
'Description
'
'Stores the selected parent when creating catalogue from scratch
'And then calls the asp page that displays the form to create a new catalogue entry from scratch
'
'Amendment history
'
'   d3  ndk     28Oct1999       Created
'   d47 tmj     24Mar2000       Passed EditsetID to Move Checks
'   d74 alm     25Oct2001       Checks that the parent does not exist is
'                               not marked for deletion

On Error GoTo ErrorHandler

    Dim oBOBasicCatalogueEntry As BOBasicCatalogueEntry
    Dim lCatID As Long
    Dim eCATLN As ECATALOGUELEVEL
    Dim lEditsetID As Long

    Dim sSource As String
    
    Set oBOBasicCatalogueEntry = New BOBasicCatalogueEntry
    
    sSource = Me.Name & ".StoreSelectedParentForCreateFromScratch"

    If m_aspRequest(gk_sNameCatalogueLevel) = ECATALOGUELEVEL.ctItem Then
        RaiseError sSource, PEItemCannotBeParent, _
                   "An Item cannot be a parent"
        Exit Sub
    End If

    lCatID = m_aspRequest(gk_sNameCatalogueID)
    eCATLN = m_aspRequest(gk_sNameCatalogueLevel)
    lEditsetID = m_aspRequest(gk_sNameEditsetID)

    If Not oBOBasicCatalogueEntry.CheckNotMoved(lCatID, eCATLN, lEditsetID) Then
        RaiseError sSource, PEItemCannotBeParent, _
            "Cannot set Catalogue (" & _
            GenerateCatalogueReferenceFromID(lCatID, eCATLN) & ")" & _
            " to be a parent as it has been moved"

        Exit Sub
    End If
            
    ' d74 Determine if the parent has been marked for deletion
    If oBOBasicCatalogueEntry.CheckCatalogueParentMarkedForDeletion(lCatID, eCATLN) Then
        RaiseError sSource, EPROCATINTERNALERROR.PECatalogue, _
            gk_sErrorCatalogueParentMarkedForDeletion & _
            " (Catalogue Ref - " & GenerateCatalogueReferenceFromID(lCatID, eCATLN) & _
            ", Catalogue Level - " & GetCatalogueLevelDescription( _
                                GetCatalogueLevelCode(eCATLN)) & ")<br>"
        Exit Sub
    End If

    m_aspResponse.Redirect _
        gk_sLinkDisplayCreateNewCatalogueFromScratchScreen & "?" & m_aspRequest.QueryString
    
    m_objObjectContext.SetComplete
    
    Exit Sub
ErrorHandler:
    m_objObjectContext.SetAbort
    ShowError sSource
End Sub
Public Sub StoreSelectedParentForMoveCatalogue()
'
'Description
'
'Stores the selected parent when creating catalogue from scratch
'And then calls the asp page that displays the form to create a new catalogue entry from scratch
'
'Amendment history
'
'   d3  ndk     28Oct1999       Created
'   d74 alm     25Oct2001       Checks that the parent does not exist is
'                               not marked for deletion
On Error GoTo ErrorHandler

    Dim oBOBasicCatalogueEntry As BOBasicCatalogueEntry
    Dim lCatID As Long
    Dim eCATLN As ECATALOGUELEVEL

    Dim sSource As String
    
    Set oBOBasicCatalogueEntry = New BOBasicCatalogueEntry
    
    sSource = Me.Name & ".StoreSelectedParentForMoveCatalogue"
    
    lCatID = m_aspRequest(gk_sNameCatalogueID)
    eCATLN = m_aspRequest(gk_sNameCatalogueLevel)
    
    ' d74 Determine if the parent has been marked for deletion
    If oBOBasicCatalogueEntry.CheckCatalogueParentMarkedForDeletion(lCatID, eCATLN) Then
        RaiseError sSource, EPROCATINTERNALERROR.PECatalogue, _
            gk_sErrorCatalogueParentMarkedForDeletion & _
            " (Catalogue Ref - " & GenerateCatalogueReferenceFromID(lCatID, eCATLN) & _
            ", Catalogue Level - " & GetCatalogueLevelDescription( _
                                GetCatalogueLevelCode(eCATLN)) & ")<br>"
        Exit Sub
    End If

    m_aspResponse.Redirect _
        gk_sLinkMoveCatalogueToNewParent & "?" & m_aspRequest.QueryString
    
    m_objObjectContext.SetComplete
    
    Exit Sub
ErrorHandler:
    m_objObjectContext.SetAbort
    ShowError sSource
End Sub

Public Sub DisplayCreateNewCatalogueFromScratchScreen()
'
'Description
'
'Calls the methods to display the page to select the parent
'for the new catalogue entry being created
'
'Amendment history
'   d1  ndk     25Aug1999       Integrated from VWEditset

On Error GoTo ErrorHandler

    Dim sSource As String
    Dim lParentCatalogueID As Long, lEditsetID As Long, eParentCatalogueLevel As ECATALOGUELEVEL
    
    sSource = Me.Name & ".DisplayCreateNewCatalogueFromScratchScreen"
    
    'Get editset id from the session
    lEditsetID = GetDataFromSession(dlCreateCatalogueEditSetID)
    
    'Get the remaining details from the query string
    lParentCatalogueID = m_aspRequest.QueryString(gk_sParamCatalogueID)
    eParentCatalogueLevel = m_aspRequest.QueryString(gk_sParamCatalogueLevel)
    
    DisplayEditSummaryDetailsScreen _
        etLoadFromNoneCreateCatalogueEntry, lEditsetID, , , , , ecCreateCatalogueFromScatch, lParentCatalogueID, eParentCatalogueLevel
    
    m_objObjectContext.SetComplete
    
    Exit Sub
ErrorHandler:
    m_objObjectContext.SetAbort
    ShowError sSource
End Sub

Public Sub DisplayCreateNewCatalogueByCopyingScreen()

'Description
'
'Calls the methods to display the page to select the parent
'for the new catalogue entry being created
'
'Amendment history
'   d1  ndk     25Aug1999       Integrated from VWEditset
'   d3  ndk     28Oct1999       Create by copying now permits copying only at the same level (except piece and item)
'

On Error GoTo ErrorHandler

    Dim sSource As String
    Dim lEditsetID As Long
    Dim lParentCatalogueID As Long, eParentCatalogueLevel As ECATALOGUELEVEL
    Dim lCatalogueToBeCopiedId As Long, eCatalogueToBeCopiedLevel As ECATALOGUELEVEL

    sSource = Me.Name & ".DisplayCreateNewCatalogueByCopyingScreen"
        
    'Get the data from the query string
    lParentCatalogueID = m_aspRequest.QueryString(gk_sParamCatalogueID)
    eParentCatalogueLevel = m_aspRequest.QueryString(gk_sParamCatalogueLevel)
    
    'get the data from the session store (store when the user chooses the catalogue to be copied)
    lEditsetID = GetDataFromSession(dlCreateCatalogueEditSetID)
    lCatalogueToBeCopiedId = GetDataFromSession(dlCreateCatalogueCatalogueToBeCopiedID)
    eCatalogueToBeCopiedLevel = GetDataFromSession(diCreateCatalogueCatalogueToBeCopiedLevel)
    
    'check whether catalogue to be copied can be created directly under the parent choosen
    Select Case eCatalogueToBeCopiedLevel
        Case ECATALOGUELEVEL.ctLetterCode
            RaiseError sSource, PECannotCreateLetterCodeByCopying, _
            " Can not create a " & _
                GetCatalogueLevelDescription(GetCatalogueLevelCode(ctLetterCode)) & _
                " by copying"
            
        Case ECATALOGUELEVEL.ctDivision
            'parent has to be a lettercode
            If eParentCatalogueLevel <> ECATALOGUELEVEL.ctLetterCode Then
                 RaiseError sSource, PEInvalidParentLevelForCreateByCopying, _
                    "Only a " & _
                    GetCatalogueLevelDescription(GetCatalogueLevelCode(ctLetterCode)) & _
                    " can be a parent for division"
            End If
        Case ECATALOGUELEVEL.ctClass
            'parent has to be a lettercode or division
            If eParentCatalogueLevel <> ECATALOGUELEVEL.ctLetterCode _
                And eParentCatalogueLevel <> ECATALOGUELEVEL.ctDivision Then
                    RaiseError sSource, PEInvalidParentLevelForCreateByCopying, _
                        "Only a " & _
                            GetCatalogueLevelDescription(GetCatalogueLevelCode(ctLetterCode)) & _
                            " or " & _
                            GetCatalogueLevelDescription(GetCatalogueLevelCode(ctDivision)) & _
                            " can be a parent for " & _
                            GetCatalogueLevelDescription(GetCatalogueLevelCode(ctClass))
            End If
        Case ECATALOGUELEVEL.ctHeaders
            'parent has to be a class
            If eParentCatalogueLevel <> ECATALOGUELEVEL.ctClass Then
                 RaiseError sSource, PEInvalidParentLevelForCreateByCopying, _
                    "Only a " & _
                    GetCatalogueLevelDescription(GetCatalogueLevelCode(ctClass)) & _
                    " can be a parent for " & _
                    GetCatalogueLevelDescription(GetCatalogueLevelCode(ctHeaders))
            End If
        Case ECATALOGUELEVEL.ctSubHeaders
            'parent has to be a header
            If eParentCatalogueLevel <> ECATALOGUELEVEL.ctHeaders Then
                 RaiseError sSource, PEInvalidParentLevelForCreateByCopying, _
                    "Only a " & _
                    GetCatalogueLevelDescription(GetCatalogueLevelCode(ctHeaders)) & _
                    " can be a parent for " & _
                    GetCatalogueLevelDescription(GetCatalogueLevelCode(ctSubHeaders))
            End If
        Case ECATALOGUELEVEL.ctPiece
            'Note : When you are copying from a piece, you could create another piece or an item
            'parent has to be a subheader, header or class (for piece) or piece (for item)
            If eParentCatalogueLevel <> ECATALOGUELEVEL.ctSubHeaders _
                And eParentCatalogueLevel <> ECATALOGUELEVEL.ctHeaders _
                And eParentCatalogueLevel <> ECATALOGUELEVEL.ctClass _
                And eParentCatalogueLevel <> ECATALOGUELEVEL.ctPiece Then
                
                RaiseError sSource, PEInvalidParentLevelForCreateByCopying, _
                    "Only a " & _
                    GetCatalogueLevelDescription(GetCatalogueLevelCode(ctSubHeaders)) & "," & _
                    GetCatalogueLevelDescription(GetCatalogueLevelCode(ctHeaders)) & "," & _
                    GetCatalogueLevelDescription(GetCatalogueLevelCode(ctClass)) & "," & _
                    GetCatalogueLevelDescription(GetCatalogueLevelCode(ctPiece)) & _
                    " can be a parent for the catalogue entry you are creating"
            End If
        Case ECATALOGUELEVEL.ctItem
            'parent has to be a piece
            If eParentCatalogueLevel <> ECATALOGUELEVEL.ctPiece Then
                 RaiseError sSource, PEInvalidParentLevelForCreateByCopying, _
                    "Only a " & _
                    GetCatalogueLevelDescription(GetCatalogueLevelCode(ctPiece)) & _
                    " can be a parent for an " & _
                    GetCatalogueLevelDescription(GetCatalogueLevelCode(ctItem))

            End If
    End Select
    
    'Call the method to display the screen to create an entry
    DisplayEditSummaryDetailsScreen _
        etLoadFromNoneCreateCatalogueEntry, lEditsetID, , , , , ecCreateCatalogueByCopying, lParentCatalogueID, eParentCatalogueLevel, lCatalogueToBeCopiedId, eCatalogueToBeCopiedLevel
    
    m_objObjectContext.SetComplete
    
    Exit Sub

ErrorHandler:
    m_objObjectContext.SetAbort
    ShowError sSource
End Sub


Public Sub DisplayCreateNewLetterCodeScreen()
'
'Description
'
'Calls the methods to display the page to create a new letter code
'
'Amendment history
'   d1  ndk     25Aug1999       Integrated from VWEditset

On Error GoTo ErrorHandler

    Dim sSource As String
    Dim lEditsetID As Long
        
    sSource = Me.Name & ".DisplayCreateNewLetterCodeScreen"
    
    'Get the data from the query string
    lEditsetID = m_aspRequest.QueryString(gk_sParamEditSetID)
    
    'Call the method to display the screen to create an entry
    DisplayEditSummaryDetailsScreen _
         ECatalogueEditFormLoadType.etLoadFromNoneCreateCatalogueEntry, _
         lEditsetID, , , , , _
         ECatalogueEntryEditType.ecCreateLetterCodeFromScratch
    
    m_objObjectContext.SetComplete
    
    Exit Sub

ErrorHandler:
    m_objObjectContext.SetAbort
    ShowError sSource
    
End Sub



Public Sub SetBrowserUsage(t_eBrowserUsage As ECATALOGUEBROWSERUSAGE)
' Sets the browser usage. This is used so we can tell what 'mode' the
' browser is working in
'   d1  bcj     22Jan20001

m_eBrowserUsage = t_eBrowserUsage

End Sub

Public Sub AddCataloguesToEditSet()
' Adds catalogues seleced to the editset
'
'   d1  ndk     25Aug1999       Moved from VWEditset to here (Created by KSK)

On Error GoTo ErrHandler
    Dim lNoOfCheckBoxes As Integer
    Dim lCatalogueID() As Long, eCatalogueLevels() As ECATALOGUELEVEL
    Dim i As Integer
    Dim sCatRefStored As String
    Dim iNoOfSelectedCatalogues As Integer
    Dim sSource As String
    Dim lEditsetID  As Long
    Dim sEditsetParam  As String
    
    Dim oBOBasicCatalogueEntry As BOBasicCatalogueEntry
    
    sSource = "VWCatalogueEntry.AddCataloguesToEditSet()"
    
    Set oBOBasicCatalogueEntry = New BOBasicCatalogueEntry
    
    
    ' d59 bcj   Added condition for enhancement 2, enhancement set 2
    
    If m_eBrowserUsage = eEditorialParentSelectionForAddFromExisting Then
    
        ' If the browser is in add catalogues to entry set, then the EDID
        ' is already stored in the session, so we just need to extract it
        lEditsetID = SessionManager.GetDataFromSession(dlCreateCatalogueEditSetID)
        
        ' We do not perform the error checking here, though. We do this
        ' outside this 'if' clause
        
    Else
        
        ' Otherwise, the EDID will be present in the querystring
        
        If m_aspRequest(gk_sParamEditSetID) = "" Then
            RaiseError sSource, EPROCATINTERNALERROR.peEditSetIDInvalue, _
                        "Editset ID not found"
        End If
          
          lEditsetID = CLng(m_aspRequest(gk_sParamEditSetID))
          
    End If
    
    
  
    If lEditsetID <= 0 Then
        RaiseError sSource, EPROCATINTERNALERROR.peEditSetIDInvalue, _
                            "Editset ID not found"
    End If
    sEditsetParam = "?" & gk_sParamEditSetID & "=" & lEditsetID
    
   ' d59    Added condition.
   ' Enhancement 2, enhancement set 1
    If m_eBrowserUsage = eEditorialParentSelectionForAddFromExisting Then
        ReDim Preserve lCatalogueID(1)
        ReDim Preserve eCatalogueLevels(1)
        lCatalogueID(1) = m_aspRequest.Item(gk_sParamCatalogueID)
        eCatalogueLevels(1) = m_aspRequest.Item(gk_sParamCatalogueLevel)
   
        ' Make sure that we have valid catalogue details
        ' d58 changed lCatalogueID(1) > 0 with lCatalogueID(1) <> 0
        '     as pieces can be negative.
        
        If lCatalogueID(1) <> 0 Then
            oBOBasicCatalogueEntry.AddCataloguesIntoEditSet lEditsetID, lCatalogueID, eCatalogueLevels

        Else
            RaiseError sSource, EPROCATINTERNALERROR.peInvalidCatalogueReference, _
                                                    "Catalogue ID not recognised"
        End If
   
   Else
   
        'Get the number of check boxes
        lNoOfCheckBoxes = IIf(IsEmpty(m_aspRequest(ksNameCheckBoxCount)), 0, m_aspRequest(ksNameCheckBoxCount))
        
        'Loop through all the chek boxed and pick the checked ones and separate the values in the value field
        ' of the checkbox as Catalogue ID and Level
        iNoOfSelectedCatalogues = 0
        For i = 1 To lNoOfCheckBoxes
            
            sCatRefStored = m_aspRequest(ksNameCheckBoxes & Trim(CStr(i)))
            If Not IsEmpty(sCatRefStored) And sCatRefStored <> "" Then
                iNoOfSelectedCatalogues = iNoOfSelectedCatalogues + 1
                ReDim Preserve lCatalogueID(iNoOfSelectedCatalogues), eCatalogueLevels(iNoOfSelectedCatalogues)
                lCatalogueID(iNoOfSelectedCatalogues) = CLng(Mid(sCatRefStored, 1, InStr(sCatRefStored, gk_sDataDelimiter) - 1))
                eCatalogueLevels(iNoOfSelectedCatalogues) = GetCatalogueLevelEnum( _
                                            Mid(sCatRefStored, InStr(sCatRefStored, gk_sDataDelimiter) + _
                                            Len(gk_sDataDelimiter)))
            End If
        Next i
        'If there were catalogues selected then call the appropriate business object to add the catalogue
        ' into the editset
        If iNoOfSelectedCatalogues > 0 Then
            oBOBasicCatalogueEntry.AddCataloguesIntoEditSet lEditsetID, lCatalogueID, eCatalogueLevels
        End If
    
    End If
    
    m_objObjectContext.SetComplete
    
    'After adding the catalogue redirect to the page to display the editset details
    m_aspResponse.Redirect gk_sLinkDisplayEditSetDetails & sEditsetParam
    
    Set oBOBasicCatalogueEntry = Nothing
    Exit Sub
ErrHandler:

    Set oBOBasicCatalogueEntry = Nothing
    
    m_objObjectContext.SetAbort
    ShowError sSource
End Sub





Private Sub InitialiseStyles()
'
'Description
'
'Initialises the styles
'
'Amendment history
'   d1  ndk     25Aug1999       Integrated from VWEditset
'   d24 ndk     27Dec1999       Changed to VI code use customizable styles
'

On Error GoTo ErrorHandler

    Dim sSource As String
    
    sSource = Me.Name & ".InitialiseStyles"
    
    'Cell style for the left heading
    m_typCellStyleLeftHeading = m_oFWHTML.GetCellStyle(CELLGeneralLabelStyle)
    m_typCellStyleLeftHeading.sWidth = "25%"
    m_typCellStyleLeftHeading.lCOLSPAN = 1
    m_typCellStyleLeftHeading.eVALIGN = VALMiddle
    
    'Cell style for the right details
    m_typCellStyleRightDetails = m_oFWHTML.GetCellStyle(CELLGeneralContentStyle)
    m_typCellStyleRightDetails.sWidth = "100%"
    m_typCellStyleRightDetails.lCOLSPAN = 1
    m_typCellStyleRightDetails.eVALIGN = VALMiddle

    'Cell style for ok cancel button
    m_typCellStyleForOKCancel = m_oFWHTML.GetCellStyle(CELLGeneralLabelStyle)
    m_typCellStyleForOKCancel.sWidth = "100%"
    m_typCellStyleForOKCancel.lCOLSPAN = 4
    m_typCellStyleForOKCancel.eVALIGN = VALBottom

    'font style for left headings
    m_typFontStyleLeftHeading = m_oFWHTML.GetFontStyle(FONTReadWriteLabelStyle)
    m_typFontStyleLeftHeading.iFontSize = 2
    
    'font style for right details
    m_typFontStyleRightDetails = m_oFWHTML.GetFontStyle(FONTReadWriteDataStyle)
    m_typFontStyleRightDetails.iFontSize = 2
    
    'Table style for contents
    m_typTableStyleContents = m_oFWHTML.GetTableStyle(TABContentStyle)
    m_typTableStyleContents.sWidth = "100%"
    m_typTableStyleContents.lCELLPADDING = 0
    m_typTableStyleContents.lCELLSPACING = 0
    
    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub

Private Sub DisplayHeaderDetails( _
        ByVal t_lEditSetID As Long, _
        ByVal t_lCatalogueID As Long, _
        ByVal t_eCatalogueLevelID As ECATALOGUELEVEL, _
        ByVal t_eCurrentTab As ETabType, _
        ByVal t_eSectionType As ESectionType, _
        Optional ByVal t_lCurrentEditSetID As Long = 0)

'
'Description
'
'Displays the details title, covering date, reference on top of display catalogue details
'
'Amendment history
'   d1  ndk     25Aug1999       Integrated from VWEditset
'   d6  ndk     15Nov1999       Displaying scope and content at the top for piece and item
'   d24 ndk     27Dec1999       Changed to VI code use customizable styles
'   d25 ndk     27Dec1999       Added tooltips
'   d34 jcc     10Feb2000       removed covering dates from catalogue entry heading
'   d36 jcc     18Feb2000       and remove the -from- bit too!
'   d42 tmj     02Mar2000       Fixed links, now point to current catalogue
'   d48 daj     27Mar2000       Tidied up the audit trail and fixed defect 766 - Audit trail 'back' links not working
'                               once you move off the summary tab
'   d56 daj     04Jan2001       Added Covering Dates to the header details
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String, sHeader As String
    Dim oBOSummaryCatalogueDetails As BOSummaryCatalogueDetails, typCellStyle As CellStyle
    Dim sCatalogueReference As String, sParamString As String
    Dim sTitleScopeAndContent As String
    Dim sCoveringDates As String
    
    'd42 parameters to hold current editset information
    Dim lLinkEditsetID As Long
    
    sSource = Me.Name & ".DisplayHeaderDetails"
    
    'Instantiate the objects
    Set oBOSummaryCatalogueDetails = New BOSummaryCatalogueDetails
    
    'set the properties
    oBOSummaryCatalogueDetails.EditsetID = t_lEditSetID
    oBOSummaryCatalogueDetails.CatalogueID = t_lCatalogueID
    oBOSummaryCatalogueDetails.CatalogueLevel = t_eCatalogueLevelID
    oBOSummaryCatalogueDetails.SectionType = t_eSectionType
    
    'Load the contents
    oBOSummaryCatalogueDetails.Load
    
    m_oFWTable.Reset
    
    'Level details heading 1
    m_oFWTable.StartRow
    
    'Catalogue reference
    sCatalogueReference = GenerateCatalogueReferenceFromID(t_lCatalogueID, t_eCatalogueLevelID)
    
    sHeader = _
        GetCatalogueLevelDescription(GetCatalogueLevelCode(t_eCatalogueLevelID)) & " " & _
        m_oFWAppData.GetDisplayInformation( _
        EEditorialResourceFile.erfTabDisplayGeneralDetailsCaptionDetailsFor) & " " & _
        sCatalogueReference

    typCellStyle = m_oFWHTML.GetCellStyle(CELLNormalStyle)
    typCellStyle.sWidth = "90%"
    typCellStyle.eALIGN = HALLeft
    
    m_oFWTable.AddCellAsHTML _
        m_oFWHTML.Paragraph(m_oFWHTML.Strong(m_oFWHTML.Font(sHeader, COLOURForegroundGeneralLabel, 4))), typCellStyle
        
    typCellStyle.sWidth = "10%"
    
    'Hyperlink back to editset
    
    ' d48 - Use the current edit set ID if available
    If Not t_lCurrentEditSetID = 0 Then
        lLinkEditsetID = t_lCurrentEditSetID
    Else
        lLinkEditsetID = t_lEditSetID
    End If
    
    m_oFWTable.AddCellAsHTML _
        m_oFWHTML.Hyperlink(gk_sLinkDisplayEditSetDetails & _
            "?" & gk_sParamEditSetID & "=" & lLinkEditsetID, _
            m_oFWAppData.GetDisplayInformation( _
            EEditorialResourceFile.erfTabDisplayGeneralDetailsHyperLinkBackToEditSet), _
            gk_sMainFrame, , GetToolTip(ETOOLTIP.ttBackToEditSet)), typCellStyle
        
    m_oFWTable.EndRow

    m_oFWTable.StartRow
    
    typCellStyle.sWidth = "90%"
    typCellStyle.eALIGN = HALLeft
    m_oFWTable.AddCellAsHTML " ", typCellStyle
    
    typCellStyle.sWidth = "10%"
    If t_eSectionType = SECeditorial Then
        'Hyperlink Audit trail
        m_oFWTable.AddCellAsHTML _
            m_oFWHTML.Hyperlink(gk_sLinkDisplayAuditTrailForCatalogue & _
                "?" & gk_sParamCatalogueID & "=" & t_lCatalogueID & _
                "&" & gk_sParamCatalogueLevel & "=" & t_eCatalogueLevelID & _
                "&" & gk_sParamEditSetID & "=" & t_lEditSetID, _
            m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfTabDisplayGeneralDetailsHyperLinkAuditTrail) _
            , gk_sMainFrame, , GetToolTip(ETOOLTIP.ttAuditTrail)), typCellStyle
    Else
        'Hyperlink - Back to editorial catalogue
        sParamString = gk_sLinkDisplayViewEditorialCatalogueDetails & _
            "?" & gk_sParamEditSetID & "=" & lLinkEditsetID & _
            "&" & gk_sParamCatalogueID & "=" & t_lCatalogueID & _
            "&" & gk_sParamCatalogueLevel & "=" & t_eCatalogueLevelID & _
            "&" & ksParamLoadType & "=" & ECatalogueEditFormLoadType.etLoadFromDatabaseEditCatalogueEntry & _
            "&" & ksParamCatalogueEntryEditType & "=" & ECatalogueEntryEditType.ecEditCatalogueEntry & _
            "&" & gk_sParamSectionType & "=" & ESectionType.SECeditorial & _
            "&" & gk_sParamDetailsTabNumber & "=" & ETabType.etTabSummary

        m_oFWTable.AddCellAsHTML _
            m_oFWHTML.Hyperlink(sParamString, "Back to " & sCatalogueReference, gk_sMainFrame, , GetToolTip(ETOOLTIP.ttBackToEditSetCatalogue)), typCellStyle
    End If
    
    m_oFWTable.EndRow
        
    If t_eCatalogueLevelID = ECATALOGUELEVEL.ctPiece Or t_eCatalogueLevelID = ECATALOGUELEVEL.ctItem Then
        'scope and content of piece and item
        sTitleScopeAndContent = VariantStringAsString(oBOSummaryCatalogueDetails.ScopeAndContent)
    Else    'for other levels it is title
        sTitleScopeAndContent = VariantStringAsString(oBOSummaryCatalogueDetails.Title)
    End If
    
    sTitleScopeAndContent = TruncateTitle(sTitleScopeAndContent)
    
    'Level details heading 2
    m_oFWTable.StartRow
    m_oFWTable.AddCellAsHTML _
            VariantStringAsString(sTitleScopeAndContent)
    m_oFWTable.EndRow
    
    ' d56 - Add the Covering Dates under the reference line
    m_oFWTable.StartRow
    
    ' Check that we have covering dates to use
    If IsNull(oBOSummaryCatalogueDetails.CoveringDates) Or _
        oBOSummaryCatalogueDetails.CoveringDates = "" Then
        sCoveringDates = gk_sNoCoveringDatesText
    Else
        sCoveringDates = oBOSummaryCatalogueDetails.CoveringDates
    End If
    
    ' Only take the first 50 characters
    sCoveringDates = Mid(sCoveringDates, 1, 50)
    m_oFWTable.AddCellAsString sCoveringDates
    m_oFWTable.EndRow
    
    m_sHtml = m_oFWTable.GetHTML
    
    m_aspResponse.Write m_sHtml
    
    Set oBOSummaryCatalogueDetails = Nothing
    
    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub

Private Sub DisplayHeaderTab( _
        ByVal t_lEditSetID As Long, _
        ByVal t_lCatalogueID As Long, _
        ByVal t_eCatalogueLevelID As ECATALOGUELEVEL, _
        ByVal t_eCurrentTab As ETabType, _
        ByVal t_eSectionType As ESectionType, _
        ByVal t_bViewMode As Boolean, _
        Optional ByVal t_lCurrentEditSetID As Long = 0, _
        Optional t_bDeleteEditSet As Boolean = False)
    
'
'Description
'
'Displays the tab while displaying catalogue details
'
'Amendment history
'   d1  ndk     25Aug1999       Integrated from VWEditset
'   d3  ndk     28Oct1999       Added move function
'   d5  ndk     11Nov1999       Edit button disabled for users with no rights
'   d25 ndk     27Dec1999       Added tooltips and disabled edit button
'   d32 dtm     09Feb2000       Fixed defect in DisplayHeaderTab : IF A AND B OR C was supposed to be IF A AND (B OR C) but evaluated to IF (A AND B) OR C
'   d48 daj     27Mar2000       Tidied up the audit trail and fixed defect 766 - Audit trail 'back' links not working
'                               once you move off the summary tab
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim iCounter As Integer, sParamString As String, bCanEditCatalogue As Boolean
    
    sSource = Me.Name & ".DisplayHeaderTab"
    
    'Reset the table
    m_oFWTable.Reset
    
    m_oFWTable.StartRow
    m_sHtml = ""
    
    ' d48 - Added 'current' editset ID so that it is preserved across tabs
    sParamString = "?" & gk_sParamEditSetID & "=" & t_lEditSetID & _
        "&" & gk_sParamCatalogueID & "=" & t_lCatalogueID & _
        "&" & gk_sParamCatalogueLevel & "=" & t_eCatalogueLevelID & _
        "&" & ksParamLoadType & "=" & ECatalogueEditFormLoadType.etLoadFromDatabaseEditCatalogueEntry & _
        "&" & ksParamCatalogueEntryEditType & "=" & ECatalogueEntryEditType.ecEditCatalogueEntry & _
        "&" & gk_sParamSectionType & "=" & t_eSectionType & _
        "&" & gk_sParamCurrEditSetID & "=" & t_lCurrentEditSetID & _
        "&" & gk_sParamDetailsTabNumber & "="

    'Context
    m_oFWTable.AddCellAsHTML _
        m_oFWHTML.Hyperlink(IIf(t_bViewMode And t_eCurrentTab = ETabType.etTabContext, "", gk_sLinkDisplayViewEditorialCatalogueDetails & sParamString & ETabType.etTabContext), _
        m_oFWHTML.Image("", gk_sImagesPath & IIf(t_eCurrentTab = ETabType.etTabContext, gk_sImageTabSelectedContext, gk_sImageTabContext)), _
        gk_sMainFrame, , GetToolTip(ETOOLTIP.ttViewContextDetails))
        
    'summary
    m_oFWTable.AddCellAsHTML _
        m_oFWHTML.Hyperlink(IIf(t_bViewMode And t_eCurrentTab = ETabType.etTabSummary, "", gk_sLinkDisplayViewEditorialCatalogueDetails & sParamString & ETabType.etTabSummary), _
        m_oFWHTML.Image("", gk_sImagesPath & IIf(t_eCurrentTab = ETabType.etTabSummary, gk_sImageTabSelectedSummary, gk_sImageTabSummary)), _
        gk_sMainFrame, , GetToolTip(ETOOLTIP.ttViewSummaryDetails))
        
    'Access
    m_oFWTable.AddCellAsHTML _
        m_oFWHTML.Hyperlink(IIf(t_bViewMode And t_eCurrentTab = ETabType.etTabAccess, "", gk_sLinkDisplayViewEditorialCatalogueDetails & sParamString & ETabType.etTabAccess), _
        m_oFWHTML.Image("", gk_sImagesPath & IIf(t_eCurrentTab = ETabType.etTabAccess, gk_sImageTabSelectedAccess, gk_sImageTabAccess)), _
        gk_sMainFrame, , GetToolTip(ETOOLTIP.ttViewAccessDetails))
    
    'Content
    m_oFWTable.AddCellAsHTML _
        m_oFWHTML.Hyperlink(IIf(t_bViewMode And t_eCurrentTab = ETabType.etTabContent, "", gk_sLinkDisplayViewEditorialCatalogueDetails & sParamString & ETabType.etTabContent), _
        m_oFWHTML.Image("", gk_sImagesPath & IIf(t_eCurrentTab = ETabType.etTabContent, gk_sImageTabSelectedContent, gk_sImageTabContent)), _
        gk_sMainFrame, , GetToolTip(ETOOLTIP.ttViewContentDetails))
    
    'Admin history
    m_oFWTable.AddCellAsHTML _
        m_oFWHTML.Hyperlink(IIf(t_bViewMode And t_eCurrentTab = ETabType.etTabAdminHistory, "", gk_sLinkDisplayViewEditorialCatalogueDetails & sParamString & ETabType.etTabAdminHistory), _
        m_oFWHTML.Image("", gk_sImagesPath & IIf(t_eCurrentTab = ETabType.etTabAdminHistory, gk_sImageTabSelectedAdminHistory, gk_sImageTabAdminHistory)), _
        gk_sMainFrame, , GetToolTip(ETOOLTIP.ttViewAdminhistoryDetails))
    
    'Index terms
    m_oFWTable.AddCellAsHTML _
        m_oFWHTML.Hyperlink(IIf(t_bViewMode And t_eCurrentTab = ETabType.etTabIndexTerms, "", gk_sLinkDisplayViewEditorialCatalogueDetails & sParamString & ETabType.etTabIndexTerms), _
        m_oFWHTML.Image("", gk_sImagesPath & IIf(t_eCurrentTab = ETabType.etTabIndexTerms, gk_sImageTabSelectedIndexTerms, gk_sImageTabIndexTerms)), _
        gk_sMainFrame, , GetToolTip(ETOOLTIP.ttViewIndexTermDetails))
    
    m_oFWTable.EndRow
    
    'Security check for editing catalogues
    
    m_oBOEditorialPermissions.UserID = m_lLoggedUserID
    m_oBOEditorialPermissions.EditsetID = t_lEditSetID
    
    'init
    bCanEditCatalogue = True
    
    ' d74 If the item is in a delete edit set then it can not be edited
    If t_bDeleteEditSet = True Then
        bCanEditCatalogue = False
    End If
    
    If Not t_bViewMode Then 'editing
        bCanEditCatalogue = False
    End If
    
    'no editing for audit trail
    If bCanEditCatalogue Then
        If t_eSectionType = SECAuditTrial Then
            bCanEditCatalogue = False
        End If
    End If
    
    'user rights
    If bCanEditCatalogue Then
        If Not m_oBOEditorialPermissions.Check(EUserCanEditCatalogueEntryOnEditSet) = _
            ESecurityCheckResult.ESRSecurityCheckSuccessful Then
            
            bCanEditCatalogue = False
        End If
    End If
    
    'For the edit/move button
    m_oFWTable.StartRow
    For iCounter = 1 To 6
        If iCounter = t_eCurrentTab Then
            If t_eCurrentTab = ETabType.etTabContext Then    'Move (for context)
                '   d32 dtm     09Feb2000       Fixed defect in DisplayHeaderTab : IF A AND B OR C was supposed to be IF A AND (B OR C) but evaluated to IF (A AND B) OR C
                If bCanEditCatalogue And _
                    (t_eCatalogueLevelID = ECATALOGUELEVEL.ctClass Or t_eCatalogueLevelID = ECATALOGUELEVEL.ctPiece) Then
                    
                    m_oFWTable.AddCellAsHTML _
                        m_oFWHTML.Hyperlink(gk_sLinkStartMoveCatalogueEntry & sParamString, _
                        m_oFWHTML.Image("", gk_sImagesPath & gk_sImageMoveButton), _
                        gk_sMainFrame, , GetToolTip(ETOOLTIP.ttMoveCatalogue))
                Else    'no link
                    m_oFWTable.AddCellAsHTML _
                        m_oFWHTML.Hyperlink("", _
                        m_oFWHTML.Image("", gk_sImagesPath & gk_sImageMoveButton), _
                        gk_sMainFrame)
                End If
            Else
                If bCanEditCatalogue Then
                    'Edit button
                    m_oFWTable.AddCellAsHTML _
                        m_oFWHTML.Hyperlink(gk_sLinkDisplayEditCatalogueDetailsScreen & sParamString & t_eCurrentTab, _
                        m_oFWHTML.Image("", gk_sImagesPath & gk_sImageEdit), _
                        gk_sMainFrame, , GetToolTip(ETOOLTIP.ttEditCatalogue))
                Else    'No hyperlink on the edit button
                    m_oFWTable.AddCellAsHTML _
                        m_oFWHTML.Hyperlink("", _
                        m_oFWHTML.Image("", gk_sImagesPath & gk_sImageEditDisabled), _
                        gk_sMainFrame)
                
                End If
            End If
        Else
            m_oFWTable.AddCellAsHTML gk_sHTMLNonBreakSpace
        End If
    Next
    
    m_sHtml = m_sHtml & m_oFWHTML.Paragraph(m_oFWTable.GetHTML)
    
    m_aspResponse.Write m_sHtml
    
    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub
Public Sub DisplayEditCatalogueDetailsScreen()
'
'Description
'
'Retrieves the editsetid,catalogueid,cataloguelevelid,tabid  and calls
'appropriate procedure to display the details for that tab
'
'Amendment history
'   d1 ndk     25Aug1999       Created

    On Error GoTo ErrorHandler

    'Variables
    Dim sSource As String
    Dim lEditsetID As Long, lCatalogueID As Long, eCatalogueLevelID As ECATALOGUELEVEL
    Dim eDetailsTabNumber As ETabType, eLoadType As ECatalogueEditFormLoadType, lLinkTermID As Long
    Dim ELinkTermType As ELinkTermType, iMainCollectionKey As Integer, iLinkTermLevelID As Integer
    Dim eEntryEditType As ECatalogueEntryEditType, lParentCatalogueID As Long, iParentLevelID As Integer
    Dim lCatalogueToBeCopiedFromID As Long, iCatalogueLevelToBeCopiedFromID As Integer
    
    sSource = Me.Name & ".DisplayEditCatalogueDetailsScreen"
    
    'Get the details from the request object
    lEditsetID = m_aspRequest.QueryString(gk_sParamEditSetID)
    lCatalogueID = m_aspRequest.QueryString(gk_sParamCatalogueID)
    eCatalogueLevelID = m_aspRequest.QueryString(gk_sParamCatalogueLevel)     'Level of the catalogue like 1- lettercode, 2 -division
    eDetailsTabNumber = m_aspRequest.QueryString(gk_sParamDetailsTabNumber) ''tab number like 1 - Context, 2 - Summary
    eLoadType = m_aspRequest.QueryString(ksParamLoadType)
    eEntryEditType = m_aspRequest.QueryString(ksParamCatalogueEntryEditType)

    'Link term details
    ELinkTermType = m_aspRequest.QueryString(ksParamLinkTermType)
    iLinkTermLevelID = m_aspRequest.QueryString(ksParamLinkTermLevelID)
    lLinkTermID = m_aspRequest.QueryString(ksParamLinkTermID)
    iMainCollectionKey = m_aspRequest.QueryString(ksParamMainCollectionKey)
    
    'Catalogue context details
    lParentCatalogueID = m_aspRequest.QueryString(ksParamParentCatalogueID)
    iParentLevelID = m_aspRequest.QueryString(ksParamParentCatalogueLevel)
    lCatalogueToBeCopiedFromID = m_aspRequest.QueryString(ksParamCatalogueToBeCopiedID)
    iCatalogueLevelToBeCopiedFromID = m_aspRequest.QueryString(ksParamCatalogueToBeCopiedLevel)
    
    If eEntryEditType = ECatalogueEntryEditType.ecEditCatalogueEntry Then
        'Editing a catalogue entry
        If eLoadType = ECatalogueEditFormLoadType.etLoadFromDatabaseEditCatalogueEntry Then
            'This happens when the user clicks on the edit or move button
        
            'Lock the catalogue entry, if the lock fails, an error is raised
            LockCatalogueEntry eDetailsTabNumber, lCatalogueID, eCatalogueLevelID, lEditsetID
            
        End If
        
        'Display the title, covering dates etc
        DisplayHeaderDetails lEditsetID, lCatalogueID, eCatalogueLevelID, eDetailsTabNumber, ESectionType.SECeditorial
           
        'Displays the tab at the top
        DisplayHeaderTab lEditsetID, lCatalogueID, eCatalogueLevelID, eDetailsTabNumber, ESectionType.SECeditorial, False
    
        Select Case eDetailsTabNumber
            
            Case ETabType.etTabSummary      'Display Summary details tab
                    
                DisplayEditSummaryDetailsScreen eLoadType, lEditsetID, lCatalogueID, eCatalogueLevelID, lLinkTermID, ELinkTermType
            
            Case ETabType.etTabAccess       'Access Summary details tab
                        
                DisplayEditAccessDetailsScreen eLoadType, lEditsetID, lCatalogueID, eCatalogueLevelID, lLinkTermID, ELinkTermType
            
            Case ETabType.etTabContent  ' Content Details
                        
                DisplayEditContentDetailsScreen eLoadType, lEditsetID, lCatalogueID, eCatalogueLevelID, lLinkTermID, ELinkTermType, iLinkTermLevelID, iMainCollectionKey
            
            Case ETabType.etTabAdminHistory  ' Admin History
                        
                DisplayEditAdminHistoryScreen eLoadType, lEditsetID, lCatalogueID, eCatalogueLevelID
            
            Case ETabType.etTabIndexTerms  ' Index Terms
                        
                DisplayEditIndexTermDetailsScreen eLoadType, lEditsetID, lCatalogueID, eCatalogueLevelID, lLinkTermID, ELinkTermType
        
        End Select
    Else
        'Creating a new catalogue
            
        DisplayEditSummaryDetailsScreen _
            eLoadType, lEditsetID, lCatalogueID, eCatalogueLevelID, lLinkTermID, ELinkTermType, eEntryEditType, _
            lParentCatalogueID, iParentLevelID, lCatalogueToBeCopiedFromID, iCatalogueLevelToBeCopiedFromID
    
    End If
    
    Exit Sub
ErrorHandler:
    ShowError sSource
    m_objObjectContext.SetAbort
End Sub

Public Sub DisplayViewEditorialCatalogueDetails()
'
'Description
'
'Retrieves the editsetid,catalogueid,cataloguelevelid,tabid  and calls
'appropriate procedure to display the details for that tab
'
'Note : An attempt is made to unlock the catalogue here, this is because the control is
'sent here when the user clicks on the cancel button while. This is done when the parameter
'bEditingCancelled is found to be true
'
'Amendment history
'   d01 ndk     25Aug1999       Created
'   d02 ndk     16Oct1999       Added functionality to unlock a catalogue
'   d16 ndk     08Dec1999       Added help file
'   d48 daj     27Mar2000       Tidied up the audit trail and fixed defect 766 - Audit trail 'back' links not working
'                               once you move off the summary tab
'   d74 alm     24Oct2001       Determines if the editset is a delete one - as all entries will be readonly
'
    On Error GoTo ErrorHandler

    Dim sSource As String
    Dim lEditsetID As Long, lCatalogueID As Long, eDetailsTabNumber As ETabType, eCatalogueLevelID As ECATALOGUELEVEL
    Dim eProcatSectionType As ESectionType, bEditingCancelled, sHTML As String, sHelpFileName As String
    Dim VWSharedCatalogueEntry As VWSharedCatalogueEntry
    
    ' d74 - Added
    Dim oBOEditSet     As BOEditSet
    Dim bDeleteEditSet As Boolean
    
    ' d48 - Stores the editset ID of the 'current' entry if this is an audit trail version
    Dim lCurrentEditsetID As Long
    
    sSource = Me.Name & ".DisplayViewEditorialCatalogueDetails"
    
    'Get the details from the request object
    lEditsetID = m_aspRequest.QueryString(gk_sParamEditSetID)
    lCatalogueID = m_aspRequest.QueryString(gk_sParamCatalogueID)
    eCatalogueLevelID = m_aspRequest.QueryString(gk_sParamCatalogueLevel)   'Level of the catalogue like 1- lettercode, 2 -division
    eDetailsTabNumber = m_aspRequest.QueryString(gk_sParamDetailsTabNumber) 'tab number like 1 - Context, 2 - Summary
    eProcatSectionType = m_aspRequest.QueryString(gk_sParamSectionType)     'Editorial or Audit trail
    bEditingCancelled = m_aspRequest.QueryString(ksParamEditingCancelled)
    
    ' d48 - Get the 'current' edit set ID too
    lCurrentEditsetID = m_aspRequest.QueryString(gk_sParamCurrEditSetID)
    
    If bEditingCancelled Then
        'The control has come here when the user clicked cancel while editing this catalogue
    
        'Release the lock
        UnLockCatalogueEntry eDetailsTabNumber, lCatalogueID, eCatalogueLevelID, lEditsetID

    End If
    
    ' d74 Determine if the editset is a delete edit set
    ' (4 is hardcoded as the edit set types do not match the values
    ' in the enumeration ECreateEditSetUsage)
    Set oBOEditSet = New BOEditSet
    oBOEditSet.EditsetID = lEditsetID
    oBOEditSet.Load
    
    If oBOEditSet.EditSetTypeID = 4 Then
        bDeleteEditSet = True
    Else
        bDeleteEditSet = False
    End If
    
    'Details of the catalogue (title,level etc)
    DisplayHeaderDetails lEditsetID, lCatalogueID, eCatalogueLevelID, eDetailsTabNumber, eProcatSectionType, lCurrentEditsetID
    
    'Displays the tab at the top
    DisplayHeaderTab lEditsetID, lCatalogueID, eCatalogueLevelID, eDetailsTabNumber, eProcatSectionType, True, lCurrentEditsetID, bDeleteEditSet
    
    'Instantiate the object
    Set VWSharedCatalogueEntry = New VWSharedCatalogueEntry
    
    'Set the properties
    VWSharedCatalogueEntry.TabType = eDetailsTabNumber
    VWSharedCatalogueEntry.CatalogueID = lCatalogueID
    VWSharedCatalogueEntry.CatalogueLevel = eCatalogueLevelID
    VWSharedCatalogueEntry.SectionType = eProcatSectionType
    VWSharedCatalogueEntry.EditsetID = lEditsetID
    
    'Call the method to display the tab
    VWSharedCatalogueEntry.DisplayViewCatalogueDetails
        
    'get the name of the help file
    Select Case eDetailsTabNumber
        Case ETabType.etTabAccess
            sHelpFileName = gk_sHelpLinkToViewAccessDetails
        Case ETabType.etTabAdminHistory
            sHelpFileName = gk_sHelpLinkToViewAdminHistoryDetails
        Case ETabType.etTabContent
            sHelpFileName = gk_sHelpLinkToViewContentDetails
        Case ETabType.etTabContext
            sHelpFileName = gk_sHelpLinkToViewAndEditContextDetails
        Case ETabType.etTabIndexTerms
            sHelpFileName = gk_sHelpLinkToViewIndexTermsDetails
        Case ETabType.etTabSummary
            sHelpFileName = gk_sHelpLinkToViewSummaryDetails
    End Select
    
    'Add the help button
    sHTML = sHTML & HTMLToAddHelpButton(sHelpFileName)
    
    'write the html
    m_aspResponse.Write sHTML
    
    'Release the object
    Set VWSharedCatalogueEntry = Nothing
    Set oBOEditSet = Nothing
    
    m_objObjectContext.SetComplete
    
    Exit Sub
ErrorHandler:
    ShowError sSource
    m_objObjectContext.SetAbort
End Sub

Public Sub SubmitCatalogueDetails()
'
'Description : Submits the catalogue details, calls routines to save details in
' database or session
'
'Amendment history
'   d01 ndk     25Aug1999       Created
'   d15 ndk     07Dec1999       Implemented EAD Editing and made related changes
'   d29 crm     02Feb2000       Updated code to provide links for Leaflets to catalogues

    On Error GoTo ErrorHandler
    
    Dim sSource  As String
    
    sSource = Me.Name & ".DisplayViewEditorialCatalogueDetails"
    
    'EAD updated and saved
    
    If m_aspRequest.Form(ksNameHiddenEADUpdatedScopeAndContent) = True Or _
        m_aspRequest.Form(ksNameHiddenEADUpdatedArrangement) = True Or _
        m_aspRequest.Form(ksNameHiddenEADUpdatedAdminHistory) = True Or _
        m_aspRequest.Form(ksNameHiddenEADUpdatedTitle) = True Or _
        m_aspRequest.Form(ksNameHiddenEADUpdatedMapDesignation) = True Then
        
        ReDisplayEditCatalogueDetailsWithUpdatedEAD
    End If
    
    'Ok Button to save
    
    If (Not IsEmpty(m_aspRequest.Form.Item(ksNameButtonOk & gk_sXCoordinate))) Or _
        (Not IsEmpty(m_aspRequest.Form.Item(ksNameButtonOk & gk_sYCoordinate))) Then
        
        'Save the catalogue details in database
        UpdateCatalogueDetails

    'Normal Add Term Buttons
    
    ElseIf (Not IsEmpty(m_aspRequest.Form.Item(ksNameButtonAddTermImmSrcPerson & gk_sXCoordinate))) Or _
        (Not IsEmpty(m_aspRequest.Form.Item(ksNameButtonAddTermImmSrcPerson & gk_sYCoordinate))) Then
        
        StartAddLinkTerms ELinkTermType.elImmedieteSourceOfAcquisitionPerson
    
    ElseIf (Not IsEmpty(m_aspRequest.Form.Item(ksNameButtonAddTermImmSrcCorpBody & gk_sXCoordinate))) Or _
        (Not IsEmpty(m_aspRequest.Form.Item(ksNameButtonAddTermImmSrcCorpBody & gk_sYCoordinate))) Then
        
        StartAddLinkTerms ELinkTermType.elImmedieteSourceOfAcquisitionCorporateBody
        
    ElseIf (Not IsEmpty(m_aspRequest.Form.Item(ksNameButtonAddTermCopiesInformation & gk_sXCoordinate))) Or _
        (Not IsEmpty(m_aspRequest.Form.Item(ksNameButtonAddTermCopiesInformation & gk_sYCoordinate))) Then
        
        StartAddLinkTerms ELinkTermType.elCopiesInformation
        
    ElseIf (Not IsEmpty(m_aspRequest.Form.Item(ksNameButtonAddTermLocationOfOriginals & gk_sXCoordinate))) Or _
        (Not IsEmpty(m_aspRequest.Form.Item(ksNameButtonAddTermLocationOfOriginals & gk_sYCoordinate))) Then
        
        StartAddLinkTerms ELinkTermType.elLocationOfOriginals
        
    ElseIf (Not IsEmpty(m_aspRequest.Form.Item(ksNameButtonAddTermCreatorNamesPerson & gk_sXCoordinate))) Or _
        (Not IsEmpty(m_aspRequest.Form.Item(ksNameButtonAddTermCreatorNamesPerson & gk_sYCoordinate))) Then
        
        StartAddLinkTerms ELinkTermType.elCreatorNamesPerson
        
    ElseIf (Not IsEmpty(m_aspRequest.Form.Item(ksNameButtonAddTermCreatorNamesCorpBody & gk_sXCoordinate))) Or _
        (Not IsEmpty(m_aspRequest.Form.Item(ksNameButtonAddTermCreatorNamesCorpBody & gk_sYCoordinate))) Then
        
        StartAddLinkTerms ELinkTermType.elCreatorNamesCorporateBody
        
    ElseIf (Not IsEmpty(m_aspRequest.Form.Item(ksNameButtonAddTermPlaceOfDeposit & gk_sXCoordinate))) Or _
        (Not IsEmpty(m_aspRequest.Form.Item(ksNameButtonAddTermPlaceOfDeposit & gk_sYCoordinate))) Then
        
        StartAddLinkTerms ELinkTermType.elPlaceOfDeposit

    ElseIf (Not IsEmpty(m_aspRequest.Form.Item(ksNameButtonAddTermPublicationNote & gk_sXCoordinate))) Or _
        (Not IsEmpty(m_aspRequest.Form.Item(ksNameButtonAddTermPublicationNote & gk_sYCoordinate))) Then
        
        StartAddLinkTerms ELinkTermType.elPublicationNote
    
    ElseIf (Not IsEmpty(m_aspRequest.Form.Item(ksNameButtonAddTermUnPublishedFindingAids & gk_sXCoordinate))) Or _
        (Not IsEmpty(m_aspRequest.Form.Item(ksNameButtonAddTermUnPublishedFindingAids & gk_sYCoordinate))) Then
        
        StartAddLinkTerms ELinkTermType.elUnPublishedFindingAids
    
    ElseIf (Not IsEmpty(m_aspRequest.Form.Item(ksNameButtonAddTermIndexTermCorporateBody & gk_sXCoordinate))) Or _
        (Not IsEmpty(m_aspRequest.Form.Item(ksNameButtonAddTermIndexTermCorporateBody & gk_sYCoordinate))) Then
        
        StartAddLinkTerms ELinkTermType.elIndexTermCorporateBody
        
    ElseIf (Not IsEmpty(m_aspRequest.Form.Item(ksNameButtonAddTermIndexTermPerson & gk_sXCoordinate))) Or _
        (Not IsEmpty(m_aspRequest.Form.Item(ksNameButtonAddTermIndexTermPerson & gk_sYCoordinate))) Then
        
        StartAddLinkTerms ELinkTermType.elIndexTermPerson
        
    ElseIf (Not IsEmpty(m_aspRequest.Form.Item(ksNameButtonAddTermIndexTermPlace & gk_sXCoordinate))) Or _
        (Not IsEmpty(m_aspRequest.Form.Item(ksNameButtonAddTermIndexTermPlace & gk_sYCoordinate))) Then
        
        StartAddLinkTerms ELinkTermType.elIndexTermPlace
        
    ElseIf (Not IsEmpty(m_aspRequest.Form.Item(ksNameButtonAddTermIndexTermSubject & gk_sXCoordinate))) Or _
        (Not IsEmpty(m_aspRequest.Form.Item(ksNameButtonAddTermIndexTermSubject & gk_sYCoordinate))) Then
        
        StartAddLinkTerms ELinkTermType.elIndexTermSubject
    'the below two are links to Leaflet Catalogue entry links
    '   d29 crm     02Feb2000
    ElseIf (Not IsEmpty(m_aspRequest.Form.Item(ksNameButtonAddTermPNoteLeaflet & gk_sXCoordinate))) Or _
        (Not IsEmpty(m_aspRequest.Form.Item(ksNameButtonAddTermPNoteLeaflet & gk_sYCoordinate))) Then
        
        StartAddLinkTerms ELinkTermType.elPublicationNoteLeaflet
        
    ElseIf (Not IsEmpty(m_aspRequest.Form.Item(ksNameButtonAddTermUnPFindingAidLeaflet & gk_sXCoordinate))) Or _
        (Not IsEmpty(m_aspRequest.Form.Item(ksNameButtonAddTermUnPFindingAidLeaflet & gk_sYCoordinate))) Then
        
        StartAddLinkTerms ELinkTermType.elUnPublishedFindingAidLeaflet
        
    'Special Add Term Buttons
        
    ElseIf InStr(m_aspRequest.Form, ksNameButtonAddTermRelatedMaterial) > 0 Then
        
        StartAddLinkTerms ELinkTermType.elRelatedMaterial
    
    ElseIf InStr(m_aspRequest.Form, ksNameButtonAddTermSeparateMaterial) > 0 Then
        
        StartAddLinkTerms ELinkTermType.elSeparateMaterial
    
    'Remove Term Buttons
        
    ElseIf InStr(m_aspRequest.Form, ksNameButtonRemoveTermImmSrcPerson) > 0 Then
        
        RemoveLinkTerm ELinkTermType.elImmedieteSourceOfAcquisitionPerson
    
    ElseIf InStr(m_aspRequest.Form, ksNameButtonRemoveTermImmSrcCorpBody) > 0 Then
        
        RemoveLinkTerm ELinkTermType.elImmedieteSourceOfAcquisitionCorporateBody
    
    ElseIf InStr(m_aspRequest.Form, ksNameButtonRemoveTermCopiesInformation) > 0 Then
        
        RemoveLinkTerm ELinkTermType.elCopiesInformation
        
    ElseIf InStr(m_aspRequest.Form, ksNameButtonRemoveTermLocationOfOriginals) > 0 Then
        
        RemoveLinkTerm ELinkTermType.elLocationOfOriginals
    
    ElseIf InStr(m_aspRequest.Form, ksNameButtonRemoveTermCreatorNamesPerson) > 0 Then
        
        RemoveLinkTerm ELinkTermType.elCreatorNamesPerson
        
    ElseIf InStr(m_aspRequest.Form, ksNameButtonRemoveTermCreatorNamesCorpBody) > 0 Then
        
        RemoveLinkTerm ELinkTermType.elCreatorNamesCorporateBody
    
    ElseIf InStr(m_aspRequest.Form, ksNameButtonRemoveTermPlaceOfDeposit) > 0 Then
        
        RemoveLinkTerm ELinkTermType.elPlaceOfDeposit

    ElseIf InStr(m_aspRequest.Form, ksNameButtonRemoveTermPublicationNote) > 0 Then
        
        RemoveLinkTerm ELinkTermType.elPublicationNote

    ElseIf InStr(m_aspRequest.Form, ksNameButtonRemoveTermUnpublishedFindingAids) > 0 Then
        
        RemoveLinkTerm ELinkTermType.elUnPublishedFindingAids

    ElseIf InStr(m_aspRequest.Form, ksNameButtonRemoveTermRelatedMaterial) > 0 Then
        
        RemoveLinkTerm ELinkTermType.elRelatedMaterial
    
    ElseIf InStr(m_aspRequest.Form, ksNameButtonRemoveTermSeparateMaterial) > 0 Then
        
        RemoveLinkTerm ELinkTermType.elSeparateMaterial

    ElseIf InStr(m_aspRequest.Form, ksNameButtonRemoveTermIndexTermCorporateBody) > 0 Then
        
        RemoveLinkTerm ELinkTermType.elIndexTermCorporateBody
        
    ElseIf InStr(m_aspRequest.Form, ksNameButtonRemoveTermIndexTermPerson) > 0 Then
        
        RemoveLinkTerm ELinkTermType.elIndexTermPerson
        
    ElseIf InStr(m_aspRequest.Form, ksNameButtonRemoveTermIndexTermPlace) > 0 Then
        
        RemoveLinkTerm ELinkTermType.elIndexTermPlace
        
    ElseIf InStr(m_aspRequest.Form, ksNameButtonRemoveTermIndexTermSubject) > 0 Then
        
        RemoveLinkTerm ELinkTermType.elIndexTermSubject
   '   d29 crm     02Feb2000
    ElseIf InStr(m_aspRequest.Form, ksNameButtonRemoveTermPNoteLeaflet) > 0 Then
        
        RemoveLinkTerm ELinkTermType.elPublicationNoteLeaflet
        
    ElseIf InStr(m_aspRequest.Form, ksNameButtonRemoveTermUnPFindingAidLeaflet) > 0 Then
        
        RemoveLinkTerm ELinkTermType.elUnPublishedFindingAidLeaflet
        
    End If
    
    m_objObjectContext.SetComplete
       
    Exit Sub
ErrorHandler:
    ShowError sSource
    m_objObjectContext.SetAbort
End Sub

Private Sub RemoveLinkTerm( _
    ByVal t_eLinkTermType As ELinkTermType)

'
'Description    : removes the link term from the page displayed
'
'Amendment history
'   d1 ndk     25Aug1999       Created
    
    On Error GoTo ErrorHandler

    'Variables
    Dim sSource As String
    Dim lTermIdToBeRemoved As Long, iStartIDStart As Integer, iIDEnd  As Integer
    Dim sParamString  As String, sFormString As String
    Dim lEditsetID As Long, lCatalogueID As Long, eCatalogueLevelID As ECATALOGUELEVEL, eDetailsTabNumber As ETabType
    Dim lParentCatalogueID As Long, eParentCatalogueLevel As ECATALOGUELEVEL
    Dim lCatalogueToBeCopiedId As Long, eCatalogueCatalogueToBeCopiedLevel As ECATALOGUELEVEL
    Dim eEntryEditType As ECatalogueEntryEditType
    Dim iCounter As Integer
    Dim iMainCollectionKey As Integer

    sSource = Me.Name & ".RemoveLinkTerm"
    
    lEditsetID = CLng(m_aspRequest.QueryString(gk_sParamEditSetID))
    lCatalogueID = CLng(m_aspRequest.QueryString(gk_sParamCatalogueID))
    eCatalogueLevelID = CInt(m_aspRequest.QueryString(gk_sParamCatalogueLevel))      'Level of the catalogue like 1- lettercode, 2 -division
    eDetailsTabNumber = CLng(m_aspRequest.QueryString(gk_sParamDetailsTabNumber)) ''tab number like 1 - Context, 2 - Summary
    eEntryEditType = m_aspRequest.QueryString(ksParamCatalogueEntryEditType)
    
    'catalogue context details (while creating new catalogue)
    lParentCatalogueID = m_aspRequest.QueryString(ksParamParentCatalogueID)
    eParentCatalogueLevel = m_aspRequest.QueryString(ksParamParentCatalogueLevel)
    lCatalogueToBeCopiedId = m_aspRequest.QueryString(ksParamCatalogueToBeCopiedID)
    eCatalogueCatalogueToBeCopiedLevel = m_aspRequest.QueryString(ksParamCatalogueToBeCopiedLevel)
    
    Select Case eDetailsTabNumber
        
        Case ETabType.etTabSummary      'Persist summary details
                
            SaveSummaryDetails False, lEditsetID, lCatalogueID, eCatalogueLevelID, eEntryEditType
        
        Case ETabType.etTabAccess       'Persist access details
                
            SaveAccessDetails False, lEditsetID, lCatalogueID, eCatalogueLevelID
            
        Case ETabType.etTabContent      'Persist content details
                
            SaveContentDetails False, lEditsetID, lCatalogueID, eCatalogueLevelID
            
        Case ETabType.etTabIndexTerms   'Persist index terms details
                
            SaveIndexTermDetails False, lEditsetID, lCatalogueID, eCatalogueLevelID
            
    End Select
    
    sFormString = UCase(m_aspRequest.Form)
    
    If t_eLinkTermType = ELinkTermType.elRelatedMaterial Or t_eLinkTermType = ELinkTermType.elSeparateMaterial Then
    
        If t_eLinkTermType = ELinkTermType.elRelatedMaterial Then
            iStartIDStart = _
                InStr(sFormString, UCase(ksNameButtonRemoveTermRelatedMaterial)) + _
                Len(ksNameButtonRemoveTermRelatedMaterial)
        
        Else
            iStartIDStart = _
                InStr(sFormString, UCase(ksNameButtonRemoveTermSeparateMaterial)) + _
                Len(ksNameButtonRemoveTermSeparateMaterial)
        End If
        
        iIDEnd = InStr(iStartIDStart, sFormString, UCase(gk_sXCoordinate))
        iMainCollectionKey = CLng(Mid(sFormString, iStartIDStart, iIDEnd - iStartIDStart))

        sParamString = _
            "?" & gk_sParamEditSetID & "=" & lEditsetID & _
            "&" & gk_sParamCatalogueID & "=" & lCatalogueID & _
            "&" & gk_sParamCatalogueLevel & "=" & eCatalogueLevelID & _
            "&" & gk_sParamDetailsTabNumber & "=" & eDetailsTabNumber & _
            "&" & ksParamLoadType & "=" & ECatalogueEditFormLoadType.etLoadFromSavedStateAndRemoveTermEditCatalogueEntry & _
            "&" & ksParamLinkTermID & "=" & 0 & _
            "&" & ksParamLinkTermType & "=" & t_eLinkTermType & _
            "&" & ksParamMainCollectionKey & "=" & iMainCollectionKey & _
            "&" & ksParamCatalogueEntryEditType & "=" & eEntryEditType
    Else
        'All other link terms (af and sfa)
        
        Select Case t_eLinkTermType
            Case ELinkTermType.elCreatorNamesPerson
                iStartIDStart = InStr(sFormString, UCase(ksNameButtonRemoveTermCreatorNamesPerson)) + Len(ksNameButtonRemoveTermCreatorNamesPerson)
            Case ELinkTermType.elCreatorNamesCorporateBody
                iStartIDStart = InStr(sFormString, UCase(ksNameButtonRemoveTermCreatorNamesCorpBody)) + Len(ksNameButtonRemoveTermCreatorNamesCorpBody)
            Case ELinkTermType.elPlaceOfDeposit
                iStartIDStart = InStr(sFormString, UCase(ksNameButtonRemoveTermPlaceOfDeposit)) + Len(ksNameButtonRemoveTermPlaceOfDeposit)
            Case ELinkTermType.elImmedieteSourceOfAcquisitionPerson
                iStartIDStart = InStr(sFormString, UCase(ksNameButtonRemoveTermImmSrcPerson)) + Len(ksNameButtonRemoveTermImmSrcPerson)
            Case ELinkTermType.elImmedieteSourceOfAcquisitionCorporateBody
                iStartIDStart = InStr(sFormString, UCase(ksNameButtonRemoveTermImmSrcCorpBody)) + Len(ksNameButtonRemoveTermImmSrcCorpBody)
            Case ELinkTermType.elCopiesInformation
                iStartIDStart = InStr(sFormString, UCase(ksNameButtonRemoveTermCopiesInformation)) + Len(ksNameButtonRemoveTermCopiesInformation)
            Case ELinkTermType.elLocationOfOriginals
                iStartIDStart = InStr(sFormString, UCase(ksNameButtonRemoveTermLocationOfOriginals)) + Len(ksNameButtonRemoveTermLocationOfOriginals)
            Case ELinkTermType.elPublicationNote
                iStartIDStart = InStr(sFormString, UCase(ksNameButtonRemoveTermPublicationNote)) + Len(ksNameButtonRemoveTermPublicationNote)
            Case ELinkTermType.elUnPublishedFindingAids
                iStartIDStart = InStr(sFormString, UCase(ksNameButtonRemoveTermUnpublishedFindingAids)) + Len(ksNameButtonRemoveTermUnpublishedFindingAids)
            Case ELinkTermType.elIndexTermCorporateBody
                iStartIDStart = InStr(sFormString, UCase(ksNameButtonRemoveTermIndexTermCorporateBody)) + Len(ksNameButtonRemoveTermIndexTermCorporateBody)
            Case ELinkTermType.elIndexTermPerson
                iStartIDStart = InStr(sFormString, UCase(ksNameButtonRemoveTermIndexTermPerson)) + Len(ksNameButtonRemoveTermIndexTermPerson)
            Case ELinkTermType.elIndexTermPlace
                iStartIDStart = InStr(sFormString, UCase(ksNameButtonRemoveTermIndexTermPlace)) + Len(ksNameButtonRemoveTermIndexTermPlace)
            Case ELinkTermType.elIndexTermSubject
                iStartIDStart = InStr(sFormString, UCase(ksNameButtonRemoveTermIndexTermSubject)) + Len(ksNameButtonRemoveTermIndexTermSubject)
            Case ELinkTermType.elPublicationNoteLeaflet
                iStartIDStart = InStr(sFormString, UCase(ksNameButtonRemoveTermPNoteLeaflet)) + Len(ksNameButtonRemoveTermPNoteLeaflet)
            Case ELinkTermType.elUnPublishedFindingAidLeaflet
                iStartIDStart = InStr(sFormString, UCase(ksNameButtonRemoveTermUnPFindingAidLeaflet)) + Len(ksNameButtonRemoveTermUnPFindingAidLeaflet)
        End Select
        
        iIDEnd = InStr(iStartIDStart, sFormString, UCase(gk_sXCoordinate))
        lTermIdToBeRemoved = CLng(Mid(sFormString, iStartIDStart, iIDEnd - iStartIDStart))
        
        sParamString = _
            "?" & gk_sParamEditSetID & "=" & lEditsetID & _
            "&" & gk_sParamCatalogueID & "=" & lCatalogueID & _
            "&" & gk_sParamCatalogueLevel & "=" & eCatalogueLevelID & _
            "&" & gk_sParamDetailsTabNumber & "=" & eDetailsTabNumber & _
            "&" & ksParamLoadType & "=" & ECatalogueEditFormLoadType.etLoadFromSavedStateAndRemoveTermEditCatalogueEntry & _
            "&" & ksParamLinkTermID & "=" & lTermIdToBeRemoved & _
            "&" & ksParamLinkTermType & "=" & t_eLinkTermType & _
            "&" & ksParamMainCollectionKey & "=" & iMainCollectionKey & _
            "&" & ksParamCatalogueEntryEditType & "=" & eEntryEditType & _
            "&" & ksParamParentCatalogueID & "=" & lParentCatalogueID & _
            "&" & ksParamParentCatalogueLevel & "=" & eParentCatalogueLevel & _
            "&" & ksParamCatalogueToBeCopiedID & "=" & lCatalogueToBeCopiedId & _
            "&" & ksParamCatalogueToBeCopiedLevel & "=" & eCatalogueCatalogueToBeCopiedLevel
    
    End If
    
    m_aspResponse.Redirect gk_sLinkDisplayEditCatalogueDetailsScreen & sParamString
    
    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub


Private Sub StartAddLinkTerms( _
    ByVal t_eLinkTermType As ELinkTermType)
'
'When the user clicks on 'Add' term button, the control comes here. Depending on the type of the link
'the user has choosen to add the control is sent to the appropriate page
'
'Amendment history
'   d01     ndk     25Aug1999       Created
'   d10     ndk     22Nov1999       When displaying browser for place of deposit , appropirate filter is set
'   d29     crm     02Feb2000       Updated code to provide links for Leaflets to catalogues
'   d31     crm     08Feb2000       Added QueryString parameter for PlaceOfDeposit
'   d40     tmj     29Feb2000       Added Editset parameter when passing through, as required in permissions
    On Error GoTo ErrorHandler

    'Variables
    Dim sSource As String, sParamString As String
    Dim lEditsetID As Long, lCatalogueID As Long, iCatalogueLevelID As Integer, eDetailsTabNumber As ETabType
    Dim lParentCatalogueID As Long, eParentCatalogueLevel As ECATALOGUELEVEL
    Dim lCatalogueToBeCopiedId As Long, eCatalogueCatalogueToBeCopiedLevel As ECATALOGUELEVEL
    Dim eEntryEditType As ECatalogueEntryEditType, eTermType As EAFTermType
    
    
    sSource = Me.Name & ".StartAddLinkTerms"
    
    'Extract the details from the querystirng
    lEditsetID = m_aspRequest.QueryString(gk_sParamEditSetID)
    lCatalogueID = m_aspRequest.QueryString(gk_sParamCatalogueID)
    iCatalogueLevelID = m_aspRequest.QueryString(gk_sParamCatalogueLevel)     'Level of the catalogue like 1- lettercode, 2 -division
    eDetailsTabNumber = m_aspRequest.QueryString(gk_sParamDetailsTabNumber) ''tab number like 1 - Context, 2 - Summary
    eEntryEditType = m_aspRequest.QueryString(ksParamCatalogueEntryEditType)
    
    'catalogue context details
    lParentCatalogueID = m_aspRequest.QueryString(ksParamParentCatalogueID)
    eParentCatalogueLevel = m_aspRequest.QueryString(ksParamParentCatalogueLevel)
    lCatalogueToBeCopiedId = m_aspRequest.QueryString(ksParamCatalogueToBeCopiedID)
    eCatalogueCatalogueToBeCopiedLevel = m_aspRequest.QueryString(ksParamCatalogueToBeCopiedLevel)
    
    'Store them in the session
    'basic details
    StoreDataInSession dsEditCatalogueEditSetID, lEditsetID
    StoreDataInSession dsEditCatalogueCatalogueID, lCatalogueID
    StoreDataInSession dsEditCatalogueCatalogueLevelID, iCatalogueLevelID
    StoreDataInSession dsEditCatalogueTabNumber, eDetailsTabNumber
    
    'Parent details of new catalogue (id and level)
    StoreDataInSession dlCreateCatalogueParentCatalogueID, lParentCatalogueID
    StoreDataInSession diCreateCatalogueParentCatalogueLevel, eParentCatalogueLevel
    
    'Catalogue to be copied from (when creating catalogue by copying)
    StoreDataInSession dlCreateCatalogueCatalogueToBeCopiedID, lCatalogueToBeCopiedId
    StoreDataInSession diCreateCatalogueCatalogueToBeCopiedLevel, eCatalogueCatalogueToBeCopiedLevel
    
    'special details
    StoreDataInSession dsEditCatalogueCurrentLinkTermBeingPicked, t_eLinkTermType   'link term type being picked
    StoreDataInSession dsEditCatalogueEntryEditType, eEntryEditType                 'creating or editing catalogue
    
    Select Case t_eLinkTermType
        Case ELinkTermType.elImmedieteSourceOfAcquisitionPerson         'Immediete Source Of Acquisition Person
            eTermType = EAFTermType.AFPerson
        Case ELinkTermType.elImmedieteSourceOfAcquisitionCorporateBody  'Immediete Source Of Acquisition Corp Body
            eTermType = EAFTermType.AFCorporateBody
        Case ELinkTermType.elCopiesInformation                          'Copies Information
            eTermType = EAFTermType.AFCorporateBody
        Case ELinkTermType.elLocationOfOriginals                        'Location Of Originals
            eTermType = EAFTermType.AFCorporateBody
        Case ELinkTermType.elCreatorNamesPerson                         'Creator Names Person
            eTermType = EAFTermType.AFPerson
        Case ELinkTermType.elCreatorNamesCorporateBody                  'Creator Names Corporate Body
            eTermType = EAFTermType.AFCorporateBody
        Case ELinkTermType.elPlaceOfDeposit                             'Place Of Deposit
            eTermType = EAFTermType.AFCorporateBody
        Case ELinkTermType.elIndexTermCorporateBody                     'IndexTerm Corporate Body
            eTermType = EAFTermType.AFCorporateBody
        Case ELinkTermType.elIndexTermPerson                            'IndexTerm Person
            eTermType = EAFTermType.AFPerson
        Case ELinkTermType.elIndexTermPlace                             'IndexTerm Place
            eTermType = EAFTermType.AFPlace
        Case ELinkTermType.elIndexTermSubject                           'IndexTerm Subject
            eTermType = EAFTermType.AFSubject
    End Select
    
    Select Case eDetailsTabNumber
        
        Case ETabType.etTabSummary      'SaveSummary details in persistence
                
            SaveSummaryDetails _
                False, lEditsetID, lCatalogueID, iCatalogueLevelID, eEntryEditType, _
                lParentCatalogueID, eParentCatalogueLevel
        
        Case ETabType.etTabAccess       'Save Access details in persistence
                
            SaveAccessDetails False, lEditsetID, lCatalogueID, iCatalogueLevelID
        
        Case ETabType.etTabContent      'Save content details in persistence
        
            SaveContentDetails False, lEditsetID, lCatalogueID, iCatalogueLevelID
        
        Case ETabType.etTabIndexTerms   'Save Index term details in persistence
                
            SaveIndexTermDetails False, lEditsetID, lCatalogueID, iCatalogueLevelID
            
    End Select
    
    If t_eLinkTermType = ELinkTermType.elPublicationNote Or t_eLinkTermType = ELinkTermType.elUnPublishedFindingAids Then
        
        'For SFAS
        
        sParamString = _
            "?" & gk_sFORMSFAListStartIndex & "=0" & _
            "&" & gk_sFORMSFAListUsage & "=" & ESFABrowserUsage.ESFAEditorialPickTerm & _
            "&" & gk_sFORMSFAListFilter & "=" & ESFAFilter.ESFAElectronicAndPrinted
    
        m_aspResponse.Redirect gk_sLinkSFABrowserForEdiorialPickTerm & sParamString
    
    ElseIf t_eLinkTermType = ELinkTermType.elRelatedMaterial Or t_eLinkTermType = ELinkTermType.elSeparateMaterial Then
        
        'For Related and Separate Material
        
        sParamString = _
            gk_sParamCatalogueID & "=" & lCatalogueID & "&" & _
            gk_sParamCatalogueLevel & "=" & iCatalogueLevelID
        
        m_aspResponse.Redirect _
            gk_sLinkDisplayMenuForRelatedSeparatedMaterial & "?" & sParamString
    '   d29 crm     02Feb2000
    ElseIf t_eLinkTermType = ELinkTermType.elPublicationNoteLeaflet Or t_eLinkTermType = ELinkTermType.elUnPublishedFindingAidLeaflet Then
        sParamString = "eUsage=" & ELeafletBrowserUsage.ELeafletCatEdit
        m_aspResponse.Redirect gk_sInformationLeafletBrowser & "?" & sParamString
        'gk_sInformationLeafletBrowser
    Else
        'For AF Terms
        
        sParamString = _
            "?" & gk_sFORMAFGroup & "=" & eTermType & _
            "&" & gk_sFORMAFListIndex & "=" & "0" & _
            "&" & gk_sParamEditSetID & "=" & CStr(lEditsetID) 'd40
            If t_eLinkTermType = ELinkTermType.elPlaceOfDeposit Then
                sParamString = sParamString & _
                    "&" & gk_sFORMAFListType & "=" & EAFStatusToDisplay.AFSTATUSPlaceOfDepositAll & _
                    "&" & gk_sDisablePlaceOfDeposit & "=" & "TRUE"  '   d22     crm
                    'the above constant is needed to disable and check the place of deposit
                    'checkbox when creating a new Index term thru the index browser
                    'check for this when displaying form for a new AF term.
            Else
                sParamString = sParamString & _
                    "&" & gk_sFORMAFListType & "=" & EAFStatusToDisplay.AFSTATUSAll
            End If
            sParamString = sParamString & _
                "&" & gk_sFORMAFListUsage & "=" & EAFBrowserUsage.BROWSEEditorialPickTerm
        
        m_aspResponse.Redirect gk_sPAGEAFList & sParamString
    End If
    
    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub
Public Sub CancelAddLinkTerm()
'
'Description :
'
'When the user clicks on cancel button when adding a link term the control comes here
'
'Amendment history
'   d3  ndk     28Oct1999   Created

    On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim sParamString As String
    Dim lEditsetID As Long, lCatalogueID As Long, eCatalogueLevelID As ECATALOGUELEVEL, eDetailsTabNumber As ETabType
    Dim eEntryEditType  As ECatalogueEntryEditType
    Dim lParentCatalogueID As Long, eParentCatalogueLevel As ECATALOGUELEVEL
    Dim lCatalogueToBeCopiedId As Long, eCatalogueCatalogueToBeCopiedLevel As ECATALOGUELEVEL
    
    sSource = Me.Name & ".CancelAddLinkTerm"
    
    'Get the details from the session store
    'Basic details
    lEditsetID = GetDataFromSession(dsEditCatalogueEditSetID, m_objObjectContext.Item(gk_sSessionObject))
    lCatalogueID = GetDataFromSession(dsEditCatalogueCatalogueID, m_objObjectContext.Item(gk_sSessionObject))
    eCatalogueLevelID = GetDataFromSession(dsEditCatalogueCatalogueLevelID, m_objObjectContext.Item(gk_sSessionObject))  'Level of the catalogue like 1- lettercode, 2 -division
    
    'Special details
    eDetailsTabNumber = GetDataFromSession(dsEditCatalogueTabNumber, m_objObjectContext.Item(gk_sSessionObject)) ''tab number like 1 - Context, 2 - Summary
    eEntryEditType = GetDataFromSession(dsEditCatalogueEntryEditType, m_objObjectContext.Item(gk_sSessionObject))   'Editing or creating
    
    'Parent details of new catalogue (id and level)
    lParentCatalogueID = GetDataFromSession(dlCreateCatalogueParentCatalogueID)
    eParentCatalogueLevel = GetDataFromSession(diCreateCatalogueParentCatalogueLevel)

    'Catalogue to be copied from (when creating catalogue by copying)
    lCatalogueToBeCopiedId = GetDataFromSession(dlCreateCatalogueCatalogueToBeCopiedID)
    eCatalogueCatalogueToBeCopiedLevel = GetDataFromSession(diCreateCatalogueCatalogueToBeCopiedLevel)
    
    sParamString = _
        "?" & gk_sParamEditSetID & "=" & lEditsetID & _
        "&" & gk_sParamCatalogueID & "=" & lCatalogueID & _
        "&" & gk_sParamCatalogueLevel & "=" & eCatalogueLevelID & _
        "&" & gk_sParamDetailsTabNumber & "=" & eDetailsTabNumber & _
        "&" & ksParamLoadType & "=" & ECatalogueEditFormLoadType.etLoadFromSavedStateNoChangeEditCatalogueEntry & _
        "&" & ksParamCatalogueEntryEditType & "=" & eEntryEditType & _
        "&" & ksParamParentCatalogueID & "=" & lParentCatalogueID & _
        "&" & ksParamParentCatalogueLevel & "=" & eParentCatalogueLevel & _
        "&" & ksParamCatalogueToBeCopiedID & "=" & lCatalogueToBeCopiedId & _
        "&" & ksParamCatalogueToBeCopiedLevel & "=" & eCatalogueCatalogueToBeCopiedLevel
    
    'Redirect to the edit catalogue page
    m_aspResponse.Redirect gk_sLinkDisplayEditCatalogueDetailsScreen & sParamString
    
    m_objObjectContext.SetComplete
    
    Exit Sub
ErrorHandler:
    m_objObjectContext.SetAbort
    ShowError sSource

End Sub
Private Sub ReDisplayEditCatalogueDetailsWithUpdatedEAD()
'
'When the user clicks on 'Update' button after editing the ead, the control comes here.
'The data in the form is persisted and the form is redisplayed for editing
'
'Amendment history
'   d15 ndk     07Dec1999       Created
    
    On Error GoTo ErrorHandler

    'Variables
    Dim sSource As String, sParamString As String
    Dim lEditsetID As Long, lCatalogueID As Long, eCatalogueLevelID As ECATALOGUELEVEL, eDetailsTabNumber As ETabType
    Dim lParentCatalogueID As Long, eParentCatalogueLevel As ECATALOGUELEVEL
    Dim lCatalogueToBeCopiedId As Long, eCatalogueCatalogueToBeCopiedLevel As ECATALOGUELEVEL
    Dim eEntryEditType As ECatalogueEntryEditType, eTermType As EAFTermType
    
    sSource = Me.Name & ".ReDisplayEditCatalogueDetailsWithUpdatedEAD"
    
    'Extract the details from the querystirng
    lEditsetID = m_aspRequest.QueryString(gk_sParamEditSetID)
    lCatalogueID = m_aspRequest.QueryString(gk_sParamCatalogueID)
    eCatalogueLevelID = m_aspRequest.QueryString(gk_sParamCatalogueLevel)     'Level of the catalogue like 1- lettercode, 2 -division
    eDetailsTabNumber = m_aspRequest.QueryString(gk_sParamDetailsTabNumber) ''tab number like 1 - Context, 2 - Summary
    eEntryEditType = m_aspRequest.QueryString(ksParamCatalogueEntryEditType)
    
    'catalogue context details
    lParentCatalogueID = m_aspRequest.QueryString(ksParamParentCatalogueID)
    eParentCatalogueLevel = m_aspRequest.QueryString(ksParamParentCatalogueLevel)
    lCatalogueToBeCopiedId = m_aspRequest.QueryString(ksParamCatalogueToBeCopiedID)
    eCatalogueCatalogueToBeCopiedLevel = m_aspRequest.QueryString(ksParamCatalogueToBeCopiedLevel)
    
    'persist the details
    Select Case eDetailsTabNumber
        
        Case ETabType.etTabSummary      'Persist summary details
                
            SaveSummaryDetails False, lEditsetID, lCatalogueID, eCatalogueLevelID, eEntryEditType
        
        Case ETabType.etTabAccess       'Persist access details
                
            SaveAccessDetails False, lEditsetID, lCatalogueID, eCatalogueLevelID
            
        Case ETabType.etTabContent      'Persist content details
                
            SaveContentDetails False, lEditsetID, lCatalogueID, eCatalogueLevelID
        
        Case ETabType.etTabAdminHistory      'Persist content details
                
            SaveAdminHistoryDetails False, lEditsetID, lCatalogueID, eCatalogueLevelID
            
        Case ETabType.etTabIndexTerms   'Persist index terms details
                
            SaveIndexTermDetails False, lEditsetID, lCatalogueID, eCatalogueLevelID
            
    End Select
    
    sParamString = _
        "?" & gk_sParamEditSetID & "=" & lEditsetID & _
        "&" & gk_sParamCatalogueID & "=" & lCatalogueID & _
        "&" & gk_sParamCatalogueLevel & "=" & eCatalogueLevelID & _
        "&" & gk_sParamDetailsTabNumber & "=" & eDetailsTabNumber & _
        "&" & ksParamLoadType & "=" & ECatalogueEditFormLoadType.etLoadFromSavedStateNoChangeEditCatalogueEntry & _
        "&" & ksParamCatalogueEntryEditType & "=" & eEntryEditType & _
        "&" & ksParamParentCatalogueID & "=" & lParentCatalogueID & _
        "&" & ksParamParentCatalogueLevel & "=" & eParentCatalogueLevel & _
        "&" & ksParamCatalogueToBeCopiedID & "=" & lCatalogueToBeCopiedId & _
        "&" & ksParamCatalogueToBeCopiedLevel & "=" & eCatalogueCatalogueToBeCopiedLevel
    
    'Redirect to the edit catalogue page
    m_aspResponse.Redirect gk_sLinkDisplayEditCatalogueDetailsScreen & sParamString
    
    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub

Public Sub AddLinkTermToCatalogue()
'
'Description :
'
'When the user selects a link term to be added to the catalogue the control comes here.
'Depending on the where the user was before selecting the link the control is sent to that edit page
'
'Amendment history
'   d1 ndk     25Aug1999       Created
'   d29 crm     02Feb2000       Updated code to provide links for Leaflets to catalogues
    On Error GoTo ErrorHandler
    Dim sSource As String
    Dim eLinkTermTypeID As Long, lLinkTermID As Long, sParamString As String
    Dim lEditsetID As Long, lCatalogueID As Long, eCatalogueLevelID As Integer, eDetailsTabNumber As ETabType
    Dim iLinkTermLevelID As Integer, iMainCollectionKey As Integer, bAddingNewRow As Boolean
    Dim eEntryEditType  As ECatalogueEntryEditType
    Dim lParentCatalogueID As Long, eParentCatalogueLevel As ECATALOGUELEVEL
    Dim lCatalogueToBeCopiedId As Long, eCatalogueCatalogueToBeCopiedLevel As ECATALOGUELEVEL
    
    sSource = Me.Name & ".AddLinkTermToCatalogue"
    
    'Get the details from the session store
    'Basic details
    lEditsetID = GetDataFromSession(dsEditCatalogueEditSetID, m_objObjectContext.Item(gk_sSessionObject))
    lCatalogueID = GetDataFromSession(dsEditCatalogueCatalogueID, m_objObjectContext.Item(gk_sSessionObject))
    eCatalogueLevelID = GetDataFromSession(dsEditCatalogueCatalogueLevelID, m_objObjectContext.Item(gk_sSessionObject))  'Level of the catalogue like 1- lettercode, 2 -division
    
    'Special details
    eDetailsTabNumber = GetDataFromSession(dsEditCatalogueTabNumber, m_objObjectContext.Item(gk_sSessionObject)) ''tab number like 1 - Context, 2 - Summary
    eLinkTermTypeID = GetDataFromSession(dsEditCatalogueCurrentLinkTermBeingPicked, m_objObjectContext.Item(gk_sSessionObject))    'type of link term
    eEntryEditType = GetDataFromSession(dsEditCatalogueEntryEditType, m_objObjectContext.Item(gk_sSessionObject))   'Editing or creating
    
    'Parent details of new catalogue (id and level)
    lParentCatalogueID = GetDataFromSession(dlCreateCatalogueParentCatalogueID)
    eParentCatalogueLevel = GetDataFromSession(diCreateCatalogueParentCatalogueLevel)
    
    'Catalogue to be copied from (when creating catalogue by copying)
    lCatalogueToBeCopiedId = GetDataFromSession(dlCreateCatalogueCatalogueToBeCopiedID)
    eCatalogueCatalogueToBeCopiedLevel = GetDataFromSession(diCreateCatalogueCatalogueToBeCopiedLevel)

    
    'For SFAs
    If eLinkTermTypeID = ELinkTermType.elPublicationNote Or eLinkTermTypeID = ELinkTermType.elUnPublishedFindingAids Then
    
        lLinkTermID = m_aspRequest.QueryString(gk_sParamSFAReferenceID)
    
    'For Related and Separate Material
    ElseIf eLinkTermTypeID = ELinkTermType.elRelatedMaterial Or eLinkTermTypeID = ELinkTermType.elSeparateMaterial Then
    
        lLinkTermID = m_aspRequest.QueryString(gk_sParamCatalogueID)
        iLinkTermLevelID = m_aspRequest.QueryString(gk_sParamCatalogueLevel)
    '   d29 crm     02Feb2000
    ElseIf eLinkTermTypeID = ELinkTermType.elPublicationNoteLeaflet Or eLinkTermTypeID = ELinkTermType.elUnPublishedFindingAidLeaflet Then
        lLinkTermID = m_aspRequest.QueryString("sLeafletID")
    Else
        'For other link term , ie those dependent on Authority files
        lLinkTermID = m_aspRequest.QueryString(gk_sFORMAFTermID)
    End If
    
    
    sParamString = _
        "?" & gk_sParamEditSetID & "=" & lEditsetID & _
        "&" & gk_sParamCatalogueID & "=" & lCatalogueID & _
        "&" & gk_sParamCatalogueLevel & "=" & eCatalogueLevelID & _
        "&" & gk_sParamDetailsTabNumber & "=" & eDetailsTabNumber & _
        "&" & ksParamLoadType & "=" & ECatalogueEditFormLoadType.etLoadFromSavedStateAndAddTermEditCatalogueEntry & _
        "&" & ksParamLinkTermID & "=" & lLinkTermID & _
        "&" & ksParamLinkTermType & "=" & eLinkTermTypeID & _
        "&" & ksParamLinkTermLevelID & "=" & iLinkTermLevelID & _
        "&" & ksParamCatalogueEntryEditType & "=" & eEntryEditType & _
        "&" & ksParamParentCatalogueID & "=" & lParentCatalogueID & _
        "&" & ksParamParentCatalogueLevel & "=" & eParentCatalogueLevel & _
        "&" & ksParamCatalogueToBeCopiedID & "=" & lCatalogueToBeCopiedId & _
        "&" & ksParamCatalogueToBeCopiedLevel & "=" & eCatalogueCatalogueToBeCopiedLevel
    
    'Redirect to the edit catalogue page
    m_aspResponse.Redirect gk_sLinkDisplayEditCatalogueDetailsScreen & sParamString
    
    m_objObjectContext.SetComplete
    
    Exit Sub
ErrorHandler:
    m_objObjectContext.SetAbort
    ShowError sSource
End Sub
Private Sub UpdateCatalogueDetails()
'
'Description
'
'Retrieves the editsetid,catalogueid,cataloguelevelid,tabid  and calls
'appropriate procedure to update the details of that tab
'
'Amendment history
'   d1  ndk     25Aug1999       Created
'   d23 ndk     21Dec1999       Added security check
'
    On Error GoTo ErrorHandler

    'Variables
    Dim sSource As String
    Dim lEditsetID As Long, lCatalogueID As Long, eCatalogueLevelID As ECATALOGUELEVEL, eDetailsTabNumber As ETabType
    Dim lParentCatalogueID As Long, eParentCatalogueLevel As ECATALOGUELEVEL
    Dim lCatalogueToBeCopiedId As Long, eCatalogueLevelToBeCopiedID As ECATALOGUELEVEL
    Dim eEntryEditType As ECatalogueEntryEditType
    Dim sParamString As String
    Dim eSecurityCheckStatus As ESecurityCheckResult
    
    sSource = Me.Name & ".UpdateCatalogueDetails"
    
    'Get the details from the request object
    lEditsetID = m_aspRequest.QueryString(gk_sParamEditSetID)
    lCatalogueID = m_aspRequest.QueryString(gk_sParamCatalogueID)
    eCatalogueLevelID = m_aspRequest.QueryString(gk_sParamCatalogueLevel)     'Level of the catalogue like 1- lettercode, 2 -division
    eDetailsTabNumber = m_aspRequest.QueryString(gk_sParamDetailsTabNumber)   'tab number like 1 - Context, 2 - Summary
    eEntryEditType = m_aspRequest.QueryString(ksParamCatalogueEntryEditType)
    lParentCatalogueID = m_aspRequest.QueryString(ksParamParentCatalogueID)
    eParentCatalogueLevel = m_aspRequest.QueryString(ksParamParentCatalogueLevel)
    lCatalogueToBeCopiedId = m_aspRequest.QueryString(ksParamCatalogueToBeCopiedID)
    eCatalogueLevelToBeCopiedID = m_aspRequest.QueryString(ksParamCatalogueToBeCopiedLevel)
    
    'perform security check
    m_oBOEditorialPermissions.EditsetID = lEditsetID
    m_oBOEditorialPermissions.UserID = m_lLoggedUserID
    
    If eEntryEditType = ecEditCatalogueEntry Then   'editing
        eSecurityCheckStatus = m_oBOEditorialPermissions.Check(EUserCanEditCatalogueEntryOnEditSet)
    Else    'creating
        eSecurityCheckStatus = m_oBOEditorialPermissions.Check(EUserCanCreateCatalogueEntryOnEditSet)
    End If
    
    If eSecurityCheckStatus <> ESecurityCheckResult.ESRSecurityCheckSuccessful Then
        RaiseError sSource, _
               GetErrorEnumForSecurityCheckResult(eSecurityCheckStatus), _
               GetDescriptionForSecurityCheckError( _
                        GetErrorEnumForSecurityCheckResult(eSecurityCheckStatus))
    End If
    
    Select Case eDetailsTabNumber
        
        Case ETabType.etTabSummary      'Update Summary details tab
                
            SaveSummaryDetails True, _
                lEditsetID, lCatalogueID, eCatalogueLevelID, eEntryEditType, _
                lParentCatalogueID, eParentCatalogueLevel, lCatalogueToBeCopiedId, eCatalogueLevelToBeCopiedID
            
        Case ETabType.etTabAccess       'Update Access details tab
                
            SaveAccessDetails True, lEditsetID, lCatalogueID, eCatalogueLevelID
        
        Case ETabType.etTabContent      'Update Content details tab
                
            SaveContentDetails True, lEditsetID, lCatalogueID, eCatalogueLevelID
        
        Case ETabType.etTabAdminHistory 'Update Admin History details tab
                
            SaveAdminHistoryDetails True, lEditsetID, lCatalogueID, eCatalogueLevelID
        
        Case ETabType.etTabIndexTerms   'Update Index Term details tab
                
            SaveIndexTermDetails True, lEditsetID, lCatalogueID, eCatalogueLevelID
            
    End Select
    
    If eEntryEditType = ECatalogueEntryEditType.ecEditCatalogueEntry Then
        'Finished Editing catalogue entry
        
        'Release the lock
        UnLockCatalogueEntry eDetailsTabNumber, lCatalogueID, eCatalogueLevelID, lEditsetID
        
        sParamString = _
                "?" & gk_sParamEditSetID & "=" & lEditsetID & _
                "&" & gk_sParamCatalogueID & "=" & lCatalogueID & _
                "&" & gk_sParamCatalogueLevel & "=" & eCatalogueLevelID & _
                "&" & gk_sParamDetailsTabNumber & "=" & eDetailsTabNumber & _
                "&" & gk_sParamSectionType & "=" & ESectionType.SECeditorial
                
        'Send control to the view catalogue details page
        m_aspResponse.Redirect gk_sLinkDisplayViewEditorialCatalogueDetails & sParamString
    Else
        'Created a new catalogue entry
        sParamString = _
                "?" & gk_sParamEditSetID & "=" & lEditsetID

        m_aspResponse.Redirect _
            gk_sLinkDisplayEditSetDetails & sParamString
    End If
    
    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub


Private Sub SaveSummaryDetails( _
    ByVal t_bSaveInDatabase As Boolean, _
    ByVal t_lEditSetID As Long, _
    ByVal t_lCatalogueID As Long, _
    ByVal t_eCatalogueLevelID As ECATALOGUELEVEL, _
    ByVal t_eEntryEditType As ECatalogueEntryEditType, _
    Optional ByVal t_lParentCatalogueID As Long, _
    Optional ByVal t_eParentCatalogueLevelID As ECATALOGUELEVEL, _
    Optional ByVal t_lCatalogueToBeCopiedFromID As Long, _
    Optional ByVal t_eCatalogueLevelToBeCopiedFromID As ECATALOGUELEVEL)

'Description : Saves the summary details in the database or persistence manager
'
'Amendment history
'   d1  ndk     25Aug1999      Created
'   d13 ndk     24Nov1999      Access condition is now updated when a editset entry is created from scratch or lettercode
'                              using default value of parents value
'   d21 ndk     15Dec1999      Replace html code  with single and double quotes
'   d56 daj     04Jan2001      Made Covering Dates non mandatory when editing
'   d60 alm     13Feb2001      Saves the Start and End dates
'   d72 alm     25Sep2001       Moved the 'Notes' field from the Content view to the Summary View
'
    On Error GoTo ErrorHandler

    'Variables
    Dim sSource As String, sValueToSave As String
    
    'Business objects
    Dim oBOCatalogueSummary As BOCatalogueSummary
        
    sSource = Me.Name & ".SaveSummaryDetails"

    'Instantiate the object
    Set oBOCatalogueSummary = New BOCatalogueSummary
    
    'Type of entry (creating or editing)
    oBOCatalogueSummary.EntryEditType = t_eEntryEditType
    
    'Edit set id and section
    oBOCatalogueSummary.EditsetID = t_lEditSetID
    oBOCatalogueSummary.SectionType = ESectionType.SECeditorial
    
    If t_eEntryEditType = ECatalogueEntryEditType.ecEditCatalogueEntry Then
        'This is when saving an updated catalogue
        
        'Catalogue id and level of the catalogue being edited
        oBOCatalogueSummary.CatalogueID = t_lCatalogueID
        oBOCatalogueSummary.CatalogueLevelID = t_eCatalogueLevelID
    Else
        'When creating a new catalogue
        If t_eEntryEditType = ECatalogueEntryEditType.ecCreateCatalogueByCopying Then
            'Catalogue id and level from which details are to be copied (while creating by copying)
            oBOCatalogueSummary.CatalogueToBeCopiedFromID = t_lCatalogueToBeCopiedFromID
            oBOCatalogueSummary.CatalogueLevelToBeCopiedFromID = t_eCatalogueLevelToBeCopiedFromID
        End If
        
        'Parent id and levelof the new catalogue
        oBOCatalogueSummary.ParentCatalogueID = t_lParentCatalogueID
        oBOCatalogueSummary.ParentCatalogueLevelID = t_eParentCatalogueLevelID
        
        'Catalogue Level for new catalogue
        oBOCatalogueSummary.NewCatalogueLevelID = GetCatalogueLevelEnum(m_aspRequest.Form(ksNameCatalogueLevel))
    
        'Reference
        oBOCatalogueSummary.Reference = m_aspRequest.Form(ksNameReference)
    
    End If
    
    'Hidden access condition (dont do level checking - it wont be available when creating)
    oBOCatalogueSummary.AccessConditionID = m_aspRequest.Form(ksNameAccessConditionID)
        oBOCatalogueSummary.AccessConditionID = _
            IIf(VariantStringAsString(oBOCatalogueSummary.AccessConditionID) = "", _
            Null, oBOCatalogueSummary.AccessConditionID)
    
    'Title
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpTitle) _
        Or t_eEntryEditType <> ECatalogueEntryEditType.ecEditCatalogueEntry Then
        
        sValueToSave = VariantStringAsString(m_aspRequest.Form(ksNameTitle))
        
        'replace html code  with single and double quotes
        sValueToSave = Replace(sValueToSave, gk_sHTMLApostropheDecimalCode, "'")
        sValueToSave = Replace(sValueToSave, gk_sHTMLQuotationDecimalCode, """")
        
        oBOCatalogueSummary.Title = sValueToSave
        

    End If
                
    'Legal Status
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpLegalStatus) _
        Or t_eEntryEditType <> ECatalogueEntryEditType.ecEditCatalogueEntry Then
        oBOCatalogueSummary.LegalStatusID = m_aspRequest.Form(ksNameLegalStatusID)
        
        'Set value to null if nothing is selectd
        oBOCatalogueSummary.LegalStatusID = IIf(oBOCatalogueSummary.LegalStatusID = -1, Null, oBOCatalogueSummary.LegalStatusID)
    End If
    
    'Languages
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpLanguage) _
        Or t_eEntryEditType <> ECatalogueEntryEditType.ecEditCatalogueEntry Then
        oBOCatalogueSummary.LanguageID = m_aspRequest.Form(ksNameLanguageID)
    
        'Set value to null if nothing is selectd
        oBOCatalogueSummary.LanguageID = IIf(oBOCatalogueSummary.LanguageID = -1, Null, oBOCatalogueSummary.LanguageID)
    
    End If
    
    'Former Reference Department (Only while editing)
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpFormerDepartmentReference) Then
        oBOCatalogueSummary.FormerReferenceDepartment = m_aspRequest.Form(ksNameFormerReferenceDepartment)
    End If
    
    'Former Reference PRO (Only while editing)
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpFormerPROReference) Then
        oBOCatalogueSummary.FormerReferencePRO = m_aspRequest.Form(ksNameFormerReferencePRO)
    End If
    
    'Map designation (Only while editing)
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpMapDesignation) Then
        
        sValueToSave = m_aspRequest.Form(ksNameMapDesignation)
        
        'replace html code  with single and double quotes
        sValueToSave = Replace(sValueToSave, gk_sHTMLApostropheDecimalCode, "'")
        sValueToSave = Replace(sValueToSave, gk_sHTMLQuotationDecimalCode, """")
        
        oBOCatalogueSummary.MapDesignation = sValueToSave
        
    End If
    
    'Creator Names
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpCreatorNames) _
        Or t_eEntryEditType <> ECatalogueEntryEditType.ecEditCatalogueEntry Then
                
        'Creator Names - Person
        oBOCatalogueSummary.CreatorNamesPerson = _
            StoreFormLinkItemsInCollection(ELinkTermType.elCreatorNamesPerson)

        'Creator Names - Corporate Body
        oBOCatalogueSummary.CreatorNamesCorporateBody = _
            StoreFormLinkItemsInCollection(ELinkTermType.elCreatorNamesCorporateBody)

    End If
    
    ' Covering Dates (including Start and End dates)
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpCoveringDates) _
        Or t_eEntryEditType <> ECatalogueEntryEditType.ecEditCatalogueEntry Then
        
        'Covering Date
        oBOCatalogueSummary.CoveringDates = m_aspRequest.Form(ksNameCoveringDate)
        ' d71 Check if RegnalDate box is ticked (although probably better way of doing this!)
        If m_aspRequest.Form(ksNameIsRegnalDate) = "" Then
            oBOCatalogueSummary.IsRegnalDate = False
        Else
            oBOCatalogueSummary.IsRegnalDate = True
        End If
        
        ' d60 - Start Date
        oBOCatalogueSummary.StartDate = m_aspRequest.Form(ksNameStartDate)
        
        ' d60 - End Date
        oBOCatalogueSummary.EndDate = m_aspRequest.Form(ksNameEndDate)
        
    End If
       
    'Physical Description Form Code (Only while editing)
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty( _
        t_eCatalogueLevelID, cpPhysicalDescription_Form) Then
        
        oBOCatalogueSummary.PhysicalDescriptionFormCode = m_aspRequest.Form(ksNamePhysicalDescriptionFormCode)
    
        'Set value to null if nothing is selectd
        oBOCatalogueSummary.PhysicalDescriptionFormCode = IIf(oBOCatalogueSummary.PhysicalDescriptionFormCode = -1, Null, oBOCatalogueSummary.PhysicalDescriptionFormCode)
    
    End If
    
    'Physical Description Record Qty (Only while editing)
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty( _
        t_eCatalogueLevelID, cpPhysicalDescription_RecordQuantity) Then

        oBOCatalogueSummary.PhysicalDescriptionRecordQuantity = m_aspRequest.Form(ksNamePhysicalDescriptionRecordQty)
    End If
    
    'Physical Description Dimensions (Only while editing)
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty( _
        t_eCatalogueLevelID, cpPhysicalDescription_Dimension) Then
        
        oBOCatalogueSummary.PhysicalDescriptionDimensions = m_aspRequest.Form(ksNamePhysicalDescriptionDimensions)
    End If
    
    'Map scale number (Only while editing)
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpMapscaleNumber) Then
        oBOCatalogueSummary.MapScaleNumber = m_aspRequest.Form(ksNameMapScaleNumber)
    End If
    
    'Physical Condition (Only while editing)
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty( _
        t_eCatalogueLevelID, cpPhysicalCondition) Then
        
        oBOCatalogueSummary.PhysicalCondition = m_aspRequest.Form(ksNamePhysicalCondition)
    End If

    'Place of Deposit
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpPlaceOfDeposit) _
        Or t_eEntryEditType <> ECatalogueEntryEditType.ecEditCatalogueEntry Then
        
        'Links
        oBOCatalogueSummary.PlaceOfDeposit = _
            StoreFormLinkItemsInCollection(ELinkTermType.elPlaceOfDeposit)
        
    End If
    
    'Piece Key Order (Only while editing Piece)
    If t_eCatalogueLevelID = ECATALOGUELEVEL.ctPiece Then
        oBOCatalogueSummary.PieceKeyOrder = m_aspRequest.Form(ksNamePieceKeyOrder)
    End If

    'Item Key Order (Only while editing Item)
    If t_eCatalogueLevelID = ECATALOGUELEVEL.ctItem Then
        oBOCatalogueSummary.ItemKeyOrder = m_aspRequest.Form(ksNameItemKeyOrder)
    End If

    'd72 Note
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpNote) Then
        oBOCatalogueSummary.Notes = m_aspRequest.Form(ksNameNote)
    End If

    If t_bSaveInDatabase Then       'save in the database
        oBOCatalogueSummary.Save
    Else    'Store the persistance manager (save the state)
        oBOCatalogueSummary.PROCATTransientInterface_SetState
    End If
    
    Set oBOCatalogueSummary = Nothing
    
    Exit Sub

ErrorHandler:
    ConvertError sSource
End Sub
Private Sub SaveAccessDetails( _
    ByVal t_bSaveInDatabase As Boolean, _
    ByVal t_lEditSetID As Long, _
    ByVal t_lCatalogueID As Long, _
    ByVal t_eCatalogueLevelID As ECATALOGUELEVEL)

'Description : Saves the access details in the database or persistence manager
'
'Amendment history
'   d1 ndk     25Aug1999       Created
'   d4  ndk    04Nov1999       saving of access conditions handled differently
'   d9  ndk    20Nov1999       Closure status was not being sent properly, now corrected
'   d20 wdp    06Dec1999       Save record status flags
'   d55 map    07Nov2000       Closure code now set to Null if form returns an empty stringStringToNullOrString
'

    On Error GoTo ErrorHandler

    On Error GoTo ErrorHandler

    'Variables
    Dim sSource As String
    
    'Business objects
    Dim oBOCatalogueAccess As BOCatalogueAccess
    
    sSource = Me.Name & ".SaveAccessDetails"

    'Instantiate the object
    Set oBOCatalogueAccess = New BOCatalogueAccess
    
    'Set the properties of the catalogue to be saved
    oBOCatalogueAccess.CatalogueID = t_lCatalogueID
    oBOCatalogueAccess.CatalogueLevelID = t_eCatalogueLevelID
    oBOCatalogueAccess.EditsetID = t_lEditSetID
    oBOCatalogueAccess.SectionType = ESectionType.SECeditorial
    
    'Access Conditions
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpAccessCondition) Then
        oBOCatalogueAccess.AccessConditionID = m_aspRequest.Form(ksNameAccessConditionID)
        
        'Set value to null if nothing is selectd
        oBOCatalogueAccess.AccessConditionID = _
            IIf(oBOCatalogueAccess.AccessConditionID = -1, Null, oBOCatalogueAccess.AccessConditionID)
    End If
    
    'Closure type and Closure code
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpClosureType) And _
        m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpClosureCode) Then
    
        'Closure type
        oBOCatalogueAccess.ClosureType = m_aspRequest.Form(ksNameClosureType)
            
        'Set value to null if nothing is selectd
        oBOCatalogueAccess.ClosureType = _
            IIf(oBOCatalogueAccess.ClosureType = -1, Null, oBOCatalogueAccess.ClosureType)
            
        'Closure code
        'd55 Added StringtoNullOrString
        oBOCatalogueAccess.ClosureCode = StringToNullOrString(m_aspRequest.Form(ksNameClosureCode))
        
    End If
    
    'Record Opening date
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpRecordOpeningDate) Then
        oBOCatalogueAccess.RecordOpeningDate = m_aspRequest.Form(ksNameRecordOpeningDate)
    End If
                
    'Closure Status
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpClosureStatus) Then
    
        oBOCatalogueAccess.ClosureStatus = m_aspRequest.Form(ksNameClosureStatus)
    
        'Set value to null if nothing is selectd
        oBOCatalogueAccess.ClosureStatus = _
            IIf(oBOCatalogueAccess.ClosureStatus = -1, Null, oBOCatalogueAccess.ClosureStatus)
    
    End If
                
    'Restrictions On Use
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpRestrictionsOnUse) Then
        oBOCatalogueAccess.RestrictionsOnUse = m_aspRequest.Form(ksNameRestrictionsOnUse)
    End If
    
    'Immediete Source of acquisition
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpImmediateSourceOfAcquisition) Then
        
        'Links - Corporate body
        oBOCatalogueAccess.ImmedieteSourceOfAcquisitionCorporateBody = _
            StoreFormLinkItemsInCollection(ELinkTermType.elImmedieteSourceOfAcquisitionCorporateBody)

        'Links - Person
        oBOCatalogueAccess.ImmedieteSourceOfAcquisitionPerson = _
            StoreFormLinkItemsInCollection(ELinkTermType.elImmedieteSourceOfAcquisitionPerson)
        
        'Unknown description
        oBOCatalogueAccess.UnknownImmedieteSourceOfAcquisition = _
            m_aspRequest.Form(ksNameUnknownImmediateSourceOfAcquisition)

    End If
    
    'Copies Information
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpCopiesInformation) Then
        
        'Links
        oBOCatalogueAccess.CopiesInformation = _
            StoreFormLinkItemsInCollection(ELinkTermType.elCopiesInformation)
            
        'Unknown description
        oBOCatalogueAccess.UnknownCopiesInformation = m_aspRequest.Form(ksNameUnknownCopiesInformation)
                
    End If
    
    'Custodial History
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpCustodialHistory) Then
        oBOCatalogueAccess.CustodialHistory = m_aspRequest.Form(ksNameCustodialHistory)
    End If
    
    'Accumulation Dates
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpAccumulationDates) Then
        oBOCatalogueAccess.AccumulationDate = m_aspRequest.Form(ksNameAccumulationDates)
    End If
    
    'Appriasal Desctruction Information
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty( _
        t_eCatalogueLevelID, cpAppraisal_DestructionInformation) Then
        oBOCatalogueAccess.AppraisalDestructionInfo = m_aspRequest.Form(ksNameAppDestructionInformation)
    End If
    
    'Accruals
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpAccruals) Then
        oBOCatalogueAccess.Accruals = m_aspRequest.Form(ksNameAccruals)
    End If
    
    'Location of originals
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpLocationOfOriginals) Then
        
        'Links
        oBOCatalogueAccess.LocationOfOriginals = _
            StoreFormLinkItemsInCollection(ELinkTermType.elLocationOfOriginals)

    End If
    
    'Record status flags
    If t_eCatalogueLevelID >= ctPiece Then
    
        If m_aspRequest.Item(gk_sFORMRecordStatusAccessionedAsCopy) = mk_sValue Then
            oBOCatalogueAccess.AccessionedAsCopy = True
        Else
            oBOCatalogueAccess.AccessionedAsCopy = False
        End If
        
        If m_aspRequest.Item(gk_sFORMRecordStatusExtractAtTransfer) = mk_sValue Then
            oBOCatalogueAccess.ExtractAtTransfer = True
        Else
            oBOCatalogueAccess.ExtractAtTransfer = False
        End If
        
        If m_aspRequest.Item(gk_sFORMRecordStatusOrderable) = mk_sValue Then
            oBOCatalogueAccess.Orderable = True
        Else
            oBOCatalogueAccess.Orderable = False
        End If
        
        If m_aspRequest.Item(gk_sFORMRecordStatusAR) = mk_sValue Then
            oBOCatalogueAccess.AR = True
        Else
            oBOCatalogueAccess.AR = False
        End If
        
    End If
    
    If t_bSaveInDatabase Then   'Save in database
        oBOCatalogueAccess.Save
    Else    'Store the persistance manager (save the state)
        oBOCatalogueAccess.PROCATTransientInterface_SetState
    End If
    
    Set oBOCatalogueAccess = Nothing
    
    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub

Private Sub SaveContentDetails( _
    ByVal t_bSaveInDatabase As Boolean, _
    ByVal t_lEditSetID As Long, _
    ByVal t_lCatalogueID As Long, _
    ByVal t_eCatalogueLevelID As ECATALOGUELEVEL)

'Description : Saves the content details in the database or persistence manager
'
'Amendment history
'   d1 ndk     25Aug1999       Created
'   d21 ndk     15Dec1999      Replace html code  with single and double quotes
'   d29 crm     02Feb2000      Updated code to provide links for Leaflets to catalogues
'   d72 alm     25Sep2001      Moved the 'Notes' field from the Content view to the Summary View
'
    On Error GoTo ErrorHandler

    'Variables
    Dim sSource As String, sValueToSave As String
    
    'Business objects
    Dim oBOCatalogueContent As BOCatalogueContent
    
    sSource = Me.Name & ".SaveContentDetails"

    'Instantiate the object
    Set oBOCatalogueContent = New BOCatalogueContent
    
    'Set the properties of the catalogue to be saved
    oBOCatalogueContent.CatalogueID = t_lCatalogueID
    oBOCatalogueContent.CatalogueLevelID = t_eCatalogueLevelID
    oBOCatalogueContent.EditsetID = t_lEditSetID
    oBOCatalogueContent.SectionType = ESectionType.SECeditorial
    
    'Scope and Content
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpScopeAndContent) Then
        
        sValueToSave = m_aspRequest.Form(ksNameScopeAndContent)
        
        'replace html code  with single and double quotes
        sValueToSave = Replace(sValueToSave, gk_sHTMLApostropheDecimalCode, "'")
        sValueToSave = Replace(sValueToSave, gk_sHTMLQuotationDecimalCode, """")
        
        oBOCatalogueContent.ScopeAndContent = sValueToSave
    End If
                
    'Arrangement
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpArrangement) Then
        
        sValueToSave = m_aspRequest.Form(ksNameArrangement)
        
        'replace html code  with single and double quotes
        sValueToSave = Replace(sValueToSave, gk_sHTMLApostropheDecimalCode, "'")
        sValueToSave = Replace(sValueToSave, gk_sHTMLQuotationDecimalCode, """")
        
        oBOCatalogueContent.Arrangement = sValueToSave
        
    End If
    
    'Publication Note
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpPublicationNote) Then
        
        'Links
        oBOCatalogueContent.PublicationNote = _
            StoreFormLinkItemsInCollection(ELinkTermType.elPublicationNote)
        '   d29 crm     02Feb2000
        oBOCatalogueContent.PublicationNoteLeaflet = _
            StoreFormLinkItemsInCollection(ELinkTermType.elPublicationNoteLeaflet)
        'Unknown description
        oBOCatalogueContent.UnknownPublicationNote = m_aspRequest.Form(ksNameUnknownPublicationNote)
    End If
    
    'Unpublished Finding Aids
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpUnpublishedFindingAids) Then
        
        'Links
        oBOCatalogueContent.UnpublishedFindingAids = _
            StoreFormLinkItemsInCollection(ELinkTermType.elUnPublishedFindingAids)
        '   d29 crm     02Feb2000
        oBOCatalogueContent.UnpublishedFindingAidsLeaflet = _
            StoreFormLinkItemsInCollection(ELinkTermType.elUnPublishedFindingAidLeaflet)
        'Unknown description
        oBOCatalogueContent.UnknownUnpublishedFindingAids = m_aspRequest.Form(ksNameUnknownUnPublishedFindingAids)
        
    End If
    
    'Related Material
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpRelatedMaterials) Then
        
        'Links
        oBOCatalogueContent.RelatedMaterial = _
            StoreFormLinkItemsInCollection(ELinkTermType.elRelatedMaterial)
            
        'Unknown description
        oBOCatalogueContent.UnknownRelatedMaterial = m_aspRequest.Form(ksNameUnknownRelatedMaterial)
        
    End If
    
    'Separate Material
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpSeperatedMaterials) Then
        
        'Links
        oBOCatalogueContent.SeparateMaterial = _
            StoreFormLinkItemsInCollection(ELinkTermType.elSeparateMaterial)

        'Unknown description
        oBOCatalogueContent.UnknownSeparateMaterial = m_aspRequest.Form(ksNameUnknownSeparateMaterial)
        
    End If
    
    If t_bSaveInDatabase Then       'Save in database
        oBOCatalogueContent.Save
    Else    'Store the persistance manager (save the state)
        oBOCatalogueContent.PROCATTransientInterface_SetState
    End If
    
    Set oBOCatalogueContent = Nothing
    
    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub

Private Sub SaveAdminHistoryDetails( _
    ByVal t_bSaveInDatabase As Boolean, _
    ByVal t_lEditSetID As Long, _
    ByVal t_lCatalogueID As Long, _
    ByVal t_eCatalogueLevelID As ECATALOGUELEVEL)

'Description : Saves the admin history details in the database
'
'Amendment history
'   d1 ndk     25Aug1999       Created
'   d21 ndk     15Dec1999      Replace html code  with single and double quotes
'

    On Error GoTo ErrorHandler

    'Variables
    Dim sSource As String, sValueToSave As String
        
    'Business objects
    Dim oBOCatalogueAdminHistory As BOCatalogueAdminHistory
    
    sSource = Me.Name & ".SaveAdminHistoryDetails"

    'Instantiate the object
    Set oBOCatalogueAdminHistory = New BOCatalogueAdminHistory
    
    'Set the properties of the catalogue to be saved
    oBOCatalogueAdminHistory.CatalogueID = t_lCatalogueID
    oBOCatalogueAdminHistory.CatalogueLevelID = t_eCatalogueLevelID
    oBOCatalogueAdminHistory.EditsetID = t_lEditSetID
    oBOCatalogueAdminHistory.SectionType = ESectionType.SECeditorial
    
    'Admin History
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty( _
        t_eCatalogueLevelID, cpAdministrative_BiographicalHistory) Then
        
        sValueToSave = m_aspRequest.Form(ksNameAdminHistory)
        
        'replace html code  with single and double quotes
        sValueToSave = Replace(sValueToSave, gk_sHTMLApostropheDecimalCode, "'")
        sValueToSave = Replace(sValueToSave, gk_sHTMLQuotationDecimalCode, """")
        
        oBOCatalogueAdminHistory.AdminHistory = sValueToSave
    End If
    
    If t_bSaveInDatabase Then       'Save in the database
        oBOCatalogueAdminHistory.Save
    Else    'Store the persistance manager (save the state)
        oBOCatalogueAdminHistory.PROCATTransientInterface_SetState
    End If
    
    Set oBOCatalogueAdminHistory = Nothing
    
    Exit Sub

ErrorHandler:
    ConvertError sSource
End Sub

Private Sub SaveIndexTermDetails( _
    ByVal t_bSaveInDatabase As Boolean, _
    ByVal t_lEditSetID As Long, _
    ByVal t_lCatalogueID As Long, _
    ByVal t_eCatalogueLevelID As ECATALOGUELEVEL)

'Description : Saves the index terms details in the database or persistence manager
'
'Amendment history
'   d1 ndk     25Aug1999       Created

    On Error GoTo ErrorHandler

    'Variables
    Dim sSource As String, iCounter As Integer, sParamString  As String, sFormField As String
    
    'Business objects
    Dim oBOCatalogueIndexTerms As BOCatalogueIndexTerms
    
    sSource = Me.Name & ".SaveIndexTermDetails"

    'Instantiate the object
    Set oBOCatalogueIndexTerms = New BOCatalogueIndexTerms
    
    'Set the properties of the catalogue to be saved
    oBOCatalogueIndexTerms.CatalogueID = t_lCatalogueID
    oBOCatalogueIndexTerms.CatalogueLevelID = t_eCatalogueLevelID
    oBOCatalogueIndexTerms.EditsetID = t_lEditSetID
    oBOCatalogueIndexTerms.SectionType = ESectionType.SECeditorial
    
    'Corporate Body
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpIndexTermCorporateName) Then
        
        oBOCatalogueIndexTerms.CorporateBodyNames = _
            StoreFormLinkItemsInCollection(ELinkTermType.elIndexTermCorporateBody)
    End If
    
    'Person
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpIndexTermPersonName) Then
        
        oBOCatalogueIndexTerms.PersonalNames = _
            StoreFormLinkItemsInCollection(ELinkTermType.elIndexTermPerson)
    End If
    
    'Place
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpIndexTermPlaceName) Then
        
        oBOCatalogueIndexTerms.Places = _
            StoreFormLinkItemsInCollection(ELinkTermType.elIndexTermPlace)
    End If
    
    'Subject
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpIndexTermSubject) Then
        
        oBOCatalogueIndexTerms.Subjects = _
            StoreFormLinkItemsInCollection(ELinkTermType.elIndexTermSubject)
    End If
        
    If t_bSaveInDatabase Then       'Save in the database
        oBOCatalogueIndexTerms.Save
    Else    'Store the persistance manager (save the state)
        oBOCatalogueIndexTerms.PROCATTransientInterface_SetState
    End If
    
    Set oBOCatalogueIndexTerms = Nothing
    Exit Sub

ErrorHandler:
    ConvertError sSource
End Sub




Private Sub DisplayEditSummaryDetailsScreen( _
            ByVal t_eLoadType As ECatalogueEditFormLoadType, _
            ByVal t_lEditSetID As Long, _
            Optional ByVal t_lCatalogueID As Long, _
            Optional ByVal t_eCatalogueLevelID As ECATALOGUELEVEL, _
            Optional ByVal t_lLinkTermID As Long, _
            Optional ByVal t_eLinkTermType As ELinkTermType, _
            Optional ByVal t_eCatalogueEntryEditType As ECatalogueEntryEditType = ecEditCatalogueEntry, _
            Optional ByVal t_lParentCatalogueID As Long, _
            Optional ByVal t_eParentCatalogueLevelID As ECATALOGUELEVEL, _
            Optional ByVal t_lCatalogueToBeCopiedFromID As Long, _
            Optional ByVal t_eCatalogueLevelToBeCopiedFromID As ECATALOGUELEVEL)

    
'Description : Displays the content screen for editing and creating a catalogue
'
'Amendment history
'   d1  ndk     25Aug1999       Created
'   d3  ndk     28Oct1999       This tab now permits modification or piece key order and item_key_order
'                               while creating by copying selection of level permitted only if copied entry is an piece
'   d4  ndk     04Nov1999       Rsources for item and piece key order were put
'   d5  ndk     11Nov1999       Create by copying drop down for levels changed
'   d9  ndk     20Nov1999       For lookups using lookup list of bolookup instead of fwreference data
'   d13 ndk     24Nov1999       Access condition is now updated when a editset entry is created from scratch or lettercode
'                               using default value of parents value
'   d14 ndk     25Nov1999       Minor changes
'   d15 ndk     07Dec1999       Implemented EAD Editing and made related changes
'   d16 ndk     08Dec1999       Added help button
'   d18 ndk     09Dec1999       Changed images for pick term buttons
'   d21 ndk     15Dec1999       Replacing single and double quotes with equivalent html codes before rendering ead data
'   d22 ndk     16Dec1999       Changed the display length for reference field when creating catalogues
'   d23 ndk     21Dec1999       Added security check
'   d28 ndk     30Dec1999       Persisting catalogue level while creating
'   d35 dtm     10Feb2000       Modified so that a * is displayed when title is empty to indicate mandatoriness
'   d36 jcc     18Feb2000       and remove the -from- bit too! Also, should be a * on title all the time, above piece
'   d38 tmj     24Feb2000       Modified to handle the copying of departments
'   d44 wdp     06Mar2000       Fixed DisplayEditSummaryDetailsScreen to cope with
'                               title OR scope-content for create new entry and
'                               be able to switch between them.
'   d45 tmj     06Mar2000       Changed so Item reference part can be null
'   d52 daj     11Apr2000       Added extra situations where the summary edit screen should toggle the DTDtype
'   d56 daj     04Jan2001       Made Covering Dates non mandatory when editing
'   d60 alm     13Feb2001       Supports the editing of start / end dates
'   d61 daj     13Feb2001       Fixed bugs with enhancement done in d56
'   d64 pnp     16Feb2001       Added Next reference functionality
'   d69 map     05Mar2001       Major amendments to the next reference functionality
'   d70 map     15May2001       Keep value of generated last reference after a user edits in xmetal
'   d72 alm     25Sep2001       Moved the 'Notes' field from DisplayEditContentDetailsScreen
'
    On Error GoTo ErrorHandler

    'Variables
    Dim sSource As String, sHTML As String, sParamString As String, sTempTitle As String
    Dim oFWRecordlist As FWRecordList, oFWValidCatalogueLevels As FWRecordList
    Dim sParentReference As String, iHiddenLevelID As Integer, sLevelName As String
    Dim sHiddenAccessCondition As String, sValueForHiddenEADField As String
    Dim eSecurityCheckStatus As ESecurityCheckResult
        
    'd69
    Dim sLastRef(ECATALOGUELEVEL.ctItem + 1) As String
    Dim sRefNote(ECATALOGUELEVEL.ctItem + 1) As String
    Dim sNextReference(ECATALOGUELEVEL.ctItem + 1) As String
        
    'Business objects
    Dim oBOCatalogueSummary As BOCatalogueSummary
    Dim oBOBasicCatalogueEntry As BOBasicCatalogueEntry
    
    sSource = Me.Name & ".DisplayEditSummaryDetailsScreen"

    'perform security check
    m_oBOEditorialPermissions.EditsetID = t_lEditSetID
    m_oBOEditorialPermissions.UserID = m_lLoggedUserID
    
    If t_eCatalogueEntryEditType = ecEditCatalogueEntry Then    'editing
        eSecurityCheckStatus = m_oBOEditorialPermissions.Check(EUserCanEditCatalogueEntryOnEditSet)
    Else    'creating
        eSecurityCheckStatus = m_oBOEditorialPermissions.Check(EUserCanCreateCatalogueEntryOnEditSet)
    End If
    
    If eSecurityCheckStatus <> ESecurityCheckResult.ESRSecurityCheckSuccessful Then
        RaiseError sSource, _
               GetErrorEnumForSecurityCheckResult(eSecurityCheckStatus), _
               GetDescriptionForSecurityCheckError( _
                        GetErrorEnumForSecurityCheckResult(eSecurityCheckStatus))
    End If

    'Instantiate the object
    Set oBOCatalogueSummary = New BOCatalogueSummary
    Set oBOBasicCatalogueEntry = New BOBasicCatalogueEntry
    
    'Set the properties of the catalogue to be displayed
    If t_eLoadType = ECatalogueEditFormLoadType.etLoadFromDatabaseEditCatalogueEntry Then
        'The user clicked on Edit button and control has come here
        
        oBOCatalogueSummary.CatalogueID = t_lCatalogueID
        oBOCatalogueSummary.CatalogueLevelID = t_eCatalogueLevelID
        oBOCatalogueSummary.EditsetID = t_lEditSetID
        oBOCatalogueSummary.SectionType = ESectionType.SECeditorial
        
        'Get the details
        oBOCatalogueSummary.Load t_eLoadType
    Else
        Select Case t_eLoadType
        
            Case ECatalogueEditFormLoadType.etLoadFromSavedStateAndAddTermEditCatalogueEntry
                'The user has selected a pick term to add and the control has come here
                
                oBOCatalogueSummary.Load t_eLoadType, t_lLinkTermID, t_eLinkTermType
            
            Case ECatalogueEditFormLoadType.etLoadFromSavedStateAndRemoveTermEditCatalogueEntry
                'The user has selected a pick term to remove and the control has come here
                
                oBOCatalogueSummary.Load t_eLoadType, t_lLinkTermID, t_eLinkTermType
                
                't_lParentCatalogueID
                't_eParentCatalogueLevelID
                
            Case ECatalogueEditFormLoadType.etLoadFromSavedStateNoChangeEditCatalogueEntry
                'The user has selected cancel operation while adding/removing a link term and the control has come here
                
                oBOCatalogueSummary.Load t_eLoadType
            
            Case ECatalogueEditFormLoadType.etLoadFromNoneCreateCatalogueEntry
                'Creating a new catalogue entry
                If t_eCatalogueEntryEditType = ECatalogueEntryEditType.ecCreateCatalogueByCopying Then
                    'Create by copying
                    oBOCatalogueSummary.CatalogueToBeCopiedFromID = t_lCatalogueToBeCopiedFromID
                    oBOCatalogueSummary.CatalogueLevelToBeCopiedFromID = t_eCatalogueLevelToBeCopiedFromID
                End If
                
                If t_eCatalogueEntryEditType = ECatalogueEntryEditType.ecCreateCatalogueFromScatch Then
                    oBOCatalogueSummary.EditsetID = t_lEditSetID
                    oBOCatalogueSummary.ParentCatalogueID = t_lParentCatalogueID
                    oBOCatalogueSummary.ParentCatalogueLevelID = t_eParentCatalogueLevelID
                End If
                'Load the details
                oBOCatalogueSummary.Load t_eLoadType, , , t_eCatalogueEntryEditType
                
                
        End Select
                
    End If
    
    'To display the details
    m_oFWTable.Reset    'Reset the table details
    
    'Apply the style to the table
    m_oFWTable.TableStyle = m_typTableStyleContents
    
    'Link for submiting the form details
    sParamString = gk_sLinkSubmitCatalogueDetails & _
        "?" & gk_sParamEditSetID & "=" & t_lEditSetID & _
        "&" & gk_sParamCatalogueID & "=" & t_lCatalogueID & _
        "&" & gk_sParamCatalogueLevel & "=" & t_eCatalogueLevelID & _
        "&" & gk_sParamDetailsTabNumber & "=" & ETabType.etTabSummary & _
        "&" & ksParamCatalogueEntryEditType & "=" & t_eCatalogueEntryEditType & _
        "&" & ksParamParentCatalogueID & "=" & t_lParentCatalogueID & _
        "&" & ksParamParentCatalogueLevel & "=" & t_eParentCatalogueLevelID & _
        "&" & ksParamCatalogueToBeCopiedID & "=" & t_lCatalogueToBeCopiedFromID & _
        "&" & ksParamCatalogueToBeCopiedLevel & "=" & t_eCatalogueLevelToBeCopiedFromID
                    
    'Start the form
    sHTML = sHTML & m_oFWHTML.FormStart(sParamString, gk_sMainFrame, gk_sHTMLDefaultFormName, MPost, True)
    
    'Hidden access condition field
    sHiddenAccessCondition = _
        m_oFWHTML.HiddenField(ksNameAccessConditionID, _
            VariantStringAsString(oBOCatalogueSummary.AccessConditionID))
    
    sHTML = sHTML & sHiddenAccessCondition
    
    Dim eLevel As ECATALOGUELEVEL
        
    'Need to find out what level this entry is
    'to select a DTD to use to launch XMetal and
    'change mandatory settings for Covering Dates
    ' d61 - Fixed defects
    If oBOCatalogueSummary.NewCatalogueLevelID = ctNone And _
        t_eCatalogueEntryEditType <> ECatalogueEntryEditType.ecEditCatalogueEntry Then
        ' For create new, the level has not yet been set, default
        ' to one level below parent
        ' This may be altered at a later stage if more information
        ' is found
        eLevel = t_eParentCatalogueLevelID + 1
    ElseIf oBOCatalogueSummary.NewCatalogueLevelID = ctNone And _
        t_eCatalogueEntryEditType = ECatalogueEntryEditType.ecEditCatalogueEntry Then
        ' d56 - If we are editing, then the level is the current level
        eLevel = t_eCatalogueLevelID
    Else
        ' Otherwise, use the new catalogue level specified in the summary object
        eLevel = oBOCatalogueSummary.NewCatalogueLevelID
    End If
    
    'Creating a new catalogue (not editing)
    If t_eCatalogueEntryEditType <> ECatalogueEntryEditType.ecEditCatalogueEntry Then
        
        'Parent Reference
        'd38 added additional clause in if statement to handle copied departments
        If Not ((t_eCatalogueEntryEditType = ECatalogueEntryEditType.ecCreateLetterCodeFromScratch) _
          Or ((t_eCatalogueEntryEditType = ecCreateCatalogueByCopying) And _
                 (t_eCatalogueLevelToBeCopiedFromID = ctLetterCode))) Then
            'No parent reference while creating lettercodes
            sParentReference = _
                GenerateCatalogueReferenceFromID(t_lParentCatalogueID, t_eParentCatalogueLevelID)
            
            DisplayOneLineHeadingForContent _
                m_oFWAppData.GetDisplayInformation( _
                EEditorialResourceFile.erfDisplayCreateCatalogueFromScratchLeftHeaderParentReference)
        
            DisplayOneLineEditForContent _
                sParentReference
        End If
        
        'Reference
        DisplayOneLineHeadingForContent _
            m_oFWAppData.GetDisplayInformation( _
            EEditorialResourceFile.erfDisplayCreateCatalogueFromScratchLeftHeaderReference)
        
        'd38 added additional clause in if statement to handle copied departments
        'd64 extended if statements
        'd69 Added below if statement structure to replace previous last reference functionality code
        '    The last reference values are written into arrays so the java script can jump between the values used in the drop down box
        
        'If creating entry...
        If (t_eCatalogueEntryEditType = ecCreateCatalogueFromScatch) Then
        '   If parent is lettercode, class or header then need to find two last references...
        '       If at lettercode or header, then need last references of levels directly below
        '       Else need last references for levels directly below, or three levels below
        
            sLastRef(t_eParentCatalogueLevelID + 1) = VariantStringAsString(GetLastUsedRef(t_eParentCatalogueLevelID + 1, t_lParentCatalogueID, t_eParentCatalogueLevelID))
            
            If sLastRef(t_eParentCatalogueLevelID + 1) = "" Then
                sRefNote(t_eParentCatalogueLevelID + 1) = m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfNoRefFound)
            Else
                'Evaluate the next reference
                If IsNumeric(sLastRef(t_eParentCatalogueLevelID + 1)) Then
                    sNextReference(t_eParentCatalogueLevelID + 1) = CStr(CLng(sLastRef(t_eParentCatalogueLevelID + 1)) + 1)
                End If
                sRefNote(t_eParentCatalogueLevelID + 1) = m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfLastRefUsed) & sLastRef(t_eParentCatalogueLevelID + 1)
            End If
            
            If ((t_eParentCatalogueLevelID = ctLetterCode) Or _
                (t_eParentCatalogueLevelID = ctClass) Or _
                (t_eParentCatalogueLevelID = ctHeaders)) Then
                
                If (t_eParentCatalogueLevelID = ctClass) Then
                    sLastRef(t_eParentCatalogueLevelID + 3) = VariantStringAsString(GetLastUsedRef(t_eParentCatalogueLevelID + 3, t_lParentCatalogueID, t_eParentCatalogueLevelID))
                    If sLastRef(t_eParentCatalogueLevelID + 3) = "" Then
                        sRefNote(t_eParentCatalogueLevelID + 3) = m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfNoRefFound)
                    Else
                        'Evaluate the next reference
                        If IsNumeric(sLastRef(t_eParentCatalogueLevelID + 3)) Then
                            sNextReference(t_eParentCatalogueLevelID + 3) = CStr(CLng(sLastRef(t_eParentCatalogueLevelID + 3)) + 1)
                        End If
                        sRefNote(t_eParentCatalogueLevelID + 3) = m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfLastRefUsed) & sLastRef(t_eParentCatalogueLevelID + 3)
                    End If
                Else
                    sLastRef(t_eParentCatalogueLevelID + 2) = VariantStringAsString(GetLastUsedRef(t_eParentCatalogueLevelID + 2, t_lParentCatalogueID, t_eParentCatalogueLevelID))
                
                    If sLastRef(t_eParentCatalogueLevelID + 2) = "" Then
                        sRefNote(t_eParentCatalogueLevelID + 2) = m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfNoRefFound)
                    Else
                        'Evaluate the next reference
                        If IsNumeric(sLastRef(t_eParentCatalogueLevelID + 2)) Then
                            sNextReference(t_eParentCatalogueLevelID + 2) = CStr(CLng(sLastRef(t_eParentCatalogueLevelID + 2)) + 1)
                        End If
                        sRefNote(t_eParentCatalogueLevelID + 2) = m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfLastRefUsed) & sLastRef(t_eParentCatalogueLevelID + 2)
                    End If
                End If
            End If
        
        'Else copying entry...
        ElseIf (t_eCatalogueEntryEditType = ecCreateCatalogueByCopying) Then
            
            sLastRef(t_eCatalogueLevelToBeCopiedFromID) = VariantStringAsString(GetLastUsedRef(t_eCatalogueLevelToBeCopiedFromID, t_lParentCatalogueID, t_eParentCatalogueLevelID))
            
            If sLastRef(t_eCatalogueLevelToBeCopiedFromID) = "" Then
                sRefNote(t_eCatalogueLevelToBeCopiedFromID) = m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfNoRefFound)
            Else
                'Evaluate the next reference
                If IsNumeric(sLastRef(t_eCatalogueLevelToBeCopiedFromID)) Then
                    sNextReference(t_eCatalogueLevelToBeCopiedFromID) = CStr(CLng(sLastRef(t_eCatalogueLevelToBeCopiedFromID)) + 1)
                End If
                sRefNote(t_eCatalogueLevelToBeCopiedFromID) = m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfLastRefUsed) & sLastRef(t_eCatalogueLevelToBeCopiedFromID)
            End If
                
        End If
        
      

            If ((t_eCatalogueEntryEditType = ECatalogueEntryEditType.ecCreateLetterCodeFromScratch) _
              Or ((t_eCatalogueEntryEditType = ecCreateCatalogueByCopying) And _
                    (t_eCatalogueLevelToBeCopiedFromID = ctLetterCode))) Then
                'Creating a lettercode
                DisplayOneLineEditForContent _
                    m_oFWHTML.InputField(ITText, ksNameReference, kiColumnsForLetterCodeReference, _
                    gk_iDBCOLLenLetterCodeReference, _
                    oBOCatalogueSummary.Reference, _
                    VCNone, True)
            Else
                'd69 Added check for changing a piece into an item - if so, then copy the piece values into the item array section
                If t_eParentCatalogueLevelID = ctPiece Then
                    If (t_eCatalogueLevelToBeCopiedFromID = ctPiece) Then
                        sRefNote(ctItem) = sRefNote(ctPiece)
                        sNextReference(ctItem) = sNextReference(ctPiece)
                        sLastRef(ctItem) = sLastRef(ctPiece)
                    End If
                    ' d70 MAP If we've already got a reference (eg, from transient store) then use it
                    If Not IsEmpty(oBOCatalogueSummary.Reference) Then
                        DisplayOneLineEditForContent _
                            m_oFWHTML.InputField(ITText, ksNameReference, kiColumnsForOtherCatalogueReference, _
                            kiSizeOtherCatalogueReferences, _
                            oBOCatalogueSummary.Reference, _
                            VCNone, False) & m_oFWHTML.Span("", , gk_sSpanPrefix + ksNameReference)
                    Else
                        DisplayOneLineEditForContent _
                            m_oFWHTML.InputField(ITText, ksNameReference, kiColumnsForOtherCatalogueReference, _
                            kiSizeOtherCatalogueReferences, _
                            sNextReference(ctItem), _
                            VCNone, False) & m_oFWHTML.Span(sRefNote(ctItem), , gk_sSpanPrefix + ksNameReference)
                    End If
                Else
                    If (t_eCatalogueEntryEditType = ecCreateCatalogueFromScatch) Then
                        ' d70 MAP If we've already got a reference (eg, from transient store) then use it
                        If Not IsEmpty(oBOCatalogueSummary.Reference) Then
                            DisplayOneLineEditForContent _
                                m_oFWHTML.InputField(ITText, ksNameReference, kiColumnsForOtherCatalogueReference, _
                                kiSizeOtherCatalogueReferences, _
                                oBOCatalogueSummary.Reference, _
                                VCNone, True) & m_oFWHTML.Span("", , gk_sSpanPrefix + ksNameReference)
                        Else
                             DisplayOneLineEditForContent _
                                m_oFWHTML.InputField(ITText, ksNameReference, kiColumnsForOtherCatalogueReference, _
                                kiSizeOtherCatalogueReferences, _
                                sNextReference(t_eParentCatalogueLevelID + 1), _
                                VCNone, True) & m_oFWHTML.Span(sRefNote(t_eParentCatalogueLevelID + 1), , gk_sSpanPrefix + ksNameReference)
                        End If
                    Else
                        If (t_eCatalogueEntryEditType = ecCreateCatalogueByCopying) Then
                            ' d70 MAP If we've already got a reference (eg, from transient store) then use it
                            If Not IsEmpty(oBOCatalogueSummary.Reference) Then
                                DisplayOneLineEditForContent _
                                    m_oFWHTML.InputField(ITText, ksNameReference, kiColumnsForOtherCatalogueReference, _
                                    kiSizeOtherCatalogueReferences, _
                                    oBOCatalogueSummary.Reference, _
                                    VCNone, True) & m_oFWHTML.Span("", , gk_sSpanPrefix + ksNameReference)
                            Else
                                DisplayOneLineEditForContent _
                                    m_oFWHTML.InputField(ITText, ksNameReference, kiColumnsForOtherCatalogueReference, _
                                    kiSizeOtherCatalogueReferences, _
                                    sNextReference(t_eCatalogueLevelToBeCopiedFromID), _
                                    VCNone, True) & m_oFWHTML.Span(sRefNote(t_eCatalogueLevelToBeCopiedFromID), , gk_sSpanPrefix + ksNameReference)
                            End If
                        End If
                End If
            End If
        End If
        

        
        'Level of the catalogue being created
        
        DisplayOneLineHeadingForContent _
            m_oFWAppData.GetDisplayInformation( _
            EEditorialResourceFile.erfDisplayCreateCatalogueFromScratchLeftHeaderLevel)

        If t_eCatalogueEntryEditType = ECatalogueEntryEditType.ecCreateLetterCodeFromScratch Then
            'Creating a lettercode
            DisplayOneLineEditForContent _
                GetCatalogueLevelDescription(GetCatalogueLevelCode(ECATALOGUELEVEL.ctLetterCode)) & _
                m_oFWHTML.HiddenField(ksNameCatalogueLevel, GetCatalogueLevelCode(ECATALOGUELEVEL.ctLetterCode))
        
        ElseIf t_eCatalogueEntryEditType = ECatalogueEntryEditType.ecCreateCatalogueByCopying Then
            'source level is piece
            If t_eCatalogueLevelToBeCopiedFromID = ECATALOGUELEVEL.ctPiece Then
                If t_eParentCatalogueLevelID = ECATALOGUELEVEL.ctPiece Then
                    iHiddenLevelID = GetCatalogueLevelCode(ECATALOGUELEVEL.ctItem)
                    sLevelName = GetCatalogueLevelDescription(GetCatalogueLevelCode(ECATALOGUELEVEL.ctItem))
                Else 'if parent is SubHeaders,Headers , Class
                    iHiddenLevelID = GetCatalogueLevelCode(ECATALOGUELEVEL.ctPiece)
                    sLevelName = GetCatalogueLevelDescription(GetCatalogueLevelCode(ECATALOGUELEVEL.ctPiece))
                End If
            Else    ' for other levels the level will be same as the level of the item copied
                iHiddenLevelID = GetCatalogueLevelCode(t_eCatalogueLevelToBeCopiedFromID)
                sLevelName = GetCatalogueLevelDescription(GetCatalogueLevelCode(t_eCatalogueLevelToBeCopiedFromID))
            End If
            
            DisplayOneLineEditForContent _
                sLevelName & _
                m_oFWHTML.HiddenField(ksNameCatalogueLevel, CStr(iHiddenLevelID))
            'reset the level
            eLevel = iHiddenLevelID
            
            'Creating from scratch
        ElseIf t_eCatalogueEntryEditType = ECatalogueEntryEditType.ecCreateCatalogueFromScatch Then
            
            
            'Get the list of sub levels
            Set oFWValidCatalogueLevels = _
                oBOBasicCatalogueEntry.ChildCatalogueLevels(GetCatalogueLevelCode(t_eParentCatalogueLevelID))

            If oFWValidCatalogueLevels.RecordCount = 1 Then
                DisplayOneLineEditForContent _
                    oFWValidCatalogueLevels.GetFieldValue(gk_sDBCOLCatalogueLevelName) & _
                    m_oFWHTML.HiddenField(ksNameCatalogueLevel, _
                    oFWValidCatalogueLevels.GetFieldValue(gk_sDBCOLTempCatalogueLevelID))
                    
                'set the default level to the first entry
                eLevel = oFWValidCatalogueLevels.GetFieldValue(gk_sDBCOLTempCatalogueLevelID)
            Else
                ' Add the script to alter the DTD and the Covering Dates mandatory value
                'd64 extended script for nect reference
                Dim sScript As String
                sScript = "<SCRIPT LANGUAGE=""javascript""> function ChangeForm() {"

                sScript = sScript & vbNewLine & "if (document.forms[0].elements(" & DoubleQuotes(ksNameCatalogueLevel) & ").value>=6) {" & _
                    "document.forms[0].elements(" & DoubleQuotes(ksNameDTDType) & ").value=" & eadScopeAndContent & _
                    ";} else {" & _
                    "document.forms[0].elements(" & DoubleQuotes(ksNameDTDType) & ").value=" & eadTitle & ";" & _
                    "}"
                
                sScript = sScript & vbNewLine & "if (document.forms[0].elements(" & DoubleQuotes(ksNameCatalogueLevel) & ").value == " & ECATALOGUELEVEL.ctHeaders & " || " & _
                    "document.forms[0].elements(" & DoubleQuotes(ksNameCatalogueLevel) & ").value == " & ECATALOGUELEVEL.ctSubHeaders & ") {" & _
                    "updateField(" & DoubleQuotes(ksNameCoveringDate) & ", false, ""text"");" & _
                    "} else {" & _
                    "updateField(" & DoubleQuotes(ksNameCoveringDate) & ", true, ""text"");" & _
                    "}"
                oFWValidCatalogueLevels.MoveFirst
                Do While Not oFWValidCatalogueLevels.IsAtEnd
                    'd69 Script now uses values loaded into arrays, rather than calling GetLastRef fn every time
                    sScript = sScript & vbNewLine & "if (document.forms[0].elements(" & DoubleQuotes(ksNameCatalogueLevel) & ").value == " & oFWValidCatalogueLevels.GetFieldValue(gk_sDBCOLTempCatalogueLevelID) & ") " & _
                        "{document.forms[0].elements(" & DoubleQuotes(ksNameReference) & ").value = '" & sNextReference(oFWValidCatalogueLevels.GetFieldValue(gk_sDBCOLTempCatalogueLevelID)) & _
                    "';document.all." & gk_sSpanPrefix + ksNameReference & ".innerText=" & DoubleQuotes(sRefNote(oFWValidCatalogueLevels.GetFieldValue(gk_sDBCOLTempCatalogueLevelID))) & ";}"
                   
                   
                    oFWValidCatalogueLevels.MoveNext
                Loop
                    'PNP Add getLastRef() functionality here
                    
                sScript = sScript & vbNewLine & "} </SCRIPT>"
                
                DisplayOneLineEditForContent _
                    m_oFWHTML.Listbox(ksNameCatalogueLevel, oFWValidCatalogueLevels, False, oBOCatalogueSummary.NewCatalogueLevelID, 1, _
                    gk_sDBCOLTempCatalogueLevelID, gk_sDBCOLCatalogueLevelName, "ChangeForm();") & sScript
                    
            End If
        End If
    End If
    
    'Title
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpTitle) _
        Or t_eCatalogueEntryEditType <> ECatalogueEntryEditType.ecEditCatalogueEntry Then
        
        If t_eCatalogueEntryEditType <> ECatalogueEntryEditType.ecEditCatalogueEntry Then
            'Creating catalogue entry (show title/scope and content)
            
            DisplayOneLineHeadingForContent _
                m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfCatalogueTitle)
        Else
            'Show title
            DisplayOneLineHeadingForContent _
                m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfTabDisplaySummaryDetailsLeftHeaderTitle)
        End If
        
        
        'We use the ScopeContent dtd for create new entries
        'at piece/item level
        Dim eEADType As EADType
        Select Case eLevel
            Case ctPiece, ctItem
                If Not t_eCatalogueEntryEditType = ecEditCatalogueEntry Then
                    eEADType = eadScopeAndContent
                Else
                    eEADType = eadTitle
                End If
            Case Else
                eEADType = eadTitle
        End Select
        
        'Render current html and insert the activex control for editing xml
        
        'get the value in string
        sValueForHiddenEADField = VariantStringAsString(oBOCatalogueSummary.Title)
        sTempTitle = sValueForHiddenEADField
        'replace single and double quotes with html code equivalent
        sValueForHiddenEADField = Replace(sValueForHiddenEADField, "'", gk_sHTMLApostropheDecimalCode)
        sValueForHiddenEADField = Replace(sValueForHiddenEADField, """", gk_sHTMLQuotationDecimalCode)

        '   d35 dtm     10Feb2000       Modified so that a * is displayed when title is empty to indicate mandatoriness
                '   d36 jcc     18Feb2000       and remove the -from- bit too! Also, should be a * on title all the time, above piece
        
        'Have a * for subheader up if no data present for title/scope-content
        'and we are on the tabbed edit cat entry section
        'or for all levels if we are doing a "create"
        If (sTempTitle = "" And eLevel < ctPiece And t_eCatalogueEntryEditType = ecEditCatalogueEntry) Or _
            (Not t_eCatalogueEntryEditType = ecEditCatalogueEntry And sTempTitle = "") Then
            sTempTitle = "*"
        End If
        
        
        'The field ksNameDTDType is needed so that we
        'can change between dtds to load for the case where
        'we have a class as parent for creating new and could be
        'creating a header (in which case we would want title dtd)
        'or a piece (when we would want scope-content)
        
        DisplayOneLineEditForContent _
            RenderXML(sTempTitle) & _
            m_oFWHTML.HiddenField(ksNameTitle, m_oFWHTML.EncodeAsHTML(VariantStringAsString(sValueForHiddenEADField))) & _
            m_oFWHTML.HiddenField(ksNameDTDType, CStr(eEADType)) & _
            m_oFWHTML.HiddenField(ksNameHiddenEADUpdatedTitle, "false") & _
            m_oFWHTML.EADEditorLauncher( _
                gk_sHTMLDefaultFormName, _
                gk_sHTMLDefaultFormName & "." & ksNameTitle, _
                gk_sHTMLDefaultFormName & "." & ksNameHiddenEADUpdatedTitle, _
                gk_sHTMLDefaultXMLEditorLauncherControlName & ksNameTitle)
            
        'blank heading
        DisplayOneLineHeadingForContent ""
        
        ' d52 - Added extra situations where there may be a listbox for
        '       the level to create
        Dim bUseDTDSwitchField As Boolean
        If (t_eParentCatalogueLevelID = ctClass Or _
            t_eParentCatalogueLevelID = ctLetterCode Or _
            t_eParentCatalogueLevelID = ctHeaders) And _
            Not t_eCatalogueEntryEditType = ecEditCatalogueEntry Then
            bUseDTDSwitchField = True
        Else
            bUseDTDSwitchField = False
        End If
        
        DisplayOneLineEditForContent _
            m_oFWHTML.ImageClick( _
                gk_sImagesPath & gk_sImageEdit, _
                GenerateEditOnclickCode(ksNameTitle, eEADType, bUseDTDSwitchField)) & _
            m_oFWHTML.ImageClick( _
                gk_sImagesPath & gk_sIMAGEUpdate, _
                GenerateUpdateOnclickCode(ksNameTitle))
    End If
                
    'Legal Status
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpLegalStatus) _
        Or t_eCatalogueEntryEditType <> ECatalogueEntryEditType.ecEditCatalogueEntry Then
    
        'Heading
        DisplayOneLineHeadingForContent _
            m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfTabDisplaySummaryDetailsLeftHeaderLegalStatus)
        
        'Get the list of legal status
        m_oBOLookUp.LookupTableID = ELOOKUPTYPE.lutLegalStatus
        
        Set oFWRecordlist = m_oBOLookUp.LookupList(False)
        
        If VariantStringAsString(oBOCatalogueSummary.LegalStatusID) = "" Then   'If null
            DisplayOneLineEditForContent _
                m_oFWHTML.Listbox(ksNameLegalStatusID, oFWRecordlist, _
                True, -1, 0, gk_sDBCOLLegalStatusID, gk_sDBCOLLegalStatusDescription)
        Else
            DisplayOneLineEditForContent _
                m_oFWHTML.Listbox(ksNameLegalStatusID, oFWRecordlist, _
                True, CLng(oBOCatalogueSummary.LegalStatusID), 0, gk_sDBCOLLegalStatusID, gk_sDBCOLLegalStatusDescription)
        End If
    End If
    
    'Languages
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpLanguage) _
        Or t_eCatalogueEntryEditType <> ECatalogueEntryEditType.ecEditCatalogueEntry Then
        
        'Heading
        DisplayOneLineHeadingForContent _
            m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfTabDisplaySummaryDetailsLeftHeaderLanguages)
        
        'Get the list of language
        m_oBOLookUp.LookupTableID = ELOOKUPTYPE.lutLanguage
        
        Set oFWRecordlist = m_oBOLookUp.LookupList(False)
        
        If VariantStringAsString(oBOCatalogueSummary.LanguageID) = "" Then   'If null
            DisplayOneLineEditForContent _
                m_oFWHTML.Listbox(ksNameLanguageID, oFWRecordlist, _
                True, -1, 0, gk_sDBCOLLanguageID, gk_sDBCOLLanguageDescription)
        Else
            DisplayOneLineEditForContent _
                m_oFWHTML.Listbox(ksNameLanguageID, oFWRecordlist, _
                True, CLng(oBOCatalogueSummary.LanguageID), 0, gk_sDBCOLLanguageID, gk_sDBCOLLanguageDescription)
        End If
    End If
    
    'Former Reference Department (Only while editing)
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpFormerDepartmentReference) Then
        'Heading
        DisplayOneLineHeadingForContent _
            m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfTabDisplaySummaryDetailsLeftHeaderFormerReferenceDepartment)
        
        'Accept data
        DisplayOneLineEditForContent _
            m_oFWHTML.InputField(ITText, ksNameFormerReferenceDepartment, 50, _
            m_oBOCatalogueElementMatrix.CataloguePropertyLength(cpFormerDepartmentReference), _
            VariantStringAsString(oBOCatalogueSummary.FormerReferenceDepartment), _
            m_oBOCatalogueElementMatrix.CataloguePropertyDataType(cpFormerDepartmentReference), False)

    End If
    
    'Former Reference PRO (Only while editing)
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpFormerPROReference) Then
        'Heading
        DisplayOneLineHeadingForContent _
            m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfTabDisplaySummaryDetailsLeftHeaderFormerReferencePRO)
        
        'Accept data
        DisplayOneLineEditForContent _
            m_oFWHTML.InputField(ITText, ksNameFormerReferencePRO, 50, _
            m_oBOCatalogueElementMatrix.CataloguePropertyLength(cpFormerPROReference), _
            VariantStringAsString(oBOCatalogueSummary.FormerReferencePRO), _
            m_oBOCatalogueElementMatrix.CataloguePropertyDataType(cpFormerPROReference), False)

    End If
    
    'Map Designation
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpMapDesignation) Then
        'Heading
        DisplayOneLineHeadingForContent _
            m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfTabDisplaySummaryDetailsLeftHeaderMapDesignation)
        
        'get the value in string
        sValueForHiddenEADField = VariantStringAsString(oBOCatalogueSummary.MapDesignation)
        
        'replace single and double quotes with html code equivalent
        sValueForHiddenEADField = Replace(sValueForHiddenEADField, "'", gk_sHTMLApostropheDecimalCode)
        sValueForHiddenEADField = Replace(sValueForHiddenEADField, """", gk_sHTMLQuotationDecimalCode)
        
        'Render current html and insert the activex control for editing xml
        DisplayOneLineEditForContent _
            RenderXML(VariantStringAsString(oBOCatalogueSummary.MapDesignation)) & _
            m_oFWHTML.HiddenField(ksNameMapDesignation, m_oFWHTML.EncodeAsHTML(sValueForHiddenEADField)) & _
            m_oFWHTML.HiddenField(ksNameHiddenEADUpdatedMapDesignation, "false") & _
            m_oFWHTML.EADEditorLauncher( _
                gk_sHTMLDefaultFormName, _
                gk_sHTMLDefaultFormName & "." & ksNameMapDesignation, _
                gk_sHTMLDefaultFormName & "." & ksNameHiddenEADUpdatedMapDesignation, _
                gk_sHTMLDefaultXMLEditorLauncherControlName & ksNameMapDesignation)
            
        'blank heading
        DisplayOneLineHeadingForContent ""
        
        'edit and update buttons
        DisplayOneLineEditForContent _
            m_oFWHTML.ImageClick( _
                gk_sImagesPath & gk_sImageEdit, _
                GenerateEditOnclickCode(ksNameMapDesignation, EADType.eadMapDesignation)) & _
            m_oFWHTML.ImageClick( _
                gk_sImagesPath & gk_sIMAGEUpdate, _
                GenerateUpdateOnclickCode(ksNameMapDesignation))
        
    End If
    
    'Creator Names
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpCreatorNames) _
        Or t_eCatalogueEntryEditType <> ECatalogueEntryEditType.ecEditCatalogueEntry Then
        
        If DisplayLinkItemForEdit _
            (cpCreatorNames, _
            m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfTabDisplaySummaryDetailsLeftHeaderCreatorNames), _
            oBOCatalogueSummary.CreatorNamesCorporateBody, , _
            oBOCatalogueSummary.CreatorNamesPerson) Then      'Some data was displayed
            
            DisplayOneLineHeadingForContent ""  'To start a new line
        End If
        
        'Add link buttons for coroprate body and person
        DisplayOneLineEditForContent _
            m_oFWHTML.SubmitImage(gk_sImagesPath & gk_sImagePickCorporateBody, ksNameButtonAddTermCreatorNamesCorpBody, _
                ksNameButtonAddTermCreatorNamesCorpBody) & _
            m_oFWHTML.SubmitImage(gk_sImagesPath & gk_sImagePickPerson, ksNameButtonAddTermCreatorNamesPerson, _
                ksNameButtonAddTermCreatorNamesPerson)

    End If
    
    'Covering Dates
    ' d56 - Made non-mandatory for headers and subheaders only
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpCoveringDates) _
        Or t_eCatalogueEntryEditType <> ECatalogueEntryEditType.ecEditCatalogueEntry Then
        
        Dim bMandatory As Boolean
        If eLevel = ctHeaders Or _
            eLevel = ctSubHeaders Then
            bMandatory = False
        Else
            bMandatory = True
        End If
        
        'Covering Date
        'Heading
        DisplayOneLineHeadingForContent _
            m_oFWAppData.GetDisplayInformation( _
            ESharedResourceFile.srfTabDisplaySummaryDetailsLeftHeaderCoveringDates)
        
        'Accept
        ' d71 Added regnal years check box
        DisplayOneLineEditForContent _
            m_oFWHTML.InputField(ITText, ksNameCoveringDate, 50, _
                gk_iDBCOLLenCoveringDate, _
                VariantStringAsString(oBOCatalogueSummary.CoveringDates), VCNone, bMandatory) & _
            m_oFWHTML.Checkbox("Regnal Date", ksNameIsRegnalDate)
            
            
    End If

    ' d60 - Start / End Date
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpCoveringDates) _
        Or t_eCatalogueEntryEditType <> ECatalogueEntryEditType.ecEditCatalogueEntry Then
        
        'Heading
        DisplayOneLineHeadingForContent _
            m_oFWAppData.GetDisplayInformation( _
            ESharedResourceFile.srfTabDisplaySummaryDetailsLeftHeaderStartEndDate)
        
        'Accept
        DisplayOneLineEditForContent _
            m_oFWHTML.InputField(ITText, ksNameStartDate, 10, _
                gk_iDBCOLLenStartDate, _
                VariantStringAsString(oBOCatalogueSummary.StartDate), VCDate) & _
            " - " & _
            m_oFWHTML.InputField(ITText, ksNameEndDate, 10, _
                gk_iDBCOLLenEndDate, _
                VariantStringAsString(oBOCatalogueSummary.EndDate), VCDate) & _
            " (" & gk_sDMYDateFormat & ")"
            
    End If

    'Physical Description Form Code
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpPhysicalDescription_Form) Then
        
        'Heading
        DisplayOneLineHeadingForContent _
            m_oFWAppData.GetDisplayInformation( _
            ESharedResourceFile.srfTabDisplaySummaryDetailsLeftHeaderPhysicalDescriptionFormCode)
        
        'Get the list of physical forms
        m_oBOLookUp.LookupTableID = ELOOKUPTYPE.lutPhysicalForm
        
        Set oFWRecordlist = m_oBOLookUp.LookupList(False)
        
        If VariantStringAsString(oBOCatalogueSummary.PhysicalDescriptionFormCode) = "" Then    'If null
            DisplayOneLineEditForContent _
                m_oFWHTML.Listbox(ksNamePhysicalDescriptionFormCode, oFWRecordlist, _
                True, -1, 0, gk_sDBCOLPhysicalDescriptionFormID, gk_sDBCOLPhysicalDescriptionFormDescription)
        Else
            DisplayOneLineEditForContent _
                m_oFWHTML.Listbox(ksNamePhysicalDescriptionFormCode, oFWRecordlist, _
                True, CLng(oBOCatalogueSummary.PhysicalDescriptionFormCode), 0, _
                gk_sDBCOLPhysicalDescriptionFormID, gk_sDBCOLPhysicalDescriptionFormDescription)
        End If
    End If
    
    'Physical Description Record Qty
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, _
        cpPhysicalDescription_RecordQuantity) Then
        
        'Heading
        DisplayOneLineHeadingForContent _
            m_oFWAppData.GetDisplayInformation( _
            ESharedResourceFile.srfTabDisplaySummaryDetailsLeftHeaderPhysicalDescriptionRecordQuantity)

        'Accept data
        DisplayOneLineEditForContent _
            m_oFWHTML.InputField(ITText, ksNamePhysicalDescriptionRecordQty, 50, _
            m_oBOCatalogueElementMatrix.CataloguePropertyLength(cpPhysicalDescription_RecordQuantity), _
            VariantStringAsString(oBOCatalogueSummary.PhysicalDescriptionRecordQuantity), _
            m_oBOCatalogueElementMatrix.CataloguePropertyDataType(cpPhysicalDescription_RecordQuantity), _
            False, False, 0, gk_lMaximumValueForIntegerInSQL7)

    End If

    'Physical Description Dimensions
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpPhysicalDescription_Dimension) Then
        'Heading
        DisplayOneLineHeadingForContent _
            m_oFWAppData.GetDisplayInformation( _
            ESharedResourceFile.srfTabDisplaySummaryDetailsLeftHeaderPhysicalDescriptionDimensions)

        'Accept data
        DisplayOneLineEditForContent _
            m_oFWHTML.InputField(ITText, ksNamePhysicalDescriptionDimensions, 50, _
            m_oBOCatalogueElementMatrix.CataloguePropertyLength(cpPhysicalDescription_Dimension), _
            VariantStringAsString(oBOCatalogueSummary.PhysicalDescriptionDimensions), _
            m_oBOCatalogueElementMatrix.CataloguePropertyDataType(cpPhysicalDescription_Dimension), False)

    End If

    'Map Scale Number
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpMapscaleNumber) Then
        'Heading
        DisplayOneLineHeadingForContent _
            m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfTabDisplaySummaryDetailsLeftHeaderMapScaleNumber)
        
        'Accept data
        DisplayOneLineEditForContent _
            m_oFWHTML.InputField(ITText, ksNameMapScaleNumber, 50, _
           m_oBOCatalogueElementMatrix.CataloguePropertyLength(cpMapscaleNumber), _
           VariantStringAsString(oBOCatalogueSummary.MapScaleNumber), _
           m_oBOCatalogueElementMatrix.CataloguePropertyDataType(cpMapscaleNumber), False, False, 0, gk_lMaximumValueForIntegerInSQL7)
            
    End If

    'Physical Condition
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpPhysicalCondition) Then
        'Heading
        DisplayOneLineHeadingForContent _
            m_oFWAppData.GetDisplayInformation( _
            ESharedResourceFile.srfTabDisplaySummaryDetailsLeftHeaderPhysicalCondition)

        'Accept data
        DisplayOneLineEditForContent _
            m_oFWHTML.InputField(ITText, ksNamePhysicalCondition, 50, _
           m_oBOCatalogueElementMatrix.CataloguePropertyLength(cpPhysicalCondition), _
           VariantStringAsString(oBOCatalogueSummary.PhysicalCondition), _
           m_oBOCatalogueElementMatrix.CataloguePropertyDataType(cpPhysicalCondition), False)

    End If
    
    'Place of Deposit
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpCreatorNames) _
        Or t_eCatalogueEntryEditType <> ECatalogueEntryEditType.ecEditCatalogueEntry Then
        
        If DisplayLinkItemForEdit _
            (cpPlaceOfDeposit, _
            m_oFWAppData.GetDisplayInformation( _
            ESharedResourceFile.srfTabDisplaySummaryDetailsLeftHeaderPlaceOfDeposit), _
            oBOCatalogueSummary.PlaceOfDeposit) Then     'Some data was displayed
            
            DisplayOneLineHeadingForContent ""  'To start a new line
        End If
        
        DisplayOneLineEditForContent _
            m_oFWHTML.SubmitImage(gk_sImagesPath & gk_sImagePickCorporateBody, _
            ksNameButtonAddTermPlaceOfDeposit, ksNameButtonAddTermPlaceOfDeposit)
    End If
    
    'Piece key order (only while editing)
    If t_eCatalogueEntryEditType = ECatalogueEntryEditType.ecEditCatalogueEntry _
        And t_eCatalogueLevelID = ECATALOGUELEVEL.ctPiece Then  'while editing piece
        
        DisplayOneLineHeadingForContent _
         m_oFWAppData.GetDisplayInformation( _
            ESharedResourceFile.srfTabDisplaySummaryDetailsLeftHeaderPieceKeyOrder)
        
        'Accept data
        DisplayOneLineEditForContent _
            m_oFWHTML.InputField(ITText, ksNamePieceKeyOrder, 50, _
           , oBOCatalogueSummary.PieceKeyOrder, VCNumber, True, False, 1, gk_lMaximumValueForIntegerInSQL7)
    End If
    
    'Item key order (only while editing)
    If t_eCatalogueEntryEditType = ECatalogueEntryEditType.ecEditCatalogueEntry _
        And t_eCatalogueLevelID = ECATALOGUELEVEL.ctItem Then  'while editing item
        
        DisplayOneLineHeadingForContent _
         m_oFWAppData.GetDisplayInformation( _
            ESharedResourceFile.srfTabDisplaySummaryDetailsLeftHeaderItemKeyOrder)
        
        'Accept data
        DisplayOneLineEditForContent _
            m_oFWHTML.InputField(ITText, ksNameItemKeyOrder, 50, _
           , oBOCatalogueSummary.ItemKeyOrder, VCNumber, True, False, 1, gk_lMaximumValueForIntegerInSQL7)
    End If
    
    'd72 Note
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpNote) Then
        'Heading
        DisplayOneLineHeadingForContent _
            m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfTabDisplaySummaryDetailsLeftHeaderNote)
        
        'Accept data
        'd72 Changed cpScopeContent to cpNote which is the correct one
        DisplayOneLineEditForContent _
            m_oFWHTML.InputField(VCNone, ksNameNote, 50, _
            m_oBOCatalogueElementMatrix.CataloguePropertyLength(cpNote), _
            VariantStringAsString(oBOCatalogueSummary.Notes), VCNone, False)
    End If
    
    'Links for the cancel button
    If t_eCatalogueEntryEditType <> ECatalogueEntryEditType.ecEditCatalogueEntry Then
        'when creating a catalogue
        sParamString = gk_sLinkDisplayEditSetDetails & _
                        "?" & gk_sParamEditSetID & "=" & t_lEditSetID
    Else
        'when editing a catalogue
        sParamString = gk_sLinkDisplayViewEditorialCatalogueDetails & _
                        "?" & gk_sParamEditSetID & "=" & t_lEditSetID & _
                        "&" & gk_sParamCatalogueID & "=" & t_lCatalogueID & _
                        "&" & gk_sParamCatalogueLevel & "=" & t_eCatalogueLevelID & _
                        "&" & gk_sParamDetailsTabNumber & "=" & ETabType.etTabSummary & _
                        "&" & gk_sParamSectionType & "=" & ESectionType.SECeditorial & _
                        "&" & ksParamEditingCancelled & "=" & True
    End If
    
    'Display OK and Cancel button
    DisplayOKCancelButtonsForEditAndCreateCatalogue (sParamString)
    
    'Release the object
    Set oBOCatalogueSummary = Nothing
    Set oFWRecordlist = Nothing
    Set oFWValidCatalogueLevels = Nothing
    Set oBOBasicCatalogueEntry = Nothing
    
    'Generate the html
    sHTML = sHTML & m_oFWHTML.Paragraph(m_oFWTable.GetHTML) & m_oFWHTML.EndTag(gk_sHTMLForm)
    

    'Add the help button
    If t_eLoadType = ECatalogueEditFormLoadType.etLoadFromNoneCreateCatalogueEntry Then
        sHTML = sHTML & HTMLToAddHelpButton(gk_sHelpLinkToCreateCatalogueEntry)
    Else
        sHTML = sHTML & HTMLToAddHelpButton(gk_sHelpLinkToEditSummaryDetails)
    End If
    
    'Show the html page
    m_aspResponse.Write sHTML
    
    Exit Sub
    
ErrorHandler:
    Set oBOCatalogueSummary = Nothing
    Set oFWRecordlist = Nothing
    Set oFWValidCatalogueLevels = Nothing
    Set oBOBasicCatalogueEntry = Nothing

    ConvertError sSource
End Sub

Private Sub DisplayEditAccessDetailsScreen( _
            ByVal t_eLoadType As ECatalogueEditFormLoadType, _
            ByVal t_lEditSetID As Long, _
            ByVal t_lCatalogueID As Long, _
            ByVal t_eCatalogueLevelID As ECATALOGUELEVEL, _
            Optional ByVal t_lLinkTermID As Long, _
            Optional ByVal t_eLinkTermType As ELinkTermType)
    

'Description : Displays the content screen for editing for the selected catalogue
'
'Amendment history
'   d01  ndk     25Aug1999       Created
'   d04  ndk     04Nov1999       display of access conditions handled differently
'   d09  ndk     20Nov1999       For lookups using lookup list of bolookup instead of fwreference data
'   d16  ndk     08Dec1999       Added help button
'   d20  wdp     10Dec1999       Added record status checkboxes
'   d22  ndk     16Dec1999       Changed edit format for record opening date
'   d23  ndk     21Dec1999       Added security check
'   d49  tmj     30Mar2000       Fixed Display of EAD elements, by encoding hidden field as HTML
'   d53  daj     23May2000       Added checking of closure code to prevent invalid use of null errors
'   d67  gdb     20Feb2001       Added more arguments when invoking inputFiled() for closure code in DisplayEditAccessDetailScreen


    On Error GoTo ErrorHandler

    'Variables
    Dim sSource As String, sHTML As String, sParamString As String
    Dim sRecordOpeningDate As String
    Dim oFWRecordlist As FWRecordList
    Dim eSecurityCheckStatus As ESecurityCheckResult
        
    'Business objects
    Dim oBOCatalogueAccess As BOCatalogueAccess
    
    sSource = Me.Name & ".DisplayEditAccessDetailsScreen"

    'perform security check
    m_oBOEditorialPermissions.EditsetID = t_lEditSetID
    m_oBOEditorialPermissions.UserID = m_lLoggedUserID
    
    eSecurityCheckStatus = m_oBOEditorialPermissions.Check(EUserCanEditCatalogueEntryOnEditSet)
    
    If eSecurityCheckStatus <> ESecurityCheckResult.ESRSecurityCheckSuccessful Then
        RaiseError sSource, _
               GetErrorEnumForSecurityCheckResult(eSecurityCheckStatus), _
               GetDescriptionForSecurityCheckError( _
                        GetErrorEnumForSecurityCheckResult(eSecurityCheckStatus))
    End If
    
    'Instantiate the object
    Set oBOCatalogueAccess = New BOCatalogueAccess
    
    'Set the properties of the catalogue to be displayed
    If t_eLoadType = ECatalogueEditFormLoadType.etLoadFromDatabaseEditCatalogueEntry Then
        'The user clicked on Edit button and control has come here
        
        oBOCatalogueAccess.CatalogueID = t_lCatalogueID
        oBOCatalogueAccess.CatalogueLevelID = t_eCatalogueLevelID
        oBOCatalogueAccess.EditsetID = t_lEditSetID
        oBOCatalogueAccess.SectionType = ESectionType.SECeditorial
        
        'Get the details
        oBOCatalogueAccess.Load t_eLoadType
    Else
        Select Case t_eLoadType
        
            Case ECatalogueEditFormLoadType.etLoadFromSavedStateAndAddTermEditCatalogueEntry
                
                'The user has selected a pick term to add and the control has come here
                oBOCatalogueAccess.Load t_eLoadType, t_lLinkTermID, t_eLinkTermType
            
            Case ECatalogueEditFormLoadType.etLoadFromSavedStateAndRemoveTermEditCatalogueEntry
            
                'The user has selected a pick term to remove and the control has come here
                oBOCatalogueAccess.Load t_eLoadType, t_lLinkTermID, t_eLinkTermType
                
            Case ECatalogueEditFormLoadType.etLoadFromSavedStateNoChangeEditCatalogueEntry
            
                'The user has selected cancel operation while adding/removing a link term and the control has come here
                oBOCatalogueAccess.Load t_eLoadType
                
        End Select
    
    End If
    
    'To display the details
    m_oFWTable.Reset    'Reset the table details
    sHTML = ""
    
    'Apply the style to the table
    m_oFWTable.TableStyle = m_typTableStyleContents
    
    'Link for submiting the form details
    sParamString = gk_sLinkSubmitCatalogueDetails & _
                    "?" & gk_sParamEditSetID & "=" & t_lEditSetID & _
                    "&" & gk_sParamCatalogueID & "=" & t_lCatalogueID & _
                    "&" & gk_sParamCatalogueLevel & "=" & t_eCatalogueLevelID & _
                    "&" & gk_sParamDetailsTabNumber & "=" & ETabType.etTabAccess
    
    'Start the form
    sHTML = sHTML & m_oFWHTML.FormStart(sParamString, gk_sMainFrame, gk_sHTMLDefaultFormName, MPost, True)
        
    'Access Condition
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpAccessCondition) Then
        'Get the list of access conditions
        
        'Heading - Access conditions
        DisplayOneLineHeadingForContent _
            m_oFWAppData.GetDisplayInformation( _
            ESharedResourceFile.srfTabDisplayAccessDetailsLeftHeaderAccessConditions)
        
        
        'Get the list of access conditions
        m_oBOLookUp.LookupTableID = ELOOKUPTYPE.lutAccessCondition
        
        Set oFWRecordlist = m_oBOLookUp.LookupList(False)
        
        If VariantStringAsString(oBOCatalogueAccess.AccessConditionID) = "" Then   'If null
            DisplayOneLineEditForContent _
                m_oFWHTML.Listbox(ksNameAccessConditionID, oFWRecordlist, True, -1, 0, _
                gk_sDBCOLAccessConditionID, gk_sDBCOLAccessConditionDescription)
        Else
            DisplayOneLineEditForContent _
                m_oFWHTML.Listbox(ksNameAccessConditionID, oFWRecordlist, True, _
                CLng(oBOCatalogueAccess.AccessConditionID), 0, gk_sDBCOLAccessConditionID, gk_sDBCOLAccessConditionDescription)
        End If
    End If
    
    'closure type and closure code
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpClosureType) And _
        m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpClosureCode) Then
            
        'Heading - Closure type
        DisplayOneLineHeadingForContent _
            m_oFWAppData.GetDisplayInformation( _
            ESharedResourceFile.srfTabDisplayAccessDetailsLeftHeaderClosureType)
        
        'Get the list of access conditions
        m_oBOLookUp.LookupTableID = ELOOKUPTYPE.lutClosureType
        
        Set oFWRecordlist = m_oBOLookUp.LookupList(False)
        
        If VariantStringAsString(oBOCatalogueAccess.ClosureType) = "" Then    'If null
            DisplayOneLineEditForContent _
                m_oFWHTML.Listbox(ksNameClosureType, oFWRecordlist, True, -1, 0, _
                gk_sDBCOLClosureType, gk_sDBCOLClosureTypeDescription)
        Else
            DisplayOneLineEditForContent _
                m_oFWHTML.Listbox(ksNameClosureType, oFWRecordlist, True, _
                oBOCatalogueAccess.ClosureType, 0, gk_sDBCOLClosureType, gk_sDBCOLClosureTypeDescription)
        End If
        
        'Heading - Closure code
        DisplayOneLineHeadingForContent _
            m_oFWAppData.GetDisplayInformation( _
            ESharedResourceFile.srfTabDisplayAccessDetailsLeftHeaderClosureCode)
        
        ' d53 - If null, use an empty string

        DisplayOneLineEditForContent _
            m_oFWHTML.InputField(ITText, ksNameClosureCode, 40, gk_lMaximumValueForIntegerInSQL7, _
            IIf(IsNull(oBOCatalogueAccess.ClosureCode), "", oBOCatalogueAccess.ClosureCode), EValidCheck.VCNumber, False, False, gk_vRangeMin, gk_vRangeMax)
        
    End If
    
        
    'Record Opening Date
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpRecordOpeningDate) Then
        'Heading
        DisplayOneLineHeadingForContent _
            m_oFWAppData.GetDisplayInformation( _
                ESharedResourceFile.srfTabDisplayEditAccessDetailsLeftHeaderRecordOpeningDateDDMMYYYY)
        
        sRecordOpeningDate = VariantDateAsDDMMYYYY(oBOCatalogueAccess.RecordOpeningDate)
        
        'Accept data
        DisplayOneLineEditForContent _
            m_oFWHTML.InputField(VCNone, ksNameRecordOpeningDate, 12, _
               m_oBOCatalogueElementMatrix.CataloguePropertyLength(cpRecordOpeningDate), _
               sRecordOpeningDate, _
               m_oBOCatalogueElementMatrix.CataloguePropertyDataType(cpRecordOpeningDate), False)
    End If
    
    'Closure Status
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpClosureStatus) Then
        'Heading
        DisplayOneLineHeadingForContent _
            m_oFWAppData.GetDisplayInformation( _
                ESharedResourceFile.srfTabDisplayAccessDetailsLeftHeaderClosureStatus)
        
        'Get the list of closure status
        m_oBOLookUp.LookupTableID = ELOOKUPTYPE.lutClosureStatus
        
        Set oFWRecordlist = m_oBOLookUp.LookupList(False)
        
        'Accept data
        If VariantStringAsString(oBOCatalogueAccess.ClosureStatus) = "" Then    'If null
            DisplayOneLineEditForContent _
                m_oFWHTML.Listbox(ksNameClosureStatus, oFWRecordlist, True, -1, 0, _
                gk_sDBCOLClosureStatus, gk_sDBCOLClosureStatusDescription)
        Else
            DisplayOneLineEditForContent _
                m_oFWHTML.Listbox(ksNameClosureStatus, oFWRecordlist, True, oBOCatalogueAccess.ClosureStatus, _
                0, gk_sDBCOLClosureStatus, gk_sDBCOLClosureStatusDescription)
        End If
        
    End If
    
    'Restrictions On Use
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpRestrictionsOnUse) Then
        'Heading
        DisplayOneLineHeadingForContent _
            m_oFWAppData.GetDisplayInformation( _
            ESharedResourceFile.srfTabDisplayAccessDetailsLeftHeaderRestrictionsOnUse)
        
        'Accept data
        DisplayOneLineEditForContent _
            m_oFWHTML.InputField(ITText, ksNameRestrictionsOnUse, 50, _
                m_oBOCatalogueElementMatrix.CataloguePropertyLength(cpRestrictionsOnUse), _
                VariantStringAsString(oBOCatalogueAccess.RestrictionsOnUse), _
                m_oBOCatalogueElementMatrix.CataloguePropertyDataType(cpRestrictionsOnUse), False)
    End If
    
    'Immediete Source of acquisition
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpImmediateSourceOfAcquisition) Then

        If DisplayLinkItemForEdit( _
            cpImmediateSourceOfAcquisition, _
            m_oFWAppData.GetDisplayInformation( _
                ESharedResourceFile.srfTabDisplayAccessDetailsLeftHeaderImmediateSourceOfAcquisition), _
            oBOCatalogueAccess.ImmedieteSourceOfAcquisitionCorporateBody, _
            oBOCatalogueAccess.UnknownImmedieteSourceOfAcquisition, _
            oBOCatalogueAccess.ImmedieteSourceOfAcquisitionPerson) Then
                        
            DisplayOneLineHeadingForContent ""
        End If
        
        'Add new link buttons (corporate boody and person)
        DisplayOneLineEditForContent _
            m_oFWHTML.SubmitImage(gk_sImagesPath & gk_sImagePickCorporateBody, ksNameButtonAddTermImmSrcCorpBody, ksNameButtonAddTermImmSrcCorpBody) & _
            m_oFWHTML.SubmitImage(gk_sImagesPath & gk_sImagePickPerson, ksNameButtonAddTermImmSrcPerson, ksNameButtonAddTermImmSrcPerson)

    End If
    
    'Custodial History - text area
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpCustodialHistory) Then
        'Heading
        DisplayOneLineHeadingForContent _
            m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfTabDisplayAccessDetailsLeftHeaderCustodialHistory)
        
        'Accept data
        DisplayOneLineEditForContent _
            m_oFWHTML.TextArea( _
                ksNameCustodialHistory, _
                VariantStringAsString(oBOCatalogueAccess.CustodialHistory), _
                kiColumnsForTextFields, _
                kiRowsForTextFields, _
                m_oBOCatalogueElementMatrix.CataloguePropertyLength(cpCustodialHistory))
        
    End If
    
    'Accumulation Dates
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpAccumulationDates) Then
        'Heading
        DisplayOneLineHeadingForContent _
            m_oFWAppData.GetDisplayInformation( _
            ESharedResourceFile.srfTabDisplayAccessDetailsLeftHeaderAccumulationDates)
        
        'Accept data
        DisplayOneLineEditForContent _
            m_oFWHTML.InputField(ITText, ksNameAccumulationDates, 50, _
               m_oBOCatalogueElementMatrix.CataloguePropertyLength(cpAccumulationDates), _
               VariantStringAsString(oBOCatalogueAccess.AccumulationDate), _
               m_oBOCatalogueElementMatrix.CataloguePropertyDataType(cpAccumulationDates), False)
    End If
    
    'Appriasal Desctruction Information -text field
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpAppraisal_DestructionInformation) Then
        'Heading
        DisplayOneLineHeadingForContent _
            m_oFWAppData.GetDisplayInformation( _
                ESharedResourceFile.srfTabDisplayAccessDetailsLeftHeaderAppDestructionInformation)
        
        'Accept data
        DisplayOneLineEditForContent _
            m_oFWHTML.TextArea( _
                ksNameAppDestructionInformation, _
                VariantStringAsString(oBOCatalogueAccess.AppraisalDestructionInfo), _
                kiColumnsForTextFields, _
                kiRowsForTextFields, _
                m_oBOCatalogueElementMatrix.CataloguePropertyLength(cpAppraisal_DestructionInformation))
    
    End If
    
    'Accruals
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpAccruals) Then
        'Heading
        DisplayOneLineHeadingForContent _
            m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfTabDisplayAccessDetailsLeftHeaderAccruals)
        
        'Accept data
        DisplayOneLineEditForContent _
            m_oFWHTML.TextArea( _
                ksNameAccruals, _
                VariantStringAsString(oBOCatalogueAccess.Accruals), _
                kiColumnsForTextFields, _
                kiRowsForTextFields, _
                m_oBOCatalogueElementMatrix.CataloguePropertyLength(cpAccruals))
    End If
    
    'Location of originals
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpLocationOfOriginals) Then
        If DisplayLinkItemForEdit( _
            cpLocationOfOriginals, _
            m_oFWAppData.GetDisplayInformation( _
            ESharedResourceFile.srfTabDisplayAccessDetailsLeftHeaderLocationOfOriginals), _
            oBOCatalogueAccess.LocationOfOriginals) Then
                        
                DisplayOneLineHeadingForContent ""  'To start a new line
        End If
        
        'Add new link button
        DisplayOneLineEditForContent _
            m_oFWHTML.SubmitImage(gk_sImagesPath & gk_sImagePickCorporateBody, _
                ksNameButtonAddTermLocationOfOriginals, ksNameButtonAddTermLocationOfOriginals)
            
    End If
    
    'Copies Information
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpCopiesInformation) Then
        If DisplayLinkItemForEdit( _
            cpCopiesInformation, _
            m_oFWAppData.GetDisplayInformation( _
                ESharedResourceFile.srfTabDisplayAccessDetailsLeftHeaderCopiesInformation), _
            oBOCatalogueAccess.CopiesInformation, _
            oBOCatalogueAccess.UnknownCopiesInformation) Then
            
            DisplayOneLineHeadingForContent ""
            
        End If
        
        'Add new link button
        DisplayOneLineEditForContent _
            m_oFWHTML.SubmitImage( _
                gk_sImagesPath & gk_sImagePickCorporateBody, ksNameButtonAddTermCopiesInformation, ksNameButtonAddTermCopiesInformation)
        
    End If
    
    'Link term for cancel button
    sParamString = gk_sLinkDisplayViewEditorialCatalogueDetails & _
                    "?" & gk_sParamEditSetID & "=" & t_lEditSetID & _
                    "&" & gk_sParamCatalogueID & "=" & t_lCatalogueID & _
                    "&" & gk_sParamCatalogueLevel & "=" & t_eCatalogueLevelID & _
                    "&" & gk_sParamDetailsTabNumber & "=" & ETabType.etTabAccess & _
                    "&" & gk_sParamSectionType & "=" & ESectionType.SECeditorial & _
                    "&" & ksParamEditingCancelled & "=" & True

                    
    'Record status flags for pitems
    If t_eCatalogueLevelID >= ctPiece Then
        DisplayRecordStatus m_oFWAppData.GetDisplayInformation( _
            ESharedResourceFile.srfTabDisplayAccessDetailsLeftHeaderRecordStatus), _
            oBOCatalogueAccess.AccessionedAsCopy, _
            oBOCatalogueAccess.ExtractAtTransfer, _
            oBOCatalogueAccess.AR, _
            oBOCatalogueAccess.Orderable
    End If
    'Display OK and Cancel button
    DisplayOKCancelButtonsForEditAndCreateCatalogue (sParamString)
    
    'Release the object
    Set oBOCatalogueAccess = Nothing
    Set oFWRecordlist = Nothing
    
    'Generate the html
    sHTML = sHTML & m_oFWHTML.Paragraph(m_oFWTable.GetHTML) & m_oFWHTML.EndTag(gk_sHTMLForm)
    
    'Add the help button
    sHTML = sHTML & HTMLToAddHelpButton(gk_sHelpLinkToEditAccessDetails)
    
    'Show the html page
    m_aspResponse.Write sHTML
    
    Exit Sub
    
ErrorHandler:
    Set oBOCatalogueAccess = Nothing
    Set oFWRecordlist = Nothing
    
    ConvertError sSource
End Sub
Private Sub DisplayEditContentDetailsScreen( _
            ByVal t_eLoadType As ECatalogueEditFormLoadType, _
            ByVal t_lEditSetID As Long, _
            ByVal t_lCatalogueID As Long, _
            ByVal t_eCatalogueLevelID As ECATALOGUELEVEL, _
            Optional ByVal t_lLinkTermID As Long, _
            Optional ByVal t_eLinkTermType As ELinkTermType, _
            Optional ByVal t_iLinkTermLevelID As Integer, _
            Optional ByVal t_iMainCollectionKey As Integer, _
            Optional ByVal t_iSubCollectionKey As Integer, _
            Optional ByVal t_bAddingNewRow As Boolean)
    
'Description : Displays the content screen for editing for the selected catalogue
'
'Amendment history
'   d01 ndk     25Aug1999       Created
'   d15 ndk     07Dec1999       Implemented EAD Editing and made related changes
'   d16 ndk     08Dec1999       Added help button
'   d21 ndk     15Dec1999       Replacing single and double quotes with equivalent html codes before rendering ead data
'   d23 ndk     21Dec1999       Added security check
'   d29 crm     02Feb2000       Updated code to provide links for Leaflets to catalogues
'   d36 jcc     18Feb2000       and remove the -from- bit too! Also, should be a * on title all the time, above piece
'   d49 tmj     30Mar2000       Fixed Display of EAD elements, by encoding hidden field as HTML
'   d72 alm     25Sep2001       Moved the 'Notes' field to DisplayEditSummaryDetailsScreen
'
    On Error GoTo ErrorHandler

    'Variables
    Dim sSource As String, sHTML As String, sParamString As String, sTempScope As String
    Dim eSecurityCheckStatus As ESecurityCheckResult
        
    'Business objects
    Dim oBOCatalogueContent As BOCatalogueContent, sValueForHiddenEADField As String
    
    sSource = Me.Name & ".DisplayEditContentDetailsScreen"

    'perform security check
    m_oBOEditorialPermissions.EditsetID = t_lEditSetID
    m_oBOEditorialPermissions.UserID = m_lLoggedUserID
    
    eSecurityCheckStatus = m_oBOEditorialPermissions.Check(EUserCanEditCatalogueEntryOnEditSet)
    
    If eSecurityCheckStatus <> ESecurityCheckResult.ESRSecurityCheckSuccessful Then
        RaiseError sSource, _
               GetErrorEnumForSecurityCheckResult(eSecurityCheckStatus), _
               GetDescriptionForSecurityCheckError( _
                        GetErrorEnumForSecurityCheckResult(eSecurityCheckStatus))
    End If

    'Instantiate the object
    Set oBOCatalogueContent = New BOCatalogueContent
    
    'Set the properties of the catalogue to be displayed
    If t_eLoadType = ECatalogueEditFormLoadType.etLoadFromDatabaseEditCatalogueEntry Then
   
        'The user clicked on Edit button and control has come here
        oBOCatalogueContent.CatalogueID = t_lCatalogueID
        oBOCatalogueContent.CatalogueLevelID = t_eCatalogueLevelID
        oBOCatalogueContent.EditsetID = t_lEditSetID
        oBOCatalogueContent.SectionType = ESectionType.SECeditorial
        
        'Get the details
        oBOCatalogueContent.Load t_eLoadType
    Else
        Select Case t_eLoadType
        
            Case ECatalogueEditFormLoadType.etLoadFromSavedStateAndAddTermEditCatalogueEntry
            
                'The user has selected a pick term to add and the control has come here
                oBOCatalogueContent.Load t_eLoadType, t_lLinkTermID, t_eLinkTermType, _
                    t_iLinkTermLevelID, t_iMainCollectionKey
            
            Case ECatalogueEditFormLoadType.etLoadFromSavedStateAndRemoveTermEditCatalogueEntry
                
                'The user has selected a pick term to remove and the control has come here
                oBOCatalogueContent.Load t_eLoadType, t_lLinkTermID, t_eLinkTermType, _
                    t_iLinkTermLevelID, t_iMainCollectionKey
                
            Case ECatalogueEditFormLoadType.etLoadFromSavedStateNoChangeEditCatalogueEntry
            
                'The user has selected cancel operation while adding/removing a link term and
                'the control has come here
                oBOCatalogueContent.Load t_eLoadType
                
        End Select
    
    End If
    'To display the details
    
    m_oFWTable.Reset    'Reset the table details
    sHTML = ""
    
    'Apply the style to the table
    m_oFWTable.TableStyle = m_typTableStyleContents
    
    'Link for submiting the form details
    sParamString = gk_sLinkSubmitCatalogueDetails & _
                    "?" & gk_sParamEditSetID & "=" & t_lEditSetID & _
                    "&" & gk_sParamCatalogueID & "=" & t_lCatalogueID & _
                    "&" & gk_sParamCatalogueLevel & "=" & t_eCatalogueLevelID & _
                    "&" & gk_sParamDetailsTabNumber & "=" & ETabType.etTabContent
    
    'Start the form
    sHTML = sHTML & m_oFWHTML.FormStart(sParamString, gk_sMainFrame, gk_sHTMLDefaultFormName, MPost, True)
        
    'Scope and Content
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpScopeAndContent) Then
        'Heading
        DisplayOneLineHeadingForContent _
            m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfTabDisplayContentDetailsLeftHeaderScopeAndContent)
        
        'get the value in string
        sValueForHiddenEADField = VariantStringAsString(oBOCatalogueContent.ScopeAndContent)
        
        'replace single and double quotes with html code equivalent
        sValueForHiddenEADField = Replace(sValueForHiddenEADField, "'", gk_sHTMLApostropheDecimalCode)
        sValueForHiddenEADField = Replace(sValueForHiddenEADField, """", gk_sHTMLQuotationDecimalCode)
        
        'd36 jcc     18Feb2000       and remove the -from- bit too! Also, should be a * on title all the time, above piece
        sTempScope = VariantStringAsString(oBOCatalogueContent.ScopeAndContent)
        If (t_eCatalogueLevelID = ctPiece Or t_eCatalogueLevelID = ctItem) And sTempScope = "" Then
            sTempScope = "*"
        End If

        'Render current html and insert the activex control for editing xml
        DisplayOneLineEditForContent _
            RenderXML(sTempScope) & _
            m_oFWHTML.HiddenField(ksNameScopeAndContent, m_oFWHTML.EncodeAsHTML(sValueForHiddenEADField)) & _
            m_oFWHTML.HiddenField(ksNameHiddenEADUpdatedScopeAndContent, "false") & _
            m_oFWHTML.EADEditorLauncher( _
                gk_sHTMLDefaultFormName, _
                gk_sHTMLDefaultFormName & "." & ksNameScopeAndContent, _
                gk_sHTMLDefaultFormName & "." & ksNameHiddenEADUpdatedScopeAndContent, _
                gk_sHTMLDefaultXMLEditorLauncherControlName & ksNameScopeAndContent)
            
        'blank heading
        DisplayOneLineHeadingForContent ""
        
        'edit and update buttons
        DisplayOneLineEditForContent _
            m_oFWHTML.ImageClick( _
                gk_sImagesPath & gk_sImageEdit, _
                GenerateEditOnclickCode(ksNameScopeAndContent, EADType.eadScopeAndContent)) & _
            m_oFWHTML.ImageClick( _
                gk_sImagesPath & gk_sIMAGEUpdate, _
                GenerateUpdateOnclickCode(ksNameScopeAndContent))
    
    End If

    'Arrangement
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpArrangement) Then
        'Heading
        DisplayOneLineHeadingForContent _
            m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfTabDisplayContentDetailsLeftHeaderArrangement)
        
        'get the value in string
        sValueForHiddenEADField = VariantStringAsString(oBOCatalogueContent.Arrangement)
        
        'replace single and double quotes with html code equivalent
        sValueForHiddenEADField = Replace(sValueForHiddenEADField, "'", gk_sHTMLApostropheDecimalCode)
        sValueForHiddenEADField = Replace(sValueForHiddenEADField, """", gk_sHTMLQuotationDecimalCode)
                
        'Render current html and insert the activex control for editing xml
        DisplayOneLineEditForContent _
            RenderXML(VariantStringAsString(oBOCatalogueContent.Arrangement)) & _
            m_oFWHTML.HiddenField(ksNameArrangement, m_oFWHTML.EncodeAsHTML(sValueForHiddenEADField)) & _
            m_oFWHTML.HiddenField(ksNameHiddenEADUpdatedArrangement, "false") & _
            m_oFWHTML.EADEditorLauncher( _
                gk_sHTMLDefaultFormName, _
                gk_sHTMLDefaultFormName & "." & ksNameArrangement, _
                gk_sHTMLDefaultFormName & "." & ksNameHiddenEADUpdatedArrangement, _
                gk_sHTMLDefaultXMLEditorLauncherControlName & ksNameArrangement)
            
        'blank heading
        DisplayOneLineHeadingForContent ""
        
        'edit and update buttons
        DisplayOneLineEditForContent _
            m_oFWHTML.ImageClick( _
                gk_sImagesPath & gk_sImageEdit, _
                GenerateEditOnclickCode(ksNameArrangement, EADType.eadArrangement)) & _
            m_oFWHTML.ImageClick( _
                gk_sImagesPath & gk_sIMAGEUpdate, _
                GenerateUpdateOnclickCode(ksNameArrangement))
        
    End If
    '   d29 crm     02Feb2000
    'Publication Note
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpPublicationNote) Then
        If DisplayLinkItemForEdit( _
            cpPublicationNote, _
            m_oFWAppData.GetDisplayInformation( _
            ESharedResourceFile.srfTabDisplayContentDetailsLeftHeaderPublicationNote), _
            oBOCatalogueContent.PublicationNote, oBOCatalogueContent.UnknownPublicationNote, oBOCatalogueContent.PublicationNoteLeaflet) Then

            DisplayOneLineHeadingForContent ""

        End If

        DisplayOneLineEditForContent _
            m_oFWHTML.SubmitImage( _
            gk_sImagesPath & gk_sImagePickSFA, ksNameButtonAddTermPublicationNote, ksNameButtonAddTermPublicationNote) & _
            m_oFWHTML.SubmitImage( _
             gk_sIMAGEAddLeafletButton, ksNameButtonAddTermPNoteLeaflet, ksNameButtonAddTermPNoteLeaflet)
    End If
    
    'Unpublished Finding Aids
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpUnpublishedFindingAids) Then
        If DisplayLinkItemForEdit( _
            cpUnpublishedFindingAids, _
            m_oFWAppData.GetDisplayInformation( _
                ESharedResourceFile.srfTabDisplayContentDetailsLeftHeaderUnpublishedFindingAids), _
            oBOCatalogueContent.UnpublishedFindingAids, oBOCatalogueContent.UnknownUnpublishedFindingAids, oBOCatalogueContent.UnpublishedFindingAidsLeaflet) Then

            DisplayOneLineHeadingForContent ""

        End If
        
        'Add new link button
        DisplayOneLineEditForContent _
            m_oFWHTML.SubmitImage( _
            gk_sImagesPath & gk_sImagePickSFA, ksNameButtonAddTermUnPublishedFindingAids, _
            ksNameButtonAddTermUnPublishedFindingAids) & _
            m_oFWHTML.SubmitImage( _
             gk_sIMAGEAddLeafletButton, ksNameButtonAddTermUnPFindingAidLeaflet, _
            ksNameButtonAddTermUnPFindingAidLeaflet)

    End If
    
    'Related Material
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpRelatedMaterials) Then
        
        DisplayRelatedSeparatedMaterialForEdit _
            cpRelatedMaterials, _
            m_oFWAppData.GetDisplayInformation( _
                ESharedResourceFile.srfTabDisplayContentDetailsLeftHeaderRelatedMaterial), _
            oBOCatalogueContent.RelatedMaterial, _
            oBOCatalogueContent.UnknownRelatedMaterial
        
    End If
    
    'Separate Material
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpSeperatedMaterials) Then
        
        DisplayRelatedSeparatedMaterialForEdit _
            cpSeperatedMaterials, _
            m_oFWAppData.GetDisplayInformation( _
                ESharedResourceFile.srfTabDisplayContentDetailsLeftHeaderSeparateMaterial), _
            oBOCatalogueContent.SeparateMaterial, _
            oBOCatalogueContent.UnknownSeparateMaterial
        
    End If
        
    'Link for the cancel button
    sParamString = gk_sLinkDisplayViewEditorialCatalogueDetails & _
                    "?" & gk_sParamEditSetID & "=" & t_lEditSetID & _
                    "&" & gk_sParamCatalogueID & "=" & t_lCatalogueID & _
                    "&" & gk_sParamCatalogueLevel & "=" & t_eCatalogueLevelID & _
                    "&" & gk_sParamDetailsTabNumber & "=" & ETabType.etTabContent & _
                    "&" & gk_sParamSectionType & "=" & ESectionType.SECeditorial & _
                    "&" & ksParamEditingCancelled & "=" & True
                    
    'Display OK and Cancel button
    DisplayOKCancelButtonsForEditAndCreateCatalogue (sParamString)
    
    'Release the object
    Set oBOCatalogueContent = Nothing
    
    'Generate the html
    sHTML = sHTML & m_oFWHTML.Paragraph(m_oFWTable.GetHTML) & m_oFWHTML.EndTag(gk_sHTMLForm)
    
    'Add the help button
    sHTML = sHTML & HTMLToAddHelpButton(gk_sHelpLinkToEditContentDetails)
    
    'Show the html page
    m_aspResponse.Write sHTML
    
    Exit Sub
    
ErrorHandler:
    Set oBOCatalogueContent = Nothing

    ConvertError sSource
End Sub
Private Sub DisplayEditAdminHistoryScreen( _
            ByVal t_eLoadType As ECatalogueEditFormLoadType, _
            ByVal t_lEditSetID As Long, _
            ByVal t_lCatalogueID As Long, _
            ByVal t_eCatalogueLevelID As ECATALOGUELEVEL)

'Description : Displays the admin history for editing for the selected catalogue
'
'Amendment history
'   d01 ndk     25Aug1999       Created
'   d15 ndk     07Dec1999       Implemented EAD Editing and made related changes
'   d16 ndk     08Dec1999       Added help button
'   d21 ndk     15Dec1999       Replacing single and double quotes with equivalent html codes before rendering ead data
'   d23 ndk     21Dec1999       Added security check
'   d49 tmj     30Mar2000       Fixed Display of EAD elements, by encoding hidden field as HTML
'
    On Error GoTo ErrorHandler

    'Variables
    Dim sSource As String, sHTML As String, sParamString As String
            
    'Business objects
    Dim oBOCatalogueAdminHistory As BOCatalogueAdminHistory, sValueForHiddenEADField As String
    Dim eSecurityCheckStatus As ESecurityCheckResult
    
    sSource = Me.Name & ".DisplayEditAdminHistoryScreen"

    'perform security check
    m_oBOEditorialPermissions.EditsetID = t_lEditSetID
    m_oBOEditorialPermissions.UserID = m_lLoggedUserID
    
    eSecurityCheckStatus = m_oBOEditorialPermissions.Check(EUserCanEditCatalogueEntryOnEditSet)
    
    If eSecurityCheckStatus <> ESecurityCheckResult.ESRSecurityCheckSuccessful Then
        RaiseError sSource, _
               GetErrorEnumForSecurityCheckResult(eSecurityCheckStatus), _
               GetDescriptionForSecurityCheckError( _
                        GetErrorEnumForSecurityCheckResult(eSecurityCheckStatus))
    End If

    'Instantiate the object
    Set oBOCatalogueAdminHistory = New BOCatalogueAdminHistory
    
    'Set the properties of the catalogue to be displayed
    If t_eLoadType = ECatalogueEditFormLoadType.etLoadFromDatabaseEditCatalogueEntry Then
        'The user clicked on Edit button and control has come here
        oBOCatalogueAdminHistory.CatalogueID = t_lCatalogueID
        oBOCatalogueAdminHistory.CatalogueLevelID = t_eCatalogueLevelID
        oBOCatalogueAdminHistory.EditsetID = t_lEditSetID
        oBOCatalogueAdminHistory.SectionType = ESectionType.SECeditorial
        
        'Get the details
        oBOCatalogueAdminHistory.Load t_eLoadType
    Else
        'it will be ECatalogueEditFormLoadType.etLoadFromSavedStateNoChangeEditCatalogueEntry
        
        oBOCatalogueAdminHistory.Load t_eLoadType
                
    End If
    
    'To display the details
    
    m_oFWTable.Reset    'Reset the table details
    sHTML = ""
    
    'Apply the style to the table
    m_oFWTable.TableStyle = m_typTableStyleContents
    
    'Link for submiting the form details
    sParamString = gk_sLinkSubmitCatalogueDetails & _
                    "?" & gk_sParamEditSetID & "=" & t_lEditSetID & _
                    "&" & gk_sParamCatalogueID & "=" & t_lCatalogueID & _
                    "&" & gk_sParamCatalogueLevel & "=" & t_eCatalogueLevelID & _
                    "&" & gk_sParamDetailsTabNumber & "=" & ETabType.etTabAdminHistory
                    
    'Start the form
    sHTML = sHTML & m_oFWHTML.FormStart(sParamString, gk_sMainFrame, gk_sHTMLDefaultFormName, MPost, True)
        
    'Admin History
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpAdministrative_BiographicalHistory) Then
        'Heading
        DisplayOneLineHeadingForContent _
            m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfTabDisplayAdminHistoryLeftHeaderAdminHistory)
        
        'get the value in string
        sValueForHiddenEADField = VariantStringAsString(oBOCatalogueAdminHistory.AdminHistory)
        
        'replace single and double quotes with html code equivalent
        sValueForHiddenEADField = Replace(sValueForHiddenEADField, "'", gk_sHTMLApostropheDecimalCode)
        sValueForHiddenEADField = Replace(sValueForHiddenEADField, """", gk_sHTMLQuotationDecimalCode)
        
        'Render current html and insert the activex control for editing xml
        DisplayOneLineEditForContent _
            RenderXML(VariantStringAsString(oBOCatalogueAdminHistory.AdminHistory)) & _
            m_oFWHTML.HiddenField(ksNameAdminHistory, m_oFWHTML.EncodeAsHTML(sValueForHiddenEADField)) & _
            m_oFWHTML.HiddenField(ksNameHiddenEADUpdatedAdminHistory, "false") & _
            m_oFWHTML.EADEditorLauncher( _
                gk_sHTMLDefaultFormName, _
                gk_sHTMLDefaultFormName & "." & ksNameAdminHistory, _
                gk_sHTMLDefaultFormName & "." & ksNameHiddenEADUpdatedAdminHistory, _
                gk_sHTMLDefaultXMLEditorLauncherControlName & ksNameAdminHistory)
            
        'blank heading
        DisplayOneLineHeadingForContent ""
        
        'edit and update buttons
        DisplayOneLineEditForContent _
            m_oFWHTML.ImageClick( _
                gk_sImagesPath & gk_sImageEdit, _
                GenerateEditOnclickCode(ksNameAdminHistory, EADType.eadAdministrativeHistory)) & _
            m_oFWHTML.ImageClick( _
                gk_sImagesPath & gk_sIMAGEUpdate, _
                GenerateUpdateOnclickCode(ksNameAdminHistory))

    End If
    
    'Link for the cancel button
    sParamString = gk_sLinkDisplayViewEditorialCatalogueDetails & _
                    "?" & gk_sParamEditSetID & "=" & t_lEditSetID & _
                    "&" & gk_sParamCatalogueID & "=" & t_lCatalogueID & _
                    "&" & gk_sParamCatalogueLevel & "=" & t_eCatalogueLevelID & _
                    "&" & gk_sParamDetailsTabNumber & "=" & ETabType.etTabAdminHistory & _
                    "&" & gk_sParamSectionType & "=" & ESectionType.SECeditorial & _
                    "&" & ksParamEditingCancelled & "=" & True

                    
    'Display OK and Cancel button
    DisplayOKCancelButtonsForEditAndCreateCatalogue (sParamString)
    
    'Release the object
    Set oBOCatalogueAdminHistory = Nothing
    
    'Generate the html
    sHTML = sHTML & m_oFWHTML.Paragraph(m_oFWTable.GetHTML) & m_oFWHTML.EndTag(gk_sHTMLForm)
    
    'Add the help button
    sHTML = sHTML & HTMLToAddHelpButton(gk_sHelpLinkToEditAdminHistoryDetails)
    
    'Show the html page
    m_aspResponse.Write sHTML
    
    Exit Sub
    
ErrorHandler:
    Set oBOCatalogueAdminHistory = Nothing
    
    ConvertError sSource
End Sub

Private Sub DisplayEditIndexTermDetailsScreen( _
            ByVal t_eLoadType As ECatalogueEditFormLoadType, _
            ByVal t_lEditSetID As Long, _
            ByVal t_lCatalogueID As Long, _
            ByVal t_eCatalogueLevelID As ECATALOGUELEVEL, _
            Optional ByVal t_lLinkTermID As Long, _
            Optional ByVal t_eLinkTermType As ELinkTermType)
    
'Description : Displays the index terms for editing for the selected catalogue
'
'Amendment history
'   d01 ndk     25Aug1999       Created
'   d16 ndk     08Dec1999       Added help button
'   d23 ndk     21Dec1999       Added security check
'
    On Error GoTo ErrorHandler

    'Variables
    Dim sSource As String, sHTML As String, sParamString As String
    
    'Business objects
    Dim oBOCatalogueIndexTerms As BOCatalogueIndexTerms
    Dim eSecurityCheckStatus As ESecurityCheckResult
    
    sSource = Me.Name & "DisplayEditIndexTermDetailsScreen"

    'perform security check
    m_oBOEditorialPermissions.EditsetID = t_lEditSetID
    m_oBOEditorialPermissions.UserID = m_lLoggedUserID
    
    eSecurityCheckStatus = m_oBOEditorialPermissions.Check(EUserCanEditCatalogueEntryOnEditSet)
    
    If eSecurityCheckStatus <> ESecurityCheckResult.ESRSecurityCheckSuccessful Then
        RaiseError sSource, _
               GetErrorEnumForSecurityCheckResult(eSecurityCheckStatus), _
               GetDescriptionForSecurityCheckError( _
                        GetErrorEnumForSecurityCheckResult(eSecurityCheckStatus))
    End If

    'Instantiate the object
    Set oBOCatalogueIndexTerms = New BOCatalogueIndexTerms
    
    'Set the properties of the catalogue to be displayed
    If t_eLoadType = ECatalogueEditFormLoadType.etLoadFromDatabaseEditCatalogueEntry Then
        
        'The user clicked on Edit button and control has come here
        oBOCatalogueIndexTerms.CatalogueID = t_lCatalogueID
        oBOCatalogueIndexTerms.CatalogueLevelID = t_eCatalogueLevelID
        oBOCatalogueIndexTerms.EditsetID = t_lEditSetID
        oBOCatalogueIndexTerms.SectionType = ESectionType.SECeditorial
        
        'Get the details
        oBOCatalogueIndexTerms.Load t_eLoadType
    Else
        Select Case t_eLoadType
        
            Case ECatalogueEditFormLoadType.etLoadFromSavedStateAndAddTermEditCatalogueEntry
                
                'The user has selected a pick term to add and the control has come here
                oBOCatalogueIndexTerms.Load t_eLoadType, t_lLinkTermID, t_eLinkTermType
            
            Case ECatalogueEditFormLoadType.etLoadFromSavedStateAndRemoveTermEditCatalogueEntry
                
                'The user has selected a pick term to remove and the control has come here
                oBOCatalogueIndexTerms.Load t_eLoadType, t_lLinkTermID, t_eLinkTermType
                
            Case ECatalogueEditFormLoadType.etLoadFromSavedStateNoChangeEditCatalogueEntry
                
                'The user has selected cancel operation while adding/removing a link term and the control has come here
                oBOCatalogueIndexTerms.Load t_eLoadType
                
        End Select
    
    End If
    'To display the details
    
    m_oFWTable.Reset    'Reset the table details
    sHTML = ""
    
    'Apply the style to the table
    m_oFWTable.TableStyle = m_typTableStyleContents
    
    'Link for submiting the form details
    sParamString = gk_sLinkSubmitCatalogueDetails & _
                    "?" & gk_sParamEditSetID & "=" & t_lEditSetID & _
                    "&" & gk_sParamCatalogueID & "=" & t_lCatalogueID & _
                    "&" & gk_sParamCatalogueLevel & "=" & t_eCatalogueLevelID & _
                    "&" & gk_sParamDetailsTabNumber & "=" & ETabType.etTabIndexTerms
                    
    'Start the form
    sHTML = sHTML & m_oFWHTML.FormStart(sParamString, gk_sMainFrame, gk_sHTMLDefaultFormName, MPost, True)
        
    'Corporate Body
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpIndexTermCorporateName) Then
       
        If DisplayLinkItemForEdit( _
            cpIndexTermCorporateName, _
            m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfTabDisplayIndexTermsLeftHeaderCorporateNames), _
            oBOCatalogueIndexTerms.CorporateBodyNames) Then
                        
            DisplayOneLineHeadingForContent ""
        End If
        
        'Add new link button
        DisplayOneLineEditForContent _
            m_oFWHTML.SubmitImage( _
            gk_sImagesPath & gk_sImagePickCorporateBody, ksNameButtonAddTermIndexTermCorporateBody, ksNameButtonAddTermIndexTermCorporateBody)

    End If
    
    'Person Name
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpIndexTermPersonName) Then
       
        If DisplayLinkItemForEdit( _
            cpIndexTermPersonName, _
            m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfTabDisplayIndexTermsLeftHeaderPersonalNames), _
            oBOCatalogueIndexTerms.PersonalNames) Then
                        
            DisplayOneLineHeadingForContent ""
        End If
        
        'Add new link button
        DisplayOneLineEditForContent _
            m_oFWHTML.SubmitImage( _
            gk_sImagesPath & gk_sImagePickPerson, ksNameButtonAddTermIndexTermPerson, ksNameButtonAddTermIndexTermPerson)

    End If
    
    'Place
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpIndexTermPlaceName) Then
       
        If DisplayLinkItemForEdit( _
            cpIndexTermPlaceName, _
            m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfTabDisplayIndexTermsLeftHeaderPlaces), _
            oBOCatalogueIndexTerms.Places) Then
                        
            DisplayOneLineHeadingForContent ""
        End If
        
        'Add new link button
        DisplayOneLineEditForContent _
            m_oFWHTML.SubmitImage( _
            gk_sImagesPath & gk_sImagePickPlace, ksNameButtonAddTermIndexTermPlace, ksNameButtonAddTermIndexTermPlace)

    End If
    
    'Subject
    If m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(t_eCatalogueLevelID, cpIndexTermSubject) Then
       
        If DisplayLinkItemForEdit( _
            cpIndexTermSubject, _
            m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfTabDisplayIndexTermsLeftHeaderSubject), _
            oBOCatalogueIndexTerms.Subjects) Then
                        
            DisplayOneLineHeadingForContent ""
        End If
        
        'Add new link button
        DisplayOneLineEditForContent _
            m_oFWHTML.SubmitImage( _
            gk_sImagesPath & gk_sImagePickSubject, ksNameButtonAddTermIndexTermSubject, ksNameButtonAddTermIndexTermSubject)

    End If
    
    'Link for the cancel button
    sParamString = gk_sLinkDisplayViewEditorialCatalogueDetails & _
                    "?" & gk_sParamEditSetID & "=" & t_lEditSetID & _
                    "&" & gk_sParamCatalogueID & "=" & t_lCatalogueID & _
                    "&" & gk_sParamCatalogueLevel & "=" & t_eCatalogueLevelID & _
                    "&" & gk_sParamDetailsTabNumber & "=" & ETabType.etTabIndexTerms & _
                    "&" & gk_sParamSectionType & "=" & ESectionType.SECeditorial & _
                    "&" & ksParamEditingCancelled & "=" & True

                    
    'Display OK and Cancel button
    DisplayOKCancelButtonsForEditAndCreateCatalogue (sParamString)
    
    'Release the object
    Set oBOCatalogueIndexTerms = Nothing
    
    'Generate the html
    sHTML = sHTML & m_oFWHTML.Paragraph(m_oFWTable.GetHTML) & m_oFWHTML.EndTag(gk_sHTMLForm)
    
    'Add the help button
    sHTML = sHTML & HTMLToAddHelpButton(gk_sHelpLinkToEditIndexTermsDetails)
    
    'Show the html page
    m_aspResponse.Write sHTML
    
    Exit Sub
    
ErrorHandler:
    Set oBOCatalogueIndexTerms = Nothing
    
    ConvertError sSource
End Sub


Private Sub DisplayOneLineHeadingAndDataForContent(ByVal t_vHeading As Variant, ByVal t_vDetails As Variant)
'
'Description : Displays the heading and content - Used while viewing
'
'Amendment history
'   d1  ndk     25Aug1999       Created

    On Error GoTo ErrorHandler
   
    Dim sSource As String
    
    sSource = Me.Name & ".DisplayOneLineHeadingAndDataForContent"
    
    'Display heading
    DisplayOneLineHeadingForContent t_vHeading
    
    'Display content
    DisplayOneLineDataForContent t_vDetails
    
    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub

Private Sub DisplayOneLineHeadingForContent(ByVal t_vHeading As Variant)
'
'Description : Displays the heading on the left side of a new row
'
'Amendment history
'   d1  ndk     25Aug1999       Created

    On Error GoTo ErrorHandler
   
    Dim sSource As String
    
    sSource = Me.Name & ".DisplayOneLineHeadingForContent"
    
    m_oFWTable.StartRow
    
    'Show the heading
    m_oFWTable.AddCellAsHTML _
        m_oFWHTML.FontWithStyle( _
            m_oFWHTML.EncodeAsHTML(CStr(t_vHeading)), m_typFontStyleLeftHeading), m_typCellStyleLeftHeading
    
    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub

Private Sub DisplayOneLineDataForContent(ByVal t_vDetails As Variant)
'
'Description : Displays the contents on the right side of the current row
'              This function is used when link items are being viewed
'
'Amendment history
'   d1  ndk     25Aug1999       Created

    On Error GoTo ErrorHandler
   
    Dim sSource As String
    
    sSource = Me.Name & ".DisplayOneLineDataForContent"
    
    'Add cell to display content in the current row
    m_oFWTable.AddCellAsHTML _
        m_oFWHTML.FontWithStyle( _
            m_oFWHTML.EncodeAsHTML(CStr(t_vDetails)), m_typFontStyleRightDetails), m_typCellStyleRightDetails
    
    m_oFWTable.EndRow
    
    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub


Private Sub DisplayOneLineEditAndDisplayForContent( _
        ByVal t_vDisplayDetails As Variant, _
        ByVal t_vRemoveButton As Variant, _
        Optional ByVal t_vHiddenField As Variant = "", _
        Optional ByVal t_vTextBox As Variant = "", _
        Optional ByVal t_vAddButton As Variant = "")
'
'Description : Displays the contents on the right side of the current row, used to display link items
'              This function is used when link items are being edited
'Amendment history
'   d1  ndk     25Aug1999       Created
'   d43 tmj     03Mar2000       Moved textbox so it appears after the text

    On Error GoTo ErrorHandler
   
    Dim sSource As String
    
    sSource = Me.Name & ".DisplayOneLineEditForContent"
        
    'Check if text box is to be displayed
    If t_vTextBox <> "" Then
        'Display the content - ie details of the link item
        m_typCellStyleRightDetails.lCOLSPAN = 1
        m_typCellStyleRightDetails.sWidth = "40%"
        m_oFWTable.AddCellAsHTML _
            m_oFWHTML.FontWithStyle( _
                m_oFWHTML.EncodeAsHTML(CStr(t_vDisplayDetails)), _
                m_typFontStyleRightDetails), m_typCellStyleRightDetails
                
        'Moved textbox so appears after text to remain consistent with LLD
        'd43 Display the text box for optional description
        m_typCellStyleRightDetails.lCOLSPAN = 1
        m_typCellStyleRightDetails.sWidth = "20%"
        m_oFWTable.AddCellAsHTML _
            m_oFWHTML.FontWithStyle( _
                CStr(t_vTextBox) & CStr(t_vHiddenField), m_typFontStyleRightDetails), _
                m_typCellStyleRightDetails
                
    Else
        'No text box
        
        'Display the content - ie details of the link item
        'The hidden field is attached to the content
        m_typCellStyleRightDetails.lCOLSPAN = 2
        m_typCellStyleRightDetails.sWidth = "40%"
        m_oFWTable.AddCellAsHTML _
            m_oFWHTML.FontWithStyle( _
            m_oFWHTML.EncodeAsHTML(CStr(t_vDisplayDetails)) & _
            CStr(t_vHiddenField), m_typFontStyleRightDetails), m_typCellStyleRightDetails
    End If
    
    'Add link button  (Only for related and separated material)
    m_typCellStyleRightDetails.lCOLSPAN = 1
    m_typCellStyleRightDetails.sWidth = "10%"
    If t_vAddButton = "" Then
        m_oFWTable.AddCellAsHTML _
            m_oFWHTML.FontWithStyle(gk_sHTMLNonBreakSpace, m_typFontStyleRightDetails), m_typCellStyleRightDetails
    Else
        m_oFWTable.AddCellAsHTML _
            m_oFWHTML.FontWithStyle(CStr(t_vAddButton), m_typFontStyleRightDetails), m_typCellStyleRightDetails
    End If
    
    'Remove link Button
    m_typCellStyleRightDetails.lCOLSPAN = 1
    m_typCellStyleRightDetails.sWidth = "10%"
    m_oFWTable.AddCellAsHTML _
        m_oFWHTML.FontWithStyle(CStr(t_vRemoveButton), m_typFontStyleRightDetails), _
        m_typCellStyleRightDetails
    
    m_oFWTable.EndRow
    
    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub
Private Sub DisplayOneLineHeadingEditAndDisplayForContent( _
        ByVal t_vHeading As Variant, _
        ByVal t_vDisplayDetails As Variant, _
        ByVal t_vRemoveButton As Variant, _
        Optional ByVal t_vHiddenField As Variant = "", _
        Optional ByVal t_vTextBox As Variant = "", _
        Optional ByVal t_vAddButton As Variant = "")

'
'Description : Displays the heading on the left side and content on the right side
'              A new row is started before displaying
'              This function is used when link items are being edited
'Amendment history
'   d1  ndk     25Aug1999       Created

    On Error GoTo ErrorHandler
   
    Dim sSource As String
    
    sSource = Me.Name & ".DisplayOneLineEditForContent"

    m_oFWTable.StartRow
    
    'Display the heading
    m_oFWTable.AddCellAsHTML _
        m_oFWHTML.FontWithStyle( _
            m_oFWHTML.EncodeAsHTML(CStr(t_vHeading)), m_typFontStyleLeftHeading), m_typCellStyleLeftHeading
    
    'call the function to display the contents
    DisplayOneLineEditAndDisplayForContent _
         t_vDisplayDetails, t_vRemoveButton, t_vHiddenField, t_vTextBox, t_vAddButton

    Exit Sub

ErrorHandler:
    ConvertError sSource
End Sub

        
Private Sub DisplayOneLineEditForContent(ByVal t_vDetails As Variant)
'
'Description : Displays the content on the right side on the current row
'              The displayed data generally is a html element like textbox
'              This function is used when catalogue entry is being edited
'
'Amendment history
'   d1  ndk     25Aug1999       Created

    On Error GoTo ErrorHandler
   
    Dim sSource As String
    
    sSource = Me.Name & ".DisplayOneLineEditForContent"

    '4 because there are max four cells while editing on the right side
    m_typCellStyleRightDetails.lCOLSPAN = 4
    
    m_oFWTable.AddCellAsHTML CStr(t_vDetails), m_typCellStyleRightDetails
    
    'End the current row
    m_oFWTable.EndRow
    
    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub


Private Sub DisplayRecordStatus(t_sHeading As String, _
                t_bAccessionedAsCopy As Boolean, _
                t_bExtractOnTransfer As Boolean, _
                t_bAR As Boolean, _
                t_bOrderable As Boolean)
'   Display the checkboxes representing the record status flags
'
'Amendment history
'   d20     WDP     10-Dec-1999      Created

    On Error GoTo ErrorHandler
   
    Dim sSource As String
    Dim sHTML As String
    
    sSource = Me.Name & ".DisplayRecordStatus"
    m_oFWTable.StartRow
    Dim sHeading As String
    sHeading = m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfTabDisplayAccessDetailsLeftHeaderRecordStatus)
    m_oFWTable.AddCellAsHTML _
        m_oFWHTML.FontWithStyle( _
            m_oFWHTML.EncodeAsHTML(t_sHeading), m_typFontStyleLeftHeading), m_typCellStyleLeftHeading
    
    Dim sAccessionAsCopyHeading As String
    Dim sOrderableHeading As String
    Dim sARHeading As String
    Dim sExtractAtTransferHeading As String
    
    sAccessionAsCopyHeading = m_oFWHTML.EncodeAsHTML(m_oFWAppData.GetDisplayInformation( _
        ESharedResourceFile.srfTabDisplayAccessDetailsAccessionedAsCopy))
    sOrderableHeading = m_oFWHTML.EncodeAsHTML(m_oFWAppData.GetDisplayInformation( _
        ESharedResourceFile.srfTabDisplayAccessDetailsOrderable))
    sARHeading = m_oFWHTML.EncodeAsHTML(m_oFWAppData.GetDisplayInformation( _
        ESharedResourceFile.srfTabDisplayAccessDetailsAR))
    sExtractAtTransferHeading = m_oFWHTML.EncodeAsHTML(m_oFWAppData.GetDisplayInformation( _
        ESharedResourceFile.srfTabDisplayAccessDetailsExtractAtTransfer))
    
    Dim sHTMLText As String
    sHTMLText = m_oFWHTML.Checkbox( _
                m_oFWHTML.FontWithStyle(sAccessionAsCopyHeading, m_typFontStyleRightDetails), _
                t_sName:=gk_sFORMRecordStatusAccessionedAsCopy, _
                t_sValue:=mk_sValue, _
                t_bChecked:=t_bAccessionedAsCopy)
    sHTMLText = sHTMLText & m_oFWHTML.Checkbox( _
                m_oFWHTML.FontWithStyle(sOrderableHeading, m_typFontStyleRightDetails), _
                t_sName:=gk_sFORMRecordStatusOrderable, _
                t_sValue:=mk_sValue, _
                t_bChecked:=t_bOrderable)
    sHTMLText = sHTMLText & m_oFWHTML.Checkbox( _
                m_oFWHTML.FontWithStyle(sARHeading, m_typFontStyleRightDetails), _
                t_sName:=gk_sFORMRecordStatusAR, _
                t_sValue:=mk_sValue, _
                t_bChecked:=t_bAR)
    sHTMLText = sHTMLText & m_oFWHTML.Checkbox( _
                m_oFWHTML.FontWithStyle(sExtractAtTransferHeading, m_typFontStyleRightDetails), _
                t_sName:=gk_sFORMRecordStatusExtractAtTransfer, _
                t_sValue:=mk_sValue, _
                t_bChecked:=t_bExtractOnTransfer)
    
    m_oFWTable.AddCellAsHTML sHTMLText, m_typCellStyleRightDetails
    
    m_oFWTable.EndRow
    
    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub

Private Sub DisplayOKCancelButtonsForEditAndCreateCatalogue(ByVal t_sCancelLink As String)
'
'Description : Displays the OK,Cancel button when a catalogue is being craeated and edited
'              OK will submit the catalogue, t_sCancelLink  is the link where the control should
'              go when cancel is clicked
'Amendment history
'   d1  ndk     25Aug1999       Created
'   d25 ndk     27Dec1999       Added tooltips

    On Error GoTo ErrorHandler
   
    Dim sSource As String
    
    sSource = Me.Name & ".DisplayOKCancelButtonsForEditAndCreateCatalogue"
    
    m_oFWTable.StartRow
    
    'Display blank heading
    DisplayOneLineHeadingForContent ""
    
    'Display Ok and Cancel Button
    m_oFWTable.AddCellAsHTML _
        m_oFWHTML.SubmitImage(gk_sImagesPath & gk_sImageOK, ksNameButtonOk, ksNameButtonOk, "") & _
        m_oFWHTML.Hyperlink(t_sCancelLink, m_oFWHTML.Image("", gk_sImagesPath & gk_sImageCancel), gk_sMainFrame, , _
        GetToolTip(ETOOLTIP.ttButtonCancelHelp)), m_typCellStyleForOKCancel
        
    m_oFWTable.EndRow
                                      
    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub

        



Private Function DisplayLinkItemForEdit( _
    ByVal t_eProperty As ECATALOGUEPROPERTIES, _
    ByVal t_sHeading As String, _
    ByVal t_colFirstSet As Collection, _
    Optional ByVal t_vUnknownDescription As Variant, _
    Optional ByVal t_colSecondSet As Collection) As Boolean
'
'Description : Displays the link items (except related and separated materials) for editing
'
'Amendment history
'   d1  ndk     25Aug1999       Created
'   d29 crm     02Feb2000       Updated code to provide links for Leaflets to catalogues
On Error GoTo ErrorHandler

    Dim sSource As String
    Dim iCounter As Integer, bDisplayedContentAgainstHeading As Boolean
    Dim bHasOptionalDescription As Boolean, bSecondRecordlistPresent As Boolean, bUnknownDescriptionPresent As Boolean
    
    'button names
    Dim sRemoveButtonName As String
    
    'field names
    Dim sUnknownDescriptionFieldName As String, sFieldNameForDescription As String, sHiddenIdFieldName As String
    
    'others
    Dim colDetails As Variant, sContent As String, vOptionalDescription As Variant, sHiddenIDFieldValue As String
    
    sSource = Me.Name & "DisplayLinkItemForEdit"
    
    'Display Headings
    DisplayOneLineHeadingForContent t_sHeading
    
    bDisplayedContentAgainstHeading = False         'Whether content is displayed against heading
    bHasOptionalDescription = False                 'Does the property has optional desc
    bSecondRecordlistPresent = False                'Does the property has two recordlists
    bUnknownDescriptionPresent = False              'Does the property has unknown description
    
    'Does the property has unknown description
    If t_eProperty = ECATALOGUEPROPERTIES.cpImmediateSourceOfAcquisition Or _
        t_eProperty = ECATALOGUEPROPERTIES.cpCopiesInformation Or _
        t_eProperty = ECATALOGUEPROPERTIES.cpPublicationNote Or _
        t_eProperty = ECATALOGUEPROPERTIES.cpUnpublishedFindingAids Then
        
        bUnknownDescriptionPresent = True
    End If
    
    'Does the property has optional desc
    If t_eProperty = cpImmediateSourceOfAcquisition Or _
        t_eProperty = cpCopiesInformation Or _
        t_eProperty = cpLocationOfOriginals Or _
        t_eProperty = cpPublicationNote Or _
        t_eProperty = cpUnpublishedFindingAids Then
        
        bHasOptionalDescription = True
    End If
    
    'Does the property has two recordlists
    If t_eProperty = ECATALOGUEPROPERTIES.cpImmediateSourceOfAcquisition Or _
        t_eProperty = ECATALOGUEPROPERTIES.cpCreatorNames Or t_eProperty = ECATALOGUEPROPERTIES.cpPublicationNote Or _
        t_eProperty = ECATALOGUEPROPERTIES.cpUnpublishedFindingAids Then
       
        bSecondRecordlistPresent = True
    End If
    
    'If unknown description is present
    If bUnknownDescriptionPresent Then      'get the name of the field for that
        Select Case t_eProperty
            Case ECATALOGUEPROPERTIES.cpImmediateSourceOfAcquisition
                sUnknownDescriptionFieldName = ksNameUnknownImmediateSourceOfAcquisition
            Case ECATALOGUEPROPERTIES.cpCopiesInformation
                sUnknownDescriptionFieldName = ksNameUnknownCopiesInformation
            Case ECATALOGUEPROPERTIES.cpPublicationNote
                sUnknownDescriptionFieldName = ksNameUnknownPublicationNote
            Case ECATALOGUEPROPERTIES.cpUnpublishedFindingAids
                sUnknownDescriptionFieldName = ksNameUnknownUnPublishedFindingAids
        End Select
                
        'Edit unknown unknown description
        DisplayOneLineEditForContent _
            m_oFWHTML.InputField(ITText, sUnknownDescriptionFieldName, _
                kiSizeLinkItemDescription, gk_iDBCOLLenLinkItemUnknownDescription, _
                VariantStringAsString(t_vUnknownDescription), VCNone, False)
        
        bDisplayedContentAgainstHeading = True
    End If

    'First Set
    iCounter = 0
    'Scan through the collection
    For Each colDetails In t_colFirstSet
        'Get the content to be displayed
        iCounter = iCounter + 1
        'Get the content, description and id fields for each of the record in the collection
        Select Case t_eProperty
            Case ECATALOGUEPROPERTIES.cpImmediateSourceOfAcquisition
                sContent = m_oVWSharedCatalogueEntry.AFTermSFADescription( _
                    colDetails, True, AFCorporateBody)
                vOptionalDescription = colDetails.sOptionalDescription
                sHiddenIdFieldName = ksNameIDImmediateSourceOfAcquisitionCorporateBody & LTrim(Str(iCounter))
                sHiddenIDFieldValue = colDetails.lID
                sFieldNameForDescription = ksNameDescriptionImmediateSourceOfAcquisitionCorporateBody & iCounter
                sRemoveButtonName = ksNameButtonRemoveTermImmSrcCorpBody & colDetails.lID
                
            Case ECATALOGUEPROPERTIES.cpCopiesInformation
                sContent = m_oVWSharedCatalogueEntry.AFTermSFADescription( _
                    colDetails, True, AFCorporateBody)
                vOptionalDescription = colDetails.sOptionalDescription
                sHiddenIdFieldName = ksNameIDCopiesInformation & LTrim(Str(iCounter))
                sHiddenIDFieldValue = colDetails.lID
                sFieldNameForDescription = ksNameDescriptionCopiesInformation & iCounter
                sRemoveButtonName = ksNameButtonRemoveTermCopiesInformation & colDetails.lID
               
            Case ECATALOGUEPROPERTIES.cpLocationOfOriginals
                sContent = m_oVWSharedCatalogueEntry.AFTermSFADescription( _
                    colDetails, True, AFCorporateBody)
                vOptionalDescription = colDetails.sOptionalDescription
                sHiddenIdFieldName = ksNameIDLocationOfOriginals & LTrim(Str(iCounter))
                sHiddenIDFieldValue = colDetails.lID
                sFieldNameForDescription = ksNameDescriptionLocationOfOriginals & iCounter
                sRemoveButtonName = ksNameButtonRemoveTermLocationOfOriginals & colDetails.lID
                
            Case ECATALOGUEPROPERTIES.cpCreatorNames
                sContent = m_oVWSharedCatalogueEntry.AFTermSFADescription( _
                    colDetails, True, AFCorporateBody)
                sHiddenIdFieldName = ksNameIDCreatorNamesCorporateBody & LTrim(Str(iCounter))
                sHiddenIDFieldValue = colDetails.lID
                sRemoveButtonName = ksNameButtonRemoveTermCreatorNamesCorpBody & colDetails.lID
        
            Case ECATALOGUEPROPERTIES.cpPlaceOfDeposit
                sContent = m_oVWSharedCatalogueEntry.AFTermSFADescription( _
                    colDetails, True, AFCorporateBody)
                sHiddenIdFieldName = ksNameIDPlaceOfDeposit & LTrim(Str(iCounter))
                sHiddenIDFieldValue = colDetails.lID
                sRemoveButtonName = ksNameButtonRemoveTermPlaceOfDeposit & colDetails.lID

            Case ECATALOGUEPROPERTIES.cpPublicationNote
                sContent = m_oVWSharedCatalogueEntry.AFTermSFADescription( _
                    colDetails, False, , True)
                vOptionalDescription = colDetails.sOptionalDescription
                sHiddenIdFieldName = ksNameIDPublicationNote & LTrim(Str(iCounter))
                sHiddenIDFieldValue = colDetails.lID
                sFieldNameForDescription = ksNameDescriptionPublicationNote & iCounter
                sRemoveButtonName = ksNameButtonRemoveTermPublicationNote & colDetails.lID

            Case ECATALOGUEPROPERTIES.cpUnpublishedFindingAids
                sContent = m_oVWSharedCatalogueEntry.AFTermSFADescription( _
                       colDetails, False, , True)
                vOptionalDescription = colDetails.sOptionalDescription
                sHiddenIdFieldName = ksNameIDUnpublishedFindingAids & LTrim(Str(iCounter))
                sHiddenIDFieldValue = colDetails.lID
                sFieldNameForDescription = ksNameDescriptionUnpublishedFindingAids & iCounter
                sRemoveButtonName = ksNameButtonRemoveTermUnpublishedFindingAids & colDetails.lID

            Case ECATALOGUEPROPERTIES.cpIndexTermCorporateName
                sContent = m_oVWSharedCatalogueEntry.AFTermSFADescription( _
                    colDetails, True, AFCorporateBody)
                sHiddenIdFieldName = ksNameIDIndexTermCorporateBody & LTrim(Str(iCounter))
                sHiddenIDFieldValue = colDetails.lID
                sRemoveButtonName = ksNameButtonRemoveTermIndexTermCorporateBody & colDetails.lID
            
            Case ECATALOGUEPROPERTIES.cpIndexTermPersonName
                sContent = m_oVWSharedCatalogueEntry.AFTermSFADescription( _
                    colDetails, True, AFPerson)
                sHiddenIdFieldName = ksNameIDIndexTermPerson & LTrim(Str(iCounter))
                sHiddenIDFieldValue = colDetails.lID
                sRemoveButtonName = ksNameButtonRemoveTermIndexTermPerson & colDetails.lID
            
            Case ECATALOGUEPROPERTIES.cpIndexTermPlaceName
                sContent = m_oVWSharedCatalogueEntry.AFTermSFADescription(colDetails, True, EAFTermType.AFPlace)
                sHiddenIdFieldName = ksNameIDIndexTermPlace & LTrim(Str(iCounter))
                sHiddenIDFieldValue = colDetails.lID
                sRemoveButtonName = ksNameButtonRemoveTermIndexTermPlace & colDetails.lID
            
            Case ECATALOGUEPROPERTIES.cpIndexTermSubject
                sContent = m_oVWSharedCatalogueEntry.AFTermSFADescription(colDetails, True, EAFTermType.AFSubject)
                sHiddenIdFieldName = ksNameIDIndexTermSubject & LTrim(Str(iCounter))
                sHiddenIDFieldValue = colDetails.lID
                sRemoveButtonName = ksNameButtonRemoveTermIndexTermSubject & colDetails.lID

        End Select
        
        'format the optional description
        vOptionalDescription = VariantStringAsString(vOptionalDescription)
        
        If bDisplayedContentAgainstHeading Then
            'Display with blank heading
            DisplayOneLineHeadingEditAndDisplayForContent _
                "", _
                sContent, _
                m_oFWHTML.SubmitImage(gk_sImagesPath & gk_sImageRemove, sRemoveButtonName, sRemoveButtonName), _
                m_oFWHTML.HiddenField(sHiddenIdFieldName, sHiddenIDFieldValue), _
                IIf(bHasOptionalDescription, _
                    m_oFWHTML.InputField(ITText, sFieldNameForDescription, , , CStr(vOptionalDescription)), _
                    "")
        Else
            'Display only the contents
            DisplayOneLineEditAndDisplayForContent _
                sContent, _
                m_oFWHTML.SubmitImage(gk_sImagesPath & gk_sImageRemove, sRemoveButtonName, sRemoveButtonName), _
                m_oFWHTML.HiddenField(sHiddenIdFieldName, sHiddenIDFieldValue), _
                IIf(bHasOptionalDescription, _
                    m_oFWHTML.InputField(ITText, sFieldNameForDescription, , , CStr(vOptionalDescription)), _
                    "")
                
            bDisplayedContentAgainstHeading = True
        End If
    
    Next
    
    'If a second recordlist is present
    If bSecondRecordlistPresent Then
        iCounter = 0
        For Each colDetails In t_colSecondSet
            'Get the content to be displayed
            iCounter = iCounter + 1
            'Get the content, description and id fields for each of the record in the collection
            Select Case t_eProperty
                
                Case ECATALOGUEPROPERTIES.cpImmediateSourceOfAcquisition
                    sContent = m_oVWSharedCatalogueEntry.AFTermSFADescription( _
                    colDetails, True, AFPerson)
                    vOptionalDescription = colDetails.sOptionalDescription
                    sHiddenIdFieldName = ksNameIDImmediateSourceOfAcquisitionPerson & LTrim(Str(iCounter))
                    sHiddenIDFieldValue = colDetails.lID
                    sFieldNameForDescription = ksNameDescriptionImmediateSourceOfAcquisitionPerson & iCounter
                    sRemoveButtonName = ksNameButtonRemoveTermImmSrcPerson & colDetails.lID
                    
                Case ECATALOGUEPROPERTIES.cpCreatorNames
                    sContent = m_oVWSharedCatalogueEntry.AFTermSFADescription( _
                    colDetails, True, AFPerson)
                    sHiddenIdFieldName = ksNameIDCreatorNamesPerson & LTrim(Str(iCounter))
                    sHiddenIDFieldValue = colDetails.lID
                    sRemoveButtonName = ksNameButtonRemoveTermCreatorNamesPerson & colDetails.lID
                '   d29 crm     02Feb2000
                Case ECATALOGUEPROPERTIES.cpPublicationNote
                    sContent = colDetails.sTitle
                    sHiddenIdFieldName = ksNameIDPublicationNoteLeaflet & LTrim(Str(iCounter))
                    sHiddenIDFieldValue = colDetails.lID
                    sFieldNameForDescription = ksNameDescPublicationNoteLeaflet & iCounter
                    sRemoveButtonName = ksNameButtonRemoveTermPNoteLeaflet & colDetails.lID
                    vOptionalDescription = colDetails.sOptionalDescription
                Case ECATALOGUEPROPERTIES.cpUnpublishedFindingAids
                    sContent = colDetails.sTitle
                    sHiddenIdFieldName = ksNameIDUnpublishedFindingAidsLeaflet & LTrim(Str(iCounter))
                    sHiddenIDFieldValue = colDetails.lID
                    sFieldNameForDescription = ksNameDescUnpublishedFindingAidsLeaflet & iCounter
                    sRemoveButtonName = ksNameButtonRemoveTermUnPFindingAidLeaflet & colDetails.lID
                    vOptionalDescription = colDetails.sOptionalDescription
            End Select
            
            'format the optional description
            vOptionalDescription = VariantStringAsString(vOptionalDescription)
                              
            If bDisplayedContentAgainstHeading Then
                'Display with blank heading
                DisplayOneLineHeadingEditAndDisplayForContent _
                    "", _
                    sContent, _
                    m_oFWHTML.SubmitImage(gk_sImagesPath & gk_sImageRemove, sRemoveButtonName, sRemoveButtonName), _
                    m_oFWHTML.HiddenField(sHiddenIdFieldName, sHiddenIDFieldValue), _
                    IIf(bHasOptionalDescription, _
                        m_oFWHTML.InputField(ITText, sFieldNameForDescription, , , CStr(vOptionalDescription)), _
                        "")
            Else
                'Display only the contents
                DisplayOneLineEditAndDisplayForContent _
                    sContent, _
                    m_oFWHTML.SubmitImage(gk_sImagesPath & gk_sImageRemove, sRemoveButtonName, sRemoveButtonName), _
                    m_oFWHTML.HiddenField(sHiddenIdFieldName, sHiddenIDFieldValue), _
                    IIf(bHasOptionalDescription, _
                        m_oFWHTML.InputField(ITText, sFieldNameForDescription, , , CStr(vOptionalDescription)), _
                        "")
                    
                bDisplayedContentAgainstHeading = True
            End If
        Next
    End If
    
    DisplayLinkItemForEdit = bDisplayedContentAgainstHeading
    
    Exit Function
ErrorHandler:
    ConvertError sSource
End Function
    


Private Function DisplayRelatedSeparatedMaterialForEdit( _
    ByVal t_eProperty As ECATALOGUEPROPERTIES, _
    ByVal t_sHeading As String, _
    ByVal t_colFirstSet As Collection, _
    ByVal t_vUnknownDescription As Variant) As Boolean
'
'Description : Displays the link items of related and separated materials for editing
'
'Amendment history
'   d1  ndk     25Aug1999       Created

On Error GoTo ErrorHandler

    Dim sSource As String
    Dim iCounter As Integer
    Dim sRemoveButtonHTML, sAddButtonHTML
    Dim sButtonHTML As String, sHiddenFieldHTML As String
    
    'types
    Dim typRelatedseparatedMaterial As RelatedSeparateMaterial
    
    'field names for data
    Dim sFieldNameForDescription As String, sUnknownDescriptionFieldName As String
    Dim sHiddenIdFieldName As String
    Dim sHiddenLevelFieldName As String
    
    'button names
    Dim sRemoveButtonName As String, sAddTermButtonName As String
            
    'for content and display
    Dim sContent As String, vOptionalDescription As Variant, sHiddenIDFieldValue As String
    Dim sHiddenLevelFieldValue As String
        
    sSource = Me.Name & ".DisplayRelatedSeparatedMaterialForEdit"
    
    'Display Headings
    DisplayOneLineHeadingForContent t_sHeading
    
    'Get the name of the field for unknown description
    If t_eProperty = ECATALOGUEPROPERTIES.cpRelatedMaterials Then
        sUnknownDescriptionFieldName = ksNameUnknownRelatedMaterial
    Else
        sUnknownDescriptionFieldName = ksNameUnknownSeparateMaterial
    End If
    
    'Edit unknown unknown description
    DisplayOneLineEditForContent _
        m_oFWHTML.InputField(ITText, sUnknownDescriptionFieldName, _
            kiSizeLinkItemDescription, gk_iDBCOLLenLinkItemUnknownDescription, _
            VariantStringAsString(t_vUnknownDescription), VCNone, False)
    
    'Scan through the collection
    For iCounter = 1 To t_colFirstSet.Count
    
        'Identify and construct the field names and button names
        If t_eProperty = ECATALOGUEPROPERTIES.cpRelatedMaterials Then
            sHiddenIdFieldName = ksNameIDRelatedMaterial & iCounter
            sHiddenLevelFieldName = ksNameLevelRelatedMaterial & iCounter
            sFieldNameForDescription = ksNameDescriptionRelatedMaterial & iCounter
            sRemoveButtonName = ksNameButtonRemoveTermRelatedMaterial & iCounter
        Else
            sHiddenIdFieldName = ksNameIDSeparateMaterial & iCounter
            sHiddenLevelFieldName = ksNameLevelSeparateMaterial & iCounter
            sFieldNameForDescription = ksNameDescriptionSeparateMaterial & iCounter
            sRemoveButtonName = ksNameButtonRemoveTermSeparateMaterial & iCounter
        End If
    
        'Get the content to be displayed
        typRelatedseparatedMaterial = t_colFirstSet.Item(iCounter)
            
        'Extract the description
        vOptionalDescription = VariantStringAsString(typRelatedseparatedMaterial.sOptionalDescription)
        
        'The links for that description
        sContent = _
        GenerateCatalogueReferenceFromID( _
            typRelatedseparatedMaterial.lCatalogueID, _
            GetCatalogueLevelCode(CInt(typRelatedseparatedMaterial.iCatalogueLevelID)))
        
            
        'values for the fields
        sHiddenIDFieldValue = typRelatedseparatedMaterial.lCatalogueID
        sHiddenLevelFieldValue = typRelatedseparatedMaterial.iCatalogueLevelID
                
        'Construct the remove button
        sRemoveButtonHTML = m_oFWHTML.SubmitImage(gk_sImagesPath & gk_sImageRemove, _
            sRemoveButtonName, sRemoveButtonName)
            
        'Hidden field (for the catalogueid and level)
        sHiddenFieldHTML = _
            m_oFWHTML.HiddenField(sHiddenIdFieldName, sHiddenIDFieldValue) & _
            m_oFWHTML.HiddenField(sHiddenLevelFieldName, sHiddenLevelFieldValue)
            
            DisplayOneLineHeadingEditAndDisplayForContent _
                "", _
                sContent, _
                sRemoveButtonHTML, _
                sHiddenFieldHTML, _
                m_oFWHTML.InputField(ITText, sFieldNameForDescription, 50, 50, CStr(vOptionalDescription))
            
    Next
    
    DisplayOneLineHeadingForContent ""      'display blank heading
    
    If t_eProperty = ECATALOGUEPROPERTIES.cpRelatedMaterials Then
        sAddTermButtonName = ksNameButtonAddTermRelatedMaterial
    Else
        sAddTermButtonName = ksNameButtonAddTermSeparateMaterial
    End If
            
    'Display the 'Add' button
    DisplayOneLineEditForContent _
        m_oFWHTML.SubmitImage(gk_sImagesPath & gk_sImageAdd, sAddTermButtonName, sAddTermButtonName)

    Exit Function
ErrorHandler:
    ConvertError sSource
End Function

Private Function StoreFormLinkItemsInCollection(t_eLinkTermType As ELinkTermType) As Collection
'
'Description : Stores the details of the link items from the form into a collection
'
'Amendment history
'   d1  ndk     25Aug1999       Created
'   d28 ndk     30Dec1999       Rectfied error - not releasing the  collection

On Error GoTo ErrorHandler

    Dim sSource As String
    Dim typGenericLinkItem As GenericLinkRecord
    Dim colCatalogueId As Collection, typRelatedSeparateMaterial As RelatedSeparateMaterial
    Dim sFieldNameForDescription  As String, sFieldNameForID As String, sFieldNameForLevelID As String
    Dim iMainCounter As Integer
    
    sSource = Me.Name & ".StoreFormLinkItemsInCollection"
    
    Set StoreFormLinkItemsInCollection = New Collection
        
    'Get the field name for id, level (for catalogues) and description for the link term
    Select Case t_eLinkTermType
        Case ELinkTermType.elImmedieteSourceOfAcquisitionPerson
            sFieldNameForID = ksNameIDImmediateSourceOfAcquisitionPerson
            sFieldNameForDescription = ksNameDescriptionImmediateSourceOfAcquisitionPerson
        Case ELinkTermType.elImmedieteSourceOfAcquisitionCorporateBody
            sFieldNameForID = ksNameIDImmediateSourceOfAcquisitionCorporateBody
            sFieldNameForDescription = ksNameDescriptionImmediateSourceOfAcquisitionCorporateBody
        Case ELinkTermType.elCopiesInformation
            sFieldNameForID = ksNameIDCopiesInformation
            sFieldNameForDescription = ksNameDescriptionCopiesInformation
        Case ELinkTermType.elLocationOfOriginals
            sFieldNameForID = ksNameIDLocationOfOriginals
            sFieldNameForDescription = ksNameDescriptionLocationOfOriginals
        Case ELinkTermType.elCreatorNamesPerson
            sFieldNameForID = ksNameIDCreatorNamesPerson
        Case ELinkTermType.elCreatorNamesCorporateBody
            sFieldNameForID = ksNameIDCreatorNamesCorporateBody
        Case ELinkTermType.elPlaceOfDeposit
            sFieldNameForID = ksNameIDPlaceOfDeposit
        Case ELinkTermType.elPublicationNote
            sFieldNameForID = ksNameIDPublicationNote
            sFieldNameForDescription = ksNameDescriptionPublicationNote
        Case ELinkTermType.elUnPublishedFindingAids
            sFieldNameForID = ksNameIDUnpublishedFindingAids
            sFieldNameForDescription = ksNameDescriptionUnpublishedFindingAids
        Case ELinkTermType.elRelatedMaterial
            sFieldNameForID = ksNameIDRelatedMaterial
            sFieldNameForLevelID = ksNameLevelRelatedMaterial
            sFieldNameForDescription = ksNameDescriptionRelatedMaterial
        Case ELinkTermType.elSeparateMaterial
            sFieldNameForID = ksNameIDSeparateMaterial
            sFieldNameForLevelID = ksNameLevelSeparateMaterial
            sFieldNameForDescription = ksNameDescriptionSeparateMaterial
        Case ELinkTermType.elIndexTermCorporateBody
            sFieldNameForID = ksNameIDIndexTermCorporateBody
        Case ELinkTermType.elIndexTermPerson
            sFieldNameForID = ksNameIDIndexTermPerson
        Case ELinkTermType.elIndexTermPlace
            sFieldNameForID = ksNameIDIndexTermPlace
        Case ELinkTermType.elIndexTermSubject
            sFieldNameForID = ksNameIDIndexTermSubject
        Case ELinkTermType.elRelatedMaterial
            sFieldNameForID = ksNameIDRelatedMaterial
        Case ELinkTermType.elPublicationNoteLeaflet
            sFieldNameForID = ksNameIDPublicationNoteLeaflet
            sFieldNameForDescription = ksNameDescPublicationNoteLeaflet
        Case ELinkTermType.elUnPublishedFindingAidLeaflet
            sFieldNameForID = ksNameIDUnpublishedFindingAidsLeaflet
            sFieldNameForDescription = ksNameDescUnpublishedFindingAidsLeaflet
    End Select
    
    'For related and separated materials
    If t_eLinkTermType = ELinkTermType.elRelatedMaterial Or t_eLinkTermType = ELinkTermType.elSeparateMaterial Then
        
        iMainCounter = 1
        
        'Main counter for different descriptions
        Do While Not IsEmpty(m_aspRequest.Form(sFieldNameForID & iMainCounter))
            
            'Get the description,id and level
            typRelatedSeparateMaterial.sOptionalDescription = m_aspRequest.Form(sFieldNameForDescription & iMainCounter)
            typRelatedSeparateMaterial.lCatalogueID = m_aspRequest.Form(sFieldNameForID & iMainCounter)
            typRelatedSeparateMaterial.iCatalogueLevelID = m_aspRequest.Form(sFieldNameForLevelID & iMainCounter)
            
            'Add the record to the collection
            StoreFormLinkItemsInCollection.Add typRelatedSeparateMaterial
            
            iMainCounter = iMainCounter + 1     'Next description
        Loop
    Else        'For other link items
        
        iMainCounter = 1
        
        'Process all the links
        Do While Not IsEmpty(m_aspRequest.Form(sFieldNameForID & iMainCounter))
            
            'Get the id
            typGenericLinkItem.lID = CLng(m_aspRequest.Form(sFieldNameForID & iMainCounter))
            
            If Not sFieldNameForDescription = "" Then
                'Description
                typGenericLinkItem.sOptionalDescription = m_aspRequest.Form(sFieldNameForDescription & iMainCounter)
            End If
            
            'Add the id and description record to the collection
            StoreFormLinkItemsInCollection.Add typGenericLinkItem
            
            iMainCounter = iMainCounter + 1 'Next id
        Loop
    End If
    
    Exit Function
ErrorHandler:
    ConvertError sSource
End Function

Public Sub DisplayAuditTrailForCatalogue()
'
'Description :Displays the audit trail for a catalogue, only the latest 2 entries are shown for any catalogue
'
'Amendment history
'   d01  ndk   25Aug1999   Created
'   d16  ndk   08Dec1999   Added help button
'   d25  ndk   27Dec1999   Added tooltips
'   d27  ndk   29Dec1999   Using resource file for audit trail related display constants
'   d42  tmj   02Mar2000   Modified so links work correctly
'   d48 daj     27Mar2000       Tidied up the audit trail and fixed defect 766 - Audit trail 'back' links not working
'                               once you move off the summary tab
'   d54 daj     20Jun2000       Removed formatting of 'View this version' link

    On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim sHTML As String, sParamString   As String, bEditSetChanged As Boolean
    Dim lCatalogueID As Long, eCatalogueLevelID As ECATALOGUELEVEL, lEditsetID  As Long
    Dim sCatalogueReference As String
    
    'd42 parameters to hold current editset information
    Dim lCurrEditsetID As Long
    
    Dim oFWRecordlist As FWRecordList
    
    Dim oBOBasicCatalogueEntry As BOBasicCatalogueEntry
    Dim ofwTable As FWTable
    Dim typCellStyle As CellStyle
    Dim typTableStyle As TableStyle
    Dim typFontStyle As FontStyle
    
    sSource = Me.Name & ".DisplayAuditTrailForCatalogue"
    
    'Read details from the query string
    lCatalogueID = m_aspRequest.QueryString(gk_sParamCatalogueID)
    eCatalogueLevelID = m_aspRequest.QueryString(gk_sParamCatalogueLevel)
    'Edit set id is required to go back to the original catalogue being viewed
    lEditsetID = m_aspRequest.QueryString(gk_sParamEditSetID)

    'Set the current parameters when they are requested
    lCurrEditsetID = lEditsetID
    
    'Generate catalogue reference
    sCatalogueReference = GenerateCatalogueReferenceFromID(lCatalogueID, eCatalogueLevelID)
    
    'Initialise the objects
    Set ofwTable = New FWTable
    Set oBOBasicCatalogueEntry = New BOBasicCatalogueEntry
    
    'Table style for heading
    typTableStyle = m_oFWHTML.GetTableStyle(TABContentStyle)
    typTableStyle.lCELLSPACING = 2
    typTableStyle.lCELLPADDING = 0
    ofwTable.TableStyle = typTableStyle
    
    'Cell and font style for the heading
    typCellStyle = m_oFWHTML.GetCellStyle(CELLNormalStyle)
    typFontStyle = m_oFWHTML.GetFontStyle(FONTNormalStyle)
    typFontStyle.bIsBold = True
    
    ofwTable.StartRow
    
    'Heading
    typCellStyle.sWidth = "30%"
    typCellStyle.eALIGN = HALLeft
    ofwTable.AddCellAsHTML _
        m_oFWHTML.FontWithStyle( _
        m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayAuditTrailCaptionHistorOfModifications) & _
          " " & sCatalogueReference, typFontStyle), typCellStyle

    ofwTable.AddCellAsHTML " ", typCellStyle
    
    'Hyperlink back to catalogue in editorial
    typCellStyle.eALIGN = HALRight
    
    sParamString = gk_sLinkDisplayViewEditorialCatalogueDetails & _
        "?" & gk_sParamEditSetID & "=" & lEditsetID & _
        "&" & gk_sParamCatalogueID & "=" & lCatalogueID & _
        "&" & gk_sParamCatalogueLevel & "=" & eCatalogueLevelID & _
        "&" & ksParamLoadType & "=" & ECatalogueEditFormLoadType.etLoadFromDatabaseEditCatalogueEntry & _
        "&" & ksParamCatalogueEntryEditType & "=" & ECatalogueEntryEditType.ecEditCatalogueEntry & _
        "&" & gk_sParamSectionType & "=" & ESectionType.SECeditorial & _
        "&" & gk_sParamDetailsTabNumber & "=" & ETabType.etTabSummary

    ofwTable.AddCellAsHTML _
        m_oFWHTML.Hyperlink(sParamString, _
            m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayAuditTrailCaptionBackTo) & _
            sCatalogueReference, _
            gk_sMainFrame, , GetToolTip(ETOOLTIP.ttBackToEditSetCatalogue)), typCellStyle
            
    ofwTable.EndRow
    
    'Store the heading details in html
    sHTML = m_oFWHTML.Paragraph(ofwTable.GetHTML)
    
    'Reset the table
    ofwTable.Reset
    
    'Table style for contents
    typTableStyle = m_oFWHTML.GetTableStyle(TABContentStyle)
    typTableStyle.lCELLSPACING = 2
    typTableStyle.lCELLPADDING = 0
    ofwTable.TableStyle = typTableStyle
    
    'Set the styles to display column headings
    typFontStyle = m_oFWHTML.GetFontStyle(FONTColumnHeadingStyle)
    typCellStyle = m_oFWHTML.GetCellStyle(CELLColumnHeadingStyle)
    
    ofwTable.StartRow
    
    'EditSet Name
    typCellStyle.sWidth = "20%"
    ofwTable.AddCellAsHTML _
        m_oFWHTML.FontWithStyle( _
        m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayAuditTrailColEditSet) _
        , typFontStyle), typCellStyle
            
    'From
    typCellStyle.sWidth = "10%"
    ofwTable.AddCellAsHTML _
        m_oFWHTML.FontWithStyle( _
        m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayAuditTrailColFrom) _
        , typFontStyle), typCellStyle
       
       
    'To
    typCellStyle.sWidth = "10%"
    ofwTable.AddCellAsHTML _
        m_oFWHTML.FontWithStyle( _
        m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayAuditTrailColTo) _
        , typFontStyle), typCellStyle
       
       
    'Comments
    typCellStyle.sWidth = "30%"
    ofwTable.AddCellAsHTML _
        m_oFWHTML.FontWithStyle( _
        m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayAuditTrailColComments) _
        , typFontStyle), typCellStyle
                
    'Role
    typCellStyle.sWidth = "10%"
    ofwTable.AddCellAsHTML _
        m_oFWHTML.FontWithStyle( _
        m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayAuditTrailColRole) _
        , typFontStyle), typCellStyle
        
    'Person
    typCellStyle.sWidth = "5%"
    ofwTable.AddCellAsHTML _
        m_oFWHTML.FontWithStyle( _
        m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayAuditTrailColPerson) _
        , typFontStyle), typCellStyle
        
    'Blank coloumn
    typCellStyle = m_oFWHTML.GetCellStyle(CELLNormalStyle)
    typCellStyle.sWidth = "15%"
    ofwTable.AddCellAsHTML _
        m_oFWHTML.FontWithStyle(gk_sHTMLNonBreakSpace, typFontStyle), typCellStyle
    
    ofwTable.EndRow
    
    'Set the styles to display details
    typFontStyle = m_oFWHTML.GetFontStyle(FONTGeneralContentStyle)
    typCellStyle = m_oFWHTML.GetCellStyle(CELLGeneralContentStyle)
            
    'Get the audit trail details
    Set oFWRecordlist = oBOBasicCatalogueEntry.AuditTrailForCatalogue(lCatalogueID, eCatalogueLevelID)
        
    If oFWRecordlist.RecordCount <> 0 Then
        
        'Go to the top of the
        oFWRecordlist.MoveFirst
        
        lEditsetID = 0
        
        'Loop through all the records
        
        ofwTable.StartRow
        
        Do While Not oFWRecordlist.IsAtEnd
        
            'On the same editset still
            bEditSetChanged = False
            
            'Re set the cell style
            typCellStyle = m_oFWHTML.GetCellStyle(CELLGeneralContentStyle)
            typCellStyle.sWidth = ""    'This is done to continue with the cellwidth of the coloumn headings
            
            'Next line would always be executed first time as lEditSetId is set to 0 at the beggining
            If lEditsetID = 0 Or lEditsetID <> oFWRecordlist.GetFieldValue(gk_sDBCOLEditSetID) Then   'Editset has changed
                
                bEditSetChanged = True
                
                'Get the new editset id and name
                lEditsetID = oFWRecordlist.GetFieldValue(gk_sDBCOLEditSetID)
                
                'Editset name
                ofwTable.AddCellAsHTML _
                    m_oFWHTML.FontWithStyle(oFWRecordlist.GetFieldValue(gk_sDBCOLEditsetName), typFontStyle), _
                    typCellStyle
                
                'Edit set start date
                ofwTable.AddCellAsHTML _
                    m_oFWHTML.FontWithStyle( _
                    VariantDateAsDDMMMYYYY(oFWRecordlist.GetFieldValue(gk_sDBCOLTempEditSetStartDate)), _
                    typFontStyle), typCellStyle
                
                'Edit set end date
                ofwTable.AddCellAsHTML _
                    m_oFWHTML.FontWithStyle( _
                    VariantDateAsDDMMMYYYY(oFWRecordlist.GetFieldValue(gk_sDBCOLTempEditSetEndDate)), _
                    typFontStyle), typCellStyle
                
                'Comments (description from edit_set)
                ofwTable.AddCellAsHTML _
                    m_oFWHTML.FontWithStyle(oFWRecordlist.GetFieldValue(gk_sDBCOLDescription), typFontStyle), _
                    typCellStyle
            Else
                ofwTable.AddCellAsHTML gk_sHTMLNonBreakSpace, typCellStyle  'Name
                ofwTable.AddCellAsHTML gk_sHTMLNonBreakSpace, typCellStyle  'from
                ofwTable.AddCellAsHTML gk_sHTMLNonBreakSpace, typCellStyle  'to
                ofwTable.AddCellAsHTML gk_sHTMLNonBreakSpace, typCellStyle  'comments
            End If
            
            'Role Name
            ofwTable.AddCellAsHTML _
                m_oFWHTML.FontWithStyle(oFWRecordlist.GetFieldValue(gk_sDBCOLRoleDescription), typFontStyle), _
                typCellStyle
            
            'Initials
            ofwTable.AddCellAsHTML _
                m_oFWHTML.FontWithStyle(oFWRecordlist.GetFieldValue(gk_sDBCOLUserInitials), typFontStyle), _
                typCellStyle
            
            'Hyperlink to catalogue details
            typCellStyle = m_oFWHTML.GetCellStyle(CELLNormalStyle)
            If bEditSetChanged Then         'Display link to the catalogue
            
                ' d48 - Removed current catID and level - they will be the same as the audit trail version
                sParamString = gk_sLinkDisplayViewEditorialCatalogueDetails & _
                    "?" & gk_sParamEditSetID & "=" & lEditsetID & _
                    "&" & gk_sParamCatalogueID & "=" & lCatalogueID & _
                    "&" & gk_sParamCatalogueLevel & "=" & eCatalogueLevelID & _
                    "&" & ksParamLoadType & "=" & ECatalogueEditFormLoadType.etLoadFromDatabaseEditCatalogueEntry & _
                    "&" & ksParamCatalogueEntryEditType & "=" & ECatalogueEntryEditType.ecEditCatalogueEntry & _
                    "&" & gk_sParamSectionType & "=" & ESectionType.SECAuditTrial & _
                    "&" & gk_sParamDetailsTabNumber & "=" & ETabType.etTabSummary & _
                    "&" & gk_sParamCurrEditSetID & "=" & lCurrEditsetID

                ' d54 - removed formatting
                ofwTable.AddCellAsHTML _
                    m_oFWHTML.Hyperlink( _
                        sParamString, _
                        m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfDisplayAuditTrailHyperLinkViewThisVersion), _
                        gk_sMainFrame, , GetToolTip(ETOOLTIP.ttViewAuditTrailCatalogue)), typCellStyle
            Else
                'Blank cell
                ofwTable.AddCellAsHTML _
                    m_oFWHTML.FontWithStyle(gk_sHTMLNonBreakSpace, typFontStyle), typCellStyle
            
            End If
            
            ofwTable.EndRow
            oFWRecordlist.MoveNext  'Next record
        Loop
    End If
            
    'Get the table details
    sHTML = sHTML & m_oFWHTML.Paragraph(ofwTable.GetHTML)
    
    'Add the help button
    sHTML = sHTML & HTMLToAddHelpButton(gk_sHelpLinkToViewAuditTrail)
    
    'Display the html
    m_aspResponse.Write sHTML
    
    'Release objects
    Set ofwTable = Nothing
    Set oFWRecordlist = Nothing
    Set oBOBasicCatalogueEntry = Nothing
    
    m_objObjectContext.SetComplete
    
    Exit Sub
    
ErrorHandler:
    m_objObjectContext.SetAbort
    ShowError sSource
End Sub

Private Sub LockCatalogueEntry( _
    ByVal t_eTabToLock As ETabType, _
    ByVal t_lCatalogueID As Long, _
    ByVal t_eCatalogueLevel As ECATALOGUELEVEL, _
    ByVal t_lEditSetID As Long)
'
'Description  : Trys to lock a catalogue in the editset, if the lock can can not be
''              obtained then an error is raised
'Amendment history
'   d2  ndk     16Oct1999       Created
'   d3  ndk     28Oct1999       Locking done at tab level

    On Error GoTo ErrorHandler
    
    Dim sSource As String
    
    'Business objects
    Dim oBOEditsetEntryLockManager As BOEditSetEntryLockManager
    
    sSource = Me.Name & ".LockCatalogueEntry"

    'Initialise the business object
    Set oBOEditsetEntryLockManager = New BOEditSetEntryLockManager
            
    'Try to obtain a lock on the catalogue entry
    If Not oBOEditsetEntryLockManager.LockEditSetCatalogueEntry( _
        t_eTabToLock, t_lCatalogueID, t_eCatalogueLevel, t_lEditSetID, _
        m_lLoggedUserID, m_objObjectContext.Item(gk_sSessionObject).SessionID) Then
    
        'Lock could not be obtained, get the details of the user who has locked the catalogue entry
        'and display appropriate message
                
        RaiseError sSource, peCatalogueAlreadyLocked, _
            "Tab " & oBOEditsetEntryLockManager.LockedTabName & " of Catalogue (" & GenerateCatalogueReferenceFromID(t_lCatalogueID, t_eCatalogueLevel) & ")" & _
            " you want to edit/move was locked by " & Trim(oBOEditsetEntryLockManager.LockedUserName) & _
            " on " & Trim(oBOEditsetEntryLockManager.LockedDateTime)
    End If
    
    'If the control comes here, it means that the catalogue was locked successfully
    
    'Release objects
    Set oBOEditsetEntryLockManager = Nothing

    Exit Sub
ErrorHandler:

    Set oBOEditsetEntryLockManager = Nothing
    
    ConvertError sSource
End Sub

Private Sub UnLockCatalogueEntry( _
    ByVal t_eTabToUnLock As ETabType, _
    ByVal t_lCatalogueID As Long, _
    ByVal t_eCatalogueLevel As ECATALOGUELEVEL, _
    ByVal t_lEditSetID As Long)
'
'Description  : Trys to unlock a catalogue in the editset, No error raised if there is no lock exising
'               This method is called when the user clicks on 'Cancel' button when editing a catalogue entry
'               When the cancel is clicked , the user is sent to the view catalogue details page'
'               In that page this method is called.
'               However as an added security the user can only unlock a catalogue only if he has locked it
'               and with the same session id (the user could log on to two pcs and work)
'Amendment history
'   d2  ndk     16Oct1999       Created
'   d3  ndk     28Oct1999       Locking done at tab level

    On Error GoTo ErrorHandler
    
    Dim sSource As String
    
    'Business objects
    Dim oBOEditsetEntryLockManager As BOEditSetEntryLockManager
    
    sSource = Me.Name & ".UnLockCatalogueEntry"

    'Initialise the business object
    Set oBOEditsetEntryLockManager = New BOEditSetEntryLockManager
            
    'Release the lock
    oBOEditsetEntryLockManager.UnlockLockEditSetCatalogueEntry _
        t_eTabToUnLock, t_lCatalogueID, t_eCatalogueLevel, t_lEditSetID, m_lLoggedUserID, _
        m_objObjectContext.Item(gk_sSessionObject).SessionID
    
    'Release objects
    Set oBOEditsetEntryLockManager = Nothing
    
    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub

Public Sub StartMoveCatalogueEntry()
'
' Starts the process of moving the catalogue entry from one catalogue to another catalogue
' Stores the current editset and other details in session
' Calls the asp page to display the catalogue browser for selection of parent
'
'Amendment history
'   d3  ndk     28Oct1999       Created
'   d16 ndk     08Dec1999       Passing the query string when redirecting to the catalogue browser
'   d23 ndk     21Dec1999       Added security check
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim lEditsetID As Long
    Dim lCatalogueID As Long
    Dim eCatalogueLevelID As ECATALOGUELEVEL
    Dim sQueryString As String
    Dim eSecurityCheckStatus As ESecurityCheckResult
    Dim eParentLevel As ECATALOGUELEVEL
    Dim lParentID As Long
    
    sSource = Me.Name & "StartMoveCatalogueEntry"
    
    'Store the current catalogue details in session
    'Get the data from the query string
    lEditsetID = m_aspRequest.QueryString(gk_sParamEditSetID)
    lCatalogueID = m_aspRequest.QueryString(gk_sParamCatalogueID)
    eCatalogueLevelID = m_aspRequest.QueryString(gk_sParamCatalogueLevel)
    
    'perform security check
    m_oBOEditorialPermissions.EditsetID = lEditsetID
    m_oBOEditorialPermissions.UserID = m_lLoggedUserID
    
    'move is possible if he can edit it
    eSecurityCheckStatus = m_oBOEditorialPermissions.Check(EUserCanEditCatalogueEntryOnEditSet)
    
    If eSecurityCheckStatus <> ESecurityCheckResult.ESRSecurityCheckSuccessful Then
        RaiseError sSource, _
               GetErrorEnumForSecurityCheckResult(eSecurityCheckStatus), _
               GetDescriptionForSecurityCheckError( _
                        GetErrorEnumForSecurityCheckResult(eSecurityCheckStatus))
    End If
    
    'Store these details in the session store
    StoreDataInSession dsEditCatalogueEditSetID, lEditsetID
    StoreDataInSession dsEditCatalogueCatalogueID, lCatalogueID
    StoreDataInSession dsEditCatalogueCatalogueLevelID, eCatalogueLevelID
    
    ' Get the parent
    GetParent eCatalogueLevelID, lCatalogueID, eParentLevel, lParentID
    
    'generate the query string to open the browser at the required level
    sQueryString = GenerateQueryStringsForCatalogueBrowser(lParentID, eParentLevel)
    
    m_aspResponse.Redirect _
        gk_sLinkDisplayCatalogueBrowserForMoveCatalogueEntryNewParentSelection & "?" & sQueryString

    Exit Sub
ErrorHandler:
    m_objObjectContext.SetAbort
    ShowError sSource
End Sub

Public Sub StartCreateNewCatalogueFromScratch()
'
' Starts the process of creation of new catalogue from scratch :-
' Stores the current editset and other details in session
' Calls the asp page to display the catalogue browser for selection of parent
'
'Amendment history
'   d3  ndk     28Oct1999       Created
'   d23 ndk     21Dec1999       Added security check

    On Error GoTo ErrorHandler
    
    Dim sSource As String, lEditsetID As Long
    Dim eSecurityCheckStatus As ESecurityCheckResult
    
    sSource = Me.Name & "StartCreateNewCatalogueFromScratch"
    
    'Get the data from the query string
    lEditsetID = m_aspRequest.QueryString(gk_sParamEditSetID)
    
    'perform security check
    m_oBOEditorialPermissions.EditsetID = lEditsetID
    m_oBOEditorialPermissions.UserID = m_lLoggedUserID
    
    eSecurityCheckStatus = m_oBOEditorialPermissions.Check(EUserCanCreateCatalogueEntryOnEditSet)
    
    If eSecurityCheckStatus <> ESecurityCheckResult.ESRSecurityCheckSuccessful Then
        RaiseError sSource, _
               GetErrorEnumForSecurityCheckResult(eSecurityCheckStatus), _
               GetDescriptionForSecurityCheckError( _
                        GetErrorEnumForSecurityCheckResult(eSecurityCheckStatus))
    End If
    
    'Store these details in the session store
    StoreDataInSession dlCreateCatalogueEditSetID, lEditsetID
    
    m_aspResponse.Redirect _
        gk_sLinkDisplayCatalogueBrowserForCreateFromScratchParentSelection
        
    Exit Sub
ErrorHandler:
    m_objObjectContext.SetAbort
    ShowError sSource
End Sub
Public Sub StartCreateNewCatalogueByCopying()
'
' Starts the process of creation of new catalogue by copying:-
' Stores the current editset and other details in session
' Calls the asp page to display the catalogue browser for selection of catalogue to be copied
'
'Amendment history
'   d3  ndk     28Oct1999       Created
'   d23 ndk     21Dec1999       Added security check
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String, lEditsetID As Long
    Dim eSecurityCheckStatus As ESecurityCheckResult

    sSource = Me.Name & "StartCreateNewCatalogueByCopying"
    
    'Store the current editset in session
    
    'Get the data from the query string
    lEditsetID = m_aspRequest.QueryString(gk_sParamEditSetID)
    
    'perform security check
    m_oBOEditorialPermissions.EditsetID = lEditsetID
    m_oBOEditorialPermissions.UserID = m_lLoggedUserID
    
    eSecurityCheckStatus = m_oBOEditorialPermissions.Check(EUserCanCreateCatalogueEntryOnEditSet)
    
    If eSecurityCheckStatus <> ESecurityCheckResult.ESRSecurityCheckSuccessful Then
        RaiseError sSource, _
               GetErrorEnumForSecurityCheckResult(eSecurityCheckStatus), _
               GetDescriptionForSecurityCheckError( _
                        GetErrorEnumForSecurityCheckResult(eSecurityCheckStatus))
    End If
    
    'Store these details in the session store
    StoreDataInSession dlCreateCatalogueEditSetID, lEditsetID

    m_aspResponse.Redirect _
        gk_sLinkDisplayCatalogueBrowserForCreateByCopyCatalogueSelection

    Exit Sub
ErrorHandler:
    m_objObjectContext.SetAbort
    ShowError sSource
End Sub



Public Sub StartAddFromExisting()
'
' Starts the process of adding an existing record to an edit set
' Stores the current editset and other details in session
' Calls the asp page to display the catalogue browser for selection of parent
'
'Amendment history
'   d57 19-Jan-2001 Ben Jessel      Created
    On Error GoTo ErrorHandler
    
    Dim sSource As String, lEditsetID As Long
    Dim eSecurityCheckStatus As ESecurityCheckResult
    
    sSource = Me.Name & "StartAddFromExisting"
    
    'Get the data from the query string
    lEditsetID = m_aspRequest.QueryString(gk_sParamEditSetID)
    
    'perform security check
    m_oBOEditorialPermissions.EditsetID = lEditsetID
    m_oBOEditorialPermissions.UserID = m_lLoggedUserID
    
    eSecurityCheckStatus = m_oBOEditorialPermissions.Check(EUserCanCreateCatalogueEntryOnEditSet)
    
    If eSecurityCheckStatus <> ESecurityCheckResult.ESRSecurityCheckSuccessful Then
        RaiseError sSource, _
               GetErrorEnumForSecurityCheckResult(eSecurityCheckStatus), _
               GetDescriptionForSecurityCheckError( _
                        GetErrorEnumForSecurityCheckResult(eSecurityCheckStatus))
    End If
    
    'Store these details in the session store
    StoreDataInSession dlCreateCatalogueEditSetID, lEditsetID
    
    m_aspResponse.Redirect _
        gk_sLinkDisplayCatalogueBrowserForAddExisting
        
    Exit Sub
ErrorHandler:
    m_objObjectContext.SetAbort
    ShowError sSource

End Sub

Public Sub MoveCatalogueToNewParent()
'
' Moves the selected catalogue entry into a new parent
'
'Amendment history
'   d3  ndk     28Oct1999       Created
'   d23 ndk     21Dec1999       Added security check
    On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim lCatalogueID As Long, eCatalogueLevelID As ECATALOGUELEVEL, lEditsetID As Long
    Dim lCurrentParentID, eCurrentParentLevelID As ECATALOGUELEVEL
    Dim lProposedParentID As Long, eProposedParentLevelID As ECATALOGUELEVEL
    Dim oBOCatalogueContext As BOCatalogueContext
    Dim eSecurityCheckStatus As ESecurityCheckResult

    sSource = Me.Name & ".MoveCatalogueToNewParent"

    'current catalogue details
    lEditsetID = GetDataFromSession(dsEditCatalogueEditSetID, m_objObjectContext.Item(gk_sSessionObject))
    lCatalogueID = GetDataFromSession(dsEditCatalogueCatalogueID, m_objObjectContext.Item(gk_sSessionObject))
    eCatalogueLevelID = GetDataFromSession(dsEditCatalogueCatalogueLevelID, m_objObjectContext.Item(gk_sSessionObject))  'Level of the catalogue like 1- lettercode, 2 -division
    
    'proposed parent catalogue details
    lProposedParentID = m_aspRequest.QueryString(gk_sParamCatalogueID)
    eProposedParentLevelID = m_aspRequest.QueryString(gk_sParamCatalogueLevel)
    
    'perform security check
    m_oBOEditorialPermissions.EditsetID = lEditsetID
    m_oBOEditorialPermissions.UserID = m_lLoggedUserID
    
    eSecurityCheckStatus = m_oBOEditorialPermissions.Check(EUserCanEditCatalogueEntryOnEditSet)
    
    If eSecurityCheckStatus <> ESecurityCheckResult.ESRSecurityCheckSuccessful Then
        RaiseError sSource, _
               GetErrorEnumForSecurityCheckResult(eSecurityCheckStatus), _
               GetDescriptionForSecurityCheckError( _
                        GetErrorEnumForSecurityCheckResult(eSecurityCheckStatus))
    End If
    
    'initialise
    Set oBOCatalogueContext = New BOCatalogueContext
    
    'set the properties
    oBOCatalogueContext.CatalogueID = lCatalogueID
    oBOCatalogueContext.CatalogueLevelID = eCatalogueLevelID
    oBOCatalogueContext.EditsetID = lEditsetID
    oBOCatalogueContext.ProposedImmedieteParentID = lProposedParentID
    oBOCatalogueContext.ProposedImmedieteParentLevelID = eProposedParentLevelID
    
    'call the method to move the catalogue entry
    oBOCatalogueContext.MoveCatalogueToNewParent
    
    m_aspResponse.Redirect _
        gk_sLinkDisplayViewEditorialCatalogueDetails & _
        "?" & gk_sParamEditSetID & "=" & lEditsetID & _
        "&" & gk_sParamCatalogueID & "=" & lCatalogueID & _
        "&" & gk_sParamCatalogueLevel & "=" & eCatalogueLevelID & _
        "&" & gk_sParamSectionType & "=" & ESectionType.SECeditorial & _
        "&" & gk_sParamDetailsTabNumber & "=" & ETabType.etTabContext
        
    'release objects
    Set oBOCatalogueContext = Nothing
    
    m_objObjectContext.SetComplete
    
    Exit Sub
ErrorHandler:
    Set oBOCatalogueContext = Nothing
    
    m_objObjectContext.SetAbort
    ShowError sSource
End Sub

Public Sub MoveMultiplePiecesToNewParent()
' Description
'
' Moves the selected pieces (and any child items) to a new parent
'
' Amendment history
'   d73 alm     19Oct2001       Created (based on MoveCatalogueToNewParent)
On Error GoTo ErrorHandler
    
    ' Declare locals
    Dim lNumOfCheckBoxes         As Long
    Dim lNumOfSelectedCatalogues As Long
    Dim sSource                  As String
    Dim sCatRefStored            As String
    Dim lCatalogueIDs()          As Long
    Dim eCatalogueLevels()       As ECATALOGUELEVEL
    Dim lCount                   As Long
    Dim lEditsetID               As Long
    Dim lProposedParentID        As Long
    Dim eProposedParentLevelID   As ECATALOGUELEVEL
    Dim eSecurityCheckStatus     As ESecurityCheckResult
    Dim oBOCatalogueContext      As BOCatalogueContext
    
    sSource = Me.Name & ".MoveMultiplePiecesToNewParent"

    ' Current catalogue details
    lEditsetID = m_aspRequest.QueryString(gk_sParamEditSetID)

    ' Proposed parent catalogue details
    lProposedParentID = m_aspRequest.QueryString(gk_sParamCatalogueID)
    eProposedParentLevelID = m_aspRequest.QueryString(gk_sParamCatalogueLevel)
    
    ' Perform security check
    m_oBOEditorialPermissions.EditsetID = lEditsetID
    m_oBOEditorialPermissions.UserID = m_lLoggedUserID

    eSecurityCheckStatus = m_oBOEditorialPermissions.Check(EUserCanEditCatalogueEntryOnEditSet)

    If eSecurityCheckStatus <> ESecurityCheckResult.ESRSecurityCheckSuccessful Then
        RaiseError sSource, _
               GetErrorEnumForSecurityCheckResult(eSecurityCheckStatus), _
               GetDescriptionForSecurityCheckError( _
                        GetErrorEnumForSecurityCheckResult(eSecurityCheckStatus))
    End If

    ' Pieces to move...
    ' Get the number of check boxes
    lNumOfCheckBoxes = IIf(IsEmpty(m_aspRequest(gk_sFORMCheckBoxCount)), 0, m_aspRequest(gk_sFORMCheckBoxCount))
        
    ' Loop through all the chek boxed and pick the checked ones and separate the values in the value field
    ' of the checkbox as Catalogue ID and Level
    lNumOfSelectedCatalogues = 0
    For lCount = 1 To lNumOfCheckBoxes

        sCatRefStored = m_aspRequest(gk_sFORMCheckboxName & Trim(CStr(lCount)))
        If Not IsEmpty(sCatRefStored) And sCatRefStored <> "" Then

            lNumOfSelectedCatalogues = lNumOfSelectedCatalogues + 1
            ReDim Preserve lCatalogueIDs(lNumOfSelectedCatalogues), eCatalogueLevels(lNumOfSelectedCatalogues)
            lCatalogueIDs(lNumOfSelectedCatalogues) = CLng(sCatRefStored)
            eCatalogueLevels(lNumOfSelectedCatalogues) = ECATALOGUELEVEL.ctPiece
            
        End If
    
    Next lCount

    ' If there were pieces selected then call the appropriate business object to move
    ' the pieces to a new parent
    If lNumOfSelectedCatalogues > 0 Then
        
        ' Initialise
        Set oBOCatalogueContext = New BOCatalogueContext
        
        ' Set the properties
        oBOCatalogueContext.EditsetID = lEditsetID
        oBOCatalogueContext.ProposedImmedieteParentID = lProposedParentID
        oBOCatalogueContext.ProposedImmedieteParentLevelID = eProposedParentLevelID

        ' Call the method to move the multiple pieces
        oBOCatalogueContext.MoveMultipleCatalogueToNewParent lCatalogueIDs, eCatalogueLevels
        
        ' Release objects
        Set oBOCatalogueContext = Nothing

    End If

    ' Return to the editset
    m_aspResponse.Redirect gk_sLinkDisplayEditSetDetails & _
                           "?" & gk_sParamEditSetID & "=" & lEditsetID

    m_objObjectContext.SetComplete
    
    Exit Sub
ErrorHandler:
    Set oBOCatalogueContext = Nothing
    m_objObjectContext.SetAbort
    ShowError sSource
End Sub

Public Sub DisplayConfirmationForCatalogueDeletion()
'
'Description :Displays the the confirmation dialog box when the user selects to delete a catalogue
'
'Amendment history
'   d01  ndk   25Aug1999   Created
'   d16  ndk   08Dec1999   Added help button
'   d18  ndk   09Dec1999   Put deletion messages into resource file
'   d23  ndk   21Dec1999   Added security check
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim sHTML As String, sOKLink As String, sCancelLink As String
    Dim lCatalogueID As Long, eCatalogueLevelID As ECATALOGUELEVEL, lEditsetID  As Long
    Dim sCatalogueReference As String, eEditSetEntryTypeID As ECATALOGUEENTRYSTATUS
    Dim sWarningMessage As String
    
    Dim oBOEditorialDBUtility As BOEditorialDBUtility
    Dim ofwTable As FWTable
    Dim typCellStyle As CellStyle
    Dim typTableStyle As TableStyle
    Dim typFontStyle As FontStyle
    Dim eSecurityCheckStatus As ESecurityCheckResult
    
    sSource = Me.Name & ".DisplayConfirmationForCatalogueDeletion"
    
    'Read details from the query string
    
    lCatalogueID = m_aspRequest.QueryString(gk_sParamCatalogueID)
    eCatalogueLevelID = m_aspRequest.QueryString(gk_sParamCatalogueLevel)
    lEditsetID = m_aspRequest.QueryString(gk_sParamEditSetID)

    'perform security check
    m_oBOEditorialPermissions.EditsetID = lEditsetID
    m_oBOEditorialPermissions.UserID = m_lLoggedUserID
    
    eSecurityCheckStatus = m_oBOEditorialPermissions.Check(EUserCanDeleteCatalogueEntryOnEditSet)
    
    If eSecurityCheckStatus <> ESecurityCheckResult.ESRSecurityCheckSuccessful Then
        RaiseError sSource, _
               GetErrorEnumForSecurityCheckResult(eSecurityCheckStatus), _
               GetDescriptionForSecurityCheckError( _
                        GetErrorEnumForSecurityCheckResult(eSecurityCheckStatus))
    End If

    'Generate catalogue reference
    sCatalogueReference = GenerateCatalogueReferenceFromID(lCatalogueID, eCatalogueLevelID)
    
    'Initialise the objects
    Set ofwTable = New FWTable
    Set oBOEditorialDBUtility = New BOEditorialDBUtility
    
    'get the type of catalogue entry
    eEditSetEntryTypeID = _
        oBOEditorialDBUtility.EditSetEntryTypeID(lEditsetID, lCatalogueID, eCatalogueLevelID)
        
    'Table style for heading
    typTableStyle = m_oFWHTML.GetTableStyle(TABContentStyle)
    typTableStyle.lCELLSPACING = 2
    typTableStyle.lCELLPADDING = 0
    typTableStyle.eALIGN = HALCenter
    
    ofwTable.TableStyle = typTableStyle
    
    'Cell and font style for the heading
    typCellStyle = m_oFWHTML.GetCellStyle(CELLNormalStyle)
    typFontStyle = m_oFWHTML.GetFontStyle(FONTNormalStyle)
    typFontStyle.bIsBold = True
    
    ofwTable.StartRow
    
    'Heading
    typCellStyle.eALIGN = HALCenter
    typCellStyle.sWidth = "50%"
    typCellStyle.eALIGN = HALCenter
    
    'warning message
    Select Case eEditSetEntryTypeID
        'accessioned
        Case ECATALOGUEENTRYSTATUS.cesImported, ECATALOGUEENTRYSTATUS.cesImportedModified
            'sWarningMessage = "This catalogue entry and all its children will be permanently  removed from the system :"
            sWarningMessage = m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfConfirmDeletionForAccessionedEntry)
        'created
        Case ECATALOGUEENTRYSTATUS.cesPROCATCreated, ECATALOGUEENTRYSTATUS.cesPROCATCreatedModified
            'sWarningMessage = "This catalogue entry will be permanently  removed from the system :"
            sWarningMessage = m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfConfirmDeletionForCreatedEntry)
        'added
        Case ECATALOGUEENTRYSTATUS.cesAddedFromLive, ECATALOGUEENTRYSTATUS.cesAddedFromLiveModified
            'sWarningMessage = "All the changes made to this catalogue entry will be discarded :"
            sWarningMessage = m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfConfirmDeletionForAddedEntry)
    End Select
        
    'Delete catalogue message
    ofwTable.AddCellAsHTML _
        m_oFWHTML.FontWithStyle(sWarningMessage & sCatalogueReference, typFontStyle), typCellStyle
    
    ofwTable.EndRow
    
    ofwTable.AddBlankRow
    
    ofwTable.StartRow
    
    sOKLink = _
        gk_sLinkDeleteEditSetCatalogue & "?" & _
        gk_sParamEditSetID & "=" & lEditsetID & "&" & _
        gk_sParamCatalogueID & "=" & lCatalogueID & "&" & _
        gk_sParamCatalogueLevel & "=" & eCatalogueLevelID
        
    sCancelLink = _
        gk_sLinkDisplayEditSetDetails & "?" & _
        gk_sParamEditSetID & "=" & lEditsetID
    
    'Display Ok and Cancel Button
    ofwTable.AddCellAsHTML _
        m_oFWHTML.ImageButton(gk_sImagesPath & gk_sImageOK, ksNameButtonOk, _
                    sOKLink, gk_sMainFrame) & _
        m_oFWHTML.ImageButton(gk_sImagesPath & gk_sImageCancel, ksNameButtonCancel, _
                    sCancelLink, gk_sMainFrame), typCellStyle
    ofwTable.EndRow
    
    'Get the table details
    sHTML = sHTML & m_oFWHTML.Paragraph(ofwTable.GetHTML)
    
    'Add the help button
    sHTML = sHTML & HTMLToAddHelpButton(gk_sHelpLinkToDeleteCatalogue)
    
    'Display the html
    m_aspResponse.Write sHTML
    
    'release objects
    Set oBOEditorialDBUtility = Nothing
    
    m_objObjectContext.SetComplete
    
    Exit Sub
    
ErrorHandler:
    
    Set oBOEditorialDBUtility = Nothing

    m_objObjectContext.SetAbort
    ShowError sSource
End Sub

Public Sub DisplayConfirmationForCatalogueUnDeletion()
'
'Description: Displays the the confirmation dialog box
'             when the user selects to delete a catalogue entry from
'             an editset that contains items to remove from the
'             live catalogue
'
'Amendment history
'   d01  alm   24Oct2001   Created (based on DisplayConfirmationForCatalogueDeletion)
'
On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim sHTML As String
    Dim sOKLink As String
    Dim sCancelLink As String
    Dim lCatalogueID As Long
    Dim eCatalogueLevelID As ECATALOGUELEVEL
    Dim lEditsetID  As Long
    Dim sCatalogueReference As String
    Dim sWarningMessage As String
    
    Dim ofwTable As FWTable
    Dim typCellStyle As CellStyle
    Dim typTableStyle As TableStyle
    Dim typFontStyle As FontStyle
    Dim eSecurityCheckStatus As ESecurityCheckResult
    
    sSource = Me.Name & ".DisplayConfirmationForCatalogueUnDeletion"
    
    ' Read details from the query string
    lCatalogueID = m_aspRequest.QueryString(gk_sParamCatalogueID)
    eCatalogueLevelID = m_aspRequest.QueryString(gk_sParamCatalogueLevel)
    lEditsetID = m_aspRequest.QueryString(gk_sParamEditSetID)

    ' Perform security check
    m_oBOEditorialPermissions.EditsetID = lEditsetID
    m_oBOEditorialPermissions.UserID = m_lLoggedUserID
    
    eSecurityCheckStatus = m_oBOEditorialPermissions.Check(EUserCanDeleteCatalogueEntryOnEditSet)
    
    If eSecurityCheckStatus <> ESecurityCheckResult.ESRSecurityCheckSuccessful Then
        RaiseError sSource, _
               GetErrorEnumForSecurityCheckResult(eSecurityCheckStatus), _
               GetDescriptionForSecurityCheckError( _
                        GetErrorEnumForSecurityCheckResult(eSecurityCheckStatus))
    End If

    ' Generate catalogue reference
    sCatalogueReference = GenerateCatalogueReferenceFromID(lCatalogueID, eCatalogueLevelID)
    
    ' Initialise the objects
    Set ofwTable = New FWTable
        
    ' Table style for heading
    typTableStyle = m_oFWHTML.GetTableStyle(TABContentStyle)
    typTableStyle.lCELLSPACING = 2
    typTableStyle.lCELLPADDING = 0
    typTableStyle.eALIGN = HALCenter
    
    ofwTable.TableStyle = typTableStyle
    
    ' Cell and font style for the heading
    typCellStyle = m_oFWHTML.GetCellStyle(CELLNormalStyle)
    typFontStyle = m_oFWHTML.GetFontStyle(FONTNormalStyle)
    typFontStyle.bIsBold = True
    
    ofwTable.StartRow
    
    ' Heading
    typCellStyle.eALIGN = HALCenter
    typCellStyle.sWidth = "50%"
    typCellStyle.eALIGN = HALCenter
    
    ' Warning message
    sWarningMessage = m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfConfirmDeletionFromEditSet)
        
    ' Delete catalogue message
    ofwTable.AddCellAsHTML _
        m_oFWHTML.FontWithStyle(sWarningMessage & sCatalogueReference, typFontStyle), typCellStyle
    
    ofwTable.EndRow
    ofwTable.AddBlankRow
    ofwTable.StartRow
    
    sOKLink = _
        gk_sLinkDeleteEditSetCatalogue & "?" & _
        gk_sParamEditSetID & "=" & lEditsetID & "&" & _
        gk_sParamCatalogueID & "=" & lCatalogueID & "&" & _
        gk_sParamCatalogueLevel & "=" & eCatalogueLevelID
        
    sCancelLink = _
        gk_sLinkDisplayEditSetDetails & "?" & _
        gk_sParamEditSetID & "=" & lEditsetID
    
    ' Display Ok and Cancel Button
    ofwTable.AddCellAsHTML _
        m_oFWHTML.ImageButton(gk_sImagesPath & gk_sImageOK, ksNameButtonOk, _
                    sOKLink, gk_sMainFrame) & _
        m_oFWHTML.ImageButton(gk_sImagesPath & gk_sImageCancel, ksNameButtonCancel, _
                    sCancelLink, gk_sMainFrame), typCellStyle
    ofwTable.EndRow
    
    ' Get the table details
    sHTML = sHTML & m_oFWHTML.Paragraph(ofwTable.GetHTML)
    
    ' Add the help button
    sHTML = sHTML & HTMLToAddHelpButton(gk_sHelpLinkToUnDeleteCatalogue)
    
    ' Display the html
    m_aspResponse.Write sHTML
    
    m_objObjectContext.SetComplete
    
    Exit Sub
    
ErrorHandler:
    m_objObjectContext.SetAbort
    ShowError sSource
End Sub

Public Sub DeleteEditSetCatalogue()
'
' delete the selected catalogue from the editset
'
'Amendment history
'   d3   ndk     28Oct1999       Created
'   d23  ndk     21Dec1999       Added security check
'

    On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim lCatalogueID As Long, eCatalogueLevelID As ECATALOGUELEVEL, lEditsetID As Long
    Dim oBODeleteCatalogue As BODeleteCatalogue
    Dim eSecurityCheckStatus As ESecurityCheckResult
    
    sSource = Me.Name & "DeleteEditSetCatalogue"
    
    'current catalogue details
    lEditsetID = m_aspRequest.QueryString(gk_sParamEditSetID)
    lCatalogueID = m_aspRequest.QueryString(gk_sParamCatalogueID)
    eCatalogueLevelID = m_aspRequest.QueryString(gk_sParamCatalogueLevel)
    
    'perform security check
    m_oBOEditorialPermissions.EditsetID = lEditsetID
    m_oBOEditorialPermissions.UserID = m_lLoggedUserID
    
    eSecurityCheckStatus = m_oBOEditorialPermissions.Check(EUserCanDeleteCatalogueEntryOnEditSet)
    
    If eSecurityCheckStatus <> ESecurityCheckResult.ESRSecurityCheckSuccessful Then
        RaiseError sSource, _
               GetErrorEnumForSecurityCheckResult(eSecurityCheckStatus), _
               GetDescriptionForSecurityCheckError( _
                        GetErrorEnumForSecurityCheckResult(eSecurityCheckStatus))
    End If
    
    'initialise
    Set oBODeleteCatalogue = New BODeleteCatalogue
    
    'call the method to delete  the catalogue entry
    
    oBODeleteCatalogue.EditsetID = lEditsetID
    oBODeleteCatalogue.CatalogueID = lCatalogueID
    oBODeleteCatalogue.CatalogueLevelID = eCatalogueLevelID
    
    oBODeleteCatalogue.DeleteEditSetCatalogue
    
    m_aspResponse.Redirect _
        gk_sLinkDisplayEditSetDetails & _
        "?" & gk_sParamEditSetID & "=" & lEditsetID
    
    'release objects
    Set oBODeleteCatalogue = Nothing
    
    m_objObjectContext.SetComplete
    
    Exit Sub
ErrorHandler:
    Set oBODeleteCatalogue = Nothing
    
    m_objObjectContext.SetAbort
    ShowError sSource
End Sub


Private Function GenerateUpdateOnclickCode( _
    t_sFieldName) As String
'
' Function returns the code to be used in the on click of update button
' which is used to update the value from the changes made in the editor to the browser
'
'   d18 ndk     09Dec1999       Created
'   d50 daj     05Apr2000       Changed GenerateUpdateOnclickCode to call VB rather than the control directly
'                               (see FWHTML.EADEditorLauncher)
   
    GenerateUpdateOnclickCode = gk_sNameUpdateFieldFunction & gk_sHTMLDefaultXMLEditorLauncherControlName & t_sFieldName & "(); "
    
End Function

Private Function GenerateEditOnclickCode( _
    t_sFieldName As String, _
    t_eEADType As EADType, _
    Optional t_bUseHiddenSwitchField As Boolean = False) As String
'
' Function returns code to be used in the on click event of the edit button
' This button is used to launch the XML editor to edit the contents in the hidden field.
'
'   d18 ndk     09Dec1999       Created
'   d44 wdp     06Mar2000       Fixed DisplayEditSummaryDetailsScreen to cope with
'                               title OR scope-content for create new entry and
'                               be able to switch between them.
    
    'For create new entry we need to be able to switch between
    'title and scope dtds as we can create headers/pieces below
    'classes
    Dim sStr As String
    
    If t_bUseHiddenSwitchField = True Then
        sStr = ksNameDTDType & ".value"
    Else
        sStr = t_eEADType
    End If
        
    GenerateEditOnclickCode = _
        " " & _
        gk_sHTMLDefaultXMLEditorLauncherControlName & t_sFieldName & _
        ".Launch( " & _
        t_sFieldName & ".value, " & sStr
    
    GenerateEditOnclickCode = GenerateEditOnclickCode & "); "

End Function

    
Public Sub AddReadOnlyEntryIntoEditSet()
'
' Adds a read only entry into the editorial as an editable entry
'
'Amendment history
'   d19 ndk     10Dec1999       Created
'   d23 ndk     21Dec1999       Added security check
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim lCatalogueID As Long, eCatalogueLevelID As ECATALOGUELEVEL, lEditsetID As Long
    Dim oBOBasicCatalogueEntry As BOBasicCatalogueEntry
    Dim oBOEditorialDBUtility  As BOEditorialDBUtility
    Dim arCatalogueIDs(1 To 1) As Long, arCatalogueLevels(1 To 1) As ECATALOGUELEVEL
    Dim eSecurityCheckStatus As ESecurityCheckResult
        
    sSource = Me.Name & "AddReadOnlyEntryIntoEditSet"

    'current catalogue details
    lEditsetID = m_aspRequest.QueryString(gk_sParamEditSetID)
    lCatalogueID = m_aspRequest.QueryString(gk_sParamCatalogueID)
    eCatalogueLevelID = m_aspRequest.QueryString(gk_sParamCatalogueLevel)
    
    'perform security check
    m_oBOEditorialPermissions.EditsetID = lEditsetID
    m_oBOEditorialPermissions.UserID = m_lLoggedUserID
    
    eSecurityCheckStatus = m_oBOEditorialPermissions.Check(EUserCanAddCatalogueEntryToEditSet)
    
    If eSecurityCheckStatus <> ESecurityCheckResult.ESRSecurityCheckSuccessful Then
        RaiseError sSource, _
               GetErrorEnumForSecurityCheckResult(eSecurityCheckStatus), _
               GetDescriptionForSecurityCheckError( _
                        GetErrorEnumForSecurityCheckResult(eSecurityCheckStatus))
    End If
    
    'initialise
    Set oBOBasicCatalogueEntry = New BOBasicCatalogueEntry
    Set oBOEditorialDBUtility = New BOEditorialDBUtility
    
    'check whether the catalogue is already present in any other editset as
    'an editable entry
    
    If oBOEditorialDBUtility.CataloguePresentAndEditableInAnyOtherEditset( _
        lEditsetID, lCatalogueID, eCatalogueLevelID) Then
        
        RaiseError sSource, PECanNotAddCatalogueAlreadyPresentAsEditable, _
            "The catalogue you are trying to add in editorial is already is present as editable entry in other editset, it was not added"
    End If
    
    'put the ids in an array
    arCatalogueIDs(1) = lCatalogueID
    arCatalogueLevels(1) = eCatalogueLevelID
    
    'call the method to add this entry into the editset as added entry
    oBOBasicCatalogueEntry.AddCataloguesIntoEditSet _
        lEditsetID, arCatalogueIDs, arCatalogueLevels
    
    m_aspResponse.Redirect _
        gk_sLinkDisplayEditSetDetails & _
        "?" & gk_sParamEditSetID & "=" & lEditsetID
    
    
    Set oBOBasicCatalogueEntry = Nothing
    Set oBOEditorialDBUtility = Nothing
    
    m_objObjectContext.SetComplete
    
    Exit Sub
ErrorHandler:

    Set oBOBasicCatalogueEntry = Nothing
    Set oBOEditorialDBUtility = Nothing

    m_objObjectContext.SetAbort
    ShowError sSource
End Sub

Public Sub DisplayCreateNewCatalogueByCopyingDepartmentScreen()
'Description
'
'Calls the methods to display the page to display page to allow a department
'to be copied.
'
'Amendment history
'   d38     tmj     23Feb2000

On Error GoTo ErrorHandler

    Dim sSource As String
    Dim lEditsetID As Long
    Dim lCatalogueToBeCopiedId As Long, eCatalogueToBeCopiedLevel As ECATALOGUELEVEL

    sSource = Me.Name & ".DisplayCreateNewCatalogueByCopyingDepartmentScreen"
            
    'get the data from the session store (store when the user chooses the catalogue to be copied)
    lEditsetID = GetDataFromSession(dlCreateCatalogueEditSetID)
    lCatalogueToBeCopiedId = GetDataFromSession(dlCreateCatalogueCatalogueToBeCopiedID)
    eCatalogueToBeCopiedLevel = GetDataFromSession(diCreateCatalogueCatalogueToBeCopiedLevel)
        
    'Call the method to display the screen to create an entry
    DisplayEditSummaryDetailsScreen _
        etLoadFromNoneCreateCatalogueEntry, lEditsetID, , , , , ecCreateCatalogueByCopying, , , lCatalogueToBeCopiedId, eCatalogueToBeCopiedLevel
    
    m_objObjectContext.SetComplete
    
    Exit Sub

ErrorHandler:
    m_objObjectContext.SetAbort
    ShowError sSource
End Sub


' WARNING - any changes/enhancements to this should be mirrored in VWReaderCatalogueEntry.LinkToBrowser
'           where appropriate.
Public Sub LinkToBrowser()

On Error GoTo ErrorHandler

    Dim sSource As String
    sSource = "VWReaderCatalogueEntry.LinkToBrowser"
    Dim eLevel As ECATALOGUELEVEL
    Dim lID As Long
    Dim lEDID As Long
    Dim sURL As String
    Dim eBrowserUsage
    
    eBrowserUsage = m_aspRequest.Item(gk_sParamUsage)
    
    lEDID = CInt(m_aspRequest.Item(gk_sParamEditSetID))
    eLevel = CInt(m_aspRequest.Item(gk_sParamCatalogueLevel))
    lID = CLng(m_aspRequest.Item(gk_sParamCatalogueID))
    
    ' d11
    If lEDID > 0 Then
        ' This is an editorial action
        ' So we need to know what type of editoral action so that we can redirect to
        ' the right page, and also concatenate the EDID onto the end of the querystring
          
        ' Quick check that the Edid in querystring is the same as the one supplied in
        ' session - should be.
        If lEDID = SessionManager.GetDataFromSession(dlCreateCatalogueEditSetID) Then
             
             
             ' Check if this is a create from scratch operation
             ' d11
            If eBrowserUsage = ECATALOGUEBROWSERUSAGE.eEditorialParentSelectionForCreateFromScratch Then
                sURL = gk_sLinkDisplayCatalogueBrowserForCreateFromScratchParentSelection & "?" & GenerateQueryStringsForCatalogueBrowser(lID, eLevel) & "&" & gk_sParamEditSetID & "=" & lEDID
            End If
            
             If eBrowserUsage = ECATALOGUEBROWSERUSAGE.eEditorialParentSelectionForAddFromExisting Then
                sURL = gk_sLinkDisplayCatalogueBrowserForAddExisting & "?" & GenerateQueryStringsForCatalogueBrowser(lID, eLevel) & "&" & gk_sParamEditSetID & "=" & lEDID
            End If
            
             If eBrowserUsage = ECATALOGUEBROWSERUSAGE.eEditorialCopySelection Then
                sURL = gk_sLinkDisplayCatalogueBrowserForCreateByCopyCatalogueSelection & "?" & GenerateQueryStringsForCatalogueBrowser(lID, eLevel) & "&" & gk_sParamEditSetID & "=" & lEDID
            End If
               
            If eBrowserUsage = ECATALOGUEBROWSERUSAGE.eEditorialMoveCatalogueEntry Then
               sURL = gk_sLinkDisplayCatalogueBrowserForMoveCatalogueEntryNewParentSelection & "?" & GenerateQueryStringsForCatalogueBrowser(lID, eLevel) & "&" & gk_sParamEditSetID & "=" & lEDID
            End If
            
            If eBrowserUsage = ECATALOGUEBROWSERUSAGE.eEditorialRelatedSeparateMaterial Then
               sURL = gk_sLinkDisplayCatalogueBrowserForRelatedSeparateMaterial & "?" & GenerateQueryStringsForCatalogueBrowser(lID, eLevel) & "&" & gk_sParamEditSetID & "=" & lEDID
            End If
      
            If eBrowserUsage = ECATALOGUEBROWSERUSAGE.eEditorialInitialParentSelectionForCreateByCopy Then
               sURL = gk_sLinkDisplayCatalogueBrowserForCreateByCopyParentSelection & "?" & GenerateQueryStringsForCatalogueBrowser(lID, eLevel) & "&" & gk_sParamEditSetID & "=" & lEDID
            End If

        End If

        
    Else
            ' If the browser usage cannot be attained, then open up into the catalogue browser in reader mode.
            sURL = gk_sLinkDisplayCatalogueBrowser & "?" & GenerateQueryStringsForCatalogueBrowser(lID, eLevel)
      
    End If
    
    m_aspResponse.Redirect sURL
    
    m_objObjectContext.SetComplete
    
    Exit Sub
ErrorHandler:
    ShowError sSource
    m_objObjectContext.SetComplete

End Sub

