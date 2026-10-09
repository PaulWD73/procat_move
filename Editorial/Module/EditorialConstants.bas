Attribute VB_Name = "EditorialConstants"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    EditorialConstants.cls
' System:       PROCAT, PRO
' Copyright:    (C) Quidnunc Limited 2000
'
' Description:  Used to store constants related to editorial section
'
' Amendment history:
'   d1      ndk     28Jun1999       Created
'   d2      ndk     28Jun1999       Changed the constant names
'   d3      ksk     29Jun1999       Added variable names passed as query string
'                                   for EditsetId and CatalogueID
'   d4      ndk     29Jun1999       Added variable names passed as query string
'                                   for editsetname, roleid, userid
'   d5      ndk     30Jun1999       Added gk_lManagingEditorRoleID
'   d6      ndk     30Jun1999       Added gk_sNameUserIDForContributor1, gk_sNameUserIDForManagingEditor
'                                   gk_lStageContributionID
'   d7      ndk     02Jul1999       Added many new constants
'   d9      ndk     07Jul1999       Added gk_lHighestRoleLevelForEditSetStage
'   d10     ndk     08Jul1999       Added constants for form contents
'   d11     ksk     12Jul1999       Added gk_iMaxTitleDisplayChars,
'                                   Added gk_sNameCatalogueID and gk_sNameCatalogueLevel
'                                    for use to pass Catalogue ID and level
'   d12     KSK     13Jul1999       Moved module depended constants out of constants.bas
'                                   Added the new constat gk_sSectionName to be used
'                                   as the section key in the registry
'   d13     KSK     13Jul1999       Added constant for catalogue browser max page length
'   d14     ndk     15Jul1999       Added constants for edit_set_entry_types
'   d15     GS      20Jul1999       Added new constants for authority_statuses.
'   d16     GS      26Jul1999       Added new consts for leaflets.
'   d17     NDK     29Jul1999       Added gk_lEditorRoleID, changed gk_lManagingEditorRoleID to 4
'                                   Added gk_sParamEditSetID
'   d19     GS      04Aug1999        Added male female constants for gender indicator.
'   d20     GS      04Aug1999        Changed the path for the Image and Error files
'   d21     KSK     04Aug1999        Added Param global const for catalogue ID. Changed
'                                   the Param and Name variable name for editset ID
'                                   to be the same
'   d22     NDK     04Aug1999       Added  gk_sParamCatalogueLevel,gk_sParamDetailsTabNumber
'   d23     KSK     06Aug1999       Removed harded coded constant definitions to
'                                   a seperate module HardcodedIDConstants.bas
'   d24     ndk     09Aug1999       Removed few more hardcoded values
'   d25     crm     09Aug1999       Placed a constant for UserTypeID
'   d26     GS      09Aug1999       Cleaned up entries for AF
'   d27     KSK     01Oct1999       Added constant for help window name
'   d28     KSK     08Oct1999       Changed the value of the constant
'                                   gk_iMaxFilteredEditsetEntriesPerPage
'   d28     KSK     14Oct1999       Removed verity database constants names to DBconstants
'   d29     AGA     27Oct1999       Added constants for Asp parameters in AF Editorial Maintenance
'   d30     AGA     03Nov1999       Removed gk_sAFPickParameter, and d29 moved to VWEditorialAuthorityFiles.cls
'   d31     KSK     16Nov1999       Added constants for Item and piece key order incerments
'   d32     KSK     17Nov1999       Added gk_sParamFileName
'   d33     ndk     17Nov1999       Added new constants for querystrings
'   d34     KSK     17Nov1999       Added gk_sParamBatchDescription
'   d35     KSK     17Nov1999       Added a const for a DSN to DRUID database gk_sDSNToDRUID
'   d36     KSK     18Nov1999       Added a Uniqueness constraint violation constants
'   d37     KSK     18Nov1999       Changed the value of the uniqueness const
'   d38     KSK     18Nov1999       Added global param const.
'   d39     WDP     19Nov1999       Added TAG constants
'   d40     KSK     25Nov1999       Moved constants from Missing file
'   d41     NDK     25Nov1999       Added gk_sNamePlaceOfDeposit
'   d42     JES     26Nov1999       Added constants from reader for the editorial search
'                                   Moved some parameters to HTTPParameters.bas so reader can also see them
'                                   Added dummy cookie constants for search
'   d43     SPB     01Dec1999       Changed catalogue browser component name.
'   d44     CRM     10Dec1999       Added CATREF constant
'   d45     KSK     10Dec1999       Added Error description constants
'   d46     WDP     15Dec1999       Added release version constant
'   d47     JES     13Jan2000       Added some reader constants when splittin into EdSearch and RdSearch
'   d48     CRM     07Feb2000       Increment release version to v0.4
'   d49     jcc     08Feb2000       Change class to series
'   d50     spb     08Feb2000       Added gk_lMaxEditsetSizeToRetrieve - this is how many edit set entries
'                                   the sp will retrieve, so the max that will be displayed.
'   d51     CRM     14Feb2000       Increment release version to v0.5
'   d52     CRM     15Feb2000       Increment release version to v0.6
'   d53     WDP     18Feb2000       added gk_sXMLStyleSheet
'   d54     DAJ     18Feb2000       added gk_sHelpPagesPath
'   d55     DAJ     18Feb2000       added gk_sBasePath and gk_sControlPath
'   d56     DAJ     29Feb2000       Increment release version to v0.7
'   d57     TMJ     02Mar2000       Added values used to pass data on current editset when using audit trail
'   d58     TMJ     06Mar2000       Added value used in error message when parent has been moved
'   d59     DAJ     07Mar2000       Added value gk_sParamRedirection
'   d60     DAJ     08Mar2000       Added value gk_sErrorDescriptionNonLiveParent
'   d61     DAJ     08Mar2000       Incremented release version to 0.8
'   d62     TMJ     16Mar2000       Added value gk_sFORMAFinSFA
'   d63     DAJ     21Mar2000       Updated version number to 0.9
'   d64     DAJ     27Mar2000       Removed a couple of constants not used anymore for the audit trail
'   d65     DAJ     05Apr2000       Added gk_sNameUpdateFieldFunction
'   d66     DAJ     18Apr2000       Incremented version number for new acceptance test build
'   d67     DAJ     03Jul2000       Updated version number to 1.2
'   d68     DAJ     14Jul2000       Added constant for ASP request parameter for update
'   d69     DAJ     01Aug2000       Updated version number to 1.3, changed the default MaxRecords property to stop limiting
'   d70     DAJ     04Sep2000       Updated version number to 1.31
'   d71     DAJ     19Sep2000       Updated version number to 1.3.2
'   d72     DAJ     04Oct2000       Updated version number to 1.3.3
'   d73     DAJ     17Oct2000       Updated version number to 1.3.4
'   d74     DAJ     06Nov2000       Updated version number to 1.3.5
'   d75     MAP     13Dec2000       Updated version number to 1.3.6
'   d76     BCJ     16Jan2001       Added constants for enhancements 1 & 2 (enhancement set 1)
'   d77     BCJ     18Jan2001       Path changes to d76
'   d78     BCJ     18Jan2001       Added gk_sLinkCopyExistingCataloguesGoToBrowser
'                                   Added gk_sLinkStoreSelectedParentForAddExisting
'   d79     BCJ     22Jan2001       Added gk_sLinkDisplayCatalogueBrowserForAddExisting
'   d80     BCJ     24Jan2001       Corrected strore with store
'   d81     DAJ     29Jan2001       Updated version number to 1.4
'   d82     BCJ     13Feb2001       Added gk_sLinkCopyExistingParentSelectionGoToBrowser
'   d83     BCJ     13Feb2001       Added gk_sPAGEBrowseFrameset
'   d84     BCJ     14Feb2001       Added gk_sLinkMoveExistingSelection
'   d85     BCJ     14Feb2001       Added gk_sLinkDisplayGoToBrowserRelatedMaterial
'   d86     DAJ     16Feb2001       Updated version number to 1.5
'   d87     GDB     20Feb2001       added 2 new variables gk_vRangeMin , gk_vRangeMax
'   d88     DAJ     22Feb2001       Updated version number to 1.6
'   d89     ALM     23Feb2001       Moved browser constants to BOCataloueSummary
'   d90     ALM     26Feb2001       Added gk_sIMAGEGoButton
'   d91     MAP     09MAR2001       Added gk_sErrorDescriptionEntryInAnotherEditSet
'   d92     DAJ     14MAR2001       Updated version number to 1.6.1
'   d93     MAP     22Mar2001       Added gk_lErrorNumberDatabase2
'   d94     MAP     30MAR2001       Updated version number to 1.6.3 (to bring in line with reader version numbering).
'   d95     MAP     18APR2001       Updated version number to 1.6.4
'   d96     MAP     23MAY2001       Updated version number to 1.6.5
'   d97     MAP     31MAY2001       Updated version number to 1.6.6
'   d98     MAP     31MAY2001       Updated version number to 1.6.7
'   d99     MAP     22AUG2001       Added gk_sDSNToEPRO
'   d100    MAP     22AUG2001       Updated version number to 1.7.0
'   d101    ALM     28SEP2001       Updated version number to 1.7.2
'   d102    MAP     02OCT2001       Added gk_sItemMarkerIndicator for catalogue browser
'   d103    MAP     12Oct2001       Added constants for ead scheduling functionality
'   d104    ALM     23Oct2001       Updated version number to 1.7.3
'   d105    ALM     25Oct2001       Added error messages when adding entries to delete editset
'   d106    ALM     28Nov2001       Updated version number to 1.7.4
'   d107    TomH    30Aug2002       Updated version number to 1.8.0
'   d108   FrancisD 06Feb2003       Updated version number to 1.8.1
'   d109   FrancisD 02Apr2003       Updated version number to 1.8.2
'                                   Removed gk_lSFABrowserTableSize (it's now a registry setting)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Const gk_sEditorialReleaseVersion As String = "r1.8.2 d38"


'd3  ksk
Public Const gk_sNameEditsetID As String = "EDID"        'variable name for the EditsetID
Public Const gk_sNameCatalogueID As String = "CATID"        'variable name for the catalogueID

'd16    GS
Public Const gk_sLeafLetId As String = "LeafLetId"

'd4 ndk
Public Const gk_sNameEditSetName As String = "ESNAME"   'variable name for edit set name
Public Const gk_sNameRoleID As String = "ROLEID" 'variable name for role id
Public Const gk_sNameUserID As String = "USERID" 'variable name of user id (usually used for user other than the logged user)

'd6
Public Const gk_sNameUserIDForContributor1 As String = "USERIDFORCONTTRIBUTOR1" 'variable name for the cont1 on new editset
Public Const gk_sNameUserIDForManagingEditor As String = "USERIDFORMANAGINGEDITOR" 'variable name for the cont1 on new editset

'd7

Public Const gk_sNameContributor1Name As String = "NAMEOFCONTRIBUTOR1"
Public Const gk_sNameEditSetCreationDate As String = "ESCREATIONDATE"
Public Const gk_sNameEditSetCurrentStageID As String = "ESCURRENTSTAGEID"
Public Const gk_sNameEditSetCurrentStageName As String = "ESCURRENTSTAGENAME"
Public Const gk_sNameDaysInStage As String = "DAYSINSTAGE"

'd17
'   d21     KSK
'Public Const gk_sParamEditSetID As String = "EDID"     'Moved to HTTPParameters.bas    JES
'Public Const gk_sParamCatalogueID As String = "CATID"

'd22 NDK
'Public Const gk_sParamCatalogueLevel As String = "CATLN"
'Public Const gk_sParamDetailsTabNumber As String = "DTN"    'Used to pass the tab number

'd9
'This the top role level for each stage
'eg : contributor1 and editor have the role level 2
Public Const gk_lHighestRoleLevelForEditSetStage As Long = 2


'd10

Public Const gk_sNameEditSetType As String = "EDITSETTYPE"
Public Const gk_sNameEditSetStage As String = "EDITSETSTAGE"
Public Const gk_sNameEditSetDescription As String = "EDITSETDESCRIPTION"
Public Const gk_sNameStageTransitionComments As String = "STAGETRANSITIONCOMMENTS"

Public Const gk_sNameCatalogueLevel As String = "CATLN"

' d83 Added
Public Const gk_sPAGEBrowseFrameset As String = "BrowseFrame.asp"



'd12

' ODBC Data Source Name for PROCAT - this is used as the default for the registry entry
Public Const gk_sDefaultDSN As String = "DSN=PROCAT; UID=PROCATUSER; PWD=PROCATUSER"

' d99 - Added global constant to connect to EPRO database (needed for RegistryUtility
Public Const gk_sDSNToEPRO As String = "DSN=EPRO; UID=PROCATUSER; PWD=PROCATUSER;"


'd20
'Path of the Image, error pages and log file in the development envi.
Public Const gk_sBasePath = "/editorial/"
Public Const gk_sImagesPath = gk_sBasePath & "images/"
Public Const gk_sErrorPagesPath = gk_sBasePath & "ErrorPage/"
Public Const gk_sHelpPagesPath = gk_sBasePath & "Help/"
Public Const gk_sControlPath = gk_sBasePath & "Control/"

' d90
Public Const gk_sIMAGEGoButton = gk_sImagesPath & "gobutton.gif"

'Default table used to store transient data
Public Const gk_sDefaultTransientStorageTable As String = "Transient_Store"
'Default table names for tool tip and error
Public Const gk_sDefaultToolTipTable As String = "ToolTip"
Public Const gk_sDefaultErrorTable As String = "Error"

'Both reader and editorial have the same set of framenames
Public Const gk_sMainFrame As String = "Main"
Public Const gk_sTopFrame As String = "Top"
Public Const gk_sContentFrame As String = "Content"

' FWStringBuffer class padding increment
Public Const gk_lStringBufferIncrement As Long = 50000

' FWRecord - maximum number of fields per record
Public Const gk_iRecordFieldsMax As Integer = 30

' FWDBHelper - max number of records returned by a search
Public Const gk_lDBMaxRecords As Long = 0

'Data delimitors are used by business objecct to serialise the state information
Public Const gk_sDataDelimiter  As String = "%$@*(!@"
'Instance delimitors are used to delimit the instance reference, sessionid, boname while storing in transience manager
Public Const gk_sInstanceDelimiter  As String = "~"

Public Const gk_sSectionName As String = "EditorialData"

'   d13     KSK
'This constant is used to find the list of children to display in a page for catalogue browser
Public Const gk_iMaxChildrenPerPage As Integer = 15


'd14

Public Const gk_iEditSetEntryTypeCreated As Integer = 1
Public Const gk_iEditSetEntryTypeAccession As Integer = 2
Public Const gk_iEditSetEntryTypeModifiedCreated As Integer = 3
Public Const gk_iEditSetEntryTypeModifiedAccession As Integer = 4

'd19    GS
Public Const gk_sMaleGender As String = "M"
Public Const gk_sFemaleGender As String = "F"

'   d25     CRM
Public Const gk_lUserTypeID As Long = 1

'd26        GS
Public Const ksYesPlaceOfDeposit As String = "Y"
Public Const ksNOPlaceOfDeposit As String = "N"

'd27
Public Const gk_sHelpWindowName As String = "Editorial Help"
Public Const gk_iMaxFilteredEditsetEntriesPerPage As Integer = 25


'   d31     KSK
'Increment for new piece key order
Public Const gk_iIncrementForPieceKeyOrder As Integer = 1000
'Increment for new item key order
Public Const gk_iIncrementForItemKeyOrder As Integer = 1000


'   d32     KSK
'Constant for Parameter name for filenames being passed
Public Const gk_sParamFileName As String = "FN"

'd33       NDK
Public Const gk_sParamCreateNewSFA As String = "CreateNewSFA"
Public Const gk_sParamSFAReferenceID As String = "SFAReferenceID"
Public Const gk_sParamAddButtonsForEditorialMaintain As String = "bAddButtnsForEditorialMaintain"

'd34 KSK
'Used to store the batch description as query string
Public Const gk_sParamBatchDescription As String = "BatchDesc"
' d103 Added - Used to store Start time and date for scheduled job
Public Const gk_sParamJobStartTime As String = "StartTime"
Public Const gk_sParamJobStartDate As String = "StartDate"
Public Const gk_sParamJobType As String = "JobType"


'd35 KSK
'Global constant to connect to DRUID database
Public Const gk_sDSNToDRUID As String = "DSN=DRUID; UID=PROCATUSER; PWD=PROCATUSER;"


'   d38     KSK
Public Const gk_sParamBatchID As String = "BID"

Public Const gk_sTAGEditsetID As String = "editsetID"
Public Const gk_sTAGAction As String = "action"
Public Const gk_sTAGCollection As String = "collection"
Public Const gk_sTAGLettercode As String = "lettercode"
Public Const gk_sTAGBulkFile As String = "bulkfile"
Public Const gk_sTAGEntry As String = "entry"


'   d40
'Index browser parameters
Public Const gk_sFORMAFGroup As String = "fldAFGroup"
Public Const gk_sFORMAFListString As String = "fldAFListString"
Public Const gk_sFORMAFListIndex As String = "fldAFListIndex"
Public Const gk_sFORMAFTermID As String = "fldAFTermID"
Public Const gk_sFORMAFListType As String = "fldAFListType"
Public Const gk_sFORMAFListTermType As String = "fldAFListTermType"
Public Const gk_sFORMAFListUsage As String = "fldAFListUsage"
Public Const gk_sFORMAFTermValue As String = "flfAFTermValue"
Public Const gk_sFORMAFinSFA As String = "inSFA"

'Frame names in the Reader
Public Const gk_sFRAMELeftFrame As String = "left"
Public Const gk_sFRAMEBottomFrame As String = "bottom"
Public Const gk_sFRAMERightFrame As String = "right"

'Length of int field
Public Const gk_lMaximumValueForIntegerInSQL7 As Long = 2147483647
Public Const gk_lErrorNumberDuplicateKeyValueForCollection As Long = 457

'd87
Public Const gk_vRangeMin As Variant = 0
Public Const gk_vRangeMax As Variant = 9999

'SFA details constants
Public Const gk_sFORMSFADetailsID As String = "lID"
Public Const gk_sFORMSFAListStartIndex As String = "lStartIndex"
Public Const gk_sFORMSFAListUsage As String = "eUsage"
Public Const gk_sFORMSFAListFilter As String = "eFilter"



'params
Public Const gk_sParamSFAType As String = "SFAType"
Public Const gk_sParamSFATitle As String = "SFATitle"
Public Const gk_sParamSectionType As String = "SectionType"


Public Const gk_sComponentCatalogueBrowser As String = "PROCATRdCatBr.VWCatalogueBrowser"

'd41
Public Const gk_sNamePlaceOfDeposit As String = "PLACEOFDEPOSIT"

'd42
Public Const gk_sSORTSTRINGReference As String = "catalogue reference"
Public Const gk_sSORTSTRINGDate As String = "covering dates"
Public Const gk_sSORTSTRINGRelevance As String = "relevance"
Public Const gk_sSORTSTRINGFormerReference As String = "former departmental reference"
'd42    dummy reader constants for search
Public Const gk_sPROCATCookieCurrentSearch As String = ""
Public Const gk_sPROCATCookieName As String = ""
Public Const gk_sPROCATCookieCurrentSearchPage = ""
Public Const gk_sPAGERdLeaflet As String = ""
Public Const gk_sFRAMELeafletframe As String = ""
Public Const gk_sFORMLeafletID As String = ""

'   d44     CRM     10Dec1999
Public Const gk_sNameCatalogueReference As String = "CATREF"

'd45 KSK

'The number of matches to display on each page
Public Const gk_lPSBrowserTableSize As Integer = 4
'Just get all the matches (for guided search results)
Public Const gk_lPSSearchGetAllMatches As Integer = -1




'DATABASE error related constants

'   d36     KSK
Public Const gk_lErrorNumberDatabase As Long = -2147217900
' d93 MAP Added to catch primary constraint due to inserting records that already exist
Public Const gk_lErrorNumberDatabase2 As Long = -2147155865
Public Const gk_sErrorDescriptionUniqueKeyValue As String = "UNIQUE KEY"


' d60 - Added
Public Const gk_sErrorDescriptionNonLiveParent As String = "Cannot add this entry as a parent, it has not been released to live - "

'd45 KSK
Public Const gk_sErrorDescriptionPrimaryKeyValue As String = "PRIMARY KEY"
Public Const gk_sErrorDescriptionDuplicateKeyValue As String = "duplicate Key"
Public Const gk_sErrorDescriptionAddDuplicateEditableCatalogue As String = "Cannot add a catalogue already present in an editset as editable"
Public Const gk_sErrorDescriptionParentMoved As String = "Cannot add catalogue as a parent has been moved in an editset"

'd105
Public Const gk_sErrorDeleteCatalogueIsWrongLevel As String = "Cannot add catalogue entry to the editset for deletion as it is ineligible<br>"
Public Const gk_sErrorDeleteCatalogueHasReferences As String = "Cannot add catalogue entry to editset for deletion as it has dependant catalogue entries<br>"
Public Const gk_sErrorDeleteCatalogueItemHasReferences As String = "Cannot add catalogue entry to editset for deletion as it is a referenced item<br>"
Public Const gk_sErrorCatalogueParentMarkedForDeletion As String = "Cannot use catalogue entry as a parent as it is marked for deletion<br>"

Public Const gk_sPROCATCookieSessionID As String = "SessionID"
Public Const gk_sPROCATCookieUserID As String = "UserID"
Public Const gk_sPROCATCookieDorisUserID As String = "DorisUserID"
Public Const gk_sPROCATCookieUserType As String = "UserType"

Public Const gk_sDummyDivisionTitle As String = "Series without a division parent"
Public Const gk_sDummyHeaderTitle As String = "Pieces without a sub-series parent"
Public Const gk_sDummySubHeaderTitle As String = "Pieces without a sub-sub-series parent"

' d50
' Max size of entire edit set list to retrieve -
' ie how many records will be retrieved from the sp results set.
Public Const gk_lMaxEditsetSizeToRetrieve As Long = 100000


Public Const gk_sXMLStyleSheet As String = "EditorialXMLRenderer.xsl"
Public Const gk_sAppname As String = "PROCAT Editorial"
Public Const gk_lDefaultNewRecordStatus As Long = gk_lFLAGNotOrderable

'd7 Constants used to pass details of Current Editset when going through audit trail
Public Const gk_sParamCurrEditSetID As String = "CEDID"

' d59 - Added
Public Const gk_sParamRedirection As String = "RedirectURL"

Public Const gk_sMasterIDXLogFile As String = "PROCAT\Log\MasterIndexingLogFile.log"

' d65 - Added
Public Const gk_sNameUpdateFieldFunction As String = "UpdateField"

Public Const gk_sVerityErrorText As String = ": Error"

' d68 - Added
Public Const gk_sParamUpdateEADList As String = "bUpdate"

' d85 - bcj Added
Public Const gk_sLinkDisplayGoToBrowserRelatedMaterial As String = "DisplayCatalogueDetailsGoToBrowserRelatedDetail.asp"
' d80 corrected strore with store
Public Const gk_sLinkStoreSelectedParentForAddExisting As String = "StoreSelectedParentForAddExisting.asp"

'd78 - bcj Added
Public Const gk_sLinkDisplayCatalogueBrowserForAddExisting As String = "DisplayCatalogueBrowserForAddExisting.asp"

'd91 MAP Added
Public Const gk_sErrorDescriptionEntryInAnotherEditSet As String = "Entry exists in another edit set"

'd95 MAP Added to mirror values in Reader constants
Public Const gk_sMarkerTop As String = "0"
Public Const gk_sMarkerNoSCN As String = "-1"
Public Const gk_sMarkerBottom As String = "-1"
'd102 Added to mirror value in Reader constants - perhaps I should have put them elsewhere!
Public Const gk_sItemMarkerIndicator As String = "*"



