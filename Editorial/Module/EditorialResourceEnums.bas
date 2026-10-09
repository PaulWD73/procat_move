Attribute VB_Name = "EditorialResourceEnums"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    EditorialResourceEnums.cls
' System:       PROCAT, PRO
' Copyright:    (C) Quidnunc Limited
'
' Description:  Used to declare the Editorial application data resource file
'               enumeration. Will contain both shared and editorial enums
'
' Amendment history:
'   d1      ksk     02Jul1999       Created
'   d2      ndk     07Jul1999       Added enums related to edit set details for one user, all users, other users
'   d3      ksk     08Jul1999       Added enum for Errors
'   d4      ksk     09Jul1999       Added tooltip help for catalogue browser
'   d5      ksk     12JUL1999       Added Invalid error for NextPage and Catalogue Reference
'   d6      ndk     15Jul1999       Added new enums
'   d7      ksk     15JUL1999       Added ENUM in Errors for missing records in ext tables
'   d8      ndk     15Jul1999       minor changes in enum number
'   d9      ndk     16Jul1999       Added new enum for total in edit set history

'   d10     Geedha Sivanadian   22July1999  Added enum entries for AF files.
'   d11     crm     10Aug1999       Added new enum for UserMaintenance module resources
'   d12     ndk     11Aug1999       Added enums for catalogue details tab
'   d13     ksk     17Aug1999       Added enums for bulkupdate
'   d14     ksk     01Oct1999       Added tooltip enums for help button and filter resources
'   d15     ksk     08Oct1999       Added tooltip for remove button
'   d16     aga     12Oct1999       Added Enums for Lookup Maintenance resource strings
'   d17     ndk     21Oct1999       Moved shared enums to a separete file, move catalogue display enums to shared enums
'   d18     ksk     21Oct1999       Added tooltip for readonly and imported image
'   d19     ksk     22Oct1999       Added tooltip for the - image
'   d20     aga     12Oct1999       Added one constants in Enums for Lookup Maintenance resource strings
'   d21     ksk     18Nov1999       Added tooltip const for load into edit set from Batch
'   d22     ksk     22Nov1999       Added tooltip for help on viewing log file contents
'   d23     jes     26Nov1999       Added some more reader file enums for search
'   d24     ndk     09Dec1999       Added enums for deletion confirmation messages for catalogue
'   d25     ksk     10Dec1999       Added tooltip for druid data preperation link
'   d26     crm     27Dec1999       Added tooltip for Information Leaflets
'   d27     ndk     28Dec1999       Added tooltips
'   d28     ndk     29Dec1999       Added resource strings for audit trail and sfa
'   d29     aga     13Jan2000       Added more resource strings for AF maintenance
'   d30     spb     08Feb2000       Added tooltip for 'add read only entry to edit set'
'   d31     tmj     22Feb2000       Added 2 enums in EEditorialResourceFile used in a confirm box
'   d32     daj     28Feb2000       Added tooltip for the parse log
'   d33     daj     08Mar2000       Added erfConfirmAbandonEditSet
'   d34     daj     21Mar2000       Added erfCreateLoadEditSetHeader, erfCreateUpdateEditSetHeader
'   d35     gbd     08Feb2001       Added erfDisplayEditSetsHyperLinkEditorialSetsWithoutEditors
'   d36     gdb     08Feb2001       Added erfDisplayEditSetsInEditorialStageWithoutEditor
'   d37     gdb     15Feb2001       Added erfBulkEditPleaseNote and erfBulkEditValueEntryInformAboutChanges
'   d38     alm     09Oct2001       Added erfs for moving pieces
'   d39     alm     23Oct2001       Added erfs for removing non-referenced entries from the catalogue
'   d40     fjd     28Nov2002       Added erErrorBadTexts
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Enum EEditorialResourceFile
'Bulk edit constants
    erfBulkEdit = 29001
    erfBulkEditFieldSelectionHelp = 29002
    erfBulkEditFieldCaption = 29003
    erfBulkEditValueCaption = 29004
    erfBulkEditLevelCaption = 29005
    erfBulkEditValueEntryHelp = 29006
    'd37
    erfBulkEditPleaseNote = 34025
    erfBulkEditValueEntryInformAboutChanges = 34026
    erfBulkEditValueEntryInformAboutCoveringdates = 34027
    erfBulkEditValueEntryInformAboutOpeningdates = 34028
    
    
    '   d13
    erfBulkEditNoCatalogueUpdated = 29007
    erfBulkEditUpdatedCatalogueList = 29008
    'd2
    'Display edit set details for one user, all users, other users
    'Captions and hyperlinks
    
    erfDisplayEditSetsHeader1 = 21001
    erfDisplayEditSetsHeader2 = 21002
    erfDisplayEditSetsColoursIndicate = 21003
    erfDisplayEditSetsReadyToApprove = 21004
    erfDisplayEditSetsWorkInProgress = 21005
    erfDisplayEditSetsNotAtYourStage = 21006
    erfDisplayEditSetsNotAtUsersStage = 21007
    erfDisplayEditSetsClickOnEditSet = 21008
    erfDisplayEditSetsHyperLinkMyEditSets = 21009
    erfDisplayEditSetsHyperLinkEditSetsForOtherPeople = 21010
    erfDisplayEditSetsHyperLinkEditorialSetsWithoutEditors = 34023
    erfDisplayEditSetsHyperLinkAllEditSets = 21011
    erfDisplayEditSetsWhoseEditSetToShow = 21012
    erfDisplayEditSetsOrShowMe = 21013
    erfDisplayEditSetsCurrentEditSets = 21014
    erfDisplayEditSetsInEditorialStageWithoutEditor = 34024
    
    'Display edit set details for one user, all users, other users
    'Coloumn headers
    
    erfDisplayEditSetsColName = 21016
    erfDisplayEditSetsColType = 21017
    erfDisplayEditSetsColContributor1 = 21018
    erfDisplayEditSetsColYourRole = 21019
    erfDisplayEditSetsColUsersRole = 21020
    erfDisplayEditSetsColCreatedOn = 21021
    erfDisplayEditSetsColWorkStage = 21022
    erfDisplayEditSetsColDaysInCurrentStage = 21023
    erfDisplayEditSetsColWorkStatus = 21024
    erfDisplayEditSetsHyperLinkHistory = 21025
    
    
    'Display Users and Roles for editset
    'Captions and headers
    
    erfDisplayUsersAndRolesHeader1 = 21031
    erfDisplayUsersAndRolesFooter1 = 21032
    erfDisplayUsersAndRolesDisplayValidUsers = 21033
    
    
    'Display Users and Roles for editset
    'Coloumn headers
    
    erfDisplayUsersAndRolesColPerson = 21036
    erfDisplayUsersAndRolesColRole = 21037
    erfDisplayUsersAndRolesColRoleActiveInStages = 21038
    
    'Display Edit Set History
    'Captions and headers
    
    erfDisplayEditSetHistoryHeader1 = 21041
    erfDisplayEditSetHistoryHeader2A = 21042
    erfDisplayEditSetHistoryHeader2B = 21043
    
    'Display Edit Set History
    'coloumn headers
    
    erfDisplayEditSetHistoryColStage = 21046
    erfDisplayEditSetHistoryColStageStart = 21047
    erfDisplayEditSetHistoryColStageEnd = 21048
    erfDisplayEditSetHistoryColDaysInStage = 21049
    erfDisplayEditSetHistoryColComments = 21050
    
    'd9
    'Display Edit Set History
    'Footer
    
    erfDisplayEditSetHistoryFooterDaysTotal = 21051
    
    'Create Empty Edit Set
    'captions and error messages
    erfCreateEmptyEditSetHeader1 = 21061
    erfCreateEmptyEditSetErrorCantCreateCont1 = 21062
    erfCreateEmptyEditSetErrorCantCreateMngEditor = 21063
    
    ' d34 - Added
    erfCreateLoadEditSetHeader = 21064
    erfCreateUpdateEditSetHeader = 34021
    
    'Create Empty Edit Set
    'Coloumn headers
    
    erfCreateEmptyEditSetColName = 21066
    erfCreateEmptyEditSetColType = 21067
    erfCreateEmptyEditSetColComments = 21068
    
    'Display Edit Set Details

    erfDisplayEditSetDetailsEditSet = 21071
    erfDisplayEditSetDetailsOwnedBy = 21072
    erfDisplayEditSetDetailsCreatedOn = 21073
    erfDisplayEditSetDetailsHasBeenIn = 21074
    erfDisplayEditSetDetailsStageFor = 21075
    erfDisplayEditSetDetailsDays = 21076
    erfDisplayEditSetDetailsDepartments = 21077
    erfDisplayEditSetDetailsDivisions = 21078
    erfDisplayEditSetDetailsSeries = 21079
    erfDisplayEditSetDetailsSubSeries = 21080
    erfDisplayEditSetDetailsSubSubSeries = 21081
    erfDisplayEditSetDetailsPieces = 21082
    erfDisplayEditSetDetailsItems = 21083
    erfDisplayEditSetDetailsLeftHeaderContents = 21084
    erfDisplayEditSetDetailsLeftHeaderFilter = 21085
    erfDisplayEditSetDetailsHyperLinkBulkUpdate = 21086
        
    'Editset entries display
    erfLevelName = 21087
    erfStatus = 21088
    erfView = 21089
    erfOrder = 21090

    'Display Next stages for edit set
    
    erfDisplayNextStagesForEditSetMessageNotEnoughRights = 21101
    erfDisplayNextStagesForEditSetMessageApproveAndAdvance = 21102
    erfDisplayNextStagesForEditSetMessageToTheNextStage = 21103
    erfDisplayNextStagesForEditSetMessageTheEditSetIsCurrentlyIn = 21104
    erfDisplayNextStagesForEditSetMessageStage = 21105
    erfDisplayNextStagesForEditSetMessageSelectANewStage = 21106
    erfDisplayNextStagesForEditSetLeftHeaderNextStage = 21107
    erfDisplayNextStagesForEditSetLeftHeaderComments = 21108
    
    'Display previous stages for edit set
    
    erfDisplayPreviousStagesForEditSetMessageNotEnoughRights = 21121
    erfDisplayPreviousStagesForEditSetMessageApproveAndSendBack = 21122
    erfDisplayPreviousStagesForEditSetMessageToThePreviousStage = 21123
    erfDisplayPreviousStagesForEditSetMessageTheEditSetIsCurrentlyIn = 21124
    erfDisplayPreviousStagesForEditSetMessageStage = 21125
    erfDisplayPreviousStagesForEditSetMessageSelectANewStage = 21126
    erfDisplayPreviousStagesForEditSetLeftHeaderPreviousStage = 21127
    erfDisplayPreviousStagesForEditSetLeftHeaderComments = 21128
    
    'Display create new catalogue menu
    
    erfDisplayCreateCatalogueMenuHeader1 = 21141
    erfDisplayCreateCatalogueMenuHeader2 = 21142
    erfDisplayCreateCatalogueMenuHyperLinkFromScratch = 21143
    erfDisplayCreateCatalogueMenuHyperLinkByCopying = 21144
    erfDisplayCreateCatalogueMenuHyperLinkNewLetterCode = 21145
    
    'Create new catalogue entry from scratch
    
    erfDisplayCreateCatalogueFromScratchHeader1 = 21151
    erfDisplayCreateCatalogueFromScratchLeftHeaderParentReference = 21152
    erfDisplayCreateCatalogueFromScratchLeftHeaderReference = 21153
    erfDisplayCreateCatalogueFromScratchLeftHeaderLevel = 21154
    erfDisplayCreateCatalogueFromScratchLeftHeaderTitle = 21155
    erfDisplayCreateCatalogueFromScratchLeftHeaderLegalStatus = 21156
    erfDisplayCreateCatalogueFromScratchLeftHeaderLanguage = 21157
    erfDisplayCreateCatalogueFromScratchLeftHeaderCreatorNames = 21158
    erfDisplayCreateCatalogueFromScratchLeftHeaderCoveringDates = 21159
    erfDisplayCreateCatalogueFromScratchLeftHeaderStartDate = 21160
    erfDisplayCreateCatalogueFromScratchLeftHeaderEndDate = 21161
    erfDisplayCreateCatalogueFromScratchLeftHeaderPlaceOfDeposit = 21162
    
    'Create new catalogue entry by copying
    
    erfDisplayCreateCatalogueByCopyingHeader1 = 21171
    erfDisplayCreateCatalogueByCopyingLeftHeaderParentReference = 21172
    erfDisplayCreateCatalogueByCopyingLeftHeaderReference = 21173
    erfDisplayCreateCatalogueByCopyingLeftHeaderLevel = 21174
    erfDisplayCreateCatalogueByCopyingLeftHeaderTitle = 21175
    erfDisplayCreateCatalogueByCopyingLeftHeaderLegalStatus = 21176
    erfDisplayCreateCatalogueByCopyingLeftHeaderLanguage = 21177
    erfDisplayCreateCatalogueByCopyingLeftHeaderCreatorNames = 21178
    erfDisplayCreateCatalogueByCopyingLeftHeaderCoveringDates = 21179
    erfDisplayCreateCatalogueByCopyingLeftHeaderStartDate = 21180
    erfDisplayCreateCatalogueByCopyingLeftHeaderEndDate = 21181
    erfDisplayCreateCatalogueByCopyingLeftHeaderPlaceOfDeposit = 21182
    
    'Create new letter code
    
    erfDisplayCreateLetteCodeHeader1 = 21191
    erfDisplayCreateLetteCodeLeftHeaderParentReference = 21192  'Not used just for consistency
    erfDisplayCreateLetteCodeLeftHeaderReference = 21193
    erfDisplayCreateLetteCodeLeftHeaderLevel = 21194
    erfDisplayCreateLetteCodeLeftHeaderTitle = 21195
    erfDisplayCreateLetteCodeLeftHeaderLegalStatus = 21196
    erfDisplayCreateLetteCodeLeftHeaderLanguage = 21197
    erfDisplayCreateLetteCodeLeftHeaderCreatorNames = 21198
    erfDisplayCreateLetteCodeLeftHeaderCoveringDates = 21199
    erfDisplayCreateLetteCodeLeftHeaderStartDate = 21200
    erfDisplayCreateLetteCodeLeftHeaderEndDate = 21201
    erfDisplayCreateLetteCodeLeftHeaderPlaceOfDeposit = 21202
    
    
    'Catalgue Browser for edit set
    
    'Create from scratch
    erfCatalogueBrowserForEditSetCreateFromScratchHeader = 21211
    erfCatalogueBrowserForEditSetCreateFromScratchSelectParent = 21212
    
    'Create by copying
    erfCatalogueBrowserForEditSetCopyCatalogueEntryHeader = 21216
    erfCatalogueBrowserForEditSetCopyCatalogueEntrySelectCatalogue = 21217
    erfCatalogueBrowserForEditSetCopyCatalogueEntrySelectParent = 21218
    
    'Captions/Headings for the Tabs to view/edit catalogue details
    
    'General
    
    erfTabDisplayGeneralDetailsHyperLinkBackToEditSet = 21301
    erfTabDisplayGeneralDetailsHyperLinkAuditTrail = 21302
    erfTabDisplayGeneralDetailsCaptionDetailsFor = 21303
    erfTabDisplayGeneralDetailsCaptionFrom = 21304
    
   
    'Deletion of catalogue confirmation messages
    erfConfirmDeletionForCreatedEntry = 21401
    erfConfirmDeletionForAccessionedEntry = 21402
    erfConfirmDeletionForAddedEntry = 21403
    ' d33 - Added
    erfConfirmAbandonEditSet = 34020

    'Text used in the display of a confirm box when adding a leaflet
    erfConfirmOverwriteLeafletOne = 21501
    erfConfirmOverwriteLeafletTwo = 21502

    'Filter Constants
    erfFilterHeading = 26001
    erfFilterHelpText = 26002
    erfFilterDescriptionText = 26003
    
    'Audit trail related
    erfDisplayAuditTrailCaptionHistorOfModifications = 26501    '"History of Modifications to "
    erfDisplayAuditTrailCaptionBackTo = 26502  '"Back to "
    erfDisplayAuditTrailColEditSet = 26503      'Edit set
    erfDisplayAuditTrailColFrom = 26504     'from
    erfDisplayAuditTrailColTo = 26505       'to
    erfDisplayAuditTrailColComments = 26506 'comments
    erfDisplayAuditTrailColRole = 26507 'role
    erfDisplayAuditTrailColPerson = 26508   'person
    erfDisplayAuditTrailHyperLinkViewThisVersion = 26509    'view this version
    
    'SFA
    erfDisplayCreateEditSfaCaptionModifyElectronicSFA = 26551
    erfDisplayCreateEditSfaCaptionModifyPrintedSFA = 26552
    erfDisplayCreateEditSfaCaptionCreateElectronicSFA = 26553
    erfDisplayCreateEditSfaCaptionCreatePrintedSFA = 26554
    
    'd38 Move Pieces
    erfDisplayEditSetDetailsHyperLinkMovePieces = 35012
    erfMoveHeader = 35013
    erfSortHierarchy = 35006
    erfSortKeyOrder = 33075
    
    'd39 Delete Catalogue Entries
    erfDisplayDeleteEditSetHeader = 35014
    erfConfirmDeletionFromEditSet = 35015
    
End Enum


Public Enum EReaderResourceFile
'Catalogue Browser
    rrfCatalogueSummaryHeader = 11001
    
'Errors
    rrfErrorInvalidFieldSearchTerm = 19001
    rrfErrorInvalidFieldFirstDate = 19002
    rrfErrorInvalidFieldLastDate = 19003
    rrfErrorInvalidFieldBreadthRestriction = 19004
    rrfErrorInvalidDateRange = 19101
    
'   Advanced/Field search
    rrfAdvancedSearchIntro = 15000
    rrfFieldSearchIntro = 15001
    rrfSelectConstraintsLabel = 15002
    rrfSelectFieldsLabel = 15003
    rrfFieldSearchFieldsDesc = 15004
    rrfAdvancedSearchFieldsDesc = 15005
    rrfAdvancedSearchTitle = 15006
    rrfFieldSearchTitle = 15007

End Enum

'd10
Public Enum EAuthorityFilesMaintResourceFile
'Authority Files Maitnenance
    afmrAuthorityFilesMaintenance = 27001
    afmrTermsRaisedthroughEditSets = 27002
    afmrEditSetName = 27003
    afmrContributor1 = 27004
    afmrCreatedon = 27005
    afmrStage = 27006
    afmrDaysinStage = 27007
    afmrNumberOfTermsForApproval = 27008
    afmrTermsRaisedthroughLeaflets = 27009
    afmrLeafletName = 27010
    afmrLiveTerms = 27012
    afmrInvisibleTerms = 27013
    afmrIncompleteTerms = 27014
    afmrInvisibleComment = 27015
    afmrInCompleteComment = 27016
    afmrNotUsedTerms = 27017
    afmrNotUsedTermsComment = 27018
    afmrNotUsedTermsLabel = 27019
    afmrIndexTermsDetailHeading1 = 27020
    afmrIndexTermsDetailHeading2 = 27021
    afmrType = 27022
    afmrReference = 27023
    afmrCreatedBy = 27024
    afmrStatus = 27025
    afmrEditThisTerm = 27026
    afmrLeafLetHeading2 = 27027
    afmrRequireApproval = 27028
    afmrSpecificIndexTermHeading = 27029
    afmrSurname = 27030
    afmrAuthorityStatus = 27031
    afmrGenderIndicator = 27032
    afmrTitle = 27033
    afmrAdditionalElementsOfName = 27034
    afmrPretitle = 27035
    afmrForename = 27036
    afmrAlternativeStatus = 27037
    afmrDateOfBirth = 27038
    afmrUncertainBirthdateCode = 27039
    afmrBiogHistory = 27040
    afmrValidation = 27041
    afmrDeathDate = 27042
    afmrUncertainDeathDateCode = 27043
    afmrApproveComplete = 27044
    afmrApproveInComplete = 27045
    afmrReject = 27046
    afmrName = 27047
    afmrParish = 27048
    afmrTown = 27049
    afmrCountry = 27050
    afmrEndDate = 27051
    afmrGridReferences = 27052
    afmrUncertainStartDateCode = 27053
    afmrUncertainEndDateCode = 27054
    afmrCounty = 27055
    afmrStartDate = 27056
    afmrHistory = 27057
    afmrDefinition = 27059
    
    
    afmrRemitAndFunction = 27061
    afmrVariant = 27062
    afmrJurisdiction = 27063
    afmrPlaceOfDeposit = 27064
    afmrNationalPlaceOfDepositCode = 27065
    
    'd29
    afmrLiveNonPreferredComment = 27066
    afmrNotUsedPreferredComment = 27067
    afmrNotUsedCandidateComment = 27068
    afmrLiveTermsComment = 27069
    afmrLiveNonPreferred = 27070
    afmrNotUsedPreferred = 27071
    afmrNotUsedCandidate = 27072
    
    
End Enum


'Tool tip enumerations
Public Enum ETOOLTIP
'Button messages
    ttButtonOKHelp = 1
    
    ttButtonCancelHelp = 2
    ttButtonMoreHelp = 3
    ttButtonPageUpHelp = 4
    ttButtonPageDownHelp = 5
    ttButtonPlusHelp = 6
'Browser Help
    ttLettercodeListSortByReference = 7
    ttLettercodeListSortByTitle = 8
    ttCatalogueDetails = 9
    ttCatalogueSummaryDetails = 10

'Editset entry Status
    ttNewCatalogue = 11
    ttAddedCatalogue = 12
    ttModifiedCatalogue = 13

'   d13     ksk     17Aug1999
    ttButtonDone = 14
'   d14
    ttButtonHelpHelp = 15
    ttButtonFilterHelp = 16
    
'   d15     ksk     08Oct1999
    ttButtonRemove = 17
    
'   d18     ksk     21Oct1999
    ttImportedCatalogue = 18
    ttReadOnlyCatalogue = 19

'   d19     ksk     22Oct1999
    ttButtonMinusHelp = 20
    
'   d21     ksk     18Nov1999
    ttLoadBatchIntoEditset = 21
'   d22     ksk     22Nov1999
    ttImportLogFileView = 22
    ttLoadLogFileView = 23
'   d25     KSK 10Dec1999
    ttPrepareDRUIDData = 24
'   d26     crm     27Dec1999
    ttEditLeaflet = 25
    ttChangeDocForLeaflet = 26
    ttEditPopularSearch = 27
    ttDeletePopularSearch = 28
    ttBackToPopSearchBrowser = 29
    ttCancelDelete = 30
    ttAddNewUser = 31
    
'   d27 ndk
    
    ttShowMyEditSets = 32
    ttShowOthersEditSets = 33
    ttShowAllEditSets = 34
    ttShowEditSetDetails = 35
    ttShowEditSetHistory = 36
    ttFilter = 37
    ttAddEntry = 38
    ttCreateEntry = 39
    ttNextStage = 40
    ttPreviousStage = 41
    ttEditRoles = 42
    ttAbandonEditSet = 43
    ttViewCatalogueDetails = 44
    ttEditCatalogue = 45
    ttMoveCatalogue = 46
    ttPickCorporateBody = 47
    ttPickPerson = 48
    ttPickPlace = 49
    ttPickSubject = 50
    ttBackToEditSet = 51
    ttAuditTrail = 52
    ttBackToEditSetCatalogue = 53
    ttViewAuditTrailCatalogue = 54
    ttStartAddRole = 55
    ttRemoveRole = 56
    ttFinishedEditingRoles = 57
    ttSelectUserForRole = 58
    ttCreateEntryFromScratch = 59
    ttCreateEntryByCopying = 60
    ttCreateLettercode = 61
    ttAddSFA = 62
    ttBacktoMaintainancemenu = 63
    ttEditSFADetails = 64
    ttDeleteSFA = 65
    ttBackToSFAList = 66
    
    ttViewContextDetails = 67
    ttViewSummaryDetails = 68
    ttViewAccessDetails = 69
    ttViewContentDetails = 70
    ttViewAdminhistoryDetails = 71
    ttViewIndexTermDetails = 72
    ttButtonAddReadOnly = 73
    
    ' d32 - Tooltip for the parse log link
    ttParseLogFileView = 74
End Enum



Public Enum EErrorResource
    'CLIENT SIDE JAVASCRIPT error description
    erErrorEndbeforeStart = 1
    erErrorBadReals = 2
    erErrorBadInts = 3
    erErrorBadDates = 4
    erErrorMissings = 5
    erErrorBadMoneys = 6
    erErrorOutOfRanges = 7
    erErrorBadDivisors = 8
    erErrorBadReal = 9
    erErrorBadInt = 10
    erErrorBadMoney = 11
    erErrorBadDate = 12
    erErrorMinimum = 13
    erErrorMaximum = 14
    erErrorBadDivisor = 15
    erErrorLength = 20
    
    'd3      ksk
    'Editset errors
    erInvalidEditSet = 16
    'd4      ksk
    erInvalidNextPage = 17
    erInvalidCatalogueReference = 18
    '   d7      ksk
    erNoRecordsForCatalogueSummary = 19
    
    'd14 ksk
    erAccessDenied = 21
    
    'd21
    erCustodialHistoryRowNotPresent = 22
    erAppDestInformationRowNotPresent = 23
    erAccrualsRowNotPresent = 24
    erTitleRowNotPresent = 25
    erArrangementRowNotPresent = 26
    erScopeAndContentRowNotPresent = 27
    erNoteRowNotPresent = 28
    erAdminHistoryRowNotPresent = 29
    
    'd40
    erErrorBadTexts = 30

End Enum

 '   d11      crm
Public Enum EUserMaintainResource
    umUserMaintenance = 28501
    umAddUsers = 28502
    umUserName = 28503
    umDetailsof = 28504
    umModifyUserDetailsAndClick = 28505
    umOK = 28506
    umToUpdate = 28507
    umInitials = 28508
    umNTLogOn = 28509
    umType = 28510
    umUserRolesCheckUnCheckToAddRemoveRoleFromUser = 28511
    umDeleteStatus = 28512
    umUnableToCompleteRequestUserPartOfExisting = 28513
    umEditSet = 28514
    umIncomplete = 28515
    umYouhavenotenteredrequiredfields = 28516
    umUserUpdateStatus = 28517
    umRoles = 28518
    umIsPartOfEditSetRemoveUser = 28519
    umNewUser = 28520
    umAddUserDetailsAndClick = 28521
    umUser = 28522
    umHasReplaced = 28523
    umAsManagingEditor = 28524
    umInvalidPrivileges = 28525
    umNeedToBeAdministrator = 28526
    umInvalidUser = 28527
    umDoYouWantToMake = 28528
    umQuestionMark = 28529
    umDoYouWantToReplace = 28530
    umWith = 28531
    umUseBrowsersBackButtonToModify = 28532
End Enum

'd16  aga  -  Enums for resource file for Lookup Maintenance
Public Enum ELookupMaintainResource
    LMMaintenance = 28000
    LMLanguageLookup = 28001
    LMLegalStatusLookup = 28002
    LMPhysicalFormLookup = 28004
    LMClosureStatusLookup = 28005
    LMAccessConditionLookup = 28006
    LMInvalidLookup = 28010
    LMLanguage = 28021
    LMLegalStatus = 28022
    LMPhysicalForm = 28024
    LMClosureStatus = 28025
    LMAccessCondition = 28026
    LMCode = 28027
    LMNoRecords = 28101
    LMListOfExisting = 28102
    LMAddNew = 28103
    LMModify = 28104
    LMAlreadyExists = 28121
    LMDuplicateRecordExistAsDeleted = 28122  'd20
    LMNoLookupExistsWithThis = 28123
End Enum

