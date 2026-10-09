Attribute VB_Name = "ASPLinkConstants"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    ASPLinkConstants.cls
' System:       PROCAT, PRO
' Copyright:    (C) Quidnunc Limited
'
' Description:  Used to store constants for the ASP Pages and images.
'
' Amendment history:
'   d1      ksk     07Jun1999      Created
'   d1      ndk     07Jun1999      Created constant for homepage, login error
'   d2      jpz     11Jun1999      Added constant for DisplayAuthorityTerms.asp
'   d3      jpz     15Jun1999      Added constant for DisplayAuthorityFiles.asp
'                                  and DisplayEditTerms
'   d3      jpz     15Jun1999      Added constants for imagebuttons
'   d4      ndk     22Jun1999      Added constants for displaying editsets
'   d5      ndk     23Jun1999      Added constants for displaying and editing roles,images
'   d6      jpz     23Jun1999      Modified the constants for for DisplayAuthorityTerms.asp,
'                                  DisplayAuthorityFiles.asp and DisplayEditTerms.
'   d7      ndk     23Jun1999      Removed path reference from imageconstants,
'                                   Added constants for displaying  and editing roles
'   d8      ksk     23Jun1999      Added link for catalogue Browser and organised
'                                   the links and images
'   d9      ndk     28Jun1999      Replaced gk_sLinkUpdateStatus with gk_sLinkUpdateEditSetStatus
'   d10     ndk     02Jul1999      Added many new constants (images and pages)
'   d11     ksk     02Jul1999      Added bulkupdate asp
'   d12     ndK     08Jul1999      Added gk_sLinkDisplayEditSetHistory ,
'                                  gk_sLinkDisplayValidUsersForTheRole,gk_sLinkCreateEmptyEditSet ,gk_sLinkDisplayNextStagesForEditSet
'                                  gk_sLinkDisplayPreviousStagesForEditSet, gk_sLinkMoveEditSetToNextStage, gk_sLinkMoveEditSetToPreviousStage
'   d13     ksk     02Jul1999      Added images
'   d14     ksk     12Jul1999      Added Link  gk_LinkSelectCatalogueReference
'   d15     ksk     12Jul1999      Added link gk_sLinkDisplayCatalogueSearchResult
'   d16     ndk     15Jul1999      Added links for creating catalogue entries
'   d17     ksk     15Jul1999      Integrating Maintanance work done by aga and crm
'   d18     crm     01Aug1999      Added constant for Message Delimiter for BOUser and VWUserEditSet classes
'   d19     GS      02Aug1999      Added link constants for authority files.
'   d20     ndk     08Aug1999      Added link and image constants for catalogue details (view and edit)
'   d21     GS      11Aug1999       Added images paths for AF.
'   d22     crm     12Aug1999      Added link constants for Leaflet maintenance module
'   d23     ndk     12Aug1999      Added link for UpdateCatalogueDetails.asp
'   d24     ndk     16Aug1999      Added image constants for checkboxes
'   d25     ksk     01Oct1999      Added constants for help image and other links
'   d26     ksk     08Oct1999      Added Help link for editset list and filter display
'   d27     aga    12Oct1999       Added help link for User, Lookup  & AF maintenance. Constant for invisible image
'   d28     ndk     13Oct1999      Added new link terms
'   d29     aga     27Oct1999      Added constant for Undelete button image
'                                           Added file name constant for "LookupUndelete.asp"
'                                           Added file name const for AFPickTermFromIndexBrowser.asp
'                                           of Editorial Maintenance
'   d30     aga     29Oct1999      Added file name consts for AF Add term related asp files
'   d31     ksk     29Oct1999      Added const for asp files to add terms and SFAs
'                                  for bulkupdate
'   d32     ksk     02Nov1999      Added const for pick term images, and Pick SFA image
'   d33     aga     03Nov1999      Added maintenance path(gk_sMaintenancePagesPath) to AF, Lookup, User maintenance pagess
'   d34     ksk     04Nov1999      Added const for cancel of selection of terms
'   d35     ksk     10Nov1999      Added const for abandon of editset
'   d36     ksk     12Nov1999      Added link constants for accession
'   d37     ndk     12Nov1999      Added new link constants
'   d38     ksk     12Nov1999      Added link constants for accession
'   d39     ksk     17Nov1999      Changed the name for accession confirmation
'   d40     ndk     17Nov1999      Added new link constants for delete catalogues,move
'                                  catalogues , and sfa
'   d41     ksk     17Nov1999      Added the static HTML link to chosed the type
'                                  of accession function
'   d42     ksk     18Nov1999      Added the asp link for staring load into editset
'   d43     crm     19Nov1999      (integrated by NDK) Added const for Information Leaflet module links
'   d44     ksk     22Nov1999      Removed constatns that are not link related
'   d45     ksk     23Nov1999      Added constant for showing the accession log file
'   d46     ksk     25Nov1999      Moved constants from missing file
'   d47     jes     26Nov1999      Copied some image constants here for the search
'   d48     jes     2Dec1999       Added select all image button
'   d49     ksk     08Dec1999      Added the update button link
'   d50     ndk     08Dec1999      Added help button constants
'   d51     crm     09Dec1999      Added help button constants
'   d52     ndk     09Dec1999      Added enums and help file constants
'   d53     ndk     09Dec1999      Changed the asp file associated with constant gk_sPAGESFAEdDetails
'                                  Added help file constants
'   d54     ksk    10Dec1999       Added the missing constant
'   d55     crm    10Dec1999       Added ASP link constants for Popular Search maintenance
'                                  and link constants for help
'   d56     crm    10Dec1999       Added a missing const for InfoLeaflet Browser
'   d57     aga    11Dec1999       Added HTML help constants for AF maintenance
'   d58     ksk    11Dec1999       Added the start DRUID data prepare asp.
'   d59     aga    15Dec1999       Added filename consts for AF preferred term reject confirmation
'   d60     spb    15Dec1999       Added some editorial parsing page links
'                                  and a refresh button image.
'   d61     ndk    23Dec1999       Added constant gk_sLinkSFAMaintenance
'   d62     ndk    28Dec1999       Added gk_sImageEditDisabled
'   d63     jes    13Jan1999       Added some index browser constants
'   d64     aga    21Jan1999       Added constant for maintenance home page
'   d65     jes    25Jan2000       Added undo pick term button
'   d65     jcc    07Feb2000       Change help file src
'   d66     spb    08Feb2000       Added 'add read only entry' image
'                                   (same as added entry status icon)
'   d67     dtm    09feb2000       Added 'select all' image
'   d68     wdp    09feb2000       removed 'select all' image to HTMLConsts as
'                                  it is refered to in both reader and editorial code
'   d69     daj    14feb2000       Changed all 'vnice' button names back to 'nice'
'   d70     crm    15Feb2000       Changed gk_sImageAddReadOnlyButton to iconicimportbutton.gif
'   d71     daj    18Feb2000       Tidied up the Help references/images and removed all
'                                   references to Maintenance path
'   d72     daj    18Feb2000       Split up AddTerm and AddLeaflet buttons
'   d73     daj    22Feb2000       Added link to the EAD Editor Control
'   d74     tmj    22Feb2000       Added link to ChooseSearchMethod page
'   d75     tmj    23Feb2000       Added link to enabled copying of departments
'   d76     daj    24Feb2000       Removed the EDA Editor Control link (moved to HTMLConsts)
'   d77     daj    24Feb2000       Added link for nice back button
'   d78     tmj    28Feb2000       Added link for initial select parent page
'   d79     tmj    14Mar2000       Added gk_sLinkAddLeafletToBulkUpdate
'   d80     daj    10Jul2000       Added gk_sPAGEDeleteLeaflet
'   d81     daj    10Jul2000       Added gk_sLinkStartUpdatingIntoEditset, gk_sLinkCreateEmptyEditSetForUpdate,
'                                   gk_sLinkDisplayCreateEmptyEditSetForUpdate
'   d82     gdb    08Feb2001       Added gk_sLinkDisplayEditorialSetsWithoutEditor
'   d83     alm    23Feb2001       Moved browser constants to BrowserConstants
'   d84     alm    26Sep2001       Added gk_sPAGEFieldSearch, gk_sPAGEAdvancedSearch, gk_sPAGESearchFrameset
'   d85     map    09Oct2001       Added ASP pages for EAD and DRUID job scheduling functionality
'   d86     alm    09Oct2001       Added ASP pages for moving multiple pieces
'   d87     alm    19Oct2001       Added constants for the move multiple pieces help pages
'   d88     alm    23Oct2001       Added constants for the display of a delete edit set
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''LINK CONSTANT'''''''''''''''''''''''''''''''''''''''''

'General
Public Const gk_sLinkHome As String = "Home.htm"

'   d2      jpz     11June1999
'   d6      jpz     23Jun1999
Public Const gk_sLinkDisplayAuthorityTerms As String = "DisplayAuthorityTerms.asp"
Public Const gk_sLinkDisplayAuthorityFiles As String = "DisplayAuthorityFiles.asp"

'd19        2Aug1999
Public Const gk_sLinkDisplaySpecificIndexTermDetails As String = "DisplaySpecificIndexTermDetails.asp"
Public Const gk_sLinkUpdateIndexTerms As String = "UpdateIndexTerms.asp"

'd59
Public Const gk_sLinkAFPreferredTermRejectConfirm As String = "RejectConfirmation.asp"
Public Const gk_sLinkAFRejectPreferredTerm As String = "RejectPreferredTerm.asp"

'd29  File name to call index browser from AF Editorial maintenance
Public Const gk_sIndexBrowserAFPickTermPage As String = "AFPickTermFromIndexBrowser.asp"

' d30   AF Add term file names
Public Const gk_sAddAFTermResults As String = "AddAFTermResults.asp"
Public Const gk_sPAGEAFAddTerm As String = "AddNewAFTerm.asp"

'd4
Public Const gk_sLinkDisplayEditSetsForOneUser As String = "DisplayEditSetsForOneUser.asp"
Public Const gk_sLinkDisplayEditSetsForAllUsers As String = "DisplayEditSetsForAllUsers.asp"
Public Const gk_sLinkUpdateEditSetStatus As String = "UpdateEditSetStatus.asp"
Public Const gk_sLinkDisplayUserNames As String = "DisplayUserNames.asp"

' d82
Public Const gk_sLinkDisplayEditorialSetsWithoutEditor As String = "DisplayEditSetWithoutEditors.asp"


'd5
Public Const gk_sLinkDisplayUsersAndRolesForEditSet As String = "DisplayUsersAndRolesForEditSet.asp"

'd7
Public Const gk_sLinkAddUpdateRoleOnEditSet As String = "AddUpdateRoleOnEditSet.asp"
Public Const gk_sLinkRemoveRoleFromEditSet As String = "RemoveRoleFromEditSet.asp"

'd19
Public Const gk_sLinkEditorialAFMaintMain As String = "DisplayEditorialAFMaintMain.asp"
'This page displays the tabbed interface yet to be coded
Public Const gk_sLinkCatalogueDetails As String = "CatalogueDetails.asp"

Public Const gk_sLinkDisplayBulkUpdateFields As String = "DisplayBulkUpdateFields.asp"
Public Const gk_sLinkDisplayBulkUpdateValue As String = "DisplayBulkUpdateValue.asp"

'd86
Public Const gk_sLinkDisplayChooseParentOfPieces As String = "DisplayChooseParentOfPieces.asp"
Public Const gk_sLinkDisplayChoosePiecesToMove As String = "DisplayChoosePiecesToMove.asp"
Public Const gk_sLinkDisplayChooseParentForPieces As String = "DisplayChooseParentForPieces.asp"
Public Const gk_sLinkDisplayMoveMultiplePieces As String = "MoveMultiplePieces.asp"

'd10
Public Const gk_sLinkDisplayEditSetDetails As String = "DisplayEditSetDetails.asp"

'd11
Public Const gk_sLinkBulkUpdateLevels As String = "BulkUpdateLevels.asp"

'd12

Public Const gk_sLinkDisplayEditSetHistory As String = "DisplayEditSetHistory.asp"
Public Const gk_sLinkDisplayValidUsersForTheRole As String = "DisplayValidUsersForTheRole.asp"
Public Const gk_sLinkCreateEmptyEditSet As String = "CreateEmptyEditSet.asp"
Public Const gk_sLinkDisplayNextStagesForEditSet = "DisplayNextStagesForEditSet.asp"
Public Const gk_sLinkDisplayPreviousStagesForEditSet As String = "DisplayPreviousStagesForEditSet.asp"
Public Const gk_sLinkMoveEditSetToNextStage As String = "MoveEditSetToNextStage.asp"
Public Const gk_sLinkMoveEditSetToPreviousStage As String = "MoveEditSetToPreviousStage.asp"

'd14
Public Const gk_LinkSelectCatalogueReference As String = "SelectCatalogue.asp"


'd15
Public Const gk_sLinkDisplayCatalogueSearchResult As String = "DisplayCatalogueSearchResult.asp"
Public Const gk_sLinkDisplayCatalogueEntrySearch As String = "BasicSearch.asp"
Public Const gk_sLinkAddCataloguesToEditset As String = "AddCataloguesToEditSet.asp"

'd16

Public Const gk_sLinkDisplayCreateNewCatalogueMenu As String = "DisplayCreateNewCatalogueMenu.asp"
Public Const gk_sLinkDisplayCatalogueBrowserForParentSelectionOfCreationFromScratch As String = "DisplayCatalogueBrowserForParentSelectionOfCreationFromScratch.asp"
Public Const gk_sLinkDisplayCopyCataloguePage As String = "DisplayCopyCataloguePage.asp"
Public Const gk_sLinkDisplayCreateNewLetterCodeScreen As String = "DisplayCreateNewLetterCodeScreen.asp"
Public Const gk_sLinkDisplaySelectParentPage As String = "DisplaySelectParentPage.asp"
Public Const gk_sLinkDisplayCreateNewCatalogueFromScratchScreen As String = "DisplayCreateNewCatalogueFromScratchScreen.asp"
Public Const gk_sLinkDisplayCreateNewCatalogueByCopyingScreen As String = "DisplayCreateNewCatalogueByCopyingScreen.asp"
'd74
Public Const gk_sLinkChooseSearchMethod As String = "ChooseSearchMethod.asp"

'd82
'Public Const gk_sLinkDisplayCatalogueBrowserForAddExisting As String = "DisplayCatalogueBrowserForAddExisting.asp"

Public Const gk_sLinkCreateNewCatalogueFromScratch As String = "CreateNewCatalogueFromScratch.Asp"
Public Const gk_sLinkCreateNewCatalogueByCopying As String = "CreateNewCatalogueByCopying.Asp"
Public Const gk_sLinkCreateNewLetterCode As String = "CreateNewLetterCode.Asp"


'd17
Public Const gk_sLookupAdd As String = "LookupAdd.asp"
Public Const gk_sLookupModify As String = "LookupModify.asp"
Public Const gk_sLookupRemove As String = "LookupRemove.asp"
Public Const gk_sLookupDisplay As String = "DisplayLookup.asp"
Public Const gk_sLookupUpdate As String = "LookupUpdate.asp"
'd29
Public Const gk_sLookupUndelete As String = "LookupUndelete.asp"

Public Const gk_sLinkUserDetail As String = "DisplayUserDetail.asp"
Public Const gk_sLinkNewUser As String = "NewUserEntryScreen.asp"
Public Const gk_sLinkUpdateUser As String = "UserUpdateDisplay.asp"
Public Const gk_sLinkDeleteUser As String = "DeleteUser.asp"
Public Const gk_sLinkShowUsers As String = "ShowAllUsersScreen.asp"

'd20

Public Const gk_sLinkDisplayCatalogueDetails As String = "DisplayCatalogueDetails.asp"
Public Const gk_sLinkDisplayEditCatalogueDetailsScreen As String = "DisplayEditCatalogueDetailsScreen.asp"
Public Const gk_sLinkDisplayMoveCatalogueScreen As String = "DisplayMoveCatalogueScreen.asp"

'd23

Public Const gk_sLinkUpdateCatalogueDetails As String = "UpdateCatalogueDetails.asp"

'   d22     crm
Public Const gk_sNewLeafletPage As String = "DisplayNewLeafletFilesSetup.asp"
Public Const gk_sDisplayStaticLeafletDetails As String = "StaticLeafletDetails.asp"
Public Const gk_sDisplayAllLeafletFiles As String = "DisplayAllLeafletFiles.asp"
Public Const gk_sDisplayAllLeaflets As String = "DisplayAllLeaflets.asp"

Public Const gk_sDisplayFilterForm As String = "DisplayFilterForm.asp"

Public Const gk_sLinkDisplayCatalogueRemoveConformationScreen As String = "ConfirmDeletionOfCatalogue.asp"


'd28
Public Const gk_sLinkDisplayCatalogueBrowserForParentSelectionOfCreationByCopying As String = _
    "DisplayCatalogueBrowserForParentSelectionOfCreationByCopying.asp"
Public Const gk_sLinkAddLinkTermToCatalogue As String = "AddLinkTermToCatalogue.asp"
Public Const gk_sLinkSubmitCatalogueDetails As String = "SubmitCatalogueDetails.asp"
Public Const gk_sPAGEAFList As String = "AuthorityFileList.asp"
Public Const gk_sLinkCatalogueList As String = "CatalogueList.asp"
Public Const gk_sLinkSFAList As String = "SFAList.asp"
Public Const gk_sLinkDisplayViewEditorialCatalogueDetails As String = "DisplayViewEditorialCatalogueDetails.asp"
Public Const gk_sLinkDisplayAuditTrailForCatalogue As String = "DisplayAuditTrailForCatalogue.asp"

'   d31     ksk     29Oct1999
Public Const gk_sLinkAddLinkTermToBulkUpdate As String = "AddIndexTermForBulkupdate.asp"
Public Const gk_sLinkAddSFAToBulkUpdate As String = "AddSFAForBulkupdate.asp"
'd79
Public Const gk_sLinkAddLeafletToBulkUpdate As String = "AddLeafletForBulkUpdate.asp"

'   d34     ksk
Public Const gk_sLinkCancelIndextermSelectionForBulkUpdate As String = "CancelIndexTermSelectionForBulkupdate.asp"


'   d35     ksk
Public Const gk_sLinkDisplayConformationForAbandon As String = "DisplayConformationForAbandon.asp"
Public Const gk_sLinkAbandonEditset As String = "AbandonEditset.asp"
'   d36    KSK
Public Const gk_sLinkDisplayAccessionBatchEntry As String = "DisplayAccessionBatchDescriptionEntryPage.asp"

'   d38   KSK
Public Const gk_sLinkAccessionStarted As String = "StartAccessionFromEAD.asp"


'   d41     KSK
Public Const gk_sLinkDisplayLoadHome As String = "DisplayLoadHome.html"

'   d42     KSK
Public Const gk_sLinkStartLoadingIntoEditset As String = "StartLoadingIntoEditset.asp"
Public Const gk_sLinkDisplayBatchList As String = "DisplayListToloadIntoEditset.asp"
Public Const gk_sLinkCreateEmptyEditSetForLoad As String = "CreateEmptyEditSetForLoad.asp"
Public Const gk_sLinkDisplayCreateEmptyEditSetForLoad As String = "DisplayCreateEmptyEditSetScreenForLoad.asp"

'd37 NDK

Public Const gk_sLinkDisplayPrintedSFAList As String = "DisplayPrintedSFAList.asp"
Public Const gk_sLinkDisplayElectronicSFAList As String = "DisplayElectronicSFAList.asp"
Public Const gk_sLinkCreateUpdateSFA As String = "CreateUpdateSFA.asp"
Public Const gk_sLinkDisplayCreateNewSFAScreen As String = "DisplayCreateNewSFAScreen.asp"
Public Const gk_sLinkDisplayEditSFADetailsScreen As String = "DisplayEditSFADetailsScreen.asp"
Public Const gk_sLinkDeleteSFA As String = "DeleteSFA.asp"
Public Const gk_sLinkDisplayConfirmationForSFADeletion As String = "DisplayConfirmationForSFADeletion.asp"
Public Const gk_sLinkDisplayConfirmationForCatalogueDeletion As String = "DisplayConfirmationForCatalogueDeletion.asp"


'   d43     crm
Public Const gk_sAddTermInformationLeaflet As String = "AddTermInformationLeaflet.asp"
Public Const gk_sLinkCancelAddTermInformationLeaflet As String = "CancelAddTermInformationLeaflet.asp"
Public Const gk_sDisplayIndexTermList As String = "DisplayIndexTermList.asp"
Public Const gk_sDisplayLeafletDetail As String = "DisplayLeafletDetail.asp"
Public Const gk_sDisplayEditLeafletDetail As String = "DisplayEditLeafletDetail.asp"
Public Const gk_sSaveLeafletDetail As String = "SaveLeafletDetail.asp"
'   d56     crm    10Dec1999
Public Const gk_sInformationLeafletBrowser As String = "InformationLeafletBrowser.asp"


'd45
Public Const gk_sLinkDisplayAccessionLogFile As String = "DisplayAccessionLogFile.asp"


'd40
Public Const gk_sLinkStartCreateNewCatalogueFromScratch As String = "StartCreateNewCatalogueFromScratch.asp"
Public Const gk_sLinkStartCreateNewCatalogueByCopying As String = "StartCreateNewCatalogueByCopying.asp"
Public Const gk_sLinkDisplayCatalogueBrowserForInitialCreateByCopyParentSelection As String = "DisplayCatalogueBrowserForInitialCreateByCopyParentSelection.asp"
Public Const gk_sLinkMoveCatalogueToNewParent As String = "MoveCatalogueToNewParent.asp"
Public Const gk_sLinkStartMoveCatalogueEntry As String = "StartMoveCatalogueEntry.asp"
Public Const gk_sLinkDeleteEditSetCatalogue As String = "DeleteEditSetCatalogue.asp"
Public Const gk_sLinkSFABrowserForEdiorialPickTerm As String = "SFABrowserForEdiorialPickTerm.asp"
Public Const gk_sLinkSFABrowserForMaintenance As String = "SFABrowserForMaintenance.asp"
Public Const gk_sLinkDisplayCreateNewCatalogueByCopyingDepartmentScreen As String = "DisplayCreateNewCatalogueByCopyingDepartmentScreen.asp"

'd52
Public Const gk_sLinkDisplayMenuForRelatedSeparatedMaterial As String = "DisplayMenuForRelatedSeparatedMaterial.asp"
Public Const gk_sLinkFindCatalogueAndStartCatalogueBrowser As String = "FindCatalogueAndStartCatalogueBrowser.asp"
Public Const gk_sLinkSelectCatalogueForRelatedSeparateMaterial As String = "DisplayCatalogueBrowserForRelatedSeparateMaterial.asp"

'   d54     ksk
Public Const gk_sLinkAddReadOnlyEntryIntoEditSet As String = "AddReadOnlyEntryIntoEditSet.asp"


'd46
Public Const gk_sPAGEAFTermSearch As String = "AFTermSearch.asp"
Public Const gk_sPAGEAFTermDetail As String = "AFTermDetail.asp"
Public Const gk_sPAGEAFPickTerm As String = "AFPickTerm.asp"
Public Const gk_sPAGEAFCancelPickTerm As String = "CancelAddLinkTerm.asp"

'SFA Page constants
Public Const gk_sPAGESFARdDetails As String = "SFARdDetails.asp"
Public Const gk_sPAGESFAEdDetails As String = "DisplaySFAForViewWithEditButton.asp"
Public Const gk_sPAGESFAList As String = "SFABrowser.asp"
Public Const gk_sPAGESFADetailHelp As String = "SFADetailsHelp.asp"

'Reader page names
Public Const gk_sPAGEBasicSearchResults As String = "BasicSearchResults.asp"
Public Const gk_sPAGENotDone As String = "NotDone.htm"
Public Const gk_sPAGEMoreResults As String = "MoreResults.asp"
Public Const gk_sPAGERefineBasicSearchFrameset As String = "RefineBasicSearchFrameset.asp"
Public Const gk_sPAGESortResults As String = "SortResults.asp"
Public Const gk_sPAGEError As String = "Error.asp"
Public Const gk_sPAGESearchWithinHits As String = "SearchWithinHits.asp"
Public Const gk_sPAGEBackToSummary As String = "BackToSummary.asp"
Public Const gk_sPAGEAFListFrame As String = "AFListFrame.asp"
Public Const gk_sPAGECatalogueDetails As String = "DisplayCatalogueDetails.asp"
Public Const gk_sPAGEBasicSearch As String = "BasicSearch.asp"

Public Const gk_sPAGEPSList As String = "PSBrowser.asp"
Public Const gk_sPagePSRunSearch As String = "PSRunSearch.asp"
Public Const gk_sPagePSEditSearch As String = "PSEditSearch.asp"

'd84
Public Const gk_sPAGEFieldSearch As String = "AdvancedSearch.asp?fldSearchType=Field"
Public Const gk_sPAGEAdvancedSearch As String = "AdvancedSearch.asp"
Public Const gk_sPAGESearchFrameset As String = "SearchFrameset.asp"

'ASP Link constants for Popular Search maintenance
'   d55     crm    10Dec1999
Public Const gk_sDisplaySavedSearches As String = "DisplaySavedSearches.asp"
Public Const gk_sDisplayCreatePopularSearch As String = "DisplayCreatePopularSearch.asp"
Public Const gk_sSavePopularSearch As String = "SavePopularSearch.asp"
Public Const gk_sPopularSearchMaintenance As String = "PopularSearchMaintenance.htm"
Public Const gk_sDisplayPopularSearches As String = "DisplayPopularSearches.asp"
Public Const gk_sDisplayPopularSearchDetail As String = "DisplayPopularSearchDetail.asp"
Public Const gk_sDeletePopularSearch As String = "DeletePopularSearch.asp"
Public Const gk_sChangeSearchString As String = "ChangeSearchString.asp"

'   d58     ksk
Public Const gk_sLinkStartDRUIDPrepare As String = "StartPrepareDRUIDData.asp"

'''''''''''''''''''''''''''''''''ERROR PAGES'''''''''''''''''''''''''''''''''''''''''

'Error
Public Const gk_sLinkErrorLogin As String = "LoginError.htm"
Public Const gk_sLinkErrorAccessDenied As String = "AccessDenied.htm"


'''''''''''''''''''''''''''''''''IMAGE CONSTANTS'''''''''''''''''''''''''''''''''''''''''

'   d3      jpz     15June1999
Public Const gk_sImageOK As String = "niceOKbutton.GIF"
Public Const gk_sImageCancel As String = "nicecancelbutton.GIF"
Public Const gk_sImagePickTerm As String = "nicepicktermbutton.GIF"

'd5
Public Const gk_sImageRemove As String = "RemoveButton.Gif"
Public Const gk_sImageAdd As String = "AddButton.Gif"
Public Const gk_sImageDone As String = "DoneButton.Gif"
Public Const gk_sImageUpdateStatus As String = "UpdateStatusButton.Gif"

'd8
Public Const gk_sImageOpen As String = "Open.gif"
Public Const gk_sImageClose As String = "Close.gif"
Public Const gk_sImageScrollUP As String = "1ScrollUp.gif"
Public Const gk_sImageScrollDown As String = "1ScrollDown.gif"
Public Const gk_sImageMoreButton As String = "nicemorebutton.gif"

'd10

Public Const gk_sImageFilter As String = "FilterButton.gif"
Public Const gk_sImageAbandon As String = "AbandonButton.gif"
Public Const gk_sImageAddEntry As String = "AddEntryButton.gif"
Public Const gk_sImageCreateEntry As String = "CreateEntryButton.gif"
Public Const gk_sImageNextStage As String = "NextStageButton.gif"
Public Const gk_sImageSendBack As String = "SendBackButton.Gif"
Public Const gk_sImageEditRoles As String = "EditRolesButton.Gif"


'd13
Public Const gk_sImageNew As String = "newitem.gif"
Public Const gk_sImageChecked As String = "tick.gif"
Public Const gk_sImageAccessioned As String = "iconicaddedbutton.gif"
Public Const gk_sImageAddedFromLive As String = "iconicaddedbutton.gif"
Public Const gk_sImageImported As String = "iconimportedbutton.gif"
Public Const gk_sImageReadOnly As String = "iconreadonlybutton.gif"
'   d70     crm    15Feb2000
Public Const gk_sImageAddReadOnlyButton As String = "iconicimportbutton.gif"



'   d17     ksk
Public Const gk_sImageEdit As String = "niceEditButton.gif"

'   d18     crm
Public Const gk_sMessageDelimiter As String = "@#$%!"

'd20

Public Const gk_sImageTabAccess As String = "TabAccess.gif"
Public Const gk_sImageTabContent As String = "TabContent.gif"
Public Const gk_sImageTabContext As String = "TabContext.gif"
Public Const gk_sImageTabAdminHistory As String = "TabAdminHistory.gif"
Public Const gk_sImageTabIndexTerms As String = "TabIndexTerms.gif"
Public Const gk_sImageTabSummary As String = "TabSummary.gif"

Public Const gk_sImageTabSelectedAccess As String = "TabSelectedAccess.gif"
Public Const gk_sImageTabSelectedContent As String = "TabSelectedContent.gif"
Public Const gk_sImageTabSelectedContext As String = "TabSelectedContext.gif"
Public Const gk_sImageTabSelectedAdminHistory As String = "TabSelectedAdminHistory.gif"
Public Const gk_sImageTabSelectedIndexTerms As String = "TabSelectedIndexTerms.gif"
Public Const gk_sImageTabSelectedSummary As String = "TabSelectedSummary.gif"

Public Const gk_sImageMoveButton As String = "niceMoveButton.gif"


'd21
Public Const gk_sImageApproveComplete As String = "approve.gif"
Public Const gk_sImageApproveIncomplete As String = "incomplete.gif"
Public Const gk_sImageReject As String = "Reject.gif"
'd27
Public Const gk_sImageInvisible As String = "Invisible.gif"

'   d24
Public Const gk_sImageCheckBoxCheckedDisabled = "CheckedDisabled.gif"
Public Const gk_sImageCheckBoxUnCheckedDisabled = "UnCheckedDisabled.gif"
Public Const gk_sImageCheckBoxCheckedEnabled = "CheckedEnabled.gif"
Public Const gk_sImageCheckBoxUnCheckedEnabled = "UnCheckedEnabled.gif"
                
                
'd25
Public Const gk_sImageHelp As String = "help.gif"

Public Const gk_sImageRemoveButton As String = "IconicRemovebutton.GIF"

'd29
Public Const gk_sImageUndelete As String = "Undeletebutton.gif"

'd32
Public Const gk_sImagePickPerson As String = "PickPerson.gif"
Public Const gk_sImagePickPlace As String = "PickPlace.gif"
Public Const gk_sImagePickSubject As String = "PickSubject.gif"
Public Const gk_sImagePickCorporateBody As String = "PickCorp.gif"
Public Const gk_sImagePickSFA As String = "PickSFA.gif"


'd49
Public Const gk_sIMAGEUpdate As String = "DoneButton.Gif"


'd46
Public Const gk_sIMAGEURLPrevButton As String = gk_sImagesPath & "niceprevbutton.Gif"
Public Const gk_sIMAGEURLNextButton As String = gk_sImagesPath & "nicenextButton.gif"
Public Const gk_sIMAGESmallGo As String = gk_sImagesPath & "nicegobutton.gif"
Public Const gk_sIMAGEURLCancelButton As String = gk_sImagesPath & "NiceCancelButton.gif"
Public Const gk_sIMAGEURLAddTermButton As String = gk_sImagesPath & "AddTermButton.gif"
Public Const gk_sIMAGEURLPageHelp As String = gk_sImagesPath & gk_sImageHelp



'd47
Public Const gk_sIMAGELeftBracketButton As String = gk_sImagesPath & "leftbracket.gif"
Public Const gk_sIMAGERightBracketButton As String = gk_sImagesPath & "rightbracket.gif"
Public Const gk_sIMAGENoBracketButton As String = gk_sImagesPath & "nobracket.gif"
Public Const gk_sIMAGEAddTermButton As String = gk_sIMAGEURLAddTermButton
Public Const gk_sIMAGEAddLeafletButton As String = gk_sImagesPath & "niceAddLeafletButton.gif"
Public Const gk_sIMAGERemoveTermButton As String = gk_sImagesPath & "removeterm.gif"
Public Const gk_sIMAGEURLSearchButton As String = gk_sImagesPath & "niceSearchbutton.gif"
Public Const gk_sIMAGEPickTermButton As String = gk_sImagesPath & "pickterm.gif"
Public Const gk_sIMAGELargeLogo As String = gk_sImagesPath & "pro_ani2.gif"

Public Const gk_sFORMSavedSearchID As String = "frmSavedSearchID"
Public Const gk_sFRAMEHelpFrame As String = "help"
Public Const gk_sIMAGEURLRefineButton As String = gk_sImagesPath & "niceRefineSearchbutton.gif"
Public Const gk_sIMAGEURLUpLevelButton As String = gk_sImagesPath & "NiceSummaryButton.gif"


'd48
Public Const gk_sIMAGEURLSelectAllButton As String = "niceSelectAllButton.gif"

'd60
Public Const gk_sImageReload As String = "NiceReloadButton.Gif"

'd62
Public Const gk_sImageEditDisabled As String = "nicedisablededitbutton.gif"

Public Const gk_sIMAGEBackButton As String = gk_sImagesPath & "niceBackButton.gif"

'd25
'''''''''''''''''''''''''''''''''HTML HELP PAGE CONSTANTS'''''''''''''''''''''''''''''''''''''''''


Public Const gk_sPAGESearchResultsHelp As String = gk_sHelpPagesPath & "HelpSearchResults.htm"
Public Const gk_sPAGEAFTermDetailHelp As String = gk_sHelpPagesPath & "HelpAFTermDetail.htm"

Public Const gk_sHelpLinkToBulkUpdateFieldSelection As String = gk_sHelpPagesPath & "HelpBulkupdateFieldSelection.htm"
Public Const gk_sHelpLinkToBulkUpdateValueEntry As String = gk_sHelpPagesPath & "BulkupdateValueEntry.htm"
'   d26     ksk     08Oct1999
Public Const gk_sHelpLinkToFilteredEditsetEntryList As String = gk_sHelpPagesPath & "HelpFilteredEditsetEntryList.htm"
Public Const gk_sHelpLinkToFiltereForm As String = gk_sHelpPagesPath & "HelpFilterForm.htm"

'   d27     aga     12Oct1999    ASP Link constants
Public Const gk_sHelpLinkToMaintLookup As String = gk_sHelpPagesPath & "HelpMaintMaintLookup.htm"
Public Const gk_sHelpLinkToMaintUsers As String = gk_sHelpPagesPath & "HelpMaintUsers.htm"

'd56
Public Const gk_sHelpMAFMainPage As String = gk_sHelpPagesPath & "HelpMAFMainPage.htm"
Public Const gk_sHelpMAFIndexTermsPage As String = gk_sHelpPagesPath & "HelpMAFIndexTermsPage.htm"
Public Const gk_sHelpMAFTermDetailsPage As String = gk_sHelpPagesPath & "HelpMAFIndexTermDetailsPage.htm"
Public Const gk_sHelpMAFAddTermPage As String = gk_sHelpPagesPath & "HelpMAFAddTermPage.htm"

'd50
'edit tabs
Public Const gk_sHelpLinkToEditSummaryDetails As String = gk_sHelpPagesPath & "HelpEditSummaryDetails.htm"
Public Const gk_sHelpLinkToEditAccessDetails As String = gk_sHelpPagesPath & "HelpEditAccessDetails.htm"
Public Const gk_sHelpLinkToEditContentDetails As String = gk_sHelpPagesPath & "HelpEditContentDetails.htm"
Public Const gk_sHelpLinkToEditAdminHistoryDetails As String = gk_sHelpPagesPath & "HelpEditAdminHistoryDetails.htm"
Public Const gk_sHelpLinkToEditIndexTermsDetails As String = gk_sHelpPagesPath & "HelpEditIndexTermsDetails.htm"

'view tabs
Public Const gk_sHelpLinkToViewSummaryDetails As String = gk_sHelpPagesPath & "HelpViewSummaryDetails.htm"
Public Const gk_sHelpLinkToViewAccessDetails As String = gk_sHelpPagesPath & "HelpViewAccessDetails.htm"
Public Const gk_sHelpLinkToViewContentDetails As String = gk_sHelpPagesPath & "HelpViewContentDetails.htm"
Public Const gk_sHelpLinkToViewAdminHistoryDetails As String = gk_sHelpPagesPath & "HelpViewAdminHistoryDetails.htm"
Public Const gk_sHelpLinkToViewIndexTermsDetails As String = gk_sHelpPagesPath & "HelpViewIndexTermsDetails.htm"

'others
Public Const gk_sHelpLinkToViewAndEditContextDetails As String = gk_sHelpPagesPath & "HelpViewAndEditContextDetails.htm"
Public Const gk_sHelpLinkToCreateCatalogueMenu As String = gk_sHelpPagesPath & "HelpCreateCatalogueMenu.htm"
Public Const gk_sHelpLinkToCreateCatalogueEntry As String = gk_sHelpPagesPath & "HelpCreateCatalogueEntry.htm"
Public Const gk_sHelpLinkToViewAuditTrail As String = gk_sHelpPagesPath & "HelpViewAuditTrail.htm"
Public Const gk_sHelpLinkToDeleteCatalogue As String = gk_sHelpPagesPath & "HelpDeleteCatalogue.htm"
Public Const gk_sHelpLinkToEditSetsForOneUser As String = gk_sHelpPagesPath & "HelpEditSetsForOneUser.htm"
Public Const gk_sHelpLinkToEditSetsForAllUsers As String = gk_sHelpPagesPath & "HelpEditSetsForAllUsers.htm"
Public Const gk_sHelpLinkToUserNames As String = gk_sHelpPagesPath & "HelpUserNames.htm"
Public Const gk_sHelpLinkToUsersAndRolesOnEditSet As String = gk_sHelpPagesPath & "HelpUsersAndRolesOnEditSet.htm"
Public Const gk_sHelpLinkToValidUsersForRole As String = gk_sHelpPagesPath & "HelpValidUsersForRole.htm"
Public Const gk_sHelpLinkToCreateEmptyEditSet As String = gk_sHelpPagesPath & "HelpCreateEmptyEditSet.htm"
Public Const gk_sHelpLinkToEditSetHistory As String = gk_sHelpPagesPath & "HelpEditSetHistory.htm"
Public Const gk_sHelpLinkToAbandnon As String = gk_sHelpPagesPath & "HelpAbandon.htm"
Public Const gk_sHelpLinkToNextStages As String = gk_sHelpPagesPath & "HelpNextStages.htm"
Public Const gk_sHelpLinkToPreviousStages As String = gk_sHelpPagesPath & "HelpPreviousStages.htm"
Public Const gk_sHelpLinkToCreateSFA As String = gk_sHelpPagesPath & "HelpCreateSFA.htm"
Public Const gk_sHelpLinkToEditSFA As String = gk_sHelpPagesPath & "HelpEditSFA.htm"
Public Const gk_sHelpLinkToViewEditorialSFA As String = gk_sHelpPagesPath & "HelpViewEditorialSFA.htm"
Public Const gk_sHelpLinkToDeleteSFA As String = gk_sHelpPagesPath & "HelpDeleteSFA.htm"
Public Const gk_sHelpLinkToSFABrowser As String = gk_sHelpPagesPath & "HelpSFABrowser.htm"

'   d51     crm     09Dec1999     Leaflet Maintenance module
Public Const gk_sHelpLinkToDisplayAllLeafletFiles As String = gk_sHelpPagesPath & "HelpDisplayAllLeafletFiles.htm"
Public Const gk_sHelpLinkToDisplayLeafletDetail As String = gk_sHelpPagesPath & "HelpDisplayLeafletDetail.htm"
Public Const gk_sHelpLinkToDisplayEditLeafletDetail As String = gk_sHelpPagesPath & "HelpDisplayEditLeafletDetail.htm"


'd52
Public Const gk_sHelpLinkToMenuForRelatedSeparatedMaterial As String = gk_sHelpPagesPath & "HelpMenuForRelatedSeparatedMaterial.htm"

'help constants for Popular Search maintenance
'   d55     crm    10Dec1999
Public Const gk_sHelpLinkToDisplaySavedSearches As String = gk_sHelpPagesPath & "HelpToDisplaySavedSearches.htm"
Public Const gk_sHelpLinkToDisplayCreatePopularSearch As String = gk_sHelpPagesPath & "HelpToDisplayCreatePopularSearch.htm"
Public Const gk_sHelpLinkToDisplayPopularSearches As String = gk_sHelpPagesPath & "HelpToDisplayPopularSearches.htm"
Public Const gk_sHelpLinkToDisplayPopularSearchDetail As String = gk_sHelpPagesPath & "HelpToDisplayPopularSearchDetail.htm"

' d60 EAD parsing constants
Public Const gk_sLinkParseOnly As String = "EADParseOnly.asp"
Public Const gk_sLinkDisplayEADParseLogging As String = "EADParseLog.asp"
Public Const gk_sLinkDisplayEADFileList As String = "DisplayEADFileList.asp"


'
Public Const gk_sLinkSFAMaintenance  As String = "SFAMaintenance.htm"
'd64
Public Const gk_sMaintenanceHome As String = "Maintenance.htm"

Public Const gk_sIMAGEIndexButton As String = ""

Public Const gk_sLinkCatalogueToCopy As String = "StoreSelectedCatalogueForCreateByCopy.asp"
Public Const gk_sLinkMoveCatalogue As String = "StoreSelectedParentForMoveCatalogue.asp"
Public Const gk_sLinkSelectParentForCreateFromScratch As String = "StoreSelectedParentForCreateFromScratch.asp"
Public Const gk_sLinkSelectParentForCreateByCopy As String = "StoreSelectedParentForCreateByCopy.asp"

Public Const gk_sIMAGEUndoPickTermButton As String = gk_sImagesPath & "niceUndoPickTermButton.gif"

Public Const gk_sPAGEEditorialErrorReport As String = "EditorialErrorReport.asp"
Public Const gk_sPARAMErrorNumber As String = "ErrorNum"
Public Const gk_sPARAMErrorDesc As String = "ErrorDesc"
Public Const gk_sPARAMErrorExtraText As String = "ErrorExtraText"
Public Const gk_sPARAMErrorSource As String = "ErrorSource"
Public Const gk_sPARAMErrorLinkBack As String = "linkback"

' d80 - Added
Public Const gk_sPAGEDeleteLeaflet As String = "DeleteLeaflet.asp"
Public Const gk_sImageDeleteLeaflet As String = "niceDeleteLeaflet.gif"

' d81 - Added
Public Const gk_sLinkStartUpdatingIntoEditset As String = "StartUpdatingIntoEditset.asp"
Public Const gk_sLinkCreateEmptyEditSetForUpdate As String = "CreateEmptyEditSetForUpdate.asp"
Public Const gk_sLinkDisplayCreateEmptyEditSetForUpdate As String = "DisplayCreateEmptyEditSetScreenForUpdate.asp"

' d85 - Added
Public Const gk_sLinkConfirmJobCreation As String = "ConfirmScheduledJobCreation.asp"
Public Const gk_sLinkCheckForScheduledJob As String = "CheckForScheduledJob.asp"
Public Const gk_sLinkUpdateScheduledJobs As String = "UpdateScheduledJobs.asp"
Public Const gk_sLinkDisplayScheduledJobs As String = "DisplayScheduledJobList.asp"
Public Const gk_sLinkDisplayCreateEmptyEditSetScreenForDRUIDLoad As String = "DisplayCreateEmptyEditSetScreenForDRUIDLoad.asp"
Public Const gk_sLinkCreateEmptyEditSetForDRUIDLoad As String = "CreateEmptyEditSetForDRUIDLoad.asp"
Public Const gk_sLinkStartDRUIDLoadIntoEditset As String = "StartDRUIDLoadIntoEditset.asp"
Public Const gk_sLinkUpdateDRUIDBatchStatusToScheduled As String = "UpdateDRUIDBatchStatusToScheduled.asp"

'd87 Added
Public Const gk_sHelpLinkParentOfPiecesEditset As String = gk_sHelpPagesPath & "HelpParentOfPiecesEditset.htm"
Public Const gk_sHelpLinkParentForPiecesEditset  As String = gk_sHelpPagesPath & "HelpParentForPiecesEditset.htm"
Public Const gk_sHelpLinkChoosePiecesEditset  As String = gk_sHelpPagesPath & "HelpChoosePiecesEditset.htm"

'd88
Public Const gk_sLinkDisplayDeleteEditSetDetails As String = "DisplayDeleteEditSetDetails.asp"
Public Const gk_sLinkDisplayConfirmationForCatalogueUnDeletion As String = "DisplayConfirmationForCatalogueUnDeletion.asp"
Public Const gk_sHelpLinkToUnDeleteCatalogue As String = gk_sHelpPagesPath & "HelpUnDeleteCatalogue.htm"
Public Const gk_sHelpLinkToDeleteEditsetEntryList As String = gk_sHelpPagesPath & "HelpDeleteEditsetEntryList.htm"
