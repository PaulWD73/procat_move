Attribute VB_Name = "PROCATEditorErrorHandler"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    PROCATEditorErrorHandler.bas
' System:       PROCAT, PRO
' Copyright:    (C) Quidnunc Limited
'
' Description:  This Module handles the Specific User specific Error Types,
'                and displays a user frielndly error message
'
' Amendment history:
'   d1      ksk     01Feb1999       Created
'   d2      ndk     04Jun1999       Added functionality to handle Invalid Login in PROCATEditorialErrorHandling
'   d3      ndk     02Jul1999       Added path name to the error page in PROCATEditorialErrorHandling
'   d4      ksk     06JUL1999       Moved EProcatInternalError definition from Errohandler
'   d5      ksk     08JUL1999       Added code to handle the INVALID EDIT SET error
'   d6      ksk     12JUL1999       Added Invalid error for NextPage and Catalogue Reference
'   d7      ndk     14JUL1999       Added provision to handle unhandled errors, added parameter source
'                                   added option explicit
'   d8      Geedha Sivanadian   22Jul1999   Added provision to handle errors in AF files maintenance.
'   d9      ksk     28JUL1999       Changed few constants for the internal error.
'                                   Added Constant SerarchValueNotFound. Split
'                                   the error constants as User and System.
'                                   Added method GeneralErrorHandler to handle module specific errors
'   d10      ksk     28JUL1999      Added extra enumeration for Datatype mismatch and length problems
'                                   Added constants for BoReleaseController PEInvalidReleaseMode
'   d11      ksk     05Aug1999      For mandatory values missing in an object.
'   d12      ksk     06Aug1999      Added constants property setting violating the business rule
'                                   and no value for the property from the database.
'   d13      ksk     08Oct1999      Added code and enum to handle access denied error
'   d14      ndk     16Oct1999      Added code and enum to handle catalogue locking error
'                                   code and enum for errors raised when required field are not present for catalogues
'   d15      aga     28Oct1999      Added code and enum to handle deletion of last procat manager
'   d16      aga     28Oct1999      Added code and enum to handle deletion of rejected AF term
'   d17      aga     29Oct1999      Added code and enum to handle Uniqueness of User (NT Logon, User name & Initials)
'   d18      ndk     01Nov1999      Added new enums and code to handle them
'   d19      ndk     02Nov1999      Added new enums and code to handle them
'   d20      ksk     02Nov1999      Added general handler for PECatalogue, peBulkUpdate, peEditSet
'   d21      ndk     03Nov1999      Added peTryingToAddLinkToItSelf
'   d22      ndk     12Nov1999      Added new enums
'                                   Added an element - peErrorWithCompleteDescription- in enum , the error descriptions of all enums for this
'                                   and after will be shown without the source or error number
'   d23      ksk     15Nov1999      Added error handler for EAD Import.
'   d24      ndk     15Nov1999      Added more errorhandlers for move catalogue entry
'   d25      ndk     15Nov1999      error handlers for catalogue reference validation for new catalogues
'   d26      ndk     22Nov1999      Added new error enum peEntryNotFoundInLiveTables
'   d27      ndk     25Nov1999      Added new error enums
'   d28      jes     26Nov1999      Added some more error enums to integrate the search
'   d29      ndk     09Dec1999      Added error enums
'   d30      ndk     10Dec1999      Added error enum
'   d31      ndk     17Dec1999      Added enum ESQLServerNativeErrors to store sql server native error numbers
'                                   Added new enums to EPROCATINTERNALERROR
'   d31      ndk     18Dec1999      Moved enum ESQLServerNativeErrors to ErrorHandling module
'   d32      ndk     22Dec1999      Added PELastProcatInternalError, this will be the last enum in the error enum element
'   d33      ndk     28Dec1999      Added new error enums
'   d34      jes     25Jan2000      Added new error enums
'   d35      crm     07Feb2000      Added extra argument for GeneralErrorHandler
'   d36      crm     07Feb2000      Added enum instance for EPROCATINTERNALERROR
'   d37      crm     14Feb2000      Added enum instance for EPROCATINTERNALERROR
'   d38      crm     15Feb2000      Chnage to display description only for ESwith candidate terms not
'                                   allowed to pre-release stage
'   d39      crm     16Feb2000      Added enum instance
'   d40      WDP     23Feb2000      Altered error handling at highest level
'                                   to redirect to error page with error details
'                                   in the http parameters. This prevents further calls
'                                   being made on components from asp pages being processed
'                                   and enables error handling to be done in a separate
'                                   component as in reader.
'   d41     tmj     06Mar2000       Added PEItemCannotBeParent and PEParentMoved
'   d42     tmj     07Mar2000       Added PEUserInAuditTrail
'   d43     tmj     17Mar2000       Added PEDupPopSearchName
'   d44     tmj     21Mar2000       Added PEUserCreatedIndexTerm
'   d45     daj     27Mar2000       Added PECannotAddNonLiveParent
'   d46     daj     15Jun2000       Added peUSERCannotAddDuplicateLookupValue
'   d47     adf     10Jul2000       Added enums for EPROCATINTERNALERROR and code to handle
'                                   (Export Errors)
'   d48     daj     10Jan2001       Added error code for wildcard Go To search problems
'   d50     alm     26Jan2001       Added PEUSERInvalidReference - Invalid reference
'   d51     bcj     01Feb2001       Added  peUSEREmptyReference
'   d52     map     09Mar2001       Added peInvalidLeafletStructure
'   d53     map     09Mar2001       Added peEntryInAnotherEditSet
'   d54     map     14Mar2001       Added peDRUIDLoadAlreadyStarted
'   d55     alm     19Mar2001       Added the option t_sLinkBackPage - although it is not
'                                   actually used in Editorial
'   d56      map     11Oct2001       Added peScheduler
'   d57     alm     17Oct2001       Added PENoValidParentToMovePiecesFrom,
'                                   PENoValidParentToMovePiecesTo, PENoLivePieces
'   d58     map     22Oct2001       Added user errors to throw if user inputs incorrect time or date when scheduling
'   d59     alm     14Nob2001       Now takes an optional param t_sLinkBackPage which
'                                   specifies how many pages a back button should go
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Private Variables
Option Explicit

Private Const klCustomErrorStartValue As Long = 65535
Private Const klUserErrorStartFromCustomError As Long = 100

'd4 ksk 06JUL1999
' Components are supposed to return errors above vbObjectError + 512.
' So PEInternalError is -2147220992 !
'PERD - PE Reference Data
'PETM - PE Transience Manager
'PEDB - PE Database
'PEOC - PE Object Context
Public Enum EPROCATINTERNALERROR
    ' System errors
    PEGeneralInternalError = vbObjectError + klCustomErrorStartValue
    peRDRecordListEmpty
    PEReRaised
    peOCUnableToGetObjectContext
    
    ' User errors
    PEDBConnectionTimeout = vbObjectError + klCustomErrorStartValue + klUserErrorStartFromCustomError
    peDBInvalidUser
    PETMInvalidInstanceReference
    
    'd31
    PEDBSQLServerErrorUniqueIndexViolation
    PEDBSQLServerErrorUniqueConstraintViolation
    PEDBSQLServerErrorInvalidObjectName


    'd5      ksk     08JUL1999
    peEditSetIDInvalue
    'd6      ksk
    'd9 changed the Constant name
    peInvalidNextPage
    peInvalidCatalogueReference
    'd8     GS
    peAuthorityFilesMaintenance
    'd9 ksk
    peRDSearchValueNotFound
    
    'Invalid release value
    PEInvalidReleaseMode
    
    
    'd12    ksk
    pePropertyValueInDatabase
    peNoPropertyAccessViolation
    
    peAccessDenied

    ' d48 - Added
    PEUSERInvalidGoToWildcardSearch
    
    'NOTE : d22
    'put your error code after this if only its description is to be displayed
    '
    peErrorWithCompleteDescription
    
    peCanNotDeleteSFAWithReferences
    peDuplicateReferenceKey
    PEInvalidEditSetHistory
    PEDuplicateEditSetName
    PEInvalidParentLevelForCreateByCopying
    PECannotCreateLetterCodeByCopying
    PECanNotMoveCataloguePresentInMoreThanOneEditset
    PECanNotMoveCataloguesOtherThanClassOrPiece
    PECanNotMoveCatalogueToParentOfDifferentLevel
    peTryingToAddLinkToItSelf
    PEAFEditorTryingToChangeOtherRole
    PENoProofReaderPresentOnEditSet
    PENotAValidForwardTransition
    PENotAValidReverseTransition
    PEOpertionNotValidAtThisStage
    PEOtherUserDoesNotHaveRoleOnEditSet
    PEProofReaderTryingToChangeOtherRole
    PERoleDoesNotBelongToAllStagesOtherRoleBelongs
    PERoleDoesNotBelongToStage
    PERoleIsNotHigherOrEqualToOtherRole
    PERoleIsNotTheHighestRoleForThisStage
    PETryingToAddRoleNeverToBeAdded
    PETryingToRemoveReplaceOnlyRole
    PEUserDoesNotHaveRoleOrHisRoleNotPermittedToPerformThisOperation
    PEUserTryingToChangeOtherUsersCheckBox
    PEUserBeingAddedAlreadyHasARoleOnEditSet
    peCustodialHistoryRowNotPresent
    peAppDestInformationRowNotPresent
    peAccrualsRowNotPresent
    peTitleRowNotPresent
    peArrangementRowNotPresent
    peScopeAndContentRowNotPresent
    peNoteRowNotPresent
    peAdminHistoryRowNotPresent
    peDuplicateLinkItemChoosen
    peInvalidCombinationOfClosureTypeAndCode
    peCatalogueAlreadyLocked
    PECatalogue
    peBulkUpdate
    peCanNotDeleteLastProcatManager
    peCanNotDeleteRejectedAFTerm
    'd17    For checking duplicate user
    peDuplicateUserName
    peDuplicateUserNTLogon
    peDuplicateUserInitials
    peEditSet
    peMandatoryValuesMissing
    'd10     ksk for invalid data types
    peIncorrectDataType
    peIncorrectLength
    
    'd23
    PEImport
    
    'd24
    pETryingToDeleteNotToBeDeletedEntries
    PECanNotMoveCatalogueToInvalidParent
    pECanNotMoveCatalogueToExisintParent
    
    'd25
    peInvalidReferenceForNewCatalogue
    
    'd26
    peEntryNotFoundInLiveTables

    'd27
    peDefaultValuesForCataloguesRowNotPresent
    peDefaultValuesForCataloguesMultipleRowsPresent
    
    'd28
    PEUSERInvalidDepthRestrictions
    PEUSERInvalidBrackets
    PEUSERInvalidDateField
    PEUSERInvalidSearchTermField
    PEUSERInvalidDateRange
    PEUSERInvalidGoToReference
    PEUSERInvalidBreadthRestrictions
    PEUSERInvalidPosIntegerField
    
    'd29
    peMoreThanOneCatalogueFoundForRelatedSeparateMaterial
    peNoCatalogueFoundForRelatedSeparateMaterial
    
    'd30
    PECanNotAddCatalogueAlreadyPresentAsEditable
    
    '
    PEDBGeneralUniqueConstraintViolation
    PEDBGeneralInvalidObjectName
    
    'd33
    peTitleCanNotBeEmpty
    peScopeAndContentCanNotBeEmpty
    
    'd41     tmj     06Mar2000
    PEItemCannotBeParent
    
    'd42     tmj     07Mar2000
    PEUserInAuditTrail
    
    'd43     tmj     17Mar2000
    PEDupPopSearchName
    
    'd44    tmj     21Mar2000
    PEUserCreatedIndexTerm
    
    ' d45 - Added
    PECannotAddNonLiveParent
    
    PEUSERVerityCollectionProblem
    
    ' d47 - Added
    peUSERInvalidExportFileName
    peUSERMissingExportDeptRef
    peUSERNoExportDataSelected
    
    'Note : Put all your errors before the enum : PELastProcatInternalError
    'And update klCustomErrorStartValue
    PELastProcatInternalError
    
    'd34
    peUSERDorisTerminalNTLoginFound
    PEUSEROverBookmarkLimit
    PEUSEROverSavedSearchLimit
    PEUSERBadFTQuery
    PEUSERNoSearchToRun
    peUSERNoSearchTerms
    peUSERSearchNoLongerInTransientStore
    peUSEREmptyReference
    peUSERNoSearchFieldSelected
'   d36      crm     07Feb2000
    PEUSEREmptySearchTermField
'   d37      crm     14Feb2000
    PECanNotMoveESWithCandidateTermsToPreRelease
'   d39      crm     16Feb2000
    PECanNotMakeNonPrefTermIcompApprovWithPrefTermIcompApprov

'   d46 - Added
    peUSERCannotAddDuplicateLookupValue
    
    ' d50 - Added
    PEUSERInvalidReference
    
    'd52 Added
    peInvalidLeafletStructure
    
    'd53 Added
    peEntryInAnotherEditSet
    
    'd54 Added
    peDRUIDLoadAlreadyStarted
    
    ' d56 - Added
    peScheduler
    
    ' d57 - Added
    PENoValidParentToMovePiecesFrom
    PENoValidParentToMovePiecesTo
    PENoLivePieces
    
    ' d58 - Added
    peUSERInvalidTime
    peUSERInvalidDate
    
End Enum



Public Sub GeneralErrorHandler(ByVal t_lErrorNumber As Long, _
                               ByVal t_sErrorSource As String, _
                               ByVal t_sErrorDescription As String, _
                               Optional t_sExtraText As String, _
                               Optional t_sLinkbackPage As String)

'   d9      ksk     28JUL1999       Created the general error handler that calls the editorial
'                                   error handler
'   d35     crm    07Feb2000      Added extra argument for compatibility with reader functionality
'                                   NOTE: Could use the extra argument in the future
'                                   if we need to
'   d40      WDP     23Feb2000      Altered error handling at highest level
'                                   to redirect to error page with error details
'                                   in the http parameters. This prevents further calls
'                                   being made on components from asp pages being processed
'                                   and enables error handling to be done in a separate
'                                   component as in reader.
'  d55      alm    19Mar2001        Added the option t_sLinkBackPage - although it is not
'                                   actually used in Editorial
'  d59      alm    14Nob2001        Now takes an optional param t_sLinkBackPage which
'                                   specifies how many pages a back button should go
    Dim aspResponse As Response
    Set aspResponse = GetObjectContext.Item("Response")
    
    #If PROCATDEBUGMODE Then
    'Generate a debug message
    'do not redirect, use old method instead
        PROCATEditorialErrorHandling t_lErrorNumber, t_sErrorSource, t_sErrorDescription
    
    #Else
    'Redirect to the display error page
        Dim sRedirectURL As String
        sRedirectURL = gk_sPAGEEditorialErrorReport & "?" & gk_sPARAMErrorNumber & "=" & CLng(t_lErrorNumber) & _
                        "&" & gk_sPARAMErrorSource & "=" & URLEncode(t_sErrorSource) & _
                        "&" & gk_sPARAMErrorDesc & "=" & URLEncode(t_sErrorDescription) & _
                        "&" & gk_sPARAMErrorExtraText & "=" & URLEncode(t_sExtraText)
                        
        ' d59 Append to the redirect string the number of pages any back button
        ' rendered on the page should go
        If Not IsEmpty(t_sLinkbackPage) Then
            sRedirectURL = sRedirectURL & "&" & gk_sPARAMErrorLinkBack & "=" & URLEncode(t_sLinkbackPage)
        End If
        
        aspResponse.Redirect sRedirectURL
    #End If
    

End Sub
                               
Public Sub PROCATEditorialErrorHandling(ByVal t_lErrorNumber As Long, _
                               ByVal t_sErrorSource As String, _
                               ByVal t_sErrorDescription As String, _
                               Optional ByVal t_sLinkbackPage As String = "")
'   d2      ndk     04Jun199        Completely rewritten
'   d3      ndk     02Jul1999       Added path name to the error page
'   d5      ksk     08JUL1999       Added call to the VWError object to
'                                   display a error page for invalid edit set
'   d7      ndk     14JUL1999       Provision made for unhandled errors,added parameter source
'   d9      ksk     28JUL1999       Split the logging of error as user error and system error.
'   d15     aga     28Oct1999       Added code  to handle deletion of last procat manager
'   d16     aga     28Oct1999       Added code  to handle rejection (deletion) of AF Term
'   d17     aga     29Oct1999       Added code to handle Uniqueness of User (NT Logon, User name & Initials)
'   d22     ndk     12Nov1999       For all errors above or equal the value of  peErrorWithCompleteDescription
'                                   only the description will be displayed
'   d32     ndk     22Dec1999      errors treated as description only errors if the error value is less than or equal to
'                                   PELastProcatInternalError
'   d50     alm     26Jan2001       Added PEUSERInvalidReference - Invalid reference
'   d52     map     14Mar2001       (Added this comment as wasn't here) Added peInvalidLeafletStructure
'   d53     map     14Mar2001       Added peEntryInAnotherEditSet
'   d54     map     14Mar2001       Added peDRUIDLoadAlreadyStarted
'   d59     alm     14Nov2001       Now takes an optional param t_sLinkBackPage which
'                                   specifies how many pages a back button should go
'What happens if an error occurs in this function
    'Yet to work on that
    
    Dim ocObjectContext As ObjectContext
    Dim oreResponse As Response
    Dim oVWErrHandler As VWError
    Dim sMessage As String

    sMessage = "[" & t_sErrorSource & "], " & _
               CStr(t_lErrorNumber) & ", " & t_sErrorDescription

        
    Set ocObjectContext = GetObjectContext()
    If ocObjectContext Is Nothing Then
        'This part needs to be worked on
        'Raise an appropriate error
        'Check again , you may go into a recursive loop
    End If
    Set oreResponse = ocObjectContext.Item(gk_sResponseObject)
    
    Set oVWErrHandler = New VWError
    
    'd32
        If t_lErrorNumber >= EPROCATINTERNALERROR.peErrorWithCompleteDescription And _
        t_lErrorNumber <= EPROCATINTERNALERROR.PELastProcatInternalError Then
        oVWErrHandler.DisplayErrorDescriptionOnly t_sErrorDescription, t_sLinkbackPage
    Else
        
        Select Case t_lErrorNumber
        
            'd5  Invalid edit set ID then show the message in the page
            Case EPROCATINTERNALERROR.peEditSetIDInvalue
                oVWErrHandler.DisplayInvalidEditsetPage t_sLinkbackPage
            
            'd6  Invalid Next page entry
            Case EPROCATINTERNALERROR.peInvalidNextPage
                oVWErrHandler.DisplayInvalidNextPage t_sLinkbackPage
            
            'd6  Invalid catalogue reference value
            Case EPROCATINTERNALERROR.peInvalidCatalogueReference
                oVWErrHandler.DisplayInvalidCatalogueReferencePage t_sLinkbackPage
            
            'The login name was invalid, unable to login etc
            Case EPROCATINTERNALERROR.peDBInvalidUser
                'Abandon the session
                ocObjectContext.Item(gk_sSessionObject).Abandon
                'Redirect to error page
                oreResponse.Redirect (gk_sErrorPagesPath & gk_sLinkErrorLogin)
            
            
            'd13 ksk  Added code and enum to handle access denied error
            Case EPROCATINTERNALERROR.peAccessDenied
                oVWErrHandler.DisplayErrorDescriptionOnly t_sErrorDescription, t_sLinkbackPage
                
            'd57
            Case PENoValidParentToMovePiecesFrom, _
                 PENoValidParentToMovePiecesTo, _
                 PENoLivePieces
    
                oVWErrHandler.DisplayErrorDescriptionOnly t_sErrorDescription, t_sLinkbackPage
                
                
            'dx
            Case peUSERInvalidDate, _
                 peUSERInvalidTime
                 
                 oVWErrHandler.DisplayErrorDescriptionOnly t_sErrorDescription, t_sLinkbackPage
                
            
            'd51 Added peUSEREmptyReference
            'd52 Added peInvalidLeafletStructure
            'd53 Added peEntryInAnotherEditSet
            'd54 Added peDRUIDLoadAlreadyStarted
            'd56 Added peScheduler
            'Search errors
            Case EPROCATINTERNALERROR.PEUSERBadFTQuery, EPROCATINTERNALERROR.PEUSEREmptySearchTermField, _
                    EPROCATINTERNALERROR.PEUSERInvalidBrackets, EPROCATINTERNALERROR.PEUSERInvalidBreadthRestrictions, _
                    EPROCATINTERNALERROR.PEUSERInvalidDateField, EPROCATINTERNALERROR.PEUSERInvalidDateField, _
                    EPROCATINTERNALERROR.PEUSERInvalidDateRange, EPROCATINTERNALERROR.PEUSERInvalidDepthRestrictions, _
                    EPROCATINTERNALERROR.peUSERNoSearchTerms, EPROCATINTERNALERROR.PECanNotMoveESWithCandidateTermsToPreRelease, _
                    EPROCATINTERNALERROR.PECanNotMakeNonPrefTermIcompApprovWithPrefTermIcompApprov, peUSERSearchNoLongerInTransientStore, _
                    peUSERNoSearchFieldSelected, PEUSERInvalidPosIntegerField, PEUSERVerityCollectionProblem, peUSERCannotAddDuplicateLookupValue, _
                    EPROCATINTERNALERROR.peUSERInvalidExportFileName, EPROCATINTERNALERROR.peUSERMissingExportDeptRef, EPROCATINTERNALERROR.peUSERNoExportDataSelected, _
                    PEUSERInvalidGoToWildcardSearch, PEUSERInvalidReference, peUSEREmptyReference, peInvalidLeafletStructure, peDRUIDLoadAlreadyStarted, peEntryInAnotherEditSet, _
                    peScheduler
                    '   d38      crm     15Feb2000
                    '   d39      crm     16Feb2000
                    '   d47      adf     10Jul2000
                oVWErrHandler.DisplayErrorDescriptionOnly t_sErrorDescription, t_sLinkbackPage
            
            Case Else       'd7 For unhandled errors
                oVWErrHandler.DisplayUnhandledErrorPage sMessage, t_sLinkbackPage
                
        End Select
    End If
    Set oreResponse = Nothing
    Set ocObjectContext = Nothing
End Sub

