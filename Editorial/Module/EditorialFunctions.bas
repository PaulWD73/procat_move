Attribute VB_Name = "EditorialFunctions"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    EditorialFunctions.cls
' System:       PROCAT, PRO
' Copyright:    (C) Quidnunc Limited
'
' Description:  global functions used just by editorial
'
' Amendment history:
'   d1      WDP     22Oct1999     Created
'   d2      KSK     27Oct1999     Validation method to validate the closure
'                                   code value given the closure type value.
'                                   Added function NullToSQLNull.
'   d3      NDK     01Nov1999     Using variantstringasstrring instead of variantlongasstring which was used by mistake
'                                 in method IsValidClosureTypeAndCode
'   d4      ndk     12Nov1999     Added some new functions related to error handling of security violations
'   d5      KSK     17Nov1999     Changed to condition in IsValidClosureTypeAndCode
'                                 to validate empty value for closure code
'   d6      WDP     18Nov1999     Added Tag and FindTaggedData
'   d7      JES     26Nov1999     Added GetSessionID() for search
'   d8      WDP     01Dec1999     Added LogReleaseMessage
'   d9      WDP     20Jan2000     Added CheckLogFileFroErrors from BOCatalogueIndexer
'   d10     JCC     07Feb2000     Altered some of the text strings for grammar etc
'   d11     CRM     14Feb2000     Added error desc for editset with candidate terms moving into pre-release
'   d12     WDP     24Feb2000     Added BackButtonHTML
'   d13     WDP     25Feb2000     Strip all tags in GetSortedTitle
'   d14     WDP     02Mar2000     Added GetPitemIDFromParentIDAndKeyOrder
'   d15     WDP     23Mar2000     Added GetPitemIDFromParentIDAndKeyOrderSQL
'   d16     TMJ     30Mar2000     Added calls to SpecialCharacterDecode
'   d17     JCC     04Apr2000     Change closure code error message
'   d18     WDP     13Apr2000     Alter LogAction to allow dummy (blank) actions
'   d19     DAJ     17Oct2000     Changed IsValidClosureTypeAndCode
'   d20     GDB     12Feb2001     in IsValidClosureTypeAndCode, added case A and added
'                                 an elseif condition to case U
'   d21     gdb     13Feb2001     use EndDate startdate in valid closure
'   d22     gdb     13Feb2001     created IsValidOpendate
'   d23     gdb     14Feb2001     created function IsValidClosureStatus
'   d24     pnp     15Feb2001     created function GetLastUsedRef
'   d25     pnp     15Feb2001     updated GetLastUsedRef to handle null record set
'   d26     gdb     20Feb2001     Defect fixing on IsvalidClosureTypeandCode - case F - now handle null value
'   d27     alm     26Feb2001     Moved GetSessionID to GlobalFunctions
'   d28     map     05Mar2001     Changed to return zero when at lettercode level
'   d29     alm     19Oct2001     Added support for new closure code D - Retained Until
'   d30     alm     14Nov2001     Re-wrote BackButtonHTML to support a back button of multiple pages
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Const mk_sIndexInsert As String = "insert"
Private Const mk_sIndexUpdate As String = "update"
Private Const mk_sIndexDelete As String = "delete"

Private Const mk_sReleaseLogPath As String = "C:\PROCAT\Log\Release\"

Public Function GetIndexActionString(t_eAction As EIndexingAction) As String
'   Returns a string representing the indexing action
'
'   d1  WDP     22Oct1999     Created

    Dim sSource As String
    sSource = "GetIndexActionString"

    Select Case t_eAction
    Case IDXInsert:
        GetIndexActionString = mk_sIndexInsert
    Case IDXUpdate:
        GetIndexActionString = mk_sIndexUpdate
    Case IDXDelete
        GetIndexActionString = mk_sIndexDelete
    Case Else
        RaiseError sSource, PEGeneralInternalError, "invalid indexing action enum value"
    End Select

End Function


Public Function GetIndexActionEnum(t_sAction As String) As EIndexingAction
'   Returns a enum representing the indexing action string
'   Opposite of GetIndexActionString
'
'   d6  WDP     24Nov1999     Created

    Dim sSource As String
    sSource = "GetIndexActionEnum"

    Select Case t_sAction
    Case mk_sIndexInsert:
        GetIndexActionEnum = IDXInsert
    Case mk_sIndexUpdate:
        GetIndexActionEnum = IDXUpdate
    Case mk_sIndexDelete
        GetIndexActionEnum = IDXDelete
    Case Else
        RaiseError sSource, PEGeneralInternalError, "invalid indexing action string value"
    End Select

End Function
Public Function GetLastUsedRef(ByVal t_iLevel As Integer, ByVal t_iParentId, _
                        Optional ByVal t_iParentLevel As Integer) As Variant

'   d24     pnp     15Feb2001     created
'   d28     map     05Mar2001     Changed to return zero when at lettercode level
On Error GoTo ErrHandler
    
    Dim sSource As String
    Dim ofwQuery As FWDBHelper
    Dim sSql As String
    Dim ofwResult As FWRecordList
    
    sSource = "GetLastUsedRef"
    
    'd28
    If t_iParentLevel = 0 Then
        GetLastUsedRef = ""
    Else
        sSql = "EXECUTE sp_Edit_GetLastUsedRef"
        sSql = sSql & " @t_iLevel = " & CStr(t_iLevel)
        sSql = sSql & ", @t_iParentID = " & CStr(t_iParentId)
        sSql = sSql & ", @t_iParentLevel = " & CStr(t_iParentLevel)

        Set ofwQuery = New FWDBHelper

        Set ofwResult = ofwQuery.FindBySql(sSql, gk_lMaxEditsetSizeToRetrieve)
        If ofwResult.IsAtEnd Then '   d25
            GetLastUsedRef = ""
        Else
            ofwResult.MoveFirst
            GetLastUsedRef = ofwResult.GetFieldValue(gk_sDBCOLValue)
        End If
    End If
Exit Function
    
ErrHandler:
    ConvertError sSource

End Function

Public Function NullToSQLNull(ByVal t_vValue As Variant) As Variant
'This function returns the string value null if the variant value is null.
'
'   d2      KSK     27-Oct-1999
    If IsNull(t_vValue) Then
        NullToSQLNull = "null"
    Else
        NullToSQLNull = t_vValue
    End If
End Function


Public Function IsValidClosureTypeAndCode(ByVal t_vClosureType As Variant, _
                    ByVal t_vClosureCode As Variant, _
                    ByRef t_sErrorMessageForClosureTypeAndCode As String, _
                    Optional ByVal t_vStart As Variant, _
                    Optional ByVal t_vEnd As Variant _
                    ) As Boolean
'
'Description: Validates the closure type and closure code fields for piece and item
'             The error message is returned in the variable t_sErrorMessageForClosureTypeAndCode which
'             is passed by reference
'
'   d2      KSK     27Oct1999    Validation method to validate the closure
'                                   code value given the closure type value
'   d3      NDK     01Nov1999     Using variantstringasstrring instead of variantlongasstring which was used by mistake
'                                 in this method
'   d5      KSK     17Nov1999     Changed to condition in IsValidClosureTypeAndCode
'                                 to validate empty value for closure code
'   d17     JCC     04Apr2000     Change closure code error message
'   d19     DAJ     17Oct2000     Changed Case of T, S or R where the type was checked instead of the code
'   d20     GDB     12Feb2001     in IsValidClosureTypeAndCode, added case A and added
'                                  an elseif condition to case U
'   d21     gdb     13Feb2001     use EndDate startdate in valid closure
'   d26     gdb     20Feb2001     Defect fixing on IsvalidClosureTypeandCode - case F - now handle null value
'   d29     alm     19Oct2001     Added support for new closure code D - Retained Until
'
    On Error GoTo ErrorHandler

    Dim sSource As String
    Dim sClosureType As String, sClosureCode As String, sErrorMessage As String
    Dim ofwReferenceData As FWReferenceData
    
    sSource = ".IsValidClosureTypeAndCode"
    Set ofwReferenceData = New FWReferenceData
    
    'convert to string
    sClosureType = VariantStringAsString(t_vClosureType)
    sClosureCode = VariantStringAsString(t_vClosureCode)
    
    'default value
    IsValidClosureTypeAndCode = True
    
    ' d21 - convert procat date
    Dim sConvertedStartDate As Variant
    Dim sConvertedEndDate As Variant
    Dim sStartYear As String
    Dim sEndYear As String
    

    Select Case sClosureType
        Case ""
            If sClosureCode <> "" Then
                sErrorMessage = "Closure code should also be empty when closure status is empty"
                
                IsValidClosureTypeAndCode = False
            End If
        
        
        Case "N"
            If sClosureCode <> "30" Then
                sErrorMessage = "Closure code should be 30 when closure status is " & _
                ofwReferenceData.Search(gClosureType, "N")
                
                IsValidClosureTypeAndCode = False
            End If
                    
        Case "F"
        
        'd26
        If (Not sClosureCode = "") Then
        'd21
              If CLng(sClosureCode) <= 0 Then
                    sErrorMessage = "Closure code should be greater than 0 when closure status is " & _
                    ofwReferenceData.Search(gClosureType, "F")
                    
                    IsValidClosureTypeAndCode = False
        
            End If
        Else
           
            sErrorMessage = "Closure code should not be empty when closure status is " & _
                    ofwReferenceData.Search(gClosureType, "F")
                    
                    IsValidClosureTypeAndCode = False
            End If
            
        
        Case "U"
            If sClosureCode = "" Then
                sErrorMessage = "Closure code should be greater than 1800 when closure status is " & _
                ofwReferenceData.Search(gClosureType, "U")
            
                IsValidClosureTypeAndCode = False
            Else
                If CLng(sClosureCode) <= 1800 Then
                    sErrorMessage = "Closure code should be greater than 1800 when closure status is " & _
                    ofwReferenceData.Search(gClosureType, "U")
                    
                    IsValidClosureTypeAndCode = False
                'd20
                ElseIf CLng(sClosureCode) >= 2500 Then
                    sErrorMessage = "Closure code should be less than 2500 when closure status is " & _
                    ofwReferenceData.Search(gClosureType, "U")
                    
                    IsValidClosureTypeAndCode = False
                
                End If
            End If
            
        Case "I"
            If sClosureCode <> "0" Then
                sErrorMessage = "Closure code should be 0 when closure status is " & _
                ofwReferenceData.Search(gClosureType, "I")
                
                IsValidClosureTypeAndCode = False
            End If
        'd20
        Case "A"
            If sClosureCode <> "0" Then
                sErrorMessage = "Closure code should be 0 when closure status is " & _
                ofwReferenceData.Search(gClosureType, "A")
                
                IsValidClosureTypeAndCode = False
            End If

        Case "T", "S", "R"
            ' d19 - Check the Code instead of the Type (doh - the Case is selecting on the Type!)
            If sClosureCode <> "" Then
                sErrorMessage = "Closure code should be empty when closure status is " & _
                ofwReferenceData.Search(gClosureType, "T") & vbCr & _
                " or " & ofwReferenceData.Search(gClosureType, "S") & vbCr & _
                " or " & ofwReferenceData.Search(gClosureType, "R")
                
                IsValidClosureTypeAndCode = False
            End If
            
        'd29
        Case "D"
            If sClosureCode = "" Then
                sErrorMessage = "Closure code should be greater than 1800 when closure status is " & _
                ofwReferenceData.Search(gClosureType, "D")
            
                IsValidClosureTypeAndCode = False
            Else
                If CLng(sClosureCode) <= 1800 Then
                    sErrorMessage = "Closure code should be greater than 1800 when closure status is " & _
                    ofwReferenceData.Search(gClosureType, "D")
                    
                    IsValidClosureTypeAndCode = False

                ElseIf CLng(sClosureCode) >= 2500 Then
                    sErrorMessage = "Closure code should be less than 2500 when closure status is " & _
                    ofwReferenceData.Search(gClosureType, "D")
                    
                    IsValidClosureTypeAndCode = False
                
                End If
            End If
                    
    End Select
    'If the comination of type and code is invalid then, change the parameter message passed as reference
    If Not IsValidClosureTypeAndCode Then
        t_sErrorMessageForClosureTypeAndCode = sErrorMessage
    Else
        t_sErrorMessageForClosureTypeAndCode = ""
    End If
    
    Exit Function
ErrorHandler:
    ConvertError sSource
End Function
Public Function IsValidOpenDate(ByVal t_vClosureType As Variant, _
                    ByVal t_vClosureCode As Variant, _
                    ByVal t_vRecordOpeningDate, _
                    Optional ByVal t_vStart As Variant, _
                    Optional ByVal t_vEnd As Variant _
                    ) As String
'
'Description: Manipulates the Field Open Date according to the values present in Field Closure
'             Type and Field closure
'
'   d22     gdb     13Feb2001     created
'   d??     alm     02Mar2001     The record opening date is nullified if < 01 Jan 1998
'
    On Error GoTo ErrorHandler

    Dim sSource As String
    Dim sClosureType As String, sClosureCode As String, sErrorMessage As String
    Dim ofwReferenceData As FWReferenceData
    Dim bGeneratedDate As Boolean
    
    sSource = ".IsValidOpenData"
    Set ofwReferenceData = New FWReferenceData
        
    Const ksEarliestAutoDate = "01/01/1998"
    Const ksEarliestManualDate = "01/01/1753"
    
    'convert to string
    sClosureType = VariantStringAsString(t_vClosureType)
    sClosureCode = VariantStringAsString(t_vClosureCode)
    
    bGeneratedDate = False
    
    ' d21 - convert procat date
    Dim sConvertedStartDate As Variant
    Dim sConvertedEndDate As Variant
    Dim sStartYear As String
    Dim sEndYear As String
    Dim sCleanDate As String
    
    If Not (t_vStart = "") Then
    sConvertedStartDate = GlobalFunctions.ConvertProcatDateToYear(t_vStart)
    sStartYear = VariantStringAsString(sConvertedStartDate)
    End If
    
    If Not (t_vEnd = "") Then
    sConvertedEndDate = GlobalFunctions.ConvertProcatDateToYear(t_vEnd)
    sEndYear = VariantStringAsString(sConvertedEndDate)
    End If
    
    
    ' the following are variables that are being used for calculation
    ' whithin the scope of Select Case
    Dim lFinalYear As Long
    Dim lClosCode As Long
    Dim lNewDate As Long
    Dim sConvertLongToString
    
    
    Select Case sClosureType
        Case "T", "S", "R"
                
                IsValidOpenDate = "Null"
        
        
        Case "F"
        

        If Not (t_vEnd = "") Then
        
            ' get the final year value in field covering dates
            lFinalYear = sEndYear
            
            ' get the closure code
            lClosCode = CLng(sClosureCode)
            
            lNewDate = lFinalYear + lClosCode + 1
            
            sConvertLongToString = CStr(lNewDate)
            
            Dim vYearForCaseF As Variant
            vYearForCaseF = "01/01/" + sConvertLongToString
            vYearForCaseF = vYearForCaseF
           
             IsValidOpenDate = _
             Quotes(ConvertDateStringFromDDMMYYYYToMMDDYYYY(vYearForCaseF, "/"))
            
            bGeneratedDate = True
            
            ' insted If vbNullString(t_vEnd) ; so if it is null keep the data input by user
            Else
            ' if the user did not input any date
             If (t_vRecordOpeningDate = "") Then
              IsValidOpenDate = "Null"
             
             Else
              IsValidOpenDate = _
              Quotes(ConvertDateStringFromDDMMYYYYToMMDDYYYY(t_vRecordOpeningDate, "/"))
             End If
        
        End If
        
        ' case if normal closure is selected
        Case "N"
        If Not (t_vEnd = "") Then
            
            lFinalYear = sEndYear
            
            lNewDate = lFinalYear + 30 + 1
            
           
            sConvertLongToString = CStr(lNewDate)
            
            Dim vYearForCaseN As Variant
            vYearForCaseN = "01/01/" + sConvertLongToString
           
            IsValidOpenDate = _
            Quotes(ConvertDateStringFromDDMMYYYYToMMDDYYYY(vYearForCaseN, "/"))
            
            bGeneratedDate = True
    
        ' insted If vbNullString(t_vEnd) ; so if it is null keep the data input by user
        Else
            ' if the user did not input any date
             If (t_vRecordOpeningDate = "") Then
              IsValidOpenDate = "Null"
             
             Else
              IsValidOpenDate = _
              Quotes(ConvertDateStringFromDDMMYYYYToMMDDYYYY(t_vRecordOpeningDate, "/"))
             End If
        
        End If
        
        ' case if Closed Until is selected
        Case "U"
            
            ' number in closure code is converted in year date to be used
            ' in Field open date
            Dim vYearForCaseU As Variant
            
            vYearForCaseU = "01/01/" + t_vClosureCode
            
            IsValidOpenDate = _
            Quotes(ConvertDateStringFromDDMMYYYYToMMDDYYYY(vYearForCaseU, "/"))
            
            bGeneratedDate = True
        
        ' Case if Retained Until is selected
        Case "D"
            
            ' If the user did not input any date then raise an error
            If (t_vRecordOpeningDate = "") Then
             
              ' The user must enter a date - raise an error
              RaiseError sSource, PEUSERInvalidDateRange, "No Record Opening Date was specified"
             
             Else
              IsValidOpenDate = _
              Quotes(ConvertDateStringFromDDMMYYYYToMMDDYYYY(t_vRecordOpeningDate, "/"))
             End If
        
        ' default case if above cases do not match
        Case Else
                 
            ' if the user did not input any date
            If (t_vRecordOpeningDate = "") Then
             IsValidOpenDate = "Null"
             
             Else
              IsValidOpenDate = _
              Quotes(ConvertDateStringFromDDMMYYYYToMMDDYYYY(t_vRecordOpeningDate, "/"))
             End If
        
       End Select
       
       sCleanDate = Replace(IsValidOpenDate, "'", "")
       
       If Not IsValidOpenDate = "Null" Then
       
        If bGeneratedDate = True Then
                 
            ' If the record opening date is earlier than 1st Jan 1998 then just
            ' make is null
            If DateDiff("d", CDate(ksEarliestAutoDate), CDate(sCleanDate)) < 0 Then
                IsValidOpenDate = "Null"
            End If
        
        Else
                 
            ' If the record opening date is earlier than 1st Jan 1998 then just
            ' make is null
            If DateDiff("d", CDate(ksEarliestManualDate), CDate(sCleanDate)) < 0 Then
                IsValidOpenDate = "Null"
            End If
        
        End If
        
       End If
    'If the comination of type and code is invalid then, change the parameter message passed as reference
    Exit Function
    
ErrorHandler:
    ConvertError sSource
End Function

Public Function IsValidClosureStatus(ByVal t_vClosureType As Variant, _
                    ByVal t_vClosureCode As Variant, _
                    ByVal t_sRecordOpeningDate As String, _
                    ByVal t_vClosureStatus As Variant, _
                    ByRef t_sErrorMessageForClosureTypeAndCode As String, _
                    Optional ByVal t_vStart As Variant, _
                    Optional ByVal t_vEnd As Variant _
                    ) As Boolean
'
'Description: Validates the closure staus fields for piece and item
'             The error message is returned in the variable t_sErrorMessageForClosureTypeAndCode which
'             is passed by reference
'
'   d23     gdb     14Feb2001     created
'   d29     alm     19Oct2001     Added support for 'Retained Until'

    
    On Error GoTo ErrorHandler

    Dim sSource As String
    Dim sClosureType As String, sClosureCode As String, sErrorMessage As String, sClosureStatus As String
    Dim ofwReferenceData As FWReferenceData
    
    sSource = ".IsValidOpenData"
    Set ofwReferenceData = New FWReferenceData
    
    'convert to string
    sClosureType = VariantStringAsString(t_vClosureType)
    sClosureCode = VariantStringAsString(t_vClosureCode)
    sClosureStatus = VariantStringAsString(t_vClosureStatus)
    
    'default value
    IsValidClosureStatus = True
    
    'variable used for data manipulation within select case
    Dim sInterval
    Dim sDateTobeCleaned As String
    Dim sCleanDate
    
    Select Case sClosureStatus
        Case "C", "D"
                
              Select Case sClosureType
              
                Case "F", "N", "U"
                     If Not (t_sRecordOpeningDate = "Null") Then
                    ' set the interval in days
                    sInterval = "d"
                    
                    'sDateTobeCleaned & sCleanDate are used "to clean" the string from quotes so it is
                    ' ready to be casted to a Date
                    sCleanDate = Date
                    sDateTobeCleaned = Replace(t_sRecordOpeningDate, "'", "")
                    sCleanDate = CDate(sDateTobeCleaned)

                    
                    ' if the date in field open is less than today, then error
                    If (DateDiff(sInterval, Date, sCleanDate) < 0) Then
                    
                        sErrorMessage = "The field open date must be greater than today "
                        IsValidClosureStatus = False
                            Else
                            IsValidClosureStatus = True
                            End If
                    Else
                    IsValidClosureStatus = True
                    End If
                    
                Case "R", "S", "T"
                IsValidClosureStatus = True
                        

                Case "I", "A", "D" 'd29 Added support for 'Retained Until'
                ' Open Immediately and Open on transfer closure type not allowed here - hence error
                sErrorMessage = "Closure type should be " & _
                ofwReferenceData.Search(gClosureType, "T") & vbCr & _
                " or " & ofwReferenceData.Search(gClosureType, "U") & vbCr & _
                 " or " & ofwReferenceData.Search(gClosureType, "F") & vbCr & _
                  " or " & ofwReferenceData.Search(gClosureType, "N") & vbCr & _
                   " or " & ofwReferenceData.Search(gClosureType, "S") & vbCr & _
                " or " & ofwReferenceData.Search(gClosureType, "R")
                            
                IsValidClosureStatus = False
    
                ' end sClosureType select
                End Select
                
        ' this case O refers to Outer select
        Case "O"
                If Not (t_sRecordOpeningDate = "Null") Then
                
                    ' set the interval in days
                    sInterval = "d"
                    
                    'sDateTobeCleaned & sCleanDate are used "to clean" the string from quotes so it is
                    ' ready to be casted to a Date
                    sCleanDate = Date
                    sDateTobeCleaned = Replace(t_sRecordOpeningDate, "'", "")
                    sCleanDate = CDate(sDateTobeCleaned)

                    
                    ' if the date in field open is greater than today, then error
                  
                    If (DateDiff(sInterval, sCleanDate, Date) > 0) Then
                        IsValidClosureStatus = True

                    Else
                    sErrorMessage = "The field open date must be less than today date "
                    IsValidClosureStatus = False
                    End If
                 Else
                IsValidClosureStatus = True
                End If
       
       
       ' end outer select
       End Select
       
     'If the comination of type and code is invalid then, change the parameter message passed as reference
    If Not IsValidClosureStatus Then
        t_sErrorMessageForClosureTypeAndCode = sErrorMessage
    Else
        t_sErrorMessageForClosureTypeAndCode = ""
    End If
    
    Exit Function
ErrorHandler:
    ConvertError sSource
End Function
Public Function GetErrorEnumForSecurityCheckResult( _
    ByVal t_eSecurityCheckResultStatus As ESecurityCheckResult) _
    As EPROCATINTERNALERROR
'
'Description: Returns the error numbers for the security check return values
'
'   d4      ndk     12Nov1999     Moved here here from missing.bas
'   d11     CRM     14Feb2000     Added error desc for editset with candidate terms moving into pre-release
    On Error GoTo ErrorHandler

    Dim sSource As String
    
    sSource = "EditorialFunctions.GetErrorEnumForSecurityCheckResult"
    
    
    Select Case t_eSecurityCheckResultStatus
        Case ESecurityCheckResult.ESRAFEditorTryingToChangeOtherRole
            GetErrorEnumForSecurityCheckResult = EPROCATINTERNALERROR.PEAFEditorTryingToChangeOtherRole
            
        Case ESecurityCheckResult.ESRNoProofReaderPresentOnEditSet
            GetErrorEnumForSecurityCheckResult = EPROCATINTERNALERROR.PENoProofReaderPresentOnEditSet
            
        Case ESecurityCheckResult.ESRNotAValidForwardTransition
            GetErrorEnumForSecurityCheckResult = EPROCATINTERNALERROR.PENotAValidForwardTransition
            
        Case ESecurityCheckResult.ESRNotAValidReverseTransition
            GetErrorEnumForSecurityCheckResult = EPROCATINTERNALERROR.PENotAValidReverseTransition
            
        Case ESecurityCheckResult.ESROpertionNotValidAtThisStage
            GetErrorEnumForSecurityCheckResult = EPROCATINTERNALERROR.PEOpertionNotValidAtThisStage
            
        Case ESecurityCheckResult.ESROtherUserDoesNotHaveRoleOnEditSet
            GetErrorEnumForSecurityCheckResult = EPROCATINTERNALERROR.PEOtherUserDoesNotHaveRoleOnEditSet
            
        Case ESecurityCheckResult.ESRProofReaderTryingToChangeOtherRole
            GetErrorEnumForSecurityCheckResult = EPROCATINTERNALERROR.PEProofReaderTryingToChangeOtherRole
            
        Case ESecurityCheckResult.ESRRoleDoesNotBelongToAllStagesOtherRoleBelongs
            GetErrorEnumForSecurityCheckResult = EPROCATINTERNALERROR.PERoleDoesNotBelongToAllStagesOtherRoleBelongs
            
        Case ESecurityCheckResult.ESRRoleDoesNotBelongToStage
            GetErrorEnumForSecurityCheckResult = EPROCATINTERNALERROR.PERoleDoesNotBelongToStage
        
        Case ESecurityCheckResult.ESRRoleIsNotHigherOrEqualToOtherRole
            GetErrorEnumForSecurityCheckResult = EPROCATINTERNALERROR.PERoleIsNotHigherOrEqualToOtherRole
        
        Case ESecurityCheckResult.ESRRoleIsNotTheHighestRoleForThisStage
            GetErrorEnumForSecurityCheckResult = EPROCATINTERNALERROR.PERoleIsNotTheHighestRoleForThisStage
        
        Case ESecurityCheckResult.ESRTryingToAddRoleNeverToBeAdded
            GetErrorEnumForSecurityCheckResult = EPROCATINTERNALERROR.PETryingToAddRoleNeverToBeAdded
        
        Case ESecurityCheckResult.ESRTryingToRemoveReplaceOnlyRole
            GetErrorEnumForSecurityCheckResult = EPROCATINTERNALERROR.PETryingToRemoveReplaceOnlyRole
        
        Case ESecurityCheckResult.ESRUserDoesNotHaveRoleOrHisRoleNotPermittedToPerformThisOperation
            GetErrorEnumForSecurityCheckResult = EPROCATINTERNALERROR.PEUserDoesNotHaveRoleOrHisRoleNotPermittedToPerformThisOperation
        
        Case ESecurityCheckResult.ESRUserTryingToChangeOtherUsersCheckBox
            GetErrorEnumForSecurityCheckResult = EPROCATINTERNALERROR.PEUserTryingToChangeOtherUsersCheckBox
            
        Case ESecurityCheckResult.ESRUserBeingAddedAlreadyHasARoleOnEditSet
            GetErrorEnumForSecurityCheckResult = EPROCATINTERNALERROR.PEUserBeingAddedAlreadyHasARoleOnEditSet
            
        '   d11     CRM     14Feb2000
        Case ESecurityCheckResult.ESREditSetHasCandidateTerms
            GetErrorEnumForSecurityCheckResult = EPROCATINTERNALERROR.PECanNotMoveESWithCandidateTermsToPreRelease
    End Select

    Exit Function
ErrorHandler:
    ConvertError sSource
End Function


Public Function GetDescriptionForSecurityCheckError( _
    ByVal t_eSecurityCheckResultError As EPROCATINTERNALERROR) _
    As String
'
'Description: Returns the error description for the security check errors
'
'   d4      ndk     12Nov1999     Moved here here from missing.bas
'   d11     CRM     14Feb2000     Added error desc for editset with candidate terms moving into pre-release
    On Error GoTo ErrorHandler

    Dim sSource As String
    
    sSource = "EditorialFunctions.GetDescriptionForSecurityCheckError"
    
    Select Case t_eSecurityCheckResultError
        Case EPROCATINTERNALERROR.PEAFEditorTryingToChangeOtherRole
            GetDescriptionForSecurityCheckError = " AF Editor is not permitted to modify other roles "
            
        Case EPROCATINTERNALERROR.PENoProofReaderPresentOnEditSet
            GetDescriptionForSecurityCheckError = " Proof Reader should be allocated to perform this operation"
            
        Case EPROCATINTERNALERROR.PENotAValidForwardTransition
            GetDescriptionForSecurityCheckError = " This not a valid forward transition of stage"
            
        Case EPROCATINTERNALERROR.PENotAValidReverseTransition
            GetDescriptionForSecurityCheckError = " This not a valid reverse transition of stage"
            
        Case EPROCATINTERNALERROR.PEOpertionNotValidAtThisStage
            GetDescriptionForSecurityCheckError = " This operation is not valid at this stage of the editset"
            
        Case EPROCATINTERNALERROR.PEOtherUserDoesNotHaveRoleOnEditSet
            GetDescriptionForSecurityCheckError = " You do not have a role on the editset to perform this operation"
            
        Case EPROCATINTERNALERROR.PEProofReaderTryingToChangeOtherRole
            GetDescriptionForSecurityCheckError = " Proof Reader is not permitted to modify other roles "
            
        Case EPROCATINTERNALERROR.PERoleDoesNotBelongToAllStagesOtherRoleBelongs
            GetDescriptionForSecurityCheckError = " Role trying to perform the operation on other role does not belong to all stages the other role belongs"
            
        Case EPROCATINTERNALERROR.PERoleDoesNotBelongToStage
            GetDescriptionForSecurityCheckError = " Your role is not active at this stage to perform this operation"
        
        Case EPROCATINTERNALERROR.PERoleIsNotHigherOrEqualToOtherRole
            GetDescriptionForSecurityCheckError = " You cannot perform this operation trying to perform the operation because your role is less senior than the role you are acting on"
        
        Case EPROCATINTERNALERROR.PERoleIsNotTheHighestRoleForThisStage
            GetDescriptionForSecurityCheckError = " Role trying to perform this operation is not the highest role for the stage"
        
        Case EPROCATINTERNALERROR.PETryingToAddRoleNeverToBeAdded
            GetDescriptionForSecurityCheckError = " User tried to add a role which should not be added"
        
        Case EPROCATINTERNALERROR.PETryingToRemoveReplaceOnlyRole
            GetDescriptionForSecurityCheckError = " User tried to remove a role which can only be replaced"
        
        Case EPROCATINTERNALERROR.PEUserDoesNotHaveRoleOrHisRoleNotPermittedToPerformThisOperation
            GetDescriptionForSecurityCheckError = " User does not have a role on editset or his role is not permitted to perform this operation"
        
        Case EPROCATINTERNALERROR.PEUserTryingToChangeOtherUsersCheckBox
            GetDescriptionForSecurityCheckError = " User not permitted to check/uncheck other users checkbox"
    
        Case EPROCATINTERNALERROR.PEUserBeingAddedAlreadyHasARoleOnEditSet
            GetDescriptionForSecurityCheckError = " User being added already has a role on the editset"
            
        '   d11     CRM     14Feb2000
        Case EPROCATINTERNALERROR.PECanNotMoveESWithCandidateTermsToPreRelease
            GetDescriptionForSecurityCheckError = " Can not move an editset with candidate terms to pre-release stage "
    End Select

    Exit Function
ErrorHandler:
    ConvertError sSource
End Function

Public Function Tag(t_sText, t_sTag) As String
'   converts the passed text into a tagged format:
'   "<t_sTag>t_sText</t_sTag>"
'
'   d1  WDP     18Nov1999     Created

    Dim sSource As String
    sSource = "Tag"
    
    Tag = "<" & t_sTag & ">" & t_sText & "</" & t_sTag & ">"
    
End Function

Public Function FindTaggedData(t_sText, t_sTag) As String
'   searches the passed string for a tagged value
'
'   d6  WDP     18Nov1999     Created

    Dim sSource As String
    sSource = "FindTaggedData"
    Dim lStart As Long, lEnd As Long
    Dim sStartTag As String, sEndTag As String
    sStartTag = "<" & t_sTag & ">"
    sEndTag = "</" & t_sTag & ">"
    lStart = InStr(t_sText, sStartTag)
    lEnd = InStr(t_sText, sEndTag)
    
    If Not lStart = 0 Then
        lStart = lStart + Len(sStartTag)
        FindTaggedData = Mid(t_sText, lStart, lEnd - lStart)
    Else
        FindTaggedData = ""
    End If
    
End Function

Public Function CreateInstanceReference(ByVal t_ObjectName As String, _
                                        ByVal t_sInstanceReference As String) As String
'   Added editorial version of this reader function to
'   enable integration of the search
'   d7  JES     26Nov1999   Created
    CreateInstanceReference = CreateDelimitedString(t_ObjectName, t_sInstanceReference)
                                    
End Function

Public Sub LogReleaseMessage(t_sMsg As String, t_eReleaseMessageType As EReleaseMessageType)
'   Function used to write to the release log for messages and errors
'   relating to the scheduled release process
'
'   d8  WDP     01Dec1999   Created
On Error GoTo ErrorHandler

    Dim sSource As String
    sSource = "LogReleaseMessage"
    
    Dim sLogFile As String
    Dim lFileNum As Long
    Dim sMessage As String
    Dim ofwAppData As New FWApplicationData
    
    'Prepend message with the date, time and area
    sMessage = Date + Time & " : " & GetMsgTypeString(t_eReleaseMessageType) & " : " & t_sMsg
    
    'Generate the file name to be used for Logging
    sLogFile = ofwAppData.GetApplicationConfigurationInfomation(pReleaseLogPath) & Format(Now, "yyyymmdd") & "_Release.log"
    
    'Open the log file.
    lFileNum = FreeFile
    
    Open sLogFile For Append As lFileNum
    
    'Check log file opened.
    If lFileNum > 0 Then
        'Add the message to the log file.
        Print #lFileNum, sMessage
        
        'Close the log file.
        Close (lFileNum)
    End If

    'Print message in Debug window.
    Debug.Print "Log: " & sMessage

    Exit Sub
ErrorHandler:
    ConvertError sSource
                                  
End Sub

Public Function GetMsgTypeString(t_eReleaseMessageType As EReleaseMessageType) As String
'   Returns a string representing enumerated message type
'
'   d8  WDP     01Dec1999   Created
On Error GoTo ErrorHandler
    
    Dim sSource As String
    sSource = "GetMsgTypeString"
    Select Case t_eReleaseMessageType
        Case RELError
            GetMsgTypeString = "Error"
        Case RELInfo
            GetMsgTypeString = "Information"
    End Select
    Exit Function
ErrorHandler:
    ConvertError sSource
                                  
End Function

Public Function GetSortedTitle(ByVal t_sTitle As String) As String
'   Returns a String suitable for use in the lettercode_sorted_title table
'   the rule applied is any chars up to the first capital letter are dropped
'   (ignoring tags)
'
'   d8  WDP     01Dec1999   Created
'   d13 WDP     25Feb2000     Strip all tags in GetSortedTitle
'   d16 TMJ     30Mar2000     Added calls to SpecialCharacterDecode
On Error GoTo ErrorHandler
    
    Dim sSource As String
    sSource = "GetSortedtitle"
    
    Dim lPos As Long, sChar As String
    lPos = 1
    
    Do While lPos <= Len(t_sTitle)
        sChar = Mid(t_sTitle, lPos, 1)
        If sChar = "<" Then
            While Not sChar = ">"
                lPos = lPos + 1
                sChar = Mid(t_sTitle, lPos, 1)
            Wend
        End If
        If sChar >= "A" And sChar <= "Z" Then
            Exit Do
        End If
        lPos = lPos + 1
    Loop
    
    If lPos <= Len(t_sTitle) Then
        GetSortedTitle = RemoveTagsFromString(Mid(t_sTitle, lPos))
        GetSortedTitle = SpecialCharacterDecode(GetSortedTitle)
    Else
        GetSortedTitle = RemoveTagsFromString(t_sTitle)
        GetSortedTitle = SpecialCharacterDecode(GetSortedTitle)
    End If
    
    Exit Function
ErrorHandler:
    ConvertError sSource
                                  
End Function


Public Sub CheckLogFileForErrors(t_sFilename As String, Optional t_bForRelease As Boolean = False, Optional t_bRaiseError As Boolean = False)
'   This method opens the given logfile and looks for error messages
'
'   d9  WDP     20-Jan-1999     Moved from BOCatalogueIndexer
On Error GoTo ErrorHandler

    Dim sSource As String
    sSource = "CheckLogFileForErrors(" & t_sFilename & ")"
    
    Dim lFileNum As String
    Dim sText As String
    
    lFileNum = FreeFile
    Open t_sFilename For Input Access Read As lFileNum
    
    If lFileNum > 0 Then
        sText = StrConv(InputB(LOF(lFileNum), lFileNum), vbUnicode)
        Close #lFileNum
           
        Dim lPos As Long
        lPos = InStr(sText, gk_sVerityErrorText)
        
        If lPos > 0 Then
            If t_bForRelease = True Then
                LogReleaseMessage "Errors occured in indexing see: " & t_sFilename, RELError
            Else
                If t_bRaiseError = True Then
                    '+ need error type for this
                    RaiseError sSource, PEGeneralInternalError, "Errors occured in indexing see: " & t_sFilename
                Else
                    LogMessage "Indexing Errors have occurred see: " & t_sFilename, pComplete
                End If
            End If
            
        End If
    Else
        If t_bForRelease = True Then
            LogReleaseMessage "Unable to open file: " & t_sFilename, RELError
        Else
            RaiseError sSource, PEGeneralInternalError, "Couldnt open log file"
        End If
    End If
    Exit Sub
ErrorHandler:
    ConvertError sSource
    
End Sub

Public Sub LogAction( _
    t_sLogFile As String, _
    Optional t_sBulkFilePath As String = "", _
    Optional t_sCollection As String = "", _
    Optional t_eAction As EIndexingAction = IDXUpdate, _
    Optional t_sLettercode As String = "", _
    Optional t_lEditSetID As Long = -1, _
    Optional t_bDummyAction As Boolean = False)
'   creates a log of all indexing actions to be performed
'   so the indexing process can be triggered off later
'   The parameters give the info to be written to the log
'
'   d5  WDP     18-Nov-1999     Created
'   d8  WDP     18-Nov-1999     Allowed lettercode to be optional
'   d18 WDP     13Apr2000       Altered LogAction to allow dummy (blank) actions
'                               This enables log files with no entries to be created
On Error GoTo ErrorHandler

    Dim sSource As String
    sSource = "LogAction"
    
    Dim lFileNum As Long
    Dim sText As String
    
    sText = Tag(t_sBulkFilePath, gk_sTAGBulkFile) & vbCrLf & _
            Tag(t_sLettercode, gk_sTAGLettercode) & vbCrLf & _
            Tag(t_sCollection, gk_sTAGCollection) & vbCrLf & _
            Tag(GetIndexActionString(t_eAction), gk_sTAGAction) & vbCrLf
    If Not t_lEditSetID = -1 Then
        sText = sText & Tag(CStr(t_lEditSetID), gk_sTAGEditsetID) & vbCrLf
    End If
    
    sText = Tag(sText, gk_sTAGEntry) & vbCrLf
    
    If t_bDummyAction = True Then
        'for dummy (blank) logging just write a newline character
        sText = vbCrLf
    End If
    
    lFileNum = FreeFile
    Open (t_sLogFile) For Append As lFileNum
    
    'Check log file opened.
    If lFileNum > 0 Then
        'Add the message to the log file.
        Print #lFileNum, sText
        
        'Close the log file.
        Close lFileNum
    Else
        '+TO DO create an error type for this
        RaiseError sSource, PEGeneralInternalError, "Couldnt open log file."
    End If
    
    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub

Public Function BackButtonHTML(ByVal t_sLinkbackPage As String) As String
'Creates a back button that mimicks the browser back button
'Implemented using the empty HREF method so the OnClick event works with NN
'd14    jes 10Jan2000   Created
'd30    alm 14Nov2001   Re-wrote to support a back button of multiple pages
    Dim sHTML As String
    If t_sLinkbackPage = "" Then
        sHTML = "<A HREF=""javascript:history.go(-1);"">"
    Else
        sHTML = "<A HREF=""javascript:history.go(" & t_sLinkbackPage & ");"">"
    End If
    sHTML = sHTML & "<IMG BORDER=0 SRC=""" & gk_sIMAGEBackButton & """></A>"
    BackButtonHTML = sHTML
End Function

Public Function GetPitemIDFromParentIDAndKeyOrder(t_eLevel As ECATALOGUELEVEL, _
                        t_lParentID As Long, t_lKeyOrder As Long) As Long
'   Given a piece/item's parentID (class ID for piece, piece id for items) and
'   key order, return the DB ID.
'
'   d14     WDP     02Mar2000     Created
On Error GoTo ErrorHandler
    
    Dim sSource As String
    sSource = "GetPitemIDFromParentIDAndKeyOrder"
    
    Dim sSql As String
    
    sSql = GetPitemIDFromParentIDAndKeyOrderSQL(t_eLevel, t_lParentID, t_lKeyOrder)
    
    Dim ofwHelper As New FWDBHelper
    Dim ofwResult As FWRecordList
    Set ofwResult = ofwHelper.FindBySql(sSql)
    
    If Not ofwResult.RecordCount = 1 Then
        RaiseError sSource, PEGeneralInternalError, _
            "Could not find any/unique ID for pitem - level: " & t_eLevel & ", parentID: " & t_lParentID & ", keyorder: " & t_lKeyOrder
    End If
    
    GetPitemIDFromParentIDAndKeyOrder = ofwResult.GetFieldValue(gk_sDIID)
    
Exit Function
ErrorHandler:
    ConvertError sSource

End Function

Public Function GetPitemIDFromParentIDAndKeyOrderSQL(t_eLevel As ECATALOGUELEVEL, _
                        t_lParentID As Long, t_lKeyOrder As Long) As String
'   Given a piece/item's parentID (class ID for piece, piece id for items) and
'   key order, return the SQL to get the DB ID.
'
'   d15     WDP     02Mar2000     Created
On Error GoTo ErrorHandler
    
    Dim sSource As String
    sSource = "GetPitemIDFromParentIDAndKeyOrderSQL"
    
    Dim sSql As String
    Dim sTable As String, sIDCol As String, sKeyCol As String, sParentIDCol As String
    
    Select Case t_eLevel
        Case ctPiece
            sTable = gk_sDBTABLEPiece
            sIDCol = gk_sDBCOLPieceID
            sKeyCol = gk_sDBColPiecekeyOrder
            sParentIDCol = gk_sDBCOLClassID
        Case ctItem
            sTable = gk_sDBTABLEItem
            sIDCol = gk_sDBCOLItemID
            sKeyCol = gk_sDBCOLItemKeyOrder
            sParentIDCol = gk_sDBCOLPieceID
        Case Else
            RaiseError sSource, PEGeneralInternalError, "Invalid catalogue level"
    End Select
    
    sSql = "SELECT " & sIDCol & " AS " & gk_sDIID & " FROM " & sTable & " WHERE " & _
            sParentIDCol & "=" & t_lParentID & " AND " & _
            sKeyCol & "=" & t_lKeyOrder
    
    
    GetPitemIDFromParentIDAndKeyOrderSQL = sSql
    
Exit Function
ErrorHandler:
    ConvertError sSource

End Function

