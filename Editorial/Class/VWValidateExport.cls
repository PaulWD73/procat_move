VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "VWValidateExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    VWValidateExport.cls
' System:       PRO, PROCAT
' Copyright:    (C) Quidnunc Limited
'
' Description:  This class validates the export form.
'
' Amendment history:
'   Vers    Auth    Date        Comments
'   d1      ADF     4-Jul-2000  Created
'   d2      MAP     22Oct2001   Added code to handle EAD scheduling
'   d3      MAP     23Oct2001   Altered format of summary screen
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit



Implements ObjectControl

Private m_objContext As Object
Private m_aspResponse As Response
Private m_aspRequest As Request
Private m_aspServer As Server
Private m_oFWAppData As FWApplicationData

Private m_lLoggedUserID  As Long
Private m_oBOEditorialPermissions As BOEditorialPermissions

' Member variable to make sure that all the series references are under the lettercode
Private m_sCurrentLetterCode As String

' The ID object that holds all the IDs we need to build the summary
Private m_oboExportIds As BOExportIDs

' Member variables for the filename and summary filtering
Private m_bFileAlreadyExists As Boolean
Private m_sValidFilename As String
Private m_sStrippedFilename As String
Private m_sLowestLevel As String
Private m_eLowestLevel As ECATALOGUELEVEL
Private m_bLowestLevelReached

' The member variables used to filter the summary results
Private m_sIsDepartmentWanted As String
Private m_sIsDivisionWanted As String
Private m_sIsSeriesWanted As String
Private m_sIsSubSeriesWanted As String
Private m_sIsSubSubSeriesWanted As String
Private m_sIsPieceWanted As String
Private m_sIsItemWanted As String

' User stats message
Private m_boMessage As BOMessage

' Amendment history:
'   Vers    Auth    Date        Comments
'   d1      ADF     4-Jul-2000  Created
'
'
' Description:Class initialisation and MTS event handlers
'

Private Sub Initialise()

    Dim sSource As String
    sSource = "VWValidateExport.ValidateForm"

    Set m_objContext = GetObjectContext()
    
    ' Set asp objects
    Set m_aspResponse = m_objContext.Item("Response")
    Set m_aspRequest = m_objContext.Item("Request")
    Set m_aspServer = m_objContext.Item("Server")
    
    m_lLoggedUserID = GetDataFromSession(dsUserID, m_objContext.Item(gk_sSessionObject))
    If m_lLoggedUserID = 0 Then     'Not a valid user
            RaiseError sSource, EPROCATINTERNALERROR.peDBInvalidUser, "Invalid User ID : " & m_lLoggedUserID
    End If

    ' Initialise the security check object
    Set m_oBOEditorialPermissions = New BOEditorialPermissions
    
End Sub

Private Sub ObjectControl_Activate()
    Initialise
End Sub

Private Sub ObjectControl_Deactivate()
    Set m_objContext = Nothing
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = True
End Function

' Amendment history:
'   Vers    Auth    Date        Comments
'   d1      ADF     4-Jul-2000  Created
'
' Description: Start the form vaildation process by checking for null values and parsing
'              the restrictions
'


Public Sub ValidateForm()

On Error GoTo ErrorHandler

    Dim sSource As String
    sSource = "VWValidateExport.ValidateForm"

    ' The strings that will hold the restrictions specified on the export form
    Dim sFileToExportTo As String
    Dim sDeptRestriction As String
    Dim sSeriesRestriction As String
    Dim sMinPieceRestriction As String
    Dim sMaxPieceRestriction As String
    Dim eResult As ESecurityCheckResult

    ' Get the restrictions from the Export form
    sFileToExportTo = m_aspRequest(gk_sFORMFileToExportTo)
    sDeptRestriction = m_aspRequest(gk_sFORMDeptToExport)
    sSeriesRestriction = m_aspRequest(gk_sFORMSeriesToExport)
    sMinPieceRestriction = m_aspRequest(gk_sFORMLowPieceReferenceToExport)
    sMaxPieceRestriction = m_aspRequest(gk_sFORMHighPieceReferenceToExport)
    
    ' Get the check box information determing which sections of data are to be exported
    m_sIsDepartmentWanted = m_aspRequest(gk_sFORMIsDepartmentWanted)
    m_sIsDivisionWanted = m_aspRequest(gk_sFORMIsDivisionWanted)
    m_sIsSeriesWanted = m_aspRequest(gk_sFORMIsSeriesWanted)
    m_sIsSubSeriesWanted = m_aspRequest(gk_sFORMIsSubSeriesWanted)
    m_sIsSubSubSeriesWanted = m_aspRequest(gk_sFORMIsSubSubSeriesWanted)
    m_sIsPieceWanted = m_aspRequest(gk_sFORMIsPieceWanted)
    m_sIsItemWanted = m_aspRequest(gk_sFORMIsItemWanted)
    
    ' first Check is to see if they are entitled to peform this operation
    m_oBOEditorialPermissions.UserID = m_lLoggedUserID
    
    eResult = m_oBOEditorialPermissions.Check(ESecurityCheck.EUserCanExportToEAD)
    If eResult <> ESecurityCheckResult.ESRSecurityCheckSuccessful Then
        RaiseError sSource, _
               GetErrorEnumForSecurityCheckResult(eResult), _
               GetDescriptionForSecurityCheckError( _
                        GetErrorEnumForSecurityCheckResult(eResult))
    End If
    
    ' Perform a check for all required fields
    
    ' Check if the user has selected a catalogue level to export
    If m_sIsDepartmentWanted = "" And m_sIsDivisionWanted = "" And m_sIsSeriesWanted = "" And _
        m_sIsSubSeriesWanted = "" And m_sIsSubSubSeriesWanted = "" And m_sIsPieceWanted = "" And _
        m_sIsItemWanted = "" Then
        
        sSource = sSource & ": No Data Selected For Export."
        
            RaiseError sSource, peUSERNoExportDataSelected, "No data selected for export."
    End If
    
    
    Trim (sFileToExportTo)
    ' Check if the user has specified an export filename
    If (sFileToExportTo = "") Or sFileToExportTo = ".xml" Then
        sSource = sSource & ": No Export File Specified"
        
            RaiseError sSource, peUSERInvalidExportFileName, "You must specify a file to export to."
    End If
    
    'Make sure that the filename is valid
    m_sValidFilename = EncodeFileName(sFileToExportTo)
    
    'Make sure that we have
    If (m_sValidFilename = "") Then
        sSource = sSource & ": No Export File Specified"
        
            RaiseError sSource, peUSERInvalidExportFileName, "You must specify a VALID file to export to."
    End If
    
    
    ' Append an XML file extension if one has'nt been added already
    m_sValidFilename = AppendFileExtension
    
    m_sStrippedFilename = m_sValidFilename
    
    ' Check if it exists already
    If (DoesFileExist) Then
        m_bFileAlreadyExists = True
    End If
    
    
    ' Check if the user has specified a department reference
    If (sDeptRestriction = "") Then
        sSource = sSource & ": No Department Reference Specified."
        
            RaiseError sSource, peUSERMissingExportDeptRef, "You must complete the department reference."
    End If
    
 
    ' We now call the BOExportHandler to parse the restrictions.
    Dim oboExportHandler As New BOExportHandler
    oboExportHandler.Initialise
    
    
    ' Parse the department reference, tell the Handler what level we are parsing
    oboExportHandler.LevelBeingChecked = "department"
    oboExportHandler.ParseRestrictions (sDeptRestriction)
       
    ' Trim and uppercase the lettercode, and then set the Handlers lettercode to this.
    ' This means that all further entries are invalid if they do not fall within this lettercode
    m_sCurrentLetterCode = UCase(Trim(sDeptRestriction))
    oboExportHandler.CurrentLetterCode = m_sCurrentLetterCode
    
    ' Parse the series reference if it is not empty
    If Not (sSeriesRestriction = "") Then
        oboExportHandler.LevelBeingChecked = "series"
        oboExportHandler.ParseRestrictions (sSeriesRestriction)
    End If
    
    ' If there is no series restriction or there is more than one series restrcition
    ' then make sure that no piece reference is specified
    If sSeriesRestriction = "" Or oboExportHandler.NumberOfSeries > 1 Then
        If Not (sMinPieceRestriction = "" Or sMaxPieceRestriction = "") Then
           sSource = sSource & ": Invalid Piece."
            RaiseError sSource, PEUSERInvalidBreadthRestrictions, _
                "You cannot specify a piece if the series range is empty or greater than one"
        End If
    Else
        '  If they are allowed to specify a max and min piece range
        '  make sure they complete both boxes
        
        If (sMinPieceRestriction = "" And Not (sMaxPieceRestriction = "")) Or _
            (sMaxPieceRestriction = "" And Not (sMinPieceRestriction = "")) Then
            ' a full range must be specified
            sSource = sSource & ": Invalid Piece."
            RaiseError sSource, PEUSERInvalidBreadthRestrictions, _
                "You must specify both piece ranges"
        End If
        
        oboExportHandler.LevelBeingChecked = "piece"
        'Parse the minimum and maximum piece key stuff
        
        If Not (sMinPieceRestriction = "") Then
             oboExportHandler.ParseRestrictions (sMinPieceRestriction)
             oboExportHandler.ParseRestrictions (sMaxPieceRestriction)
        End If
        
        If oboExportHandler.PieceKeyMinimum > oboExportHandler.PieceKeyMaximum Then
            RaiseError sSource, PEUSERInvalidBreadthRestrictions, _
                "The piece key order for the minimum piece reference is greater than " & _
                "the piece key order for the maximum piece reference - no pieces will be found"
        End If
        
        'This now contains the ID we need to get all our summary info from the database
        Set m_oboExportIds = oboExportHandler.RetrieveIDs
        
    End If
    
    
    
    'This now contains the ID we need to get all our summary info from the database
    Set m_oboExportIds = oboExportHandler.RetrieveIDs
    ' [ Call the summary procedure from here for dev purposes
    '   eventually this will be called from ASP form ]
    StartSummary

    
Exit Sub
ErrorHandler:
    m_objContext.SetComplete
    ShowError (sSource)

End Sub

' Amendment history:
'   Vers    Auth    Date        Comments
'   d1      ADF     18-Jul-2000  Created
'   d2      MAP     09-Oct-2001 Added scheduling functionality
'   d3      MAP     23-Oct-2001 Minor amendment to screen formating
'
'
' Description: Once the form has been vaidated we can get the summary using
'               the export IDs we got during the parsing
'

Public Sub StartSummary()

On Error GoTo ErrorHandler

    ' The Source in case of an error
    
    Dim sSource As String
    Dim FWHTML As New FWHTML
    Dim sDisplayString As String
    Dim iTotalDepartMents As Long, iTotalDivisions As Long, iTotalSeries As Long
    Dim iTotalSubSeries As Long, iTotalSubSubSeries As Long, iTotalPieces As Long, iTotalItems As Long
    ' d2
    Dim sStartTime As String
    Dim sStartDate As String
    Dim ofwTable As New FWTable
    Dim ofwhtmlreference As New FWHTML
    Dim typCell As CellStyle

    ' Generate a source in case of error
    sSource = "VWValidateExport.StartSummary"
    Set m_oFWAppData = New FWApplicationData
    
    ' Create a new export summary object
    Dim oboExportSummary As New BOExportSummary
    
    
    ' d2 - Check if a valid time and date have been entered
    sStartTime = m_aspRequest(gk_sFORMJobStartTime)
    If sStartTime <> "" Then
        If Not ValidTime(sStartTime) Then
            RaiseError sSource, peUSERInvalidTime, "Scheduled start time entered is invalid."
        End If
    End If
    sStartDate = m_aspRequest(gk_sFORMJobStartDate)
    If sStartDate <> "" Then
        If Not IsDate(sStartDate) Then
            RaiseError sSource, peUSERInvalidDate, "Scheduled date entered is invalid."
        End If
    End If
    
   
    ' Get the Lowest Selected level
    FindLowestLevel
    
    ' And then get the summary count from it
    oboExportSummary.CatalogueIDS = m_oboExportIds
    oboExportSummary.GetSummary m_eLowestLevel
    
    'Put all the counts into local variables so we can filter
    iTotalDepartMents = oboExportSummary.LetterCount
    iTotalDivisions = oboExportSummary.DivisionCount
    iTotalSeries = oboExportSummary.ClassCount
    iTotalSubSeries = oboExportSummary.HeaderCount
    iTotalSubSubSeries = oboExportSummary.SubheaderCount
    iTotalPieces = oboExportSummary.PieceCount
    iTotalItems = oboExportSummary.ItemCount
    
    ' Now filter the result
    Select Case m_sLowestLevel
    
        Case m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfLettercodeLongName)
            iTotalDivisions = 0
            iTotalSeries = 0
            iTotalSubSeries = 0
            iTotalSubSubSeries = 0
            iTotalPieces = 0
            iTotalItems = 0
            
        Case m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfDivisionLongName)
            iTotalSeries = 0
            iTotalSubSeries = 0
            iTotalSubSubSeries = 0
            iTotalPieces = 0
            iTotalItems = 0
            
        Case m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfClassLongName)
            iTotalSubSeries = 0
            iTotalSubSubSeries = 0
            iTotalPieces = 0
            iTotalItems = 0
            
        Case m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfHeaderLongName)
            iTotalSubSubSeries = 0
            iTotalPieces = 0
            iTotalItems = 0
            
        Case m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfSubheaderLongName)
            iTotalPieces = 0
            iTotalItems = 0
            
        Case m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfPieceLongName)
            iTotalItems = 0
            
    End Select
    
    
    ' Now we display the results adding the * to levels that are going to be stubs
    sDisplayString = FWHTML.Bold(" Export Summary" & gk_sHTMLLineBreak & gk_sHTMLLineBreak)
    sDisplayString = sDisplayString & gk_sHTMLLineBreak & " Any thing highlighted with an asterix will be exported as a stub only." & gk_sHTMLLineBreak & gk_sHTMLLineBreak
    
    ' Department
    sDisplayString = AddAsteriskIfStub(sDisplayString, m_sIsDepartmentWanted, m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfLettercodeLongName))
 
    sDisplayString = sDisplayString & iTotalDepartMents & " " & m_oFWAppData.GetDisplayInformation( _
        EEditorialResourceFile.erfDisplayEditSetDetailsDepartments) & ","         'departments
       
    ' Division
    If Not m_bLowestLevelReached Then
        sDisplayString = AddAsteriskIfStub(sDisplayString, m_sIsDivisionWanted, m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfDivisionLongName))
    End If
    
    sDisplayString = sDisplayString & iTotalDivisions & " " & m_oFWAppData.GetDisplayInformation( _
        EEditorialResourceFile.erfDisplayEditSetDetailsDivisions) & ","    'divisions
        
    
    ' Series
    If Not m_bLowestLevelReached Then
        sDisplayString = AddAsteriskIfStub(sDisplayString, m_sIsSeriesWanted, m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfClassLongName))
    End If
    
    sDisplayString = sDisplayString & iTotalSeries & " " & m_oFWAppData.GetDisplayInformation( _
        EEditorialResourceFile.erfDisplayEditSetDetailsSeries) & ","  'series
    
 
    
    ' SubSeries
    If Not m_bLowestLevelReached Then
        sDisplayString = AddAsteriskIfStub(sDisplayString, m_sIsSubSeriesWanted, m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfHeaderLongName))
    End If
    
    sDisplayString = sDisplayString & iTotalSubSeries & " " & " " & m_oFWAppData.GetDisplayInformation( _
        EEditorialResourceFile.erfDisplayEditSetDetailsSubSeries) & "," ' sub series
        
   
    
    ' Subsubseries
    If Not m_bLowestLevelReached Then
        sDisplayString = AddAsteriskIfStub(sDisplayString, m_sIsSubSubSeriesWanted, m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfSubheaderLongName))
    End If
    
    sDisplayString = sDisplayString & iTotalSubSubSeries & "  " & " " & m_oFWAppData.GetDisplayInformation( _
        EEditorialResourceFile.erfDisplayEditSetDetailsSubSubSeries) & ","   '  sub-sub-series
        
   
    
    
    ' Pieces
    If Not m_bLowestLevelReached Then
        sDisplayString = AddAsteriskIfStub(sDisplayString, m_sIsPieceWanted, m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfPieceLongName))
    End If
    
    sDisplayString = sDisplayString & iTotalPieces & "  " & " " & m_oFWAppData.GetDisplayInformation( _
        EEditorialResourceFile.erfDisplayEditSetDetailsPieces) & ","       'pieces
        
   
    
    ' Items
    If Not m_bLowestLevelReached Then
        sDisplayString = AddAsteriskIfStub(sDisplayString, m_sIsItemWanted, m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfItemLongName))
    End If
    
    sDisplayString = sDisplayString & iTotalItems & " " & " " & m_oFWAppData.GetDisplayInformation( _
        EEditorialResourceFile.erfDisplayEditSetDetailsItems)       'items
        
    

    ' Display this to the screen
    sDisplayString = FWHTML.Paragraph(sDisplayString)
    
    If (m_bFileAlreadyExists) Then
        m_sValidFilename = m_sValidFilename & gk_sHTMLLineBreak & _
        FWHTML.Bold("This file already exists, should you choose to continue it will be overwritten.")
    End If
    
    m_sValidFilename = FWHTML.Paragraph("This information will be exported to the file : " & m_sValidFilename)
    
    
    ' Now we create the buttons we want to OK the export and cancel it.
    
    'Hyperlink for the Cancel Button
    Dim cancelLink As String
    Dim sOKLink As String
    ' The URL to point the cancel to
    cancelLink = "javascript:history.back();"
    
    ' Construct the params to be passed if OK is pressed, essentially this is all the form data
    ' Put all the series ids into a string to be passed
    Dim i As Integer
    Dim lSeries As String
    Dim fwSeriesIDs As fwRecordList
    
    Set fwSeriesIDs = m_oboExportIds.SeriesIDS
    Dim iSeriesCount As Long
    
    'Put all the series IDs into a sting to be passed
    If fwSeriesIDs.RecordCount > 0 Then
        fwSeriesIDs.MoveFirst
        lSeries = CStr(fwSeriesIDs.GetFieldValue(gk_sDBCOLSeriesIDs))
        fwSeriesIDs.MoveNext
        While Not fwSeriesIDs.IsAtEnd
            lSeries = lSeries & "," & CStr(fwSeriesIDs.GetFieldValue(gk_sDBCOLSeriesIDs))
            fwSeriesIDs.MoveNext
        Wend
        iSeriesCount = fwSeriesIDs.RecordCount
    Else
        iSeriesCount = 0
    End If

    ' d2 Added Start Time and Date to end of link
    sOKLink = "TriggerExport.asp" & "?" & "IsDeptWanted" & "=" & m_sIsDepartmentWanted _
               & "&" & "IsDivWanted" & "=" & m_sIsDivisionWanted _
               & "&" & "IsSeriesWanted" & "=" & m_sIsSeriesWanted _
               & "&" & "IsSubseriesWanted" & "=" & m_sIsSubSeriesWanted _
               & "&" & "IsSubsubseriesWanted" & "=" & m_sIsSubSubSeriesWanted _
               & "&" & "IsPieceWanted" & "=" & m_sIsPieceWanted _
               & "&" & "IsItemWanted" & "=" & m_sIsItemWanted _
               & "&" & gk_sParamFileName & "=" & m_sStrippedFilename _
               & "&" & "DeptRef" & "=" & m_sCurrentLetterCode _
               & "&" & "LetterCodeDBID" & "=" & m_oboExportIds.LetterCode _
               & "&" & "NumberOfSeries" & "=" & iSeriesCount _
               & "&" & "SeriesIDs" & "=" & lSeries _
               & "&" & "MinPOrder" & "=" & m_oboExportIds.MinimumPieceKeyOrder _
               & "&" & "MaxPOrder" & "=" & m_oboExportIds.MaximumPieceKeyOrder _
               & "&" & "StartTime" & "=" & sStartTime _
               & "&" & "StartDate" & "=" & sStartDate

    
    ' d2
    If sStartTime <> "" Then
        Dim sJobHTML As String
        Set ofwTable = New FWTable
        Set ofwhtmlreference = New FWHTML
        
        typCell = ofwhtmlreference.GetCellStyle(ECellStyle.CELLHelpTextStyle)
        
        ' d3
        Dim sTempText As String
        sTempText = "A Scheduled job will be created for " & sStartTime
        If sStartDate <> "" Then
            sTempText = sTempText & ", " & sStartDate
        End If
        sTempText = sTempText & "."
        
        ofwTable.StartRow
        ofwTable.AddCellAsHTML sTempText, typCell
        ofwTable.EndRow
        
        ofwTable.AddBlankRow
        
        sJobHTML = ofwTable.GetHTML
        
    End If
        
    ' d2 - Check for a scheduled job that would over-write this file... and warn user if exists.
    Dim oboScheduler As BOScheduler
    Dim ofwResults As fwRecordList
    
    Set oboScheduler = New BOScheduler
    
    Set ofwResults = oboScheduler.GetJob(m_sStrippedFilename, jtEADExport)
    
    If ofwResults.RecordCount > 0 Then
        
        typCell = ofwhtmlreference.GetCellStyle(ECellStyle.CELLHelpTextStyle)
        ofwTable.Reset
        
        ofwTable.StartRow
        sDisplayString = "A Scheduled export job already exists for this file name. " & _
                            "Continuing will mean that the file you create now will be over-written when " & _
                            "the scheduled export takes place."
        ofwTable.AddCellAsHTML sDisplayString, typCell
        ofwTable.EndRow
        
        sJobHTML = sJobHTML & ofwTable.GetHTML
    
    End If

    'Image for the remove button
    Dim sText As String
    Dim sText2 As String
    sText = FWHTML.ImageButton(gk_sImagesPath & gk_sImageCancel, "", cancelLink)
    sText2 = FWHTML.ImageButton(gk_sImagesPath & gk_sImageOK, "", sOKLink)
    
    'Output all that to the screen and then provide a cancel button.
    ' d2 Add in possible HTML from scheduled job
    m_aspResponse.Write sDisplayString & m_sValidFilename & sJobHTML & sText2 & sText
    
    
 
Exit Sub
ErrorHandler:
    m_objContext.SetComplete
    ShowError (sSource)
End Sub


Public Function DoesFileExist() As Boolean
' Amendment history:
'   Vers    Auth    Date        Comments
'   d1      ADF     19-Jul-2000  Created
'
'
' Description: Try and open the file to see if it exists
'
   
    Set m_oFWAppData = New FWApplicationData

    Dim sFullPath As String
    sFullPath = m_oFWAppData.GetApplicationConfigurationInfomation( _
                            ECONFIGURATIONDATA.pEADAccessionDirectory) & "\" & m_sValidFilename
        
    On Error Resume Next
    Open sFullPath For Input As #1
    If Err.Number = 53 Then
        DoesFileExist = False
    Else
        DoesFileExist = True
        Close #1
    End If
        
    
End Function

Public Function IsLowestLevel(t_sCheckbox, t_sLevel As String, t_eLevel As ECATALOGUELEVEL) As String
    
    If Not t_sCheckbox = "" Then
        m_sLowestLevel = t_sLevel
        m_eLowestLevel = t_eLevel
    End If
    IsLowestLevel = t_sCheckbox

End Function

Public Function FindLowestLevel()

    Dim dummy As String
    
    dummy = IsLowestLevel(m_sIsDepartmentWanted, m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfLettercodeLongName), ctLetterCode)
    dummy = IsLowestLevel(m_sIsDivisionWanted, m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfDivisionLongName), ctDivision)
    dummy = IsLowestLevel(m_sIsSeriesWanted, m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfClassLongName), ctClass)
    dummy = IsLowestLevel(m_sIsSubSeriesWanted, m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfHeaderLongName), ctHeaders)
    dummy = IsLowestLevel(m_sIsSubSubSeriesWanted, m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfSubheaderLongName), ctSubheaders)
    dummy = IsLowestLevel(m_sIsPieceWanted, m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfPieceLongName), ctPiece)
    dummy = IsLowestLevel(m_sIsItemWanted, m_oFWAppData.GetDisplayInformation(ESharedResourceFile.srfItemLongName), ctItem)
    
End Function

Public Static Function AddAsteriskIfStub(t_sString, t_sCheckbox, t_sLevel) As String

    If t_sCheckbox = "" Then
        AddAsteriskIfStub = t_sString & "*"
    Else
        AddAsteriskIfStub = t_sString
    End If
    
    If t_sLevel = m_sLowestLevel Then
        m_bLowestLevelReached = True
    End If
End Function

Public Function AppendFileExtension()
' if the user has not  specified an xml file extenision then add it
    Dim sFinalFourCharacters As String
    
    If Len(m_sValidFilename) > 4 Then
        sFinalFourCharacters = Mid(m_sValidFilename, (Len(m_sValidFilename) - 3), Len(m_sValidFilename))
        If Not sFinalFourCharacters = ".xml" Then
            AppendFileExtension = m_sValidFilename & ".xml"
        Else
            AppendFileExtension = m_sValidFilename
        End If
    Else
        AppendFileExtension = m_sValidFilename & ".xml"
    End If
    
    
End Function
'   Description:    Check if the array containing the series is empty
'
'   Vers    Auth    Date                Created
'   d1      ADF     18-July- 2000       Created
Public Function IsArrayEmpty(t_iArray As Variant) As Boolean

    Dim lngUBound As Long
    
    On Error Resume Next
    
    lngUBound = UBound(t_iArray)
    If Err.Number <> 0 Then
        IsArrayEmpty = True
    Else
        IsArrayEmpty = False
    End If
        
End Function
