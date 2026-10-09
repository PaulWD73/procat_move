VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "VWExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    VWExport.cls
' System:       PRO, PROCAT
' Copyright:    (C) Quidnunc Limited
'
' Description:  This class is used to import into database from EAD Files
'       This class has methods called from the VB exe, which is
'       Launched by VWStartExport.
'       Not to be called by ASP.
'
'   MTSTransactionMode has to be RequiresTransaction
'
'
' Amendment history:
'   Vers    Auth        Date            Comment
'   d1      A.Farrow    24-Jul-2000     Created
'   d2      A.Farrow    09-Aug-2000     Changed code to stop getting parse catalogue entry for
'                                       the level below that selected.
'   d3      D Johnston  14-Aug-2000     Tidied up the code for release
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Implements ObjectControl
Private m_objContext As ObjectContext

'The DOM object

'The lowest level we want
Private m_iLowestLevel As Integer
Private m_lAlreadyExist() As Long

' The member variables used to filter the summary results
Private m_lLetterCodeID As Long
Private m_sIsDepartmentWanted As String
Private m_sIsDivisionWanted As String
Private m_sIsSeriesWanted As String
Private m_sIsSubSeriesWanted As String
Private m_sIsSubSubSeriesWanted As String
Private m_sIsPieceWanted As String
Private m_sIsItemWanted As String


Private m_lBatchID As Long
Private m_sFileName As String
Private m_sBatchDescription As String
Private m_oFWAppData As FWApplicationData
Private m_oFWRecordList As FWRecordList

Private m_oboExportParser As BOExportParser
Private m_oboExport As BOEADExport
Private m_oBOExportIDs As BOExportIDs
Private m_oboExportTable As BOExportTable
Private m_lObtainedDivIds() As Long

Private m_iCurrentSeriesID As Integer
Private m_iCount As Integer

Private m_oboBatchImport As BOAccessionBatch
Private m_lLoggedUserID  As Long

Private Const m_ksCallMethod As String = "CM:"
Private Const m_ksFileName As String = "FN:"
Private Const m_ksBatchID As String = "BID:"
Private Const m_ksEditSetID As String = "EID:"

Private Const m_ksEADlement As String = "ead"


Public Sub ExportToEADFile(t_lBatchID As Long)
' Description:  This class is used to get all the information out of the export table using the
'               batch id and then start to build up a structure of the children
'
' Amendment history:
'   Vers    Auth        Date            Comment
'   d1      A.Farrow    24-Jul-2000     Created
'   d3      D Johnston  14-Aug-2000     Tidied up the code for release

On Error GoTo Error_Handler
     
    Dim domXMLDoc As DOMDocument
    Dim m_domEAD As IXMLDOMElement
     
    Dim domWrappingTag As IXMLDOMElement
    Dim dscTag As IXMLDOMElement
    Dim domProcess As IXMLDOMProcessingInstruction
    
    Dim objParsedCatalogue As ParsedCatalogueEntry
    Dim sSource As String
    Dim sDirPath As String
    Dim sMessage As String
    Dim sFullPath As String, sTempPath As String
    Dim parsedCat As ParsedCatalogueEntry
    Dim sXmlVersion As String
    Dim sEADHeader As String
    Dim sFrontend As String
    Dim lNoOfSeriesIDs As Long
    Dim i As Long
    
    Dim fwSeriesIDs As FWRecordList
    
    Dim EADTags As IXMLDOMElement
    
    sSource = "VWExport.ExportToEADFile"
    LogMessage sSource, pUserErrors
    
    '  The batch id passed by asyncloader so we can get all information we need back from the
    '   the export tables
    m_lBatchID = t_lBatchID
    m_iCount = 0
        
        
    'We use the export table to get back the information we want about the export
    'This comes back in the form of a member variable record list
    Set m_oFWRecordList = m_oboExportTable.GetExportInformation(t_lBatchID)
        
        
    ' Check integrity of results
    If m_oFWRecordList.RecordCount > 1 Then
        RaiseError sSource, PEGeneralInternalError, _
            "Non unique BatchID_ID found"
    ElseIf m_oFWRecordList.RecordCount < 1 Then
        RaiseError sSource, PEGeneralInternalError, _
            "Could not find entry for the batch ID " & t_lBatchID & " in the export table"
    End If
    
    'Get the Filename back
    m_sFileName = m_oFWRecordList.GetFieldValue(gk_sDBCOLExportFilename)
    
    'Get the ID's back.
    m_oBOExportIDs.LetterCode = m_oFWRecordList.GetFieldValue(gk_sDBCOLLetterCodeID)
    m_oBOExportIDs.MinimumPieceKeyOrder = m_oFWRecordList.GetFieldValue(gk_sDBCOLMinPieceOrder)
    m_oBOExportIDs.MaximumPieceKeyOrder = m_oFWRecordList.GetFieldValue(gk_sDBCOLMaxPieceOrder)
    m_oBOExportIDs.SeriesIDS = m_oboExportTable.GetSeriesIDs(t_lBatchID)
    Set fwSeriesIDs = m_oBOExportIDs.SeriesIDS
    
    'Set the member variable lettercode
    m_lLetterCodeID = m_oFWRecordList.GetFieldValue(gk_sDBCOLLetterCodeID)
        
    ' Store exactly which levels we want to export
    m_sIsDepartmentWanted = m_oFWRecordList.GetFieldValue(gk_sDBCOLDeptWanted)
    m_sIsDivisionWanted = m_oFWRecordList.GetFieldValue(gk_sDBCOLDivWanted)
    m_sIsSeriesWanted = m_oFWRecordList.GetFieldValue(gk_sDBCOLClassWanted)
    m_sIsSubSeriesWanted = m_oFWRecordList.GetFieldValue(gk_sDBCOLHeaderWanted)
    m_sIsSubSubSeriesWanted = m_oFWRecordList.GetFieldValue(gk_sDBCOLSubheaderWanted)
    m_sIsPieceWanted = m_oFWRecordList.GetFieldValue(gk_sDBCOLPieceWanted)
    m_sIsItemWanted = m_oFWRecordList.GetFieldValue(gk_sDBCOLItemWanted)
   
    
    ' Get a DOMDocument helper object
    Set domXMLDoc = New DOMDocument
    
    ' A dummy element that will populated with the node structure
    Set m_domEAD = domXMLDoc.createElement(m_ksEADlement)
    
    ' Get the XML Version and other elements
    Set domProcess = m_oboExportParser.XMLVersionElement
    
    'The tag to wrap all this in
    Set domWrappingTag = domXMLDoc.createElement(m_ksEADlement)
    
    Set EADTags = m_oboExportParser.EADHeaderElement
    domWrappingTag.appendChild EADTags
    
    Set EADTags = m_oboExportParser.FrontMatterElement
    domWrappingTag.appendChild EADTags
    
    ' Get the lettercode to start with
    m_oboExport.Retrieve 1, m_lLetterCodeID, t_lBatchID
    Set parsedCat = m_oboExport.ParsedCatalogueEntry
    parsedCat.ID = m_lLetterCodeID
    
    ' Find the lowest level
    FindLowestLevel
    
    'Create the lettercode level
    Set m_domEAD = m_oboExportParser.CreateExportNodes(parsedCat, 1, IsWanted(1))
    LogLoadInfo m_lBatchID, "Created node for catalogue entry " & GenerateExportReference(GetEntireRefFromDBID(GetCatalogueLevelEnum(1), m_lLetterCodeID), 1)
    
    ' Get the DSC element
    Set dscTag = m_oboExportParser.DSCElement
    
    BuildChildStructure 1, m_lLetterCodeID, dscTag
    
    m_domEAD.appendChild dscTag
    
    ' Add everythiny to the wrapping tag
    domWrappingTag.appendChild m_domEAD
    domWrappingTag.normalize
    
    ' Construct the main XML document
    domXMLDoc.appendChild domProcess
    domXMLDoc.appendChild domWrappingTag
    
    ' Export the EAD to the file
    Set m_oFWAppData = New FWApplicationData
    sFullPath = m_oFWAppData.GetApplicationConfigurationInfomation( _
                            ECONFIGURATIONDATA.pEADAccessionDirectory) & "\" & m_sFileName
    sTempPath = sFullPath & ".tmp"
    
    domXMLDoc.Save (sTempPath)
    ProcessEADFile sTempPath, sFullPath
    
    LogParseMessage "", "Finished Parsing the file " & m_sFileName, pInfo, t_lBatchID
    
    m_objContext.SetComplete
    UpdateEADExportStatusToSuccess m_lBatchID
    
    Exit Sub
    
Error_Handler:
    ' The status will be set to failed from the calling EXE, so we
    ' just need to abort this transaction
    LogMessage "PROCATEdExp::VWExport:ExportToEADFile Failed", pSystemErrors
    LogMessage Err.Description, pSystemErrors
    m_objContext.SetAbort
    ConvertError sSource
End Sub


Private Sub ProcessEADFile(t_sInputFile As String, t_sOutputFile As String)
' File to take the output of the XML API Save function and process it into something
' that resembles decent EAD
'
' Amendment History:
'   Vers    Auth        Date            Comment
'   d3      D Johnston  14-Aug-2000     Created
On Error GoTo ErrorHandler

    Dim lInputFileNum As Long, lOutputFileNum As Long
    Dim sString As String
    Dim sSource As String
    
    sSource = "ProcessEADFile"
    
    ' Open the files
    lInputFileNum = FreeFile
    Open t_sInputFile For Input As lInputFileNum
    lOutputFileNum = FreeFile
    Open t_sOutputFile For Output As lOutputFileNum
    
    ' Check files actually opened
    If lInputFileNum > 0 And lOutputFileNum > 0 Then
        ' First, export the DOCTYPE line to the output file
        Print #lOutputFileNum, m_oboExportParser.DocType
        
        ' Read in the entire file
        sString = Input(LOF(lInputFileNum), lInputFileNum)
        
        If Not sString = "" Then
            ' Replace the '<' characters with < and new lines
            sString = Replace(sString, "<", vbNewLine & "<")
        
            ' But now we have to remove the newlines from the </ end tags
            sString = Replace(sString, vbNewLine & "</", "</")
        End If
        
        ' Now print it out
        Print #lOutputFileNum, sString
    End If

    ' Close the files
    Close (lInputFileNum)
    Close (lOutputFileNum)
    
    ' Delete the temporary file
    Kill t_sInputFile
    
    Exit Sub
    
ErrorHandler:
    Close (lInputFileNum)
    Close (lOutputFileNum)
    ConvertError sSource
End Sub

Public Sub UpdateEADExportStatusToFailed(t_lBatchID As Long, _
            Optional t_bFullInfo As Boolean = True, _
            Optional ByVal t_lErrorNumber As Long = 0, _
            Optional ByVal t_sErrorDescription As String = "")
'Method updates the flag in the database that the importing failed
' Look in the log file for details
'
'   d13  DAJ    08Mar2000   Added more rigorous checking before setting status to success
'   d15  DAJ    28Mar2000   Added more error reporting code

Dim sSource As String
Dim sDescription As String

    sSource = "VWExport.UpdateEADExportStatusToFailed"
        
    m_lBatchID = t_lBatchID
    
    If t_bFullInfo = True Then
        If t_lErrorNumber = 0 Then t_lErrorNumber = Err.Number
        If t_sErrorDescription = "" Then t_sErrorDescription = Err.Description
        
        sDescription = "The following error occurred" & vbNewLine
        sDescription = sDescription & " Error Number      : " & t_lErrorNumber & vbNewLine
        sDescription = sDescription & " Error Description : " & t_sErrorDescription & vbNewLine
        sDescription = sDescription & " Error Source      : " & sSource & vbNewLine
    
        LogMessage sDescription, EProfile.pSystemErrors
        LogLoadInfo m_lBatchID, sDescription, True, , pFatalError
    End If
    
    
    'Abort the transaction and update the import batch as failed
    UpdateStatusFlag m_lBatchID, EImportStatus.isEADExportFailed
    
    ' This function will be called independently from the EXE, so
    ' will have a separate transaction
    m_objContext.SetComplete
    
    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub


Public Sub UpdateEADExportStatusToSuccess(t_lBatchID As Long)
'Method updates the flag in the database that the exporting was succesfull was successful
'
'   d13  DAJ    08Mar2000   Added more rigorous checking before setting status to success
Dim sSource As String
Dim sDescription As String

    
    sSource = "VWExport.UpdateEADExportStatusToSuccess"
    m_lBatchID = t_lBatchID
    
    sDescription = "Export to EAD file " & m_sFileName & " completed." & _
                    vbNewLine
    
    LogLoadInfo m_lBatchID, sDescription, True, , pMandatoryInfo
    LogMessage sDescription, EProfile.pProfileInformation
    
    'Abort the transaction and update the import batch as failed
    
    UpdateStatusFlag m_lBatchID, EImportStatus.isEADExportSuccessful

    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub




'
' Class initialisation and MTS event handlers
'

Private Sub Initialise()
    Dim sSource As String
    sSource = "VWExport.Initialise"
    
    Set m_objContext = GetObjectContext()
    Set m_oboBatchImport = New BOAccessionBatch
    Set m_oboExport = New BOEADExport
    Set m_oboExportParser = New BOExportParser
    Set m_oboExportTable = New BOExportTable
    Set m_oBOExportIDs = New BOExportIDs

    m_oboExportParser.Initialise

End Sub

Private Sub Terminate()
    
    Set m_oFWAppData = Nothing
    Set m_objContext = Nothing
    Set m_oboBatchImport = Nothing
    Set m_oboExport = Nothing
    Set m_oboExportParser = Nothing
    Set m_oboExportTable = Nothing
    Set m_oBOExportIDs = Nothing
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



Public Sub BuildChildStructure(t_iLevel As Integer, t_lID As Long, Optional t_tempXML As IXMLDOMElement)
'
' Description:  This class calls itself to build a recurssive loop which returns the structure
'
' Amendment history:
'   Vers    Auth        Date            Comment
'   d1      A.Farrow    25-Jul-2000     Created

    'Execute the SP to return all the children that we want
    'If the level is 7 then we have reached the lowest depth that we want to reach
    If Not t_iLevel = 7 Then
        Dim ofwResult As FWRecordList
        Dim dummy
        Dim oboParsedCatalogueEntry As ParsedCatalogueEntry
        Dim tempXML As IXMLDOMElement
        
        Dim m_oboExport As New BOEADExport
       
        Set ofwResult = m_oboExportTable.GetChildren(t_lID, t_iLevel, m_lLetterCodeID, m_lBatchID, _
                                                        m_oBOExportIDs.MinimumPieceKeyOrder, _
                                                        m_oBOExportIDs.MaximumPieceKeyOrder)
    
        If ofwResult.RecordCount > 0 Then
            ofwResult.MoveFirst
            
           'We call this function on itself
            While Not ofwResult.IsAtEnd
                'We get the values of the next child to assess
                Dim iLevelNo As Integer
                Dim lID As Long
                
                iLevelNo = ofwResult.GetFieldValue(gk_sDBCOLCatalogueLevel)
                lID = ofwResult.GetFieldValue(gk_sDBCOLCatalogueID)
                If iLevelNo <= m_iLowestLevel Then
                    ' We not generate a parsed catalogue entry
                    dummy = m_oboExport.Retrieve(iLevelNo, lID, m_lBatchID)
                
                    Set oboParsedCatalogueEntry = m_oboExport.ParsedCatalogueEntry
                
                    ' make sure that the id of the catalogue is set
                    oboParsedCatalogueEntry.ID = lID
                    m_oboExportParser.Initialise
                
                    Set tempXML = m_oboExportParser.CreateExportNodes(oboParsedCatalogueEntry, iLevelNo, IsWanted(iLevelNo))
                    t_tempXML.appendChild tempXML
                    LogLoadInfo m_lBatchID, "Created node for catalogue entry " & GenerateExportReference(GetEntireRefFromDBID(GetCatalogueLevelEnum(iLevelNo), lID), iLevelNo)
                
                    If iLevelNo < m_iLowestLevel Then
                        ' Go on and do it's children
                        BuildChildStructure iLevelNo, lID, tempXML
                    End If
                End If
            
                ' Move on to the next entry in this particular list of children
                ofwResult.MoveNext
                 
            Wend
        End If
    End If

End Sub

Public Function IsAlreadyBuilt(t_lID As Long) As Boolean
'   Description:    Check if the current division has already been constructed
'
'   Vers    Auth    Date                Created
'   d1      ADF     18-July- 2000       Created
    
        Dim i As Integer
        If Not IsArrayEmpty(m_lObtainedDivIds) Then
            For i = 0 To UBound(m_lObtainedDivIds) - 1
                If m_lObtainedDivIds(i) = t_lID Then
                    IsAlreadyBuilt = True
                Else
                    IsAlreadyBuilt = False
                End If
            Next
        Else
            IsAlreadyBuilt = False
        End If
            
End Function

Public Function IsArrayEmpty(t_iArray As Variant) As Boolean
'   Description:    Check if the array containing the series is empty
'
'   Vers    Auth    Date                Created
'   d1      ADF     18-July- 2000       Created
    Dim lngUBound As Long
    
    On Error Resume Next
    
    lngUBound = UBound(t_iArray)
    If Err.Number <> 0 Then
        IsArrayEmpty = True
    Else
        IsArrayEmpty = False
    End If
        
End Function

Private Function AddDivision(t_lID As Long)
'   Description:    We need to make sure that when we have multiple restrictions that we do not
'                   add a division structure twice
'
'   Vers    Auth            Date                Created
'   d1      A.Farrow        25-July- 2000       Created
'
    ReDim Preserve m_lObtainedDivIds(m_iCount + 1)
    m_lObtainedDivIds(m_iCount) = t_lID
    m_iCount = m_iCount + 1
               
    End Function

Private Function IsWanted(t_iLevel As Integer) As Boolean
'   Description:    Decides whether or not to export just a stub
'
'   Vers    Auth            Date                Created
'   d1      A.Farrow        25-July- 2000       Created

Select Case t_iLevel

    Case 1
        If m_sIsDepartmentWanted = "False" Or t_iLevel > m_iLowestLevel Then
            IsWanted = False
        Else
            IsWanted = True
        End If
        
    Case 2
        If m_sIsDivisionWanted = "False" Or t_iLevel > m_iLowestLevel Then
            IsWanted = False
        Else
            IsWanted = True
        End If
        
    Case 3
        If m_sIsSeriesWanted = "False" Or t_iLevel > m_iLowestLevel Then
            IsWanted = False
        Else
            IsWanted = True
        End If
        
    Case 4
        If m_sIsSubSeriesWanted = "False" Or t_iLevel > m_iLowestLevel Then
            IsWanted = False
        Else
            IsWanted = True
        End If
    
    Case 5
        If m_sIsSubSubSeriesWanted = "False" Or t_iLevel > m_iLowestLevel Then
            IsWanted = False
        Else
            IsWanted = True
        End If
        
    Case 6
        If m_sIsPieceWanted = "False" Or t_iLevel > m_iLowestLevel Then
            IsWanted = False
        Else
            IsWanted = True
        End If
        
    Case 7
        If m_sIsItemWanted = "False" Or t_iLevel > m_iLowestLevel Then
            IsWanted = False
        Else
            IsWanted = True
        End If
            
    End Select
    
End Function
Public Function IsLowestLevel(t_sCheckbox, t_iLevel As Integer)
'   Description:   here we get the lowest level and don't bother exporting past that
'
'   Vers    Auth            Date                Created
'   d1      A.Farrow        25-July- 2000       Created

    If t_sCheckbox = "True" Then
        m_iLowestLevel = t_iLevel
    End If

End Function

Public Function FindLowestLevel()

    Dim dummy As String
    
    dummy = IsLowestLevel(m_sIsDepartmentWanted, 1)
    dummy = IsLowestLevel(m_sIsDivisionWanted, 2)
    dummy = IsLowestLevel(m_sIsSeriesWanted, 3)
    dummy = IsLowestLevel(m_sIsSubSeriesWanted, 4)
    dummy = IsLowestLevel(m_sIsSubSubSeriesWanted, 5)
    dummy = IsLowestLevel(m_sIsPieceWanted, 6)
    dummy = IsLowestLevel(m_sIsItemWanted, 7)
    
End Function

Private Sub UpdateEADExportStatusToStarted(t_sFilename As String, t_lBatchID As Long)
'   Description:   Let the batch table know that we have started export
'
'   Vers    Auth            Date                Created
'   d1      A.Farrow        08-Aug-2000         Created from UpdateEADImportStatusToStarted

On Error GoTo ErrorHandler

    Dim sSource As String, sMessage As String
    sSource = "VWLoad.UpdateEADImportStatusToStarted"
    
    
    m_lBatchID = t_lBatchID
    
        
    sMessage = "Starting exporting to the EAD file " & m_sFileName & "."
    LogLoadInfo m_lBatchID, sMessage
    
    LogMessage sMessage, EProfile.pProfileInformation
    
    UpdateStatusFlag m_lBatchID, EImportStatus.isEADExportStarted
    
    Exit Sub
ErrorHandler:
    ShowError sSource
End Sub
Public Sub UpdateStatusFlag(t_lBatchID As Long, t_eImportStatus As EImportStatus)
'   Description:   Update the status flag so the batch table knows what is going on
'
'   Vers    Auth            Date                Created
'   d1      A.Farrow        08-Aug-2000         Created
Dim sSource As String
Dim sDescription As String

    
    sSource = "VWExport.UpdateStatusFlag"
    
    m_oboBatchImport.Load m_lBatchID
    m_oboBatchImport.ImportStatus = t_eImportStatus
    m_oboBatchImport.UPDATE

    m_objContext.SetComplete

    Exit Sub
ErrorHandler:
    
    m_objContext.SetAbort
    
    ConvertError sSource
End Sub


Private Function GenerateExportReference(t_cCatRef As CatReferenceParts, t_iLevel As Integer) As String
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
        sRef = sRef & "-" & "d" & t_cCatRef.sDivisionRef
    End If
    If Not t_cCatRef.sClassRef = "" Then
        sRef = sRef & "-" & "c" & t_cCatRef.sClassRef
    End If
    If Not t_cCatRef.sSubclassRef = "" Then
        sRef = sRef & "-" & t_cCatRef.sSubclassRef
    End If
    If Not t_cCatRef.sHeaderRef = "" Then
        sRef = sRef & "-" & "H" & t_cCatRef.sHeaderRef
    End If
    If Not t_cCatRef.sSubheaderRef = "" Then
        sRef = sRef & "-" & "h" & t_cCatRef.sSubheaderRef
    End If
    If Not t_cCatRef.sPieceRef = "" Then
        sRef = sRef & "-" & "p" & t_cCatRef.sPieceRef
    End If
    If t_iLevel = 7 Then
        sRef = sRef & "-" & "i" & t_cCatRef.sItemRef
    End If

    
 
    
    GenerateExportReference = sRef

   

End Function

Private Sub OrderSeriesArray(ByRef lSeriesIDs As Variant)
    Dim i As Long
    Dim j As Long
    Dim tmp As Long
    Dim min As Long
    Dim max As Long
    Dim last_swap As Long
    
    max = UBound(lSeriesIDs) - 1
    min = 0
    Do While min < max
        last_swap = min - 1
        i = min + 1
        
        Do While i <= max
            If lSeriesIDs(i - 1) > lSeriesIDs(i) Then
                tmp = lSeriesIDs(i - 1)
                j = i
                Do
                    lSeriesIDs(j - 1) = lSeriesIDs(j)
                    j = j + 1
                    If j > max Then Exit Do
                Loop While lSeriesIDs(j) < tmp
                lSeriesIDs(j - 1) = tmp
                last_swap = j - 1
                i = j + 1
            Else
                i = i + 1
            End If
        Loop
        
        max = last_swap
        i = max - 1
        
        Do While i >= min
            If lSeriesIDs(i + 1) < lSeriesIDs(i) Then
                tmp = lSeriesIDs(i + 1)
                j = i
                Do
                    lSeriesIDs(j + 1) = lSeriesIDs(j)
                    j = j - 1
                    If j < min Then Exit Do
                Loop While lSeriesIDs(j) > tmp
                last_swap = j + 1
                i = j - 1
            Else
                i = i - 1
            End If
        Loop
        min = last_swap + 1
    Loop

End Sub
