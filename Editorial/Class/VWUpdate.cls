VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "VWUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    VWUpdate.cls
' System:       PRO, PROCAT
' Copyright:    (C) Quidnunc Limited
'
' Description:  This class is used to update the catalogue from EAD Files
'       This class has methods called from the VB exe, which is launched by VWStartImport.
'       ** Not to be called by ASP! **
'
'   MTSTransactionMode has to be RequiresTransaction
'   This is because the class is instantiated from a non ASP component (the EXE)
'   For the class to get hold of the object context when calling CreateObject and not using
'   server.createobject for instantiation, the class needs RequiresTransaction mode.
'
'
' Amendment history:
'   Version Name            Date            Comments
'   d1      D Johnston      19-Jul-2000     Created from VWImport
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Implements ObjectControl
Private m_objContext As ObjectContext

' Parameter variables
Private m_lBatchID As Long
Private m_lEditsetID As Long
Private m_sFileName As String
Private m_sBatchDescription As String
Private m_lLoggedUserID  As Long

' Business and helper objects
Private m_objEADParser As XMLParser
Private m_oboEADUpdate As BOEADUpdate
Private m_oboBatchImport As BOAccessionBatch
Private m_oFWAppData As FWApplicationData
Private m_objBOBasicCatEntry As BOBasicCatalogueEntry
Private m_oboLoad As BOLoad

' Other variables
Private m_iEntriesLoaded As Integer

Public Sub UpdateFromEADFile(t_sFilename As String, t_lBatchID As Long, t_lEditSetID As Long)
' Main method called from EXE to start the update
'
' Amendment history:
'   Version Name            Date            Comments
'   d1      D Johnston      19-Jul-2000     Created from VWImport

On Error GoTo Error_Handler
    
    Dim sSource As String
    Dim sDirPath As String
    Dim sMessage As String
    sSource = "PROCATEdLoader::VWUpdate.UpdateFromEADFile"
    
    LogMessage sSource, pUserErrors
    
    Dim objParsedCatalogue As ParsedCatalogueEntry
    Dim lCatalogueID As Long, eCatLevel As ECATALOGUELEVEL
    
    ' Set the number of records loaded
    m_iEntriesLoaded = 0
    
    ' Get the filename, batch and edit set IDs
    m_sFileName = t_sFilename
    m_lBatchID = t_lBatchID
    m_lEditsetID = t_lEditSetID
        
    ' Log the fact we've started
    sMessage = "Starting update from the EAD file " & m_sFileName & "."
    LogLoadInfo m_lBatchID, sMessage, , , pMandatoryInfo
    
    ' Work out the EAD directory where the file is
    LogMessage sSource & ":DirPath", pUserErrors
    sDirPath = m_oFWAppData.GetApplicationConfigurationInfomation(ECONFIGURATIONDATA.pEADAccessionDirectory)
    sDirPath = IIf(Left(sDirPath, 1) = "\", sDirPath, sDirPath & "\")
    
    'Start the Parser
    LogMessage sSource & " - About to call XMLParser.Initialise", pUserErrors
    m_objEADParser.Initialise
    
    ' Add the batch ID to the end to make the log file obvious
    If Not m_objEADParser.Parse(sDirPath, m_sFileName, , m_lBatchID) Then
        
        LogLoadInfo m_lBatchID, _
            "Could not parse the EAD file " & m_sFileName & ".", , , pFatalError
        
        RaiseError sSource, EPROCATINTERNALERROR.PEImport, _
            "Could not parse the EAD file " & m_sFileName
    End If
    
    LogMessage sSource & ":Before Loop", pUserErrors
    
    'Loop adding the catalogues into the live tables
    Do While Not m_objEADParser.IsEndOfList
        Dim bAddOK As Boolean
        
        Set objParsedCatalogue = m_objEADParser.GetCatalogueEntry(True)
        
        If objParsedCatalogue.Failed = False Then
            ' Get the entry into the specified edit set
            eCatLevel = objParsedCatalogue.Level
            
            If objParsedCatalogue.ParentEntry = True Then
                LogLoadInfo m_lBatchID, "Skipping over parent entry at level " & CStr(eCatLevel) & "...", , , pInfo
            Else
                ' Get the catalogue ID to add to the edit set
                lCatalogueID = GetDBIDFromEADRef(objParsedCatalogue.ID, eCatLevel)
                
                ' Check for a valid ID
                If lCatalogueID = 0 Then
                    RaiseError sSource, PECatalogue, "Entry '" & objParsedCatalogue.ID & "' does not exist in the database; aborting batch ..."
                End If
                
                m_objBOBasicCatEntry.AddSingleCatalogueEntryIntoEditSet m_lEditsetID, lCatalogueID, eCatLevel
                
                ' Now call the update method to overwrite that entry
                Set m_oboEADUpdate.ParsedCatalogueEntry = objParsedCatalogue
                m_oboEADUpdate.BatchID = m_lBatchID
                m_oboEADUpdate.EditsetID = m_lEditsetID
                m_oboEADUpdate.CatalogueID = lCatalogueID
                bAddOK = m_oboEADUpdate.Update
                
                ' Increment the number of entries loaded
                If bAddOK = True Then
                    m_iEntriesLoaded = m_iEntriesLoaded + 1
                Else
                    sMessage = "Failed to update catalogue entry '" & objParsedCatalogue.ID & "'; entry in edit set is unchanged."
                    LogLoadInfo m_lBatchID, sMessage, , , pNonFatalError
                End If
            End If
        Else
            ' Log failure to parse a catalogue entry
            sMessage = "Failed to parse catalogue entry '" & objParsedCatalogue.ID & "' (see parse log)."
            LogLoadInfo m_lBatchID, sMessage, , , pNonFatalError
            
            ' Added checking for invalid lettercode
            If objParsedCatalogue.Level = ctLetterCode Then
                sMessage = "Lettercode entry is invalid; cannot carry on with batch"
                LogLoadInfo m_lBatchID, sMessage, True, , pFatalError
                LogMessage sMessage, EProfile.pProfileInformation
                UpdateEADUpdateStatusToFailed m_lBatchID, False
                
                ' Simply exiting the loop should be OK; the lettercode must be the first entry
                ' in the file for hierarchecal integrity
                Exit Do
            End If
        End If
    Loop
    
    ' Import the class entries as editable
    m_oboLoad.BatchID = m_lBatchID
    m_oboLoad.EditsetID = m_lEditsetID
    m_oboLoad.MakeLoadedClassesEditable True

    ' If the batch did not load any entries, then fail it
    If m_iEntriesLoaded = 0 Then
        RaiseError sSource, PEGeneralInternalError, "Did not update any entries; batch is a failure"
    Else
        ' The transaction is OK so we can set the success status
        UpdateEADUpdateStatusToSuccess m_lBatchID
    End If

    m_objContext.SetComplete

    Exit Sub
    
Error_Handler:
    ' The status will be set to failed from the calling EXE, so we
    ' just need to abort this transaction
    m_objContext.SetAbort
    sMessage = Err.Description
    LogMessage sSource & " failed - " & sMessage, pUserErrors
    LogLoadInfo m_lBatchID, sMessage, True, , pFatalError
    ConvertError sSource
End Sub

Public Sub UpdateEADUpdateStatusToFailed(t_lBatchID As Long, _
            Optional t_bFullInfo As Boolean = True, _
            Optional ByVal t_lErrorNumber As Long = 0, _
            Optional ByVal t_sErrorDescription As String = "")
' Method updates the flag in the database that the update failed
'
' Amendment history:
'   Version Name            Date            Comments
'   d1      D Johnston      19-Jul-2000     Created from VWImport
On Error GoTo ErrorHandler

    Dim sSource As String
    Dim sDescription As String

    sSource = "VWUpdate.UpdateEADUpdateStatusToFailed"
        
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
    Else
        LogMessage t_sErrorDescription, EProfile.pSystemErrors
        LogLoadInfo m_lBatchID, t_sErrorDescription, False, , pFatalError
    End If
    
    UpdateStatusFlag m_lBatchID, EImportStatus.isEADUpdateFailed
    
    ' This function may be called independently from the EXE, so
    ' will have a separate transaction
    m_objContext.SetComplete
    
    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub


Public Sub UpdateEADUpdateStatusToSuccess(t_lBatchID As Long)
'Method updates the flag in the database that the update was successful
'
' Amendment history:
'   Version Name            Date            Comments
'   d1      D Johnston      19-Jul-2000     Created from VWImport
On Error GoTo ErrorHandler
    
    Dim sSource As String
    Dim sDescription As String
    
    sSource = "VWUpdate.UpdateEADUpdateStatusToSuccess"
    m_lBatchID = t_lBatchID
    
    sDescription = "Update from the EAD file " & m_sFileName & " completed." & _
                    vbNewLine
    
    LogLoadInfo m_lBatchID, sDescription, True, , pMandatoryInfo
    LogMessage sDescription, EProfile.pProfileInformation
    
    UpdateStatusFlag m_lBatchID, EImportStatus.isEADUpdateSuccessful

    ' This function may be called independently from the EXE, so
    ' will have a separate transaction
    m_objContext.SetComplete
    
    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub

Private Sub UpdateEADUpdateStatusToStarted(t_sFilename As String, t_lBatchID As Long)
' This method is used to set the load status and the editset Id for the batch
'   record in accession. It also sets the member variables from the parameter
'
' Amendment history:
'   Version Name            Date            Comments
'   d1      D Johnston      19-Jul-2000     Created from VWImport
On Error GoTo ErrorHandler

    Dim sSource As String, sMessage As String
    sSource = "VWLoad.UpdateEADUpdateStatusToStarted"
    
    m_lBatchID = t_lBatchID
    
    sMessage = "Starting update from the EAD file " & t_sFilename & "."
    LogLoadInfo m_lBatchID, sMessage
    
    LogMessage sMessage, EProfile.pProfileInformation
    
    UpdateStatusFlag m_lBatchID, EImportStatus.isEADUpdateStarted
    
    ' This function may be called independently from the EXE, so
    ' will have a separate transaction
    m_objContext.SetComplete
    
    Exit Sub
ErrorHandler:
    ShowError sSource
End Sub



Public Sub UpdateStatusFlag(t_lBatchID As Long, t_eImportStatus As EImportStatus)
' Method updates the flag in the database for the specified batch
'
' Amendment history:
'   Version Name            Date            Comments
'   d1      D Johnston      19-Jul-2000     Created from VWImport
On Error GoTo ErrorHandler

    Dim sSource As String
    Dim sDescription As String

    sSource = "VWUpdate.UpdateStatusFlag"
    
    m_oboBatchImport.Load m_lBatchID
    m_oboBatchImport.ImportStatus = t_eImportStatus
    m_oboBatchImport.EditsetID = m_lEditsetID
    m_oboBatchImport.Update

    m_objContext.SetComplete

    Exit Sub
ErrorHandler:
    
    m_objContext.SetAbort
    
    ConvertError sSource
End Sub



''''''''''''''''' Implementation ''''''''''''''''''''''

'
' Private functions
'


'
' Class initialisation and MTS event handlers
'

Private Sub Initialise()
    Dim sSource As String
    sSource = "VWUpdate.Initialise"
    
    Set m_objContext = GetObjectContext()
    
    Set m_oFWAppData = New FWApplicationData
    Set m_objEADParser = New XMLParser
    Set m_oboEADUpdate = New BOEADUpdate
    Set m_oboBatchImport = New BOAccessionBatch
    Set m_objBOBasicCatEntry = New BOBasicCatalogueEntry
    Set m_oboLoad = New BOLoad
    
End Sub

Private Sub Terminate()
    
    Set m_oFWAppData = Nothing
    Set m_objEADParser = Nothing
    Set m_objContext = Nothing
    Set m_oboEADUpdate = Nothing
    Set m_oboBatchImport = Nothing
    Set m_oboLoad = Nothing
    
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





