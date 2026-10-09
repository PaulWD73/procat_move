VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "VWCatalogueBulkUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    VWCatalogueBulkUpdate.cls
' System:       PROCAT, PRO
' Copyright:    (C) Quidnunc Limited
'
' Description:
'
' View object to perform the bulk update operation
'
' Amendment history:
'
' d1    ksk     07Ju11999       Created.
' d2    ksk     08Jul1999       Added an extra enumearion for Invalid EditsetID and
'                               Changed the raise error to use the new enumeration
' d3   ksk      09Jul1999       Added tooltip help for all hyper links
' d4   ksk      20Jul1999       Changed code to use the lookup methods
' d5   ksk      09Aug1999       Changed the Constant name for Closure status
' d6   ksk      13Aug1999       Call for look up list changed as the bo changed
' d7   ksk      17Aug1999       Change to display an intermediate screen with a list
'                               or catalogue that were bulk updated.
' d8   ksk      27Oct1999       Adding closure status and closure type as additional
'                               fields. And modularised the long function.
' d9   ndk      17Nov1999       Renamed ESFAFilterNone as ESFAFilterElectronicAndPrinted
' d10  ksk      18Oct1999       Changed Closure type to reflect the value in lookups
' d11  ksk      07Dec1999       Changes to display the EAD browser for select fields
' d12  ksk      14Dec1999       Added code to convert the apostrophe and quote in the ead field
' d13  wdp      24Feb2000       Transiently store the levels affected too.
' d14  tmj      13Mar2000       Changed the position of the textboxes so they are after the
'                               descriptions when an index term is added, and modified
'                               DisplayOneLineEditForContent to make it more flexible
' d15  tmj      15Mar2000       Modified so can Bulk add leaflets to Publication Notes and
'                               Unpublished Finding Aids
' d16  tmj      22Mar2000       Added code to display linked AF terms
' d17  tmj      25Mar2000       Moved checking if in collection so it works
' d18  tmj      30Mar2000       Changed so XML is rendered properly
' d19  daj      05May2000       Modified the EADEditor control code so that it works with the
'                               new non-event driven control (see VWCatalogueEntry for more details)
' d20  daj      20Jun2000       Added Admin History DTD type to the bulk update control
' d21  alm      13Feb2001       Generates a start / end date from the convering date during a bulk update
' d22  gdb      15Feb2001       Added extra header on headerHTML function
' d23  gdb      16Feb2001       added message to displat for closure type and code
' d24  gdb      16Feb2001       added warning message to DisplayBulkupdateRegularLookupFields
' d25  gdb      16Feb2001       added warning message to DisplayBulkupdateRegularFields
' d26  alm      03Oct2001       Added Accessioned As Copy, Orderable, LCI, Extract At Transfer,
'                               Creator Names, Place Of Deposit and Location Of Originals at
'                               fields that can be bulk updated
' d27  fjd      05Feb2003       Renamed LCI flag to AR
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'TO DO
'EAD editing for Title, Scope and Content and arrangement

Private Const mk_iLengthOfNumericField As Integer = 10
Private Const mk_iLengthOfTextField As Integer = 80
Private Const mk_iLengthOfUnknownDescriptionTextField As Integer = 30
Private Const mk_iNumberOfRows As Integer = 5

Private Const mk_sNameButtonOk As String = "BOK"
Private Const mk_sNameButtonCancel As String = "BCA"
Private Const mk_sNameBulkUpdateField As String = "BUF"
Private Const mk_sNameBulkUpdateValue As String = "BUV"
Private Const mk_sNameFieldList As String = "FLST"
Private Const mk_sNameLevel As String = "BUFLN"
Private Const mk_sNameValue1 As String = "BUV1"
Private Const mk_sNameValue2 As String = "BUV2"


Private Const mk_sNameHiddenIDFieldFirstLink As String = "FID"
Private Const mk_sNameHiddenIDFieldSecondLink As String = "SID"

Private Const mk_sNameDescriptionOptionalFirstLink As String = "FOD"
Private Const mk_sNameDescriptionOptionalSecondLink As String = "SOD"

Private Const mk_sNameRemoveFirstLinkButton As String = "RMVFL"
Private Const mk_sNameRemoveSecondLinkButton As String = "RMVSL"


Private Const mk_sAddCorpBodyTermButton As String = "ACB"
Private Const mk_sAddPersonTermButton As String = "APR"
Private Const mk_sAddSFAButton As String = "ASFA"
Private Const mk_sAddLeafletButton As String = "ALFL"

Private Const mk_sParamRestoreState As String = "PARAMREST"
Private Const mk_sStateRestoreValue As String = "1"

Private Const mk_sNameEditButton As String = "EADEDT"
Private Const mk_sNameUpdateButton As String = "EADUPD"

Private Const ksNameDTDType As String = "DTDType"

'The Instance ACTIVEX control Name
Private Const mk_sEADControlInstanceName As String = "BulkUpdateEADInst"
'The Hidden field that holds a flag. Used to identify which button was clicked
Private Const mk_sNameHiddenStatusField As String = "HiddenStatus"

'Private framework styles used across all methods
Private m_typCellStyleLeftHeading As CellStyle, m_typCellStyleRightDetails As CellStyle
Private m_typCellStyleForOKCancel As CellStyle
Private m_typFontStyleLeftHeading As FontStyle, m_typFontStyleRightDetails As FontStyle


Private m_oboEditorialPermission As BOEditorialPermissions
Private m_oboCatalogueBulkUpdate As BOCatalogueBulkUpdate
Private m_oBOCatalogueElementMatrix As BOCatalogueElementMatrix
Private m_oVWSharedCatalogueEntry As VWSharedCatalogueEntry
Private m_oFWAppData As FWApplicationData
Private m_ofwRefData  As FWReferenceData
Private m_ofwhtmlReference As FWHTML
Private m_lLoggedUserID As Long
Private m_eSelectedFieldID As ECATALOGUEPROPERTIES


Private m_oFWTable As FWTable

Private m_lEditsetID As Long
Private m_bSelectedLevels(gk_iNumberOfLevels) As Boolean
Private m_vUpdateValue1 As Variant
Private m_vUpdateValue2 As Variant
Private m_vUpdateValue3 As Variant 'd21
Private m_bBothTheCollectionsRequired As Boolean
Private m_colSecondSet As Collection
Private m_colFirstSet As Collection

Private m_objContext As Object
Private m_aspResponse As Response
Private m_aspRequest As Request
Private m_aspServer As Server
Private m_aspSession As Session



'Implementing the MTS object control inteface
Implements ObjectControl



''''''''''''''''' Implementation ''''''''''''''''''''''

'
' Private functions
'


'
' Class initialisation and MTS event handlers
'

Private Sub Initialise()
    Dim sSource As String
    sSource = "VWCatalogueBulkUpdate.Initialise"
    Set m_objContext = GetObjectContext()
    
    ' Set asp objects
    Set m_aspResponse = m_objContext.Item("Response")
    Set m_aspRequest = m_objContext.Item("Request")
    Set m_aspServer = m_objContext.Item("Server")
    Set m_aspSession = m_objContext.Item("Session")
    
    ' Specific initialisation
    Set m_oboCatalogueBulkUpdate = New BOCatalogueBulkUpdate
    Set m_oBOCatalogueElementMatrix = New BOCatalogueElementMatrix
    Set m_ofwhtmlReference = New FWHTML
    Set m_oFWAppData = New FWApplicationData
    Set m_ofwRefData = New FWReferenceData
    Set m_oboEditorialPermission = New BOEditorialPermissions
    Set m_oVWSharedCatalogueEntry = New VWSharedCatalogueEntry
    Set m_oFWTable = New FWTable
    
    Set m_colSecondSet = New Collection
    Set m_colFirstSet = New Collection

    m_lLoggedUserID = GetDataFromSession(dsUserID, m_objContext.Item(gk_sSessionObject))
    If m_lLoggedUserID = 0 Then     'Not a valid user
        RaiseError sSource, EPROCATINTERNALERROR.peDBInvalidUser, "Invalid User"
    End If
    Dim i As Integer
    'default to no levels selected
    For i = 1 To 7
        m_bSelectedLevels(i) = False
    Next
    
    Call InitialiseStyles
End Sub

Private Sub InitialiseStyles()
'
'Description
'
'Initialises the styles
'
' d8   ksk      27Oct1999       Created
'
On Error GoTo ErrorHandler

    Dim sSource As String
    
    sSource = "VWCatalogueBulkUpdate.InitialiseStyles"
    
    'Cell style for the left heading
    m_typCellStyleLeftHeading = m_ofwhtmlReference.GetCellStyle(ECellStyle.CELLReadWriteLabelStyle)
    
    'Cell style for the right details
    m_typCellStyleRightDetails = m_ofwhtmlReference.GetCellStyle(ECellStyle.CELLReadWriteDataStyle)

    'Cell style for ok cancel button
    m_typCellStyleForOKCancel = m_ofwhtmlReference.GetCellStyle(ECellStyle.CELLNormalStyle)

    'font style for left headings
    m_typFontStyleLeftHeading = m_ofwhtmlReference.GetFontStyle(EFontStyle.FONTReadWriteLabelStyle)
    
    'font style for right details
    m_typFontStyleRightDetails = m_ofwhtmlReference.GetFontStyle(EFontStyle.FONTReadWriteDataStyle)
    
    
    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub


Private Sub Terminate()
    Set m_objContext = Nothing
    Set m_oboCatalogueBulkUpdate = Nothing
    Set m_oBOCatalogueElementMatrix = Nothing
    Set m_ofwhtmlReference = Nothing
    Set m_oFWAppData = Nothing
    Set m_ofwRefData = Nothing
    Set m_oboEditorialPermission = Nothing
    Set m_colSecondSet = Nothing
    Set m_colFirstSet = Nothing
    Set m_oFWTable = Nothing
    Set m_oVWSharedCatalogueEntry = Nothing
End Sub

Private Sub ObjectControl_Activate()
    Initialise
End Sub

Private Sub ObjectControl_Deactivate()
    Terminate
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = True
End Function






Public Sub DisplayBulkUpdateFieldPage()
'Displays the form to select the field to be edited
'
' d2    ksk     08Jul1999
' d3   ksk      09Jul1999         Added tooltip help for all hyper links
On Error GoTo ErrHandler

    Dim typCustomType As CellStyle
    Dim typTable As TableStyle
    Dim sSource As String
    Dim sHTML As String
    Dim ofwFields As FWRecordList
    Dim sEditsetParam As String
    sSource = "VWBulkUpdate.DisplayBulkUpdateFieldPage"
    
    
    
    'Makesure that the editset ID was passes to this page
    If m_aspRequest(gk_sParamEditSetID) = "" Then
        ' d2    changed the enumeration
        RaiseError sSource, EPROCATINTERNALERROR.peEditSetIDInvalue, _
                        "Editset ID not found"
    End If
    m_lEditsetID = CLng(m_aspRequest(gk_sParamEditSetID))
    If m_lEditsetID <= 0 Then
        ' d2    changed the enumeration
        RaiseError sSource, EPROCATINTERNALERROR.peEditSetIDInvalue, _
                            "Editset ID not found"
    End If
    sEditsetParam = "?" & gk_sParamEditSetID & "=" & m_lEditsetID

    
    CheckPermissions ESecurityCheck.EUserCanBulkUpdate, sSource, ""
    
   
    sHTML = HeaderHTML
    
    'Sart the form and add the EditsetID as a querystring into the form
    sHTML = sHTML & m_ofwhtmlReference.FormStart(gk_sLinkDisplayBulkUpdateValue & sEditsetParam, _
                                                gk_sMainFrame, mk_sNameBulkUpdateField, MPost)
                                                
    'Get the bulk update fields to be displayed
    Set ofwFields = m_oBOCatalogueElementMatrix.FieldListForBulkUpdate
    
    'Set the tables format
    typTable = m_ofwhtmlReference.GetTableStyle(ETableStyle.TABContentStyle)
    m_oFWTable.TableStyle = typTable
    
    'Add the 1st row - the field to select using a list box
    m_oFWTable.StartRow
    typCustomType = m_ofwhtmlReference.GetCellStyle(ECellStyle.CELLReadWriteLabelStyle)
    typCustomType.sWidth = "25%"
    'Display the filed selection label
    m_oFWTable.AddCellAsHTML m_ofwhtmlReference.FontWithStyle(m_oFWAppData.GetDisplayInformation( _
                    EEditorialResourceFile.erfBulkEditFieldCaption), _
                    m_ofwhtmlReference.GetFontStyle(EFontStyle.FONTReadWriteLabelStyle)), typCustomType
    
    typCustomType = m_ofwhtmlReference.GetCellStyle(ECellStyle.CELLReadWriteDataStyle)
    typCustomType.sWidth = "75%"
    'Display the list box with the filed names
    m_oFWTable.AddCellAsHTML m_ofwhtmlReference.FontWithStyle( _
                m_ofwhtmlReference.Listbox(mk_sNameFieldList, ofwFields), _
                m_ofwhtmlReference.GetFontStyle(EFontStyle.FONTReadWriteDataStyle)), typCustomType
    m_oFWTable.EndRow
    
    'Add the 2nd row with the buttons in them
    m_oFWTable.StartRow
    m_oFWTable.AddEmptyCell
    
    typCustomType = m_ofwhtmlReference.GetCellStyle(ECellStyle.CELLCustomStyle)
    typCustomType.eALIGN = HALRight
    typCustomType.eBGCOLOR = COLOURNone
    
    
    m_oFWTable.AddCellAsHTML m_ofwhtmlReference.SubmitImage(gk_sImagesPath & gk_sImageOK, mk_sNameButtonOk, _
                            m_ofwRefData.Search(gToolTip, ETOOLTIP.ttButtonOKHelp)) & _
                            m_ofwhtmlReference.Hyperlink(gk_sLinkDisplayEditSetDetails & sEditsetParam, _
                                    m_ofwhtmlReference.Image(mk_sNameButtonCancel, gk_sImagesPath & gk_sImageCancel, _
                                                m_ofwRefData.Search(gToolTip, ETOOLTIP.ttButtonCancelHelp)), _
                            gk_sMainFrame, mk_sNameButtonCancel, _
                            m_ofwRefData.Search(gToolTip, ETOOLTIP.ttButtonCancelHelp))
    m_oFWTable.EndRow
    
    sHTML = sHTML & m_oFWTable.GetHTML
    sHTML = sHTML & m_ofwhtmlReference.EndTag(gk_sHTMLForm)
    
    
    sHTML = sHTML & HTMLToAddHelpButton(gk_sHelpLinkToBulkUpdateFieldSelection)
    
    GetObjectContext.Item("Response").Write sHTML
    
    'Release all the objects allocated locally
    Set ofwFields = Nothing
    Exit Sub
    
ErrHandler:
    ShowError sSource
End Sub



Public Sub DisplayBulkUpdateValuePage()
'Displays the form to select the field to be edited
'
' d2    ksk     08Jul1999
' d3   ksk      09Jul1999         Added tooltip help for all hyper links
' d4   ksk      20Jul1999         Added variables sIDFieldName and sValueFieldName
'                                 to be used by the list box and changed the usage
'                                 of static SQL select to call to lookup.
' d5   ksk      09Aug1999       Changed the Constant name for Closure status
' d8   ksk      27Oct1999       Added closure type and closure code fields
On Error GoTo ErrHandler

    Dim sSource As String
    Dim ofwLookup As FWRecordList
    Dim bIsLookup As Boolean
    Dim sIDFieldName As String
    Dim sValueFieldName As String
    Dim sEditsetParam As String
    
    
    sSource = "VWBulkUpdate.DisplayBulkUpdateValuePage"
    
    
    'Check if the Cancel button was clicked if so return to the main page
    If m_aspRequest(mk_sNameButtonCancel & gk_sXCoordinate) <> 0 Or _
        m_aspRequest(mk_sNameButtonCancel & gk_sYCoordinate) <> 0 Then
        'Code to return to the main page to be added here
        m_aspResponse.Redirect (gk_sLinkDisplayEditSetDetails)
        Exit Sub
    End If
    
    'Makesure that the editset ID was passes to this page
    If m_aspRequest(gk_sParamEditSetID) = "" Then
        ' d2    changed the enumeration
        RaiseError sSource, EPROCATINTERNALERROR.peEditSetIDInvalue, _
                            "Editset ID not found"
    End If
    m_lEditsetID = CLng(m_aspRequest(gk_sParamEditSetID))
    If m_lEditsetID <= 0 Then
        ' d2    changed the enumeration
        RaiseError sSource, EPROCATINTERNALERROR.peEditSetIDInvalue, _
                            "Editset ID not found"
        Exit Sub
    End If
    sEditsetParam = "?" & gk_sParamEditSetID & "=" & m_lEditsetID
    
    CheckPermissions ESecurityCheck.EUserCanBulkUpdate, sSource, ""
    
    'If no field is selected then return to the field selection process
    If m_aspRequest(mk_sNameFieldList) = "" Then
        DisplayBulkUpdateFieldPage
        Exit Sub
    End If
    
    m_eSelectedFieldID = CInt(m_aspRequest(mk_sNameFieldList))
    
    If m_eSelectedFieldID < 1 Then
        RaiseError sSource, EPROCATINTERNALERROR.peBulkUpdate, _
                            "Field to bulk update not set"
        Exit Sub
    End If

    FillLookupValues ofwLookup, bIsLookup, sIDFieldName, sValueFieldName
    
    DisplayBulkupdateValueContentPage sEditsetParam, ofwLookup, bIsLookup, sIDFieldName, sValueFieldName
        
    Exit Sub
ErrHandler:
    ShowError sSource
End Sub


Private Sub DisplayBulkupdateValueContentPage(t_sEditsetParam As String, _
                                    Optional ByRef t_ofwLookup As FWRecordList, _
                                    Optional t_bIsLookup As Boolean, _
                                    Optional t_sIDFieldName As String, _
                                    Optional t_sValueFieldName As String)
'This method displays the content for the page.
'
' d11  ksk      07Dec1999         Changes to display the EAD browser for select fields
' d26  alm      02Oct2001         Added support for cpAccessionedAsCopy, cpOrderable,
'                                 cpLCI, cpExtractAtTransfer
'
    Dim sSource As String
    Dim typCustomType As CellStyle
    Dim typTable As TableStyle
    Dim sHTML As String
    Dim sTempHTML As String
    Dim i As Integer
    Dim lFieldLength As Long
    Dim eDataType As EValidCheck
    
    sSource = "VWBulkUpdate.DisplayBulkupdateValueContentPage"
    
    sHTML = sHTML & HeaderHTML
    
    'Start the entry form with the editsetID as the parameter
    sHTML = sHTML & m_ofwhtmlReference.FormStart(gk_sLinkBulkUpdateLevels & t_sEditsetParam, _
                                                gk_sMainFrame, mk_sNameBulkUpdateValue, _
                                                MPost, True)
    
    'Add the Selected field ID into a hidden field
    sHTML = sHTML & m_ofwhtmlReference.HiddenField(mk_sNameFieldList, CStr(m_eSelectedFieldID))
    
    
    typTable = m_ofwhtmlReference.GetTableStyle(ETableStyle.TABContentStyle)
    m_oFWTable.TableStyle = typTable
    
    'Add the 1st row with the selected field
    m_oFWTable.StartRow
    typCustomType = m_ofwhtmlReference.GetCellStyle(ECellStyle.CELLReadOnlyLabelStyle)
    typCustomType.sWidth = "25%"
    
    m_oFWTable.AddCellAsHTML m_ofwhtmlReference.FontWithStyle(m_oFWAppData.GetDisplayInformation( _
                    EEditorialResourceFile.erfBulkEditFieldCaption), _
                    m_ofwhtmlReference.GetFontStyle(EFontStyle.FONTReadOnlyLabelStyle)), typCustomType
                    
    typCustomType = m_ofwhtmlReference.GetCellStyle(ECellStyle.CELLReadOnlyDataStyle)
    typCustomType.sWidth = "75%"
    typCustomType.lCOLSPAN = 4
    
    m_oFWTable.AddCellAsHTML m_ofwhtmlReference.FontWithStyle( _
            m_oBOCatalogueElementMatrix.DescriptionFromEnumeration _
                            (m_eSelectedFieldID), _
                            m_ofwhtmlReference.GetFontStyle(EFontStyle.FONTReadOnlyDataStyle)), _
                            typCustomType
    m_oFWTable.EndRow
    
    
    Select Case m_eSelectedFieldID
    
    ' d11  Display the EAD field with the ActiveX control
    Case ECATALOGUEPROPERTIES.cpAdministrative_BiographicalHistory, _
        ECATALOGUEPROPERTIES.cpArrangement, _
        ECATALOGUEPROPERTIES.cpMapDesignation, _
        ECATALOGUEPROPERTIES.cpScopeAndContent, _
        ECATALOGUEPROPERTIES.cpTitle
        
        DisplayBulkupdateEADFields
        
    Case ECATALOGUEPROPERTIES.cpClosureType, ECATALOGUEPROPERTIES.cpClosureCode
        DisplayForClosureTypeAndCode t_ofwLookup, t_sIDFieldName, t_sValueFieldName
        
    Case ECATALOGUEPROPERTIES.cpImmediateSourceOfAcquisition, _
            ECATALOGUEPROPERTIES.cpCopiesInformation, _
            ECATALOGUEPROPERTIES.cpPublicationNote, _
            ECATALOGUEPROPERTIES.cpUnpublishedFindingAids, _
            ECATALOGUEPROPERTIES.cpCreatorNames, _
            ECATALOGUEPROPERTIES.cpPlaceOfDeposit, _
            ECATALOGUEPROPERTIES.cpLocationOfOriginals
        
        DisplayListDataItemsEntryForm
        
    ' d26
    Case ECATALOGUEPROPERTIES.cpAccessionedAsCopy, _
         ECATALOGUEPROPERTIES.cpOrderable, _
         ECATALOGUEPROPERTIES.cpAR, _
         ECATALOGUEPROPERTIES.cpExtractAtTransfer
         
         DisplayBulkupdateCheckbox
         
    Case Else
        If t_bIsLookup Then
            DisplayBulkupdateRegularLookupFields t_ofwLookup, t_sIDFieldName, t_sValueFieldName
        Else
            DisplayBulkupdateRegularFields
        End If
    End Select
    
    'Add the 3rd row with the list catalogue levels
    m_oFWTable.StartRow
    
    'Display the lable for selecting levels
    typCustomType = m_ofwhtmlReference.GetCellStyle(ECellStyle.CELLReadWriteLabelStyle)
    typCustomType.sWidth = "25%"
    m_oFWTable.AddCellAsHTML m_ofwhtmlReference.FontWithStyle(m_oFWAppData.GetDisplayInformation( _
                            EEditorialResourceFile.erfBulkEditLevelCaption), _
                            m_ofwhtmlReference.GetFontStyle(EFontStyle.FONTReadWriteLabelStyle)), _
                            typCustomType
    
    'Build the levels checkbox list
    sTempHTML = ""
    For i = 1 To gk_iNumberOfLevels
        sTempHTML = sTempHTML & m_ofwhtmlReference.Checkbox(GetCatalogueLevelDescription(i), _
                                                    mk_sNameLevel & Trim(Str(i)), m_bSelectedLevels(i))
    Next i
    
    'Display the checkbox Levels list
    typCustomType = m_ofwhtmlReference.GetCellStyle(ECellStyle.CELLColumnDataStyle)
    typCustomType.sWidth = "75%"
    typCustomType.lCOLSPAN = 4
    m_oFWTable.AddCellAsHTML m_ofwhtmlReference.FontWithStyle(sTempHTML, _
                            m_ofwhtmlReference.GetFontStyle(EFontStyle.FONTReadWriteDataStyle)) _
                            , typCustomType
    m_oFWTable.EndRow
    
    'Add the 4the row with the images
    m_oFWTable.StartRow
    m_oFWTable.AddEmptyCell
    typCustomType = m_ofwhtmlReference.GetCellStyle(ECellStyle.CELLCustomStyle)
    typCustomType.eALIGN = HALRight
    typCustomType.eBGCOLOR = COLOURNone
    
    m_oFWTable.AddCellAsHTML m_ofwhtmlReference.SubmitImage(gk_sImagesPath & gk_sImageOK, mk_sNameButtonOk, _
                            m_ofwRefData.Search(gToolTip, ETOOLTIP.ttButtonOKHelp)) & _
                            m_ofwhtmlReference.Hyperlink(gk_sLinkDisplayEditSetDetails & t_sEditsetParam, _
                            m_ofwhtmlReference.Image(mk_sNameButtonCancel, gk_sImagesPath & gk_sImageCancel, _
                                                m_ofwRefData.Search(gToolTip, ETOOLTIP.ttButtonCancelHelp)), _
                            gk_sMainFrame, mk_sNameButtonCancel, _
                            m_ofwRefData.Search(gToolTip, ETOOLTIP.ttButtonCancelHelp))
    m_oFWTable.EndRow
    
    sHTML = sHTML & m_oFWTable.GetHTML
    sHTML = sHTML & m_ofwhtmlReference.EndTag(gk_sHTMLForm)
    
    sHTML = sHTML & HTMLToAddHelpButton(gk_sHelpLinkToBulkUpdateFieldSelection)
    
    GetObjectContext.Item("Response").Write sHTML
    
    'Release all the objects allocated locally
    Set t_ofwLookup = Nothing
    Exit Sub

ErrHandler:
    ShowError sSource
End Sub


Private Sub DisplayBulkupdateEADFields()
'This method displays the value updation area with the text box and label for regular fields
'
' d11  ksk      07Dec1999       Changes to display the EAD browser for select fields
' d12  ksk      14Dec1999       Added code to convert the apostrophe and quote in the ead field
'
' d19  daj      05May2000       Modified the EADEditor control code so that it works with the
'                               new non-event driven control (see VWCatalogueEntry for more details)
' d20  daj      20Jun2000       Added Admin History DTD type to the bulk update control

    Dim typCustomType  As CellStyle
    Dim typRowStyle As RowStyle
    Dim lFieldLength  As Long
    Dim eDataType  As EValidCheck
    Dim sTextToDisplay As String, sTempHTML As String
    
    'Restore the state back to original for display if the state was saved before
    If m_aspRequest(mk_sParamRestoreState) = mk_sStateRestoreValue Then
        
        m_oboCatalogueBulkUpdate.PROCATTransientInterface_GetState
        UpdateLocalObjectWithBusinessobjectValues True
        
        
        'Render before replacing the apostropes and quotes
        'Parse the value to display in the required format
        sTextToDisplay = RenderXML(CStr(m_vUpdateValue1))
    
        'd12
        'Replace the apostrophe character with HTML decimal code for apostrophe
        m_vUpdateValue1 = Replace(m_vUpdateValue1, "'", gk_sHTMLApostropheDecimalCode)
        'Replace the quotes character with HTML decimal code for quotes
        m_vUpdateValue1 = Replace(m_vUpdateValue1, """", gk_sHTMLQuotationDecimalCode)
    
    Else
        
        m_oboCatalogueBulkUpdate.PROCATTransientInterface_DeleteState
    
    End If

    
    'Add the 2nd row with the field with the values for entry this is a listbox for look ups,
    
    ' Row that displays the rendered text
    m_oFWTable.StartRow
        
    
    'Display the label for entering the value
    typCustomType = m_ofwhtmlReference.GetCellStyle(ECellStyle.CELLReadWriteLabelStyle)
    typCustomType.sWidth = "25%"
    m_oFWTable.AddCellAsHTML m_ofwhtmlReference.FontWithStyle(m_oFWAppData.GetDisplayInformation( _
                    EEditorialResourceFile.erfBulkEditValueCaption), _
                    m_ofwhtmlReference.GetFontStyle(EFontStyle.FONTReadWriteLabelStyle)), typCustomType
    
    
    'Display the list box/text box/text area depending on the type of the value
    typCustomType = m_ofwhtmlReference.GetCellStyle(ECellStyle.CELLReadWriteDataStyle)
    typCustomType.sWidth = "75%"
    lFieldLength = m_oBOCatalogueElementMatrix.CataloguePropertyLength(m_eSelectedFieldID)
    eDataType = m_oBOCatalogueElementMatrix.CataloguePropertyDataType(m_eSelectedFieldID)
    
    m_oFWTable.AddCellAsHTML sTextToDisplay, typCustomType
    
    m_oFWTable.EndRow

    ' We need to work out which DTD to use in XMetal
    ' d20 - Added Admin History
    Dim eEADType As EADType
    Select Case m_eSelectedFieldID
        Case cpScopeAndContent
            eEADType = eadScopeAndContent
        Case cpTitle, cpMapDesignation
            eEADType = eadTitle
        Case cpArrangement
            eEADType = eadArrangement
        Case cpAdministrative_BiographicalHistory
            eEADType = eadAdministrativeHistory
    End Select
    
    'The row that displays the button to edit and update button along with the ACTIVEX control
    m_oFWTable.StartRow
    typCustomType = m_ofwhtmlReference.GetCellStyle(ECellStyle.CELLReadWriteLabelStyle)
    typCustomType.sWidth = "25%"
    m_oFWTable.AddCellAsHTML gk_sHTMLNonBreakSpace, typCustomType
       
    'Generate the HTML display the control and the buttons and the hidden files and text
    'Hidden fields for EAD and status fields
    'd18 encode the EAD field entry
    sTempHTML = m_ofwhtmlReference.HiddenField(mk_sNameValue1, IIf(m_vUpdateValue1 = "", " ", m_ofwhtmlReference.EncodeAsHTML(CStr(m_vUpdateValue1))))
    sTempHTML = sTempHTML & m_ofwhtmlReference.HiddenField(mk_sNameHiddenStatusField, "false")
    sTempHTML = sTempHTML & m_ofwhtmlReference.HiddenField(ksNameDTDType, CStr(eEADType))
    
    'The ActiveX control to launch the XML editor
    sTempHTML = sTempHTML & m_ofwhtmlReference.EADEditorLauncher(mk_sNameBulkUpdateValue, _
                mk_sNameBulkUpdateValue & "." & mk_sNameValue1, _
                mk_sNameBulkUpdateValue & "." & mk_sNameHiddenStatusField, _
                mk_sEADControlInstanceName & mk_sNameValue1)
    
    'The edit button to display the XML editor
    sTempHTML = sTempHTML & m_ofwhtmlReference.ImageClick(gk_sImagesPath & gk_sImageEdit, GenerateEditOnclickCode(mk_sNameValue1, eEADType))
    
    'The update button to update the changed ead
    sTempHTML = sTempHTML & m_ofwhtmlReference.ImageClick(gk_sImagesPath & gk_sIMAGEUpdate, GenerateUpdateOnclickCode(mk_sNameValue1))
    
    typCustomType = m_ofwhtmlReference.GetCellStyle(ECellStyle.CELLReadWriteDataStyle)
    typCustomType.sWidth = "75%"
    m_oFWTable.AddCellAsHTML sTempHTML, typCustomType
    
    m_oFWTable.EndRow


End Sub
Private Function GenerateUpdateOnclickCode( _
    t_sFieldName) As String
'
' Function returns the code to be used in the on click of update button
' which is used to update the value from the changes made in the editor to the browser
'
' d19  daj      05May2000       Created
    
    GenerateUpdateOnclickCode = gk_sNameUpdateFieldFunction & mk_sEADControlInstanceName & t_sFieldName & "(); "
    
End Function

Private Function GenerateEditOnclickCode( _
    t_sFieldName As String, _
    t_eEADType As EADType, _
    Optional t_bUseHiddenSwitchField As Boolean = False) As String
'
' Function returns code to be used in the on click event of the edit button
' This button is used to launch the XML editor to edit the contents in the hidden field.
'
' d19  daj      05May2000       Created
    
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
        mk_sEADControlInstanceName & t_sFieldName & _
        ".Launch( " & _
        t_sFieldName & ".value, " & sStr
    
    GenerateEditOnclickCode = GenerateEditOnclickCode & "); "

End Function


Private Sub FillLookupValues(ByRef t_ofwLookup As FWRecordList, ByRef t_bIsLookup As Boolean, _
                    ByRef t_sIDFieldName As String, ByRef t_sValueFieldName As String)
' d4
'If the field is a look up field then find the list of lookup values
Dim oBOLookup As BOLookup

    Set oBOLookup = New BOLookup
    
    t_bIsLookup = False
    Select Case m_eSelectedFieldID
        Case ECATALOGUEPROPERTIES.cpPhysicalDescription_Form
            
            t_bIsLookup = True
            ' d6   ksk      13Aug1999
            oBOLookup.LookupTableID = ELOOKUPTYPE.lutPhysicalForm
            Set t_ofwLookup = oBOLookup.LookupList
            
            t_sIDFieldName = gk_sDBCOLPhysicalDescriptionFormID
            t_sValueFieldName = gk_sDBCOLPhysicalDescriptionFormDescription
        
        Case ECATALOGUEPROPERTIES.cpAccessCondition
            
            t_bIsLookup = True
            ' d6   ksk      13Aug1999
            oBOLookup.LookupTableID = ELOOKUPTYPE.lutAccessCondition
            Set t_ofwLookup = oBOLookup.LookupList
            
            t_sIDFieldName = gk_sDBCOLAccessConditionID
            t_sValueFieldName = gk_sDBCOLAccessConditionDescription
        
        Case ECATALOGUEPROPERTIES.cpClosureStatus
            
            t_bIsLookup = True
            ' d6   ksk      13Aug1999
            oBOLookup.LookupTableID = ELOOKUPTYPE.lutClosureStatus
            Set t_ofwLookup = oBOLookup.LookupList
            
            ' d5   ksk
            t_sIDFieldName = gk_sDBCOLClosureStatus
            t_sValueFieldName = gk_sDBCOLClosureStatusDescription
            
        Case ECATALOGUEPROPERTIES.cpLegalStatus
            
            t_bIsLookup = True
            ' d6   ksk      13Aug1999
            oBOLookup.LookupTableID = ELOOKUPTYPE.lutLegalStatus
            Set t_ofwLookup = oBOLookup.LookupList
            
            t_sIDFieldName = gk_sDBCOLLegalStatusID
            t_sValueFieldName = gk_sDBCOLLegalStatusDescription
        
        Case ECATALOGUEPROPERTIES.cpLanguage
            
            t_bIsLookup = True
            ' d6   ksk      13Aug1999
            oBOLookup.LookupTableID = ELOOKUPTYPE.lutLanguage
            Set t_ofwLookup = oBOLookup.LookupList
            
            t_sIDFieldName = gk_sDBCOLLanguageID
            t_sValueFieldName = gk_sDBCOLLanguageDescription
    
        Case ECATALOGUEPROPERTIES.cpClosureType, ECATALOGUEPROPERTIES.cpClosureCode
            ' d10  ksk     changed to closure type
            t_bIsLookup = True
            oBOLookup.LookupTableID = ELOOKUPTYPE.lutClosureType
            Set t_ofwLookup = oBOLookup.LookupList
            
            t_sIDFieldName = gk_sDBCOLClosureType
            t_sValueFieldName = gk_sDBCOLClosureTypeDescription
    
    End Select
    

End Sub



Private Sub DisplayBulkupdateRegularLookupFields(ByRef t_ofwLookup As FWRecordList, _
                    ByRef t_sIDFieldName As String, ByRef t_sValueFieldName As String)
'This method displays the value updation area with the text box and label for regular lookup
' fields except for closure type.
'
' d8   ksk      27Oct1999       Created
' d24  gdb      16Feb2001       added warning message to DisplayBulkupdateRegularLookupFields
Dim typCustomType  As CellStyle
Dim lFieldLength  As Long
Dim eDataType  As EValidCheck
    
  
'd24 - display warning messageif user has selected closure status
  Select Case m_eSelectedFieldID
    
    Case ECATALOGUEPROPERTIES.cpClosureStatus

   m_oFWTable.StartRow
    typCustomType = m_ofwhtmlReference.GetCellStyle(ECellStyle.CELLReadOnlyLabelStyle)
    typCustomType.sWidth = "25%"
    
    m_oFWTable.AddCellAsHTML m_ofwhtmlReference.FontWithStyle(m_oFWAppData.GetDisplayInformation( _
                    EEditorialResourceFile.erfBulkEditPleaseNote), _
                    m_ofwhtmlReference.GetFontStyle(EFontStyle.FONTReadOnlyLabelStyle)), typCustomType
                    
    typCustomType = m_ofwhtmlReference.GetCellStyle(ECellStyle.CELLReadOnlyDataStyle)
    typCustomType.sWidth = "75%"
    typCustomType.lCOLSPAN = 4
    
    m_oFWTable.AddCellAsHTML m_ofwhtmlReference.FontWithStyle(m_oFWAppData.GetDisplayInformation( _
                    EEditorialResourceFile.erfBulkEditValueEntryInformAboutChanges), _
                    m_ofwhtmlReference.GetFontStyle(EFontStyle.FONTReadOnlyLabelStyle)), _
                            typCustomType
    m_oFWTable.EndRow
   
    End Select
    
    
    'Add the 2nd row with the field with the values for entry this is a listbox for look ups,
    ' a textbox for all other field types
    m_oFWTable.StartRow
    
    'Display the label for entering the value
    typCustomType = m_ofwhtmlReference.GetCellStyle(ECellStyle.CELLReadWriteLabelStyle)
    typCustomType.sWidth = "25%"
    m_oFWTable.AddCellAsHTML m_ofwhtmlReference.FontWithStyle(m_oFWAppData.GetDisplayInformation( _
                    EEditorialResourceFile.erfBulkEditValueCaption), _
                    m_ofwhtmlReference.GetFontStyle(EFontStyle.FONTReadWriteLabelStyle)), typCustomType
    
    'Display the list box/text box/text area depending on the type of the value
    typCustomType = m_ofwhtmlReference.GetCellStyle(ECellStyle.CELLReadWriteDataStyle)
    typCustomType.sWidth = "75%"
    m_oFWTable.AddCellAsHTML m_ofwhtmlReference.FontWithStyle( _
                m_ofwhtmlReference.Listbox(mk_sNameValue1, t_ofwLookup, False, , , _
                    t_sIDFieldName, t_sValueFieldName), _
                m_ofwhtmlReference.GetFontStyle(EFontStyle.FONTReadWriteDataStyle)), _
                typCustomType
                    
    m_oFWTable.EndRow
    
End Sub


Private Sub DisplayBulkupdateRegularFields()
'This method displays the value updation area with the text box and label for regular fields
'
' d8   ksk      27Oct1999       Created
' d25  gdb      16Feb2001       added warning message

Dim typCustomType  As CellStyle
Dim lFieldLength  As Long
Dim eDataType  As EValidCheck
    
    
   ' d25-   added warning message if user has selected covering dates orrecord opening dates
  Select Case m_eSelectedFieldID
    
    Case ECATALOGUEPROPERTIES.cpCoveringDates

   m_oFWTable.StartRow
    typCustomType = m_ofwhtmlReference.GetCellStyle(ECellStyle.CELLReadOnlyLabelStyle)
    typCustomType.sWidth = "25%"
    
    m_oFWTable.AddCellAsHTML m_ofwhtmlReference.FontWithStyle(m_oFWAppData.GetDisplayInformation( _
                    EEditorialResourceFile.erfBulkEditPleaseNote), _
                    m_ofwhtmlReference.GetFontStyle(EFontStyle.FONTReadOnlyLabelStyle)), typCustomType
                    
    typCustomType = m_ofwhtmlReference.GetCellStyle(ECellStyle.CELLReadOnlyDataStyle)
    typCustomType.sWidth = "75%"
    typCustomType.lCOLSPAN = 4
    
    m_oFWTable.AddCellAsHTML m_ofwhtmlReference.FontWithStyle(m_oFWAppData.GetDisplayInformation( _
                    EEditorialResourceFile.erfBulkEditValueEntryInformAboutCoveringdates), _
                    m_ofwhtmlReference.GetFontStyle(EFontStyle.FONTReadOnlyLabelStyle)), _
                            typCustomType
    m_oFWTable.EndRow
       
    Case ECATALOGUEPROPERTIES.cpRecordOpeningDate
    
    m_oFWTable.StartRow
    typCustomType = m_ofwhtmlReference.GetCellStyle(ECellStyle.CELLReadOnlyLabelStyle)
    typCustomType.sWidth = "25%"
    
    m_oFWTable.AddCellAsHTML m_ofwhtmlReference.FontWithStyle(m_oFWAppData.GetDisplayInformation( _
                    EEditorialResourceFile.erfBulkEditPleaseNote), _
                    m_ofwhtmlReference.GetFontStyle(EFontStyle.FONTReadOnlyLabelStyle)), typCustomType
                    
    typCustomType = m_ofwhtmlReference.GetCellStyle(ECellStyle.CELLReadOnlyDataStyle)
    typCustomType.sWidth = "75%"
    typCustomType.lCOLSPAN = 4
    
    m_oFWTable.AddCellAsHTML m_ofwhtmlReference.FontWithStyle(m_oFWAppData.GetDisplayInformation( _
                    EEditorialResourceFile.erfBulkEditValueEntryInformAboutOpeningdates), _
                    m_ofwhtmlReference.GetFontStyle(EFontStyle.FONTReadOnlyLabelStyle)), _
                            typCustomType
    m_oFWTable.EndRow

    End Select
    
        
    'Add the 2nd row with the field with the values for entry this is a listbox for look ups,
    ' a textbox for all other field types
    m_oFWTable.StartRow
    
    'Display the label for entering the value
    typCustomType = m_ofwhtmlReference.GetCellStyle(ECellStyle.CELLReadWriteLabelStyle)
    typCustomType.sWidth = "25%"
    m_oFWTable.AddCellAsHTML m_ofwhtmlReference.FontWithStyle(m_oFWAppData.GetDisplayInformation( _
                    EEditorialResourceFile.erfBulkEditValueCaption), _
                    m_ofwhtmlReference.GetFontStyle(EFontStyle.FONTReadWriteLabelStyle)), typCustomType
    
    'Display the list box/text box/text area depending on the type of the value
    typCustomType = m_ofwhtmlReference.GetCellStyle(ECellStyle.CELLReadWriteDataStyle)
    typCustomType.sWidth = "75%"
    lFieldLength = m_oBOCatalogueElementMatrix.CataloguePropertyLength(m_eSelectedFieldID)
    eDataType = m_oBOCatalogueElementMatrix.CataloguePropertyDataType(m_eSelectedFieldID)
    
    If lFieldLength > gk_iMaxTextEntryInALine Then
        m_oFWTable.AddCellAsHTML m_ofwhtmlReference.FontWithStyle( _
                            m_ofwhtmlReference.TextArea(mk_sNameValue1, "", _
                                mk_iLengthOfTextField, mk_iNumberOfRows, lFieldLength, True) _
                                , m_ofwhtmlReference.GetFontStyle( _
                                EFontStyle.FONTReadWriteDataStyle)), typCustomType
    Else
        If eDataType = VCNumber Then
            m_oFWTable.AddCellAsHTML m_ofwhtmlReference.FontWithStyle( _
                    m_ofwhtmlReference.InputField(ITText, mk_sNameValue1, _
                    mk_iLengthOfNumericField, lFieldLength, "", eDataType, True, False, _
                    0, 9999999), _
                    m_ofwhtmlReference.GetFontStyle(EFontStyle.FONTReadWriteDataStyle)) _
                    , typCustomType
        Else
            m_oFWTable.AddCellAsHTML m_ofwhtmlReference.FontWithStyle( _
                    m_ofwhtmlReference.InputField(ITText, mk_sNameValue1, _
                    mk_iLengthOfTextField, lFieldLength, "", eDataType, True), _
                    m_ofwhtmlReference.GetFontStyle(EFontStyle.FONTReadWriteDataStyle)), _
                    typCustomType
        End If
    End If
    
    m_oFWTable.EndRow

End Sub

Private Sub DisplayBulkupdateCheckbox()
'This method displays the value updation area with a checkbox
'
' d26  alm      01Oct2001       Created
'
    Dim typCustomType  As CellStyle
           
    Select Case m_eSelectedFieldID
    
        Case ECATALOGUEPROPERTIES.cpAccessionedAsCopy, _
             ECATALOGUEPROPERTIES.cpAR, _
             ECATALOGUEPROPERTIES.cpExtractAtTransfer
             
            ' Start row
            m_oFWTable.StartRow
            
            ' Display the 'to' text
            typCustomType = m_ofwhtmlReference.GetCellStyle(ECellStyle.CELLReadWriteLabelStyle)
            typCustomType.sWidth = "25%"
            m_oFWTable.AddCellAsHTML m_ofwhtmlReference.FontWithStyle(m_oFWAppData.GetDisplayInformation( _
                    EEditorialResourceFile.erfBulkEditValueCaption), _
                    m_ofwhtmlReference.GetFontStyle(EFontStyle.FONTReadWriteLabelStyle)), typCustomType
            
            ' Display the checkbox
            typCustomType = m_ofwhtmlReference.GetCellStyle(ECellStyle.CELLReadWriteDataStyle)
            typCustomType.sWidth = "75%"
            typCustomType.lCOLSPAN = 4
            m_oFWTable.AddCellAsHTML m_ofwhtmlReference.FontWithStyle(m_ofwhtmlReference.Checkbox("", _
                t_sName:=mk_sNameValue1), _
                m_ofwhtmlReference.GetFontStyle(EFontStyle.FONTReadWriteDataStyle)), typCustomType

            ' End row
            m_oFWTable.EndRow
            
        Case ECATALOGUEPROPERTIES.cpOrderable
                
            ' Start row
            m_oFWTable.StartRow
            
            ' Display the 'to' text
            typCustomType = m_ofwhtmlReference.GetCellStyle(ECellStyle.CELLReadWriteLabelStyle)
            typCustomType.sWidth = "25%"
            m_oFWTable.AddCellAsHTML m_ofwhtmlReference.FontWithStyle(m_oFWAppData.GetDisplayInformation( _
                    EEditorialResourceFile.erfBulkEditValueCaption), _
                    m_ofwhtmlReference.GetFontStyle(EFontStyle.FONTReadWriteLabelStyle)), typCustomType
            
            ' Display the checkbox
            typCustomType = m_ofwhtmlReference.GetCellStyle(ECellStyle.CELLReadWriteDataStyle)
            typCustomType.sWidth = "75%"
            typCustomType.lCOLSPAN = 4
            m_oFWTable.AddCellAsHTML m_ofwhtmlReference.FontWithStyle(m_ofwhtmlReference.Checkbox("", _
                t_sName:=mk_sNameValue1, t_bChecked:=True), _
                m_ofwhtmlReference.GetFontStyle(EFontStyle.FONTReadWriteDataStyle)), typCustomType

            ' End row
            m_oFWTable.EndRow
       
    End Select

End Sub


Private Sub DisplayForClosureTypeAndCode(ByRef t_ofwLookup As FWRecordList, _
                    ByRef t_sIDFieldName As String, ByRef t_sValueFieldName As String)
' d8   ksk      27Oct1999       Created
'Displays two rows for two fields the closure type and closure code fields for bulk updation
' d23  gdb      16Feb2001       added message to displat for closure type and code

'
Dim typCustomType  As CellStyle
Dim lFieldLength  As Long
Dim eDataType  As EValidCheck

'message to display to inform user when bulk update closure type & code

   m_oFWTable.StartRow
    typCustomType = m_ofwhtmlReference.GetCellStyle(ECellStyle.CELLReadOnlyLabelStyle)
    typCustomType.sWidth = "25%"
    
    m_oFWTable.AddCellAsHTML m_ofwhtmlReference.FontWithStyle(m_oFWAppData.GetDisplayInformation( _
                    EEditorialResourceFile.erfBulkEditPleaseNote), _
                    m_ofwhtmlReference.GetFontStyle(EFontStyle.FONTReadOnlyLabelStyle)), typCustomType
                    
    typCustomType = m_ofwhtmlReference.GetCellStyle(ECellStyle.CELLReadOnlyDataStyle)
    typCustomType.sWidth = "75%"
    typCustomType.lCOLSPAN = 4
    
    m_oFWTable.AddCellAsHTML m_ofwhtmlReference.FontWithStyle(m_oFWAppData.GetDisplayInformation( _
                    EEditorialResourceFile.erfBulkEditValueEntryInformAboutChanges), _
                    m_ofwhtmlReference.GetFontStyle(EFontStyle.FONTReadOnlyLabelStyle)), _
                            typCustomType
    m_oFWTable.EndRow




    m_oFWTable.StartRow
    
    'Display the label for entering the value of Closure Type
    typCustomType = m_ofwhtmlReference.GetCellStyle(ECellStyle.CELLReadWriteLabelStyle)
    typCustomType.sWidth = "25%"
    m_oFWTable.AddCellAsHTML m_ofwhtmlReference.FontWithStyle(m_ofwRefData.Search(gCataloguePropertyMatrix, _
                    ECATALOGUEPROPERTIES.cpClosureType), _
                    m_ofwhtmlReference.GetFontStyle(EFontStyle.FONTReadWriteLabelStyle)), typCustomType
    
    'Display the lookup list box for closure type
    typCustomType = m_ofwhtmlReference.GetCellStyle(ECellStyle.CELLReadWriteDataStyle)
    typCustomType.sWidth = "75%"
    m_oFWTable.AddCellAsHTML m_ofwhtmlReference.FontWithStyle( _
                m_ofwhtmlReference.Listbox(mk_sNameValue1, t_ofwLookup, False, , , _
                    t_sIDFieldName, t_sValueFieldName), _
                m_ofwhtmlReference.GetFontStyle(EFontStyle.FONTReadWriteDataStyle)), _
                typCustomType
    
    
    m_oFWTable.EndRow

    lFieldLength = m_oBOCatalogueElementMatrix.CataloguePropertyLength(ECATALOGUEPROPERTIES.cpClosureCode)
    eDataType = m_oBOCatalogueElementMatrix.CataloguePropertyDataType(ECATALOGUEPROPERTIES.cpClosureCode)


    m_oFWTable.StartRow
    
    'Display the lable for Closure code
    typCustomType = m_ofwhtmlReference.GetCellStyle(ECellStyle.CELLReadWriteLabelStyle)
    typCustomType.sWidth = "25%"
    m_oFWTable.AddCellAsHTML m_ofwhtmlReference.FontWithStyle(m_ofwRefData.Search(gCataloguePropertyMatrix, _
                    ECATALOGUEPROPERTIES.cpClosureCode), _
                    m_ofwhtmlReference.GetFontStyle(EFontStyle.FONTReadWriteLabelStyle)), typCustomType
    
    'Display the list text box for closure code
    typCustomType = m_ofwhtmlReference.GetCellStyle(ECellStyle.CELLReadWriteDataStyle)
    typCustomType.sWidth = "75%"

    m_oFWTable.AddCellAsHTML m_ofwhtmlReference.FontWithStyle( _
            m_ofwhtmlReference.InputField(EInputType.ITText, mk_sNameValue2, _
            mk_iLengthOfNumericField, lFieldLength, "", eDataType, False, False, _
            0, 9999999), _
            m_ofwhtmlReference.GetFontStyle(EFontStyle.FONTReadWriteDataStyle)) _
            , typCustomType


    m_oFWTable.EndRow
    

    
End Sub


Private Sub DisplayListDataItemsEntryForm()
'Displays the transiently persistant pages for link data items update
' d8   ksk      27Oct1999       Created
' d15  tmj      15Mar2000       Added clauses to include add leaflet
' d26  alm      03Oct2001       Added support for creator, deposit & original
Dim sHeading As String
Dim colFirstSet As Collection, colSecondSet As Collection
    
    'Restore the state back to original for display if the state was saved before
    If m_aspRequest(mk_sParamRestoreState) = mk_sStateRestoreValue Then
        m_oboCatalogueBulkUpdate.PROCATTransientInterface_GetState
        UpdateLocalObjectWithBusinessobjectValues True
    Else
        m_oboCatalogueBulkUpdate.PROCATTransientInterface_DeleteState
    End If
    
    'Get the heading to display
    sHeading = m_oFWAppData.GetDisplayInformation( _
                    EEditorialResourceFile.erfBulkEditValueCaption)
                    
    Select Case m_eSelectedFieldID
    
    'd26
    Case ECATALOGUEPROPERTIES.cpLocationOfOriginals
        
        If DisplayLinkItemForEdit(m_eSelectedFieldID, sHeading, m_colFirstSet, m_vUpdateValue1) Then
            DisplayOneLineHeadingForContent ""  'To start a new line
        End If
        
        DisplayOneLineEditForContent _
            m_ofwhtmlReference.SubmitImage(gk_sImagesPath & gk_sImagePickCorporateBody, mk_sAddCorpBodyTermButton, "Test"), 0, 4
    
    'd26
    Case ECATALOGUEPROPERTIES.cpPlaceOfDeposit
        
        If DisplayLinkItemForEdit(m_eSelectedFieldID, sHeading, m_colFirstSet, m_vUpdateValue1) Then
            DisplayOneLineHeadingForContent ""  'To start a new line
        End If
        
        DisplayOneLineEditForContent _
            m_ofwhtmlReference.SubmitImage(gk_sImagesPath & gk_sImagePickCorporateBody, mk_sAddCorpBodyTermButton, "Test"), 0, 4
    
    'd26
    Case ECATALOGUEPROPERTIES.cpCreatorNames
        
        If DisplayLinkItemForEdit(m_eSelectedFieldID, sHeading, m_colFirstSet, m_vUpdateValue1, m_colSecondSet) Then
            DisplayOneLineHeadingForContent ""  'To start a new line
        End If
    
        'Add new link buttons (corporate boody and person)
        DisplayOneLineEditForContent _
            m_ofwhtmlReference.SubmitImage(gk_sImagesPath & gk_sImagePickCorporateBody, mk_sAddCorpBodyTermButton, "Test") & _
            m_ofwhtmlReference.SubmitImage(gk_sImagesPath & gk_sImagePickPerson, mk_sAddPersonTermButton, "Test"), 0, 4
    
    Case ECATALOGUEPROPERTIES.cpImmediateSourceOfAcquisition
        
        DisplayLinkItemForEdit m_eSelectedFieldID, sHeading, m_colFirstSet, m_vUpdateValue1, m_colSecondSet
    
        DisplayOneLineHeadingForContent ""
        'Add new link buttons (corporate boody and person)
        DisplayOneLineEditForContent _
            m_ofwhtmlReference.SubmitImage(gk_sImagesPath & gk_sImagePickCorporateBody, mk_sAddCorpBodyTermButton, "Test") & _
            m_ofwhtmlReference.SubmitImage(gk_sImagesPath & gk_sImagePickPerson, mk_sAddPersonTermButton, "Test"), 0, 4
    
    Case ECATALOGUEPROPERTIES.cpCopiesInformation
        
        DisplayLinkItemForEdit m_eSelectedFieldID, sHeading, m_colFirstSet, m_vUpdateValue1
    
        DisplayOneLineHeadingForContent ""
        DisplayOneLineEditForContent _
            m_ofwhtmlReference.SubmitImage(gk_sImagesPath & gk_sImagePickCorporateBody, mk_sAddCorpBodyTermButton, "Test"), 0, 4
    
    Case ECATALOGUEPROPERTIES.cpPublicationNote
        
        DisplayLinkItemForEdit m_eSelectedFieldID, sHeading, m_colFirstSet, m_vUpdateValue1, m_colSecondSet
        
        DisplayOneLineHeadingForContent ""
        DisplayOneLineEditForContent _
            m_ofwhtmlReference.SubmitImage(gk_sImagesPath & gk_sImagePickSFA, mk_sAddSFAButton, "Test") & _
            m_ofwhtmlReference.SubmitImage(gk_sIMAGEAddLeafletButton, mk_sAddLeafletButton, "Test"), 0, 4
    
    Case ECATALOGUEPROPERTIES.cpUnpublishedFindingAids
        
        DisplayLinkItemForEdit m_eSelectedFieldID, sHeading, m_colFirstSet, m_vUpdateValue1, m_colSecondSet
        
        DisplayOneLineHeadingForContent ""
        DisplayOneLineEditForContent _
            m_ofwhtmlReference.SubmitImage(gk_sImagesPath & gk_sImagePickSFA, mk_sAddSFAButton, "Test") & _
            m_ofwhtmlReference.SubmitImage(gk_sIMAGEAddLeafletButton, mk_sAddLeafletButton, "Test"), 0, 4
    
    End Select


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
' d8   ksk      27Oct1999       Created
' d15  tmj      15Mar2000       Added functionality to attach leaflets
' d26  alm      03Oct2001       Added support for creator, deposit and originals
On Error GoTo ErrorHandler

    Dim sSource As String
    Dim iCounter As Integer, bDisplayedContentAgainstHeading As Boolean
    Dim bHasOptionalDescription As Boolean, bSecondCollectionPresent As Boolean, bUnknownDescriptionPresent As Boolean
    
    'Objects
    Dim oVWSharedCatalogueEntry  As VWSharedCatalogueEntry
    
    'button names
    Dim sRemoveButtonName As String
    
    'field names
    Dim sUnknownDescriptionFieldName As String, sFieldNameForDescription As String, sHiddenIdFieldName As String
    
    'others
    Dim colDetails As Variant, sContent As String, vOptionalDescription As Variant, sHiddenIDFieldValue As String
    
    sSource = "VWCatalogueBulkupdate.DisplayLinkItemForEdit"
    
    'Display Headings
    DisplayOneLineHeadingForContent t_sHeading
    
    bDisplayedContentAgainstHeading = False         'Whether content is displayed against heading
    bHasOptionalDescription = False                 'Does the property has optional desc
    bSecondCollectionPresent = False                'Does the property has two recordlists
    bUnknownDescriptionPresent = False              'Does the property has unknown description
    
    'Does the property has unknown description
    If t_eProperty = ECATALOGUEPROPERTIES.cpImmediateSourceOfAcquisition Or _
        t_eProperty = ECATALOGUEPROPERTIES.cpCopiesInformation Or _
        t_eProperty = ECATALOGUEPROPERTIES.cpPublicationNote Or _
        t_eProperty = ECATALOGUEPROPERTIES.cpUnpublishedFindingAids Then
        
        bUnknownDescriptionPresent = True
    End If
    
    'Does the property has optional desc
    If t_eProperty = ECATALOGUEPROPERTIES.cpImmediateSourceOfAcquisition Or _
        t_eProperty = ECATALOGUEPROPERTIES.cpCopiesInformation Or _
        t_eProperty = ECATALOGUEPROPERTIES.cpLocationOfOriginals Or _
        t_eProperty = ECATALOGUEPROPERTIES.cpPublicationNote Or _
        t_eProperty = ECATALOGUEPROPERTIES.cpUnpublishedFindingAids Then
        
        bHasOptionalDescription = True
    End If
    
    'Check if the second collection passsed
    If Not (t_colSecondSet Is Nothing) Then
        bSecondCollectionPresent = True
    End If
    
    'If unknown description is present
    If bUnknownDescriptionPresent Then      'get the name of the field for that
        sUnknownDescriptionFieldName = mk_sNameValue1
        
        'Edit unknown unknown description
        DisplayOneLineEditForContent _
            m_ofwhtmlReference.InputField(ITText, sUnknownDescriptionFieldName, _
                mk_iLengthOfUnknownDescriptionTextField, gk_iDBCOLLenLinkItemUnknownDescription, _
                VariantStringAsString(t_vUnknownDescription), VCNone, False), 1, 3
        
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
                
                sContent = CorporateBodyDescriptionToDisplay(colDetails.lID)
                
            Case ECATALOGUEPROPERTIES.cpCopiesInformation
                
                sContent = CorporateBodyDescriptionToDisplay(colDetails.lID)
                
            Case ECATALOGUEPROPERTIES.cpPublicationNote
                
                sContent = SFADescriptionToDisplay(colDetails.lID)

            Case ECATALOGUEPROPERTIES.cpUnpublishedFindingAids
                
                sContent = SFADescriptionToDisplay(colDetails.lID)
                
            'd26
            Case ECATALOGUEPROPERTIES.cpCreatorNames
                
                sContent = CorporateBodyDescriptionToDisplay(colDetails.lID)
            
            'd26
            Case ECATALOGUEPROPERTIES.cpPlaceOfDeposit
                
                sContent = CorporateBodyDescriptionToDisplay(colDetails.lID)
                
            'd26
            Case ECATALOGUEPROPERTIES.cpLocationOfOriginals
                
                sContent = CorporateBodyDescriptionToDisplay(colDetails.lID)

        End Select
        
        vOptionalDescription = colDetails.sOptionalDescription
        sHiddenIdFieldName = mk_sNameHiddenIDFieldFirstLink & LTrim(Str(iCounter))
        sHiddenIDFieldValue = colDetails.lID
        sFieldNameForDescription = mk_sNameDescriptionOptionalFirstLink & iCounter
        sRemoveButtonName = mk_sNameRemoveFirstLinkButton & colDetails.lID
        
        'format the optional description
        vOptionalDescription = VariantStringAsString(vOptionalDescription)
        
        If bDisplayedContentAgainstHeading Then
            'Display with blank heading
            DisplayOneLineHeadingEditAndDisplayForContent _
                "", _
                sContent, _
                m_ofwhtmlReference.SubmitImage(gk_sImagesPath & gk_sImageRemove, sRemoveButtonName, sRemoveButtonName), _
                m_ofwhtmlReference.HiddenField(sHiddenIdFieldName, sHiddenIDFieldValue), _
                IIf(bHasOptionalDescription, _
                    m_ofwhtmlReference.InputField(ITText, sFieldNameForDescription, mk_iLengthOfUnknownDescriptionTextField, _
                                    gk_iDBCOLLenLinkItemUnknownDescription, CStr(vOptionalDescription), _
                                    EValidCheck.VCNone, False), _
                    "")
        Else
            'Display only the contents
            DisplayOneLineEditAndDisplayForContent _
                sContent, _
                m_ofwhtmlReference.SubmitImage(gk_sImagesPath & gk_sImageRemove, sRemoveButtonName, sRemoveButtonName), _
                m_ofwhtmlReference.HiddenField(sHiddenIdFieldName, sHiddenIDFieldValue), _
                IIf(bHasOptionalDescription, _
                    m_ofwhtmlReference.InputField(ITText, sFieldNameForDescription, mk_iLengthOfUnknownDescriptionTextField, _
                                        gk_iDBCOLLenLinkItemUnknownDescription, CStr(vOptionalDescription), _
                                        EValidCheck.VCNone, False), _
                    "")
                
            bDisplayedContentAgainstHeading = True
        End If
    
    Next
    
    'If a second recordlist is present
    If bSecondCollectionPresent Then
        iCounter = 0
        For Each colDetails In t_colSecondSet
            'Get the content to be displayed
            iCounter = iCounter + 1
            'Get the content, description and id fields for each of the record in the collection
            
            If t_eProperty = cpPublicationNote Or t_eProperty = cpUnpublishedFindingAids Then
                sContent = LeafletDescriptionToDisplay(colDetails.lID)
            Else
                sContent = PersonDescriptionToDisplay(colDetails.lID)
            End If
            
            vOptionalDescription = colDetails.sOptionalDescription
            sHiddenIdFieldName = mk_sNameHiddenIDFieldSecondLink & LTrim(Str(iCounter))
            sHiddenIDFieldValue = colDetails.lID
            sFieldNameForDescription = mk_sNameDescriptionOptionalSecondLink & iCounter
            sRemoveButtonName = mk_sNameRemoveSecondLinkButton & colDetails.lID
            
            'format the optional description
            vOptionalDescription = VariantStringAsString(vOptionalDescription)
                              
            If bDisplayedContentAgainstHeading Then
                'Display with blank heading
                DisplayOneLineHeadingEditAndDisplayForContent _
                    "", _
                    sContent, _
                    m_ofwhtmlReference.SubmitImage(gk_sImagesPath & gk_sImageRemove, sRemoveButtonName, sRemoveButtonName), _
                    m_ofwhtmlReference.HiddenField(sHiddenIdFieldName, sHiddenIDFieldValue), _
                    IIf(bHasOptionalDescription, _
                        m_ofwhtmlReference.InputField(ITText, sFieldNameForDescription, mk_iLengthOfUnknownDescriptionTextField, _
                                    gk_iDBCOLLenLinkItemUnknownDescription, CStr(vOptionalDescription), _
                                    EValidCheck.VCNone, False), _
                        "")
            Else
                'Display only the contents
                DisplayOneLineEditAndDisplayForContent _
                    sContent, _
                    m_ofwhtmlReference.SubmitImage(gk_sImagesPath & gk_sImageRemove, sRemoveButtonName, sRemoveButtonName), _
                    m_ofwhtmlReference.HiddenField(sHiddenIdFieldName, sHiddenIDFieldValue), _
                    IIf(bHasOptionalDescription, _
                        m_ofwhtmlReference.InputField(ITText, sFieldNameForDescription, mk_iLengthOfUnknownDescriptionTextField, _
                                    gk_iDBCOLLenLinkItemUnknownDescription, CStr(vOptionalDescription), _
                                    EValidCheck.VCNone, False), _
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

Public Function CancelBulkupdateAddTerm()
'Public method called when a cancell button is clicked on the index browser

' d8   ksk      27Oct1999       Created

    Dim sEditsetParam  As String
    'Get the current state of the business object
    'Reconstruct the parameters for displaying the bulkupdate screen and
    ' redirect the catalogue to the bulk update value display screen.
    
    m_oboCatalogueBulkUpdate.PROCATTransientInterface_GetState
    UpdateLocalObjectWithBusinessobjectValues True

    'Parameter to use for the regirection
    sEditsetParam = "?" & gk_sParamEditSetID & "=" & m_lEditsetID & "&" & _
                    mk_sNameFieldList & "=" & m_eSelectedFieldID & "&" & _
                    mk_sParamRestoreState & "=" & mk_sStateRestoreValue
    
    m_aspResponse.Redirect gk_sLinkDisplayBulkUpdateValue & sEditsetParam
    
End Function


Public Function AddIndexTermForBulkupdate()
'Public method called when a, Index term is selected from the Index browser
' d8   ksk      27Oct1999       Created
' d26  alm      03Oct2001       Added support for creator, deposit and originals
'
    Dim sSource As String
    Dim lIndextermID As Long
    Dim sEditsetParam As String
    Dim eTermType As EAFTermType
    Dim typLinkData As GenericLinkRecord
    Dim bIsInCollection As Boolean
    
    sSource = "VWCatalogueBulkUpdate.AddIndexTermForBulkupdate"
    
    'Parse out the values selected
    lIndextermID = m_aspRequest(gk_sFORMAFTermID)
    eTermType = m_aspRequest(gk_sFORMAFListTermType)
    
    'Bring the business object to the required state
    m_oboCatalogueBulkUpdate.PROCATTransientInterface_GetState
    UpdateLocalObjectWithBusinessobjectValues True
    
    If eTermType = AFPerson Then
        bIsInCollection = IsIDPresentInCollection(lIndextermID, m_colSecondSet)
    Else
        bIsInCollection = IsIDPresentInCollection(lIndextermID, m_colFirstSet)
    End If
    
    If Not bIsInCollection Then
        
        
        Select Case m_eSelectedFieldID
        Case ECATALOGUEPROPERTIES.cpImmediateSourceOfAcquisition, _
             ECATALOGUEPROPERTIES.cpCreatorNames
        
            If eTermType = AFPerson Then
                typLinkData.lID = lIndextermID
                m_colSecondSet.Add typLinkData, gk_sPrefixForLinkTerm & CStr(lIndextermID)
            Else
                typLinkData.lID = lIndextermID
                m_colFirstSet.Add typLinkData, gk_sPrefixForLinkTerm & CStr(lIndextermID)
            End If
        
        Case ECATALOGUEPROPERTIES.cpCopiesInformation, _
             ECATALOGUEPROPERTIES.cpPlaceOfDeposit, _
             ECATALOGUEPROPERTIES.cpLocationOfOriginals
            
                typLinkData.lID = lIndextermID
                m_colFirstSet.Add typLinkData, gk_sPrefixForLinkTerm & CStr(lIndextermID)
        
        Case Else
            RaiseError sSource, EPROCATINTERNALERROR.peBulkUpdate, _
                    "Invalid Term type for adding into bulk update"
        End Select
        
        'Update the business object and save the state
        UpdateBusinessObjectWithParsedValues True
        m_oboCatalogueBulkUpdate.PROCATTransientInterface_SetState
    
    End If
    
    'Redirect to the display page
    sEditsetParam = "?" & gk_sParamEditSetID & "=" & m_lEditsetID & "&" & _
                    mk_sNameFieldList & "=" & m_eSelectedFieldID & "&" & _
                    mk_sParamRestoreState & "=" & mk_sStateRestoreValue
    
    m_aspResponse.Redirect gk_sLinkDisplayBulkUpdateValue & sEditsetParam
    

End Function


Public Function AddSFAForBulkupdate()
'Public method called when a SFA is selected from the SFA browser
' d8   ksk      27Oct1999       Created
' d17  tmj      25Mar2000       Moved checking if in collection so it works

    
    Dim sSource As String
    Dim lSFATermID As Long
    Dim typLinkData As GenericLinkRecord
    Dim sEditsetParam As String
    Dim bIsInCollection As Boolean
    
    sSource = "VWCatalogueBulkUpdate.AddIndexTermForBulkupdate"
    
    'Parse out the values selected
    lSFATermID = m_aspRequest(gk_sFORMSFADetailsID)
    bIsInCollection = IsIDPresentInCollection(lSFATermID, m_colFirstSet)
    
    'Bring the business object to the required state
    m_oboCatalogueBulkUpdate.PROCATTransientInterface_GetState
    UpdateLocalObjectWithBusinessobjectValues True
    
    bIsInCollection = IsIDPresentInCollection(lSFATermID, m_colFirstSet)
    
    If Not bIsInCollection Then
        
        Select Case m_oboCatalogueBulkUpdate.BulkUpdateField
        
        Case ECATALOGUEPROPERTIES.cpPublicationNote
        
            typLinkData.lID = lSFATermID
            m_colFirstSet.Add typLinkData, gk_sPrefixForLinkTerm & CStr(lSFATermID)
        
        Case ECATALOGUEPROPERTIES.cpUnpublishedFindingAids
            
            typLinkData.lID = lSFATermID
            m_colFirstSet.Add typLinkData, gk_sPrefixForLinkTerm & CStr(lSFATermID)
        
        Case Else
            RaiseError sSource, EPROCATINTERNALERROR.peBulkUpdate, _
                    "Invalid SFA Term type for adding into bulk update"
        
        End Select
                
        
        'Update the business object and save the state
        UpdateBusinessObjectWithParsedValues True
        m_oboCatalogueBulkUpdate.PROCATTransientInterface_SetState
        
    End If
    
    'Redirect to the display page
    sEditsetParam = "?" & gk_sParamEditSetID & "=" & m_lEditsetID & "&" & _
                    mk_sNameFieldList & "=" & m_eSelectedFieldID & "&" & _
                    mk_sParamRestoreState & "=" & mk_sStateRestoreValue
    
    m_aspResponse.Redirect gk_sLinkDisplayBulkUpdateValue & sEditsetParam


End Function

Public Function AddLeafletForBulkupdate()
'Public method called when a Leaflet is selected from the Information Leaflet browser
' d15   tmj      14Mar2000       Created, by copying from AddSFAForBulkupdate
' d17   tmj      25Mar2000       Moved checking if in collection so it works

    
    Dim sSource As String
    Dim lLeafletID As Long
    Dim typLinkData As GenericLinkRecord
    Dim sEditsetParam As String
    Dim bIsInCollection As Boolean
    
    sSource = "VWCatalogueBulkUpdate.AddLeafletForBulkupdate"
    
    'Parse out the values selected
    lLeafletID = m_aspRequest("sLeafletID")
    
    'Bring the business object to the required state
    m_oboCatalogueBulkUpdate.PROCATTransientInterface_GetState
    UpdateLocalObjectWithBusinessobjectValues True
    
    bIsInCollection = IsIDPresentInCollection(lLeafletID, m_colSecondSet)
    
    If Not bIsInCollection Then
        
        Select Case m_oboCatalogueBulkUpdate.BulkUpdateField
        
        Case ECATALOGUEPROPERTIES.cpPublicationNote
        
            typLinkData.lID = lLeafletID
            m_colSecondSet.Add typLinkData, gk_sPrefixForLinkTerm & CStr(lLeafletID)
        
        Case ECATALOGUEPROPERTIES.cpUnpublishedFindingAids
            
            typLinkData.lID = lLeafletID
            m_colSecondSet.Add typLinkData, gk_sPrefixForLinkTerm & CStr(lLeafletID)
        
        Case Else
            RaiseError sSource, EPROCATINTERNALERROR.peBulkUpdate, _
                    "Invalid SFA Term type for adding into bulk update"
        
        End Select
                
        
        'Update the business object and save the state
        UpdateBusinessObjectWithParsedValues True
        m_oboCatalogueBulkUpdate.PROCATTransientInterface_SetState
        
    End If
    
    'Redirect to the display page
    sEditsetParam = "?" & gk_sParamEditSetID & "=" & m_lEditsetID & "&" & _
                    mk_sNameFieldList & "=" & m_eSelectedFieldID & "&" & _
                    mk_sParamRestoreState & "=" & mk_sStateRestoreValue
    
    m_aspResponse.Redirect gk_sLinkDisplayBulkUpdateValue & sEditsetParam


End Function

Public Sub BulkUpdateValue()
'Displays the form to select the field to be edited
'
' d2    ksk     08Jul1999
' d3   ksk      09Jul1999         Added tooltip help for all hyper links
' d7   ksk      17Aug1999         changes made to display the intermediate page
'                                   reflecting the changed fields
' d11  ksk      07Dec1999       Changes to display the EAD browser for select fields
' d12  ksk      14Dec1999       Added code to convert the apostrophe and quote in the ead field
' d15  tmj      15Mar2000       Added clauses to handle attaching Leaflets
' d26  alm      03Oct2001       Added support for creator, deposit and originals
'
On Error GoTo ErrHandler

    Dim sSource As String
    Dim i As Integer
    Dim sEditsetParam As String
    Dim colCatalogueList As Collection
    Dim iButtonClicked As Integer
    Dim lIDToRemove As Long

    sSource = "VWBulkUpdate.BulkUpdateValue"
    
    'Check if the Cancel button was clicked if so return to the main page
    If m_aspRequest(mk_sNameButtonCancel & gk_sXCoordinate) <> 0 Or _
        m_aspRequest(mk_sNameButtonCancel & gk_sYCoordinate) <> 0 Then
        'Code to return to the main page to be added here
        m_aspResponse.Redirect (gk_sLinkDisplayEditSetDetails)
        Exit Sub
    End If
    
    
    'Makesure that the editset ID was passes to this page
    If m_aspRequest(gk_sParamEditSetID) = "" Then
        ' d2    changed the enumeration
        RaiseError sSource, EPROCATINTERNALERROR.peEditSetIDInvalue, _
                    "Editset ID not found"
    End If
    m_lEditsetID = CLng(m_aspRequest(gk_sParamEditSetID))
    If m_lEditsetID <= 0 Then
        ' d2    changed the enumeration
        RaiseError sSource, EPROCATINTERNALERROR.peEditSetIDInvalue, _
                "Editset ID not found"
    End If
    sEditsetParam = "?" & gk_sParamEditSetID & "=" & m_lEditsetID
    
    CheckPermissions ESecurityCheck.EUserCanBulkUpdate, sSource, ""
    
    For i = 1 To gk_iNumberOfLevels
        'Fill an array of selected levels with the values of the checked Level check boxes
        m_bSelectedLevels(i) = IIf(m_aspRequest(mk_sNameLevel & Trim(Str(i))) <> "", True, False)
    Next i
    
    
    'Read the values fron the Request object
    'If no field is selected then return to the field selection process
    If m_aspRequest(mk_sNameFieldList) = "" Then
        DisplayBulkUpdateFieldPage
        Exit Sub
    End If
    
    m_eSelectedFieldID = CInt(m_aspRequest(mk_sNameFieldList))
    
    'Parse For all the values
    ParseTheButtonClicked iButtonClicked, lIDToRemove
    
    ' d26 Get the value(s), and convert to booleans if a checkbox
    Select Case m_eSelectedFieldID
            
        Case ECATALOGUEPROPERTIES.cpAccessionedAsCopy, _
             ECATALOGUEPROPERTIES.cpOrderable, _
             ECATALOGUEPROPERTIES.cpAR, _
             ECATALOGUEPROPERTIES.cpExtractAtTransfer
             
             m_vUpdateValue1 = IIf(m_aspRequest(mk_sNameValue1) <> "", True, False)
             
        Case Else
            m_vUpdateValue1 = m_aspRequest(mk_sNameValue1)
            m_vUpdateValue2 = m_aspRequest(mk_sNameValue2)
            
    End Select
        
    Select Case m_eSelectedFieldID
    
    ' d21 Added to auto-generate start / end dates
    Case ECATALOGUEPROPERTIES.cpCoveringDates
        
        ' Update the start and end dates
        UpdateStartAndEndDates m_vUpdateValue1
    
    ' d11  Added to save the ead field
    Case ECATALOGUEPROPERTIES.cpAdministrative_BiographicalHistory, _
        ECATALOGUEPROPERTIES.cpArrangement, _
        ECATALOGUEPROPERTIES.cpMapDesignation, _
        ECATALOGUEPROPERTIES.cpScopeAndContent, _
        ECATALOGUEPROPERTIES.cpTitle
                
        'd12
        'Replace HTML decimal code for apostrophe with the apostrophe character
        m_vUpdateValue1 = Replace(m_vUpdateValue1, gk_sHTMLApostropheDecimalCode, "'")
        'Replace HTML decimal code for quotes with the quotes character
        m_vUpdateValue1 = Replace(m_vUpdateValue1, gk_sHTMLQuotationDecimalCode, """")
        
        BulkUpdateEADField sEditsetParam, iButtonClicked

        Exit Sub
    
    Case ECATALOGUEPROPERTIES.cpImmediateSourceOfAcquisition, _
            ECATALOGUEPROPERTIES.cpCopiesInformation, _
            ECATALOGUEPROPERTIES.cpPublicationNote, _
            ECATALOGUEPROPERTIES.cpUnpublishedFindingAids, _
            ECATALOGUEPROPERTIES.cpCreatorNames, _
            ECATALOGUEPROPERTIES.cpPlaceOfDeposit, _
            ECATALOGUEPROPERTIES.cpLocationOfOriginals
        
        If m_eSelectedFieldID = ECATALOGUEPROPERTIES.cpImmediateSourceOfAcquisition Or _
           m_eSelectedFieldID = ECATALOGUEPROPERTIES.cpPublicationNote Or _
           m_eSelectedFieldID = ECATALOGUEPROPERTIES.cpUnpublishedFindingAids Or _
           m_eSelectedFieldID = ECATALOGUEPROPERTIES.cpCreatorNames Then
            m_bBothTheCollectionsRequired = True
        Else
            m_bBothTheCollectionsRequired = False
        End If
        
        BulkupdateLinkDataTerms sEditsetParam, iButtonClicked, lIDToRemove
        
        Exit Sub
        
    End Select
    
    
    UpdateBusinessObjectWithParsedValues False
    
    Set colCatalogueList = m_oboCatalogueBulkUpdate.BulkUpdateLevels(m_bSelectedLevels())
    
    m_aspResponse.Write GetUpdatedCatalogueDisplayPage(m_lEditsetID, m_eSelectedFieldID, _
                                                    colCatalogueList)
    
    m_objContext.SetComplete
    
    Exit Sub
    
ErrHandler:
    m_objContext.SetAbort
    ShowError sSource
End Sub



Private Sub ParseTheButtonClicked(ByRef t_iButtonTypeClicked As Integer, _
                                ByRef t_lRemoveIDValue As Long)
'Function to parse the inpot to see which button was clicked and sets the value corresponding
' to the button clicked a value in t_iButtonTypeClicked and if the remove button was clicked then
' the remove buttons corresponding ID (stored along with the name) is returned in t_lRemoveIDValue
' The values returned in t_iButtonTypeClicked  will be
    '1 For Remove from first collection set,
    '2 For Remove from second collection set
    '3 For Add corporatebody,
    '4 For Add Person,
    '5 For Add SFA,
    '6 For OK,
    '7 For Cancel
    '8 For Update for Updating the edited XML,
    '9 For Add Leaflet

    Dim sQueryString As String, sTemp As String
    Dim iStartValue As Integer, iEndValue As Integer
    
    
    sQueryString = m_aspRequest.Form
    
    
    'Check the clicking of the Remove link in the first set
    iStartValue = InStr(sQueryString, mk_sNameRemoveFirstLinkButton)
    If iStartValue > 0 Then
        iEndValue = InStr(iStartValue, sQueryString, "&")
    Else
        iEndValue = Len(sQueryString)
    End If
    If iStartValue > 0 Then
        
        t_iButtonTypeClicked = 1
        'The ID is part of the remove first link button name so seperate the ID from the button
        t_lRemoveIDValue = GetNumberStoredInButtonName(sQueryString, _
                                            mk_sNameRemoveFirstLinkButton)
        Exit Sub
    End If
        
    'Check the clicking of the Remove link in the second set
    iStartValue = InStr(sQueryString, mk_sNameRemoveSecondLinkButton)
    If iStartValue > 0 Then
        iEndValue = InStr(iStartValue, sQueryString, "&")
    Else
        iEndValue = Len(sQueryString)
    End If
    If iStartValue > 0 Then
        
        t_iButtonTypeClicked = 2
        'The ID is part of the remove first link button name so seperate the ID from the button
        t_lRemoveIDValue = GetNumberStoredInButtonName(sQueryString, _
                                            mk_sNameRemoveSecondLinkButton)
        Exit Sub
    End If

    'Check the clicking of the Add corpbody term button
    iStartValue = InStr(sQueryString, mk_sAddCorpBodyTermButton)
    If iStartValue > 0 Then
        
        t_iButtonTypeClicked = 3
        Exit Sub
    
    End If

    'Check the clicking of the Add Person term button
    iStartValue = InStr(sQueryString, mk_sAddPersonTermButton)
    If iStartValue > 0 Then
        
        t_iButtonTypeClicked = 4
        Exit Sub
    
    End If


    'Check the clicking of the Add SFA button
    iStartValue = InStr(sQueryString, mk_sAddSFAButton)
    If iStartValue > 0 Then
        
        t_iButtonTypeClicked = 5
        Exit Sub
    
    End If

    'Check the clicking of the OK button clicked
    iStartValue = InStr(sQueryString, mk_sNameButtonOk)
    If iStartValue > 0 Then
        
        t_iButtonTypeClicked = 6
        Exit Sub
    
    End If

    'Check the clicking of the Cancel button clicked
    iStartValue = InStr(sQueryString, mk_sNameButtonCancel)
    If iStartValue > 0 Then
        
        t_iButtonTypeClicked = 7
    
    End If


    'If the Hidden status flag field is set then called from update
    If Not IsEmpty(m_aspRequest(mk_sNameHiddenStatusField)) Then
        
        t_iButtonTypeClicked = 8
    
    End If
    
    'Check the clicking of the Add Leaflet button
    iStartValue = InStr(sQueryString, mk_sAddLeafletButton)
    If iStartValue > 0 Then
        
        t_iButtonTypeClicked = 9
        Exit Sub
    
    End If

End Sub

Private Function GetNumberStoredInButtonName(t_sData As String, _
                                            t_sFieldName As String) As Long
'This method searchs for t_sFieldName in t_sData  and gets the number that is present after t_sFieldName
'
' d8   ksk      27Oct1999       Created

    Dim sTemp As String
    Dim iStartValue As Integer, iEndValue As Integer
    
    'Check the clicking of the Remove link in the first set
    iStartValue = InStr(t_sData, t_sFieldName)
    If iStartValue > 0 Then
        iEndValue = InStr(iStartValue, t_sData, "&")
    Else
        iEndValue = Len(t_sData)
    End If

    'The ID is part of the remove first link button name so seperate the ID from the button
    sTemp = Mid(t_sData, iStartValue, iEndValue - iStartValue)
    iStartValue = InStr(sTemp, t_sFieldName) + Len(t_sFieldName)
    iEndValue = InStr(sTemp, ".")
    GetNumberStoredInButtonName = CLng(Mid(sTemp, iStartValue, iEndValue - iStartValue))

End Function

Private Sub BulkUpdateEADField(ByVal t_sParameter As String, ByVal t_iButtonTypeClicked As Integer)
'This routien depending on the button pressed (t_iButtonTypeClicked) does the following action
'   if the EAD change update was pressed - then transiently persists the data and re renders the XML
'   if the Ok button was pressed - then the system will save the details to the database and display the
'                                   records affected page
' The parameter (t_sParameter ) is used to redirect to display the value page
' d11  ksk      07Dec1999       Changes to display the EAD browser for select fields
'
On Error GoTo ErrHandler

    Dim sSource As String, sPageToRedirect  As String, sParameterForRedirectedPage As String
    Dim colCatalogueList As Collection

    sSource = "VWCatalogueBulkUpdate.BulkUpdateEADField"

    If t_iButtonTypeClicked = 8 Then
        StoreFormLinkItemsIntoCollections
        UpdateBusinessObjectWithParsedValues True
        
        'm_colFirstSet.Remove gk_sPrefixForLinkTerm & CStr(t_lRemoveIDValue)
        'Save the state of the Business object
        m_oboCatalogueBulkUpdate.PROCATTransientInterface_SetState
        sPageToRedirect = gk_sLinkDisplayBulkUpdateValue
        sParameterForRedirectedPage = t_sParameter & "&" & _
                    mk_sNameFieldList & "=" & m_eSelectedFieldID & "&" & _
                    mk_sParamRestoreState & "=" & mk_sStateRestoreValue
        
        m_aspResponse.Redirect sPageToRedirect & sParameterForRedirectedPage
        Exit Sub
    End If
    
    UpdateBusinessObjectWithParsedValues True
    'Bulkupdate the levels selected
    Set colCatalogueList = m_oboCatalogueBulkUpdate.BulkUpdateLevels(m_bSelectedLevels())
    
    'Remove the stae information from the store before exiting the screen
    m_oboCatalogueBulkUpdate.PROCATTransientInterface_DeleteState
    
    'Display the records affected
    m_aspResponse.Write GetUpdatedCatalogueDisplayPage(m_lEditsetID, m_eSelectedFieldID, colCatalogueList)

    m_objContext.SetComplete
    
    Exit Sub
ErrHandler:
    m_objContext.SetAbort
    ShowError sSource
End Sub

Private Sub BulkupdateLinkDataTerms(ByVal t_sParameter As String, ByVal t_iButtonTypeClicked As Integer, _
                                ByVal t_lRemoveIDValue As Long)
'Method to update and handle request for bulkupdate of link deta terms
' d8   ksk      27Oct1999       Created
' d15  tmj      15Mar2000       Added clauses to handle attaching leaflets
On Error GoTo ErrHandler

    Dim sSource As String
    Dim i As Integer
    Dim sEditsetParam As String
    Dim colCatalogueList As Collection
    
    Dim sPageToRedirect As String
    Dim sParameterForRedirectedPage As String
    
    
    sSource = "VWCatalogueBulkUpdate.BulkUpdateLinkDataTerms"
    
    Select Case t_iButtonTypeClicked
    
    '1 For Remove from first collection set,
    Case 1
        
        StoreFormLinkItemsIntoCollections
        UpdateBusinessObjectWithParsedValues True
        
        m_colFirstSet.Remove gk_sPrefixForLinkTerm & CStr(t_lRemoveIDValue)
        'Save the state of the Business object
        m_oboCatalogueBulkUpdate.PROCATTransientInterface_SetState
        sPageToRedirect = gk_sLinkDisplayBulkUpdateValue
        sParameterForRedirectedPage = t_sParameter & "&" & _
                    mk_sNameFieldList & "=" & m_eSelectedFieldID & "&" & _
                    mk_sParamRestoreState & "=" & mk_sStateRestoreValue
    
    '2 For Remove from second collection set
    Case 2
    
        StoreFormLinkItemsIntoCollections
        UpdateBusinessObjectWithParsedValues True
        
        m_colSecondSet.Remove gk_sPrefixForLinkTerm & CStr(t_lRemoveIDValue)
        'Save the state of the Business object
        m_oboCatalogueBulkUpdate.PROCATTransientInterface_SetState
        sPageToRedirect = gk_sLinkDisplayBulkUpdateValue
        sParameterForRedirectedPage = t_sParameter & "&" & _
                    mk_sNameFieldList & "=" & m_eSelectedFieldID & "&" & _
                    mk_sParamRestoreState & "=" & mk_sStateRestoreValue
    
    '3 For Add corporatebody,
    Case 3
    
        StoreFormLinkItemsIntoCollections
        UpdateBusinessObjectWithParsedValues True
        
        'Save the state of the Business object
        m_oboCatalogueBulkUpdate.PROCATTransientInterface_SetState
        sPageToRedirect = gk_sPAGEAFList
        sParameterForRedirectedPage = t_sParameter & _
            "&" & gk_sFORMAFGroup & "=" & EAFTermType.AFCorporateBody & _
            "&" & gk_sFORMAFListIndex & "=" & "0" & _
            "&" & gk_sFORMAFTermID & "=" & "1" & _
            "&" & gk_sFORMAFListType & "=" & "1" & _
            "&" & gk_sFORMAFListTermType & "=" & "1" & _
            "&" & gk_sFORMAFListUsage & "=" & EAFBrowserUsage.BROWSEEditorialBulkupdatePickTerm

    
    '4 For Add Person,
    Case 4
    
        
        StoreFormLinkItemsIntoCollections
        UpdateBusinessObjectWithParsedValues True
        
        'Save the state of the Business object
        m_oboCatalogueBulkUpdate.PROCATTransientInterface_SetState
        sPageToRedirect = gk_sPAGEAFList
        sParameterForRedirectedPage = t_sParameter & _
            "&" & gk_sFORMAFGroup & "=" & EAFTermType.AFPerson & _
            "&" & gk_sFORMAFListIndex & "=" & "0" & _
            "&" & gk_sFORMAFTermID & "=" & "1" & _
            "&" & gk_sFORMAFListType & "=" & "1" & _
            "&" & gk_sFORMAFListTermType & "=" & "1" & _
            "&" & gk_sFORMAFListUsage & "=" & EAFBrowserUsage.BROWSEEditorialBulkupdatePickTerm
    
    
    '5 For Add SFA,
    Case 5
        
        StoreFormLinkItemsIntoCollections
        UpdateBusinessObjectWithParsedValues True
        
        'Save the state of the Business object
        m_oboCatalogueBulkUpdate.PROCATTransientInterface_SetState
        '+ to be changed to SFA browser
        sPageToRedirect = gk_sPAGESFAList
        sParameterForRedirectedPage = t_sParameter & _
            "&" & gk_sFORMSFAListStartIndex & "=0" & _
            "&" & gk_sFORMSFAListUsage & "=" & ESFABrowserUsage.ESFAEditorialBulkupdatePickTerm & _
            "&" & gk_sFORMSFAListFilter & "=" & ESFAFilter.ESFAElectronicAndPrinted

    
    '6 For OK,
    Case 6
    
        StoreFormLinkItemsIntoCollections
        UpdateBusinessObjectWithParsedValues True
        
    '7 for Cancel
    Case 7
    
        'Remove the stae information from the store before exiting the screen
        m_oboCatalogueBulkUpdate.PROCATTransientInterface_DeleteState
        sPageToRedirect = gk_sLinkDisplayEditSetDetails
        sParameterForRedirectedPage = t_sParameter
    
    '9 For Add Leaflet,
    Case 9
        
        StoreFormLinkItemsIntoCollections
        UpdateBusinessObjectWithParsedValues True
        
        'Save the state of the Business object
        m_oboCatalogueBulkUpdate.PROCATTransientInterface_SetState
        sPageToRedirect = gk_sInformationLeafletBrowser
        sParameterForRedirectedPage = t_sParameter & "&" & gk_sFORMSFAListUsage & _
                                      "=" & ELeafletBrowserUsage.ELeafletBulkUpdate
    
    End Select
    
    If t_iButtonTypeClicked <> 6 Then
    
        m_aspResponse.Redirect sPageToRedirect & sParameterForRedirectedPage
        Exit Sub
    End If
    
    
    Set colCatalogueList = m_oboCatalogueBulkUpdate.BulkUpdateLevels(m_bSelectedLevels())
    
    'Remove the stae information from the store before exiting the screen
    m_oboCatalogueBulkUpdate.PROCATTransientInterface_DeleteState
    
    m_aspResponse.Write GetUpdatedCatalogueDisplayPage(m_lEditsetID, m_eSelectedFieldID, colCatalogueList)
    
    
    m_objContext.SetComplete
    
    Exit Sub
    
ErrHandler:
    m_objContext.SetAbort
    ShowError sSource
End Sub


Private Function GetUpdatedCatalogueDisplayPage(t_lEditSetID As Long, _
                                            t_eFieldUpdated As ECATALOGUEPROPERTIES, _
                                            t_colCataloguesAffected As Collection) As String
'This method returns the html to be displayed that list all the catalogues that were
'affected by the bulk update
'Parameter:
'   t_lEditsetID - The editset Id for which bulk update was done
'   t_ofwCataloguesAffected - holds the list of catalogues affected.
' d7   ksk      17Aug1999         changes made to display the intermediate page
'                                   reflecting the changed fields
' d16  tmj      22Mar2000       Changed so displayed AF term that are attached during bulk update
' d26  alm      02Oct2001       Added support for creator, deposit and originals
    Dim sEditsetParam  As String
    Dim typCellStyle As CellStyle
    Dim ofwRecord  As FWRecord
    Dim sValue As String
    Dim var As Variant
    
    If t_colCataloguesAffected Is Nothing Then
        
        GetUpdatedCatalogueDisplayPage = m_ofwhtmlReference.Paragraph( _
                                    m_ofwhtmlReference.Bold(m_oFWAppData.GetDisplayInformation _
                                        (EEditorialResourceFile.erfBulkEditNoCatalogueUpdated)))
                                        
    Else
        If t_colCataloguesAffected.Count = 0 Then
            
            GetUpdatedCatalogueDisplayPage = m_ofwhtmlReference.Bold(m_oFWAppData.GetDisplayInformation _
                                            (EEditorialResourceFile.erfBulkEditNoCatalogueUpdated))
        Else
            
            Select Case t_eFieldUpdated
                Case ECATALOGUEPROPERTIES.cpPhysicalDescription_Form
                    sValue = m_ofwRefData.Search(EGroup.gPhysicalForm, m_vUpdateValue1)
                Case ECATALOGUEPROPERTIES.cpAccessCondition
                    sValue = m_ofwRefData.Search(EGroup.gAccessCondition, m_vUpdateValue1)
                Case ECATALOGUEPROPERTIES.cpClosureStatus
                    sValue = m_ofwRefData.Search(EGroup.gClosureStatus, m_vUpdateValue1)
                Case ECATALOGUEPROPERTIES.cpLegalStatus
                    sValue = m_ofwRefData.Search(EGroup.gLegalStatus, m_vUpdateValue1)
                Case ECATALOGUEPROPERTIES.cpLanguage
                    sValue = m_ofwRefData.Search(EGroup.gLanguage, m_vUpdateValue1)
                Case ECATALOGUEPROPERTIES.cpClosureCode, ECATALOGUEPROPERTIES.cpClosureType
                    sValue = m_ofwRefData.Search(EGroup.gClosureType, m_vUpdateValue1)
                Case Else
                    If m_vUpdateValue1 = "" Then
                        sValue = "NULL"
                    Else
                        sValue = m_vUpdateValue1
                    End If
            End Select
                
            If t_eFieldUpdated = ECATALOGUEPROPERTIES.cpClosureCode Or _
                    t_eFieldUpdated = ECATALOGUEPROPERTIES.cpClosureType Then
            
                GetUpdatedCatalogueDisplayPage = m_ofwhtmlReference.Bold( _
                                                m_oBOCatalogueElementMatrix.DescriptionFromEnumeration _
                                                (ECATALOGUEPROPERTIES.cpClosureType)) & " " & _
                                                m_oFWAppData.GetDisplayInformation _
                                                (EEditorialResourceFile.erfBulkEditUpdatedCatalogueList) & _
                                                " " & m_ofwhtmlReference.Bold(sValue)
                sValue = NullToSQLNull(StringToNullOrLong(m_vUpdateValue2))
                GetUpdatedCatalogueDisplayPage = GetUpdatedCatalogueDisplayPage & " and " & _
                                                m_ofwhtmlReference.Bold( _
                                                m_oBOCatalogueElementMatrix.DescriptionFromEnumeration _
                                                (ECATALOGUEPROPERTIES.cpClosureCode)) & " as " & _
                                                m_ofwhtmlReference.Bold(sValue)
                                                
            Else
            
                GetUpdatedCatalogueDisplayPage = m_ofwhtmlReference.Bold( _
                                                m_oBOCatalogueElementMatrix.DescriptionFromEnumeration _
                                                (t_eFieldUpdated)) & " " & m_oFWAppData.GetDisplayInformation _
                                                (EEditorialResourceFile.erfBulkEditUpdatedCatalogueList) & _
                                                " " & m_ofwhtmlReference.Bold(sValue)
                                                
            End If
                        
            GetUpdatedCatalogueDisplayPage = m_ofwhtmlReference.Paragraph(GetUpdatedCatalogueDisplayPage)
            
            'd16 Added this If clause to display link term added to the entries when bulk updating
            'selecting fields where objects are attached
            If t_eFieldUpdated = ECATALOGUEPROPERTIES.cpImmediateSourceOfAcquisition Or _
               t_eFieldUpdated = ECATALOGUEPROPERTIES.cpCopiesInformation Or _
               t_eFieldUpdated = ECATALOGUEPROPERTIES.cpPublicationNote Or _
               t_eFieldUpdated = ECATALOGUEPROPERTIES.cpUnpublishedFindingAids Or _
               t_eFieldUpdated = ECATALOGUEPROPERTIES.cpCreatorNames Or _
               t_eFieldUpdated = ECATALOGUEPROPERTIES.cpPlaceOfDeposit Or _
               t_eFieldUpdated = ECATALOGUEPROPERTIES.cpLocationOfOriginals Then

                'Ensuring there is something to attach
                If Not ((m_colFirstSet.Count = 0 Or m_colFirstSet Is Nothing) _
                    And (m_colSecondSet.Count = 0 Or m_colFirstSet Is Nothing)) Then

                    typCellStyle = m_ofwhtmlReference.GetCellStyle(ECellStyle.CELLColumnHeadingStyle)
                    
                    'adding the table headers
                    m_oFWTable.StartRow
                    m_oFWTable.AddHCellAsHTML "AF Term", typCellStyle
                    
                    'd26 Display an Optional Description header if they are supported
                    If Not (t_eFieldUpdated = ECATALOGUEPROPERTIES.cpCreatorNames Or _
                            t_eFieldUpdated = ECATALOGUEPROPERTIES.cpPlaceOfDeposit) Then
                                                    
                        m_oFWTable.AddHCellAsHTML "&nbsp;&nbsp;" + "Optional Description", typCellStyle
                        
                    End If
                    
                    m_oFWTable.EndRow

                    typCellStyle = m_ofwhtmlReference.GetCellStyle(ECellStyle.CELLReadOnlyDataStyle)
                    
                    'Cycling through the collection and displaying object of type
                    'GenericLinkRecord
                    For Each var In m_colFirstSet
                        m_oFWTable.StartRow
                        Select Case t_eFieldUpdated
                            
                            'd26 - Added cpLocationOfOriginals
                            Case cpImmediateSourceOfAcquisition, _
                                 cpCopiesInformation, _
                                 cpLocationOfOriginals
                                m_oFWTable.AddCellAsHTML _
                                    CorporateBodyDescriptionToDisplay(var.lID), typCellStyle
                                m_oFWTable.AddCellAsHTML "&nbsp;&nbsp;" + var.sOptionalDescription, typCellStyle
                            Case cpPublicationNote, _
                                 cpUnpublishedFindingAids
                                m_oFWTable.AddCellAsHTML _
                                    SFADescriptionToDisplay(var.lID), typCellStyle
                                m_oFWTable.AddCellAsHTML "&nbsp;&nbsp;" + var.sOptionalDescription, typCellStyle
                            
                            'd26
                            Case cpCreatorNames, cpPlaceOfDeposit
                                m_oFWTable.AddCellAsHTML _
                                    CorporateBodyDescriptionToDisplay(var.lID), typCellStyle
                        End Select
                    Next

                    'If 2 collections used, (2 types of object attached)
                    If m_bBothTheCollectionsRequired Then

                        For Each var In m_colSecondSet
                            m_oFWTable.StartRow
                            Select Case t_eFieldUpdated
                                Case cpImmediateSourceOfAcquisition
                                    m_oFWTable.AddCellAsHTML _
                                        PersonDescriptionToDisplay(var.lID), typCellStyle
                                    m_oFWTable.AddCellAsHTML "&nbsp;&nbsp;" + var.sOptionalDescription, typCellStyle
                                Case cpPublicationNote, _
                                     cpUnpublishedFindingAids
                                    m_oFWTable.AddCellAsHTML _
                                        LeafletDescriptionToDisplay(var.lID), typCellStyle
                                    m_oFWTable.AddCellAsHTML "&nbsp;&nbsp;" + var.sOptionalDescription, typCellStyle
                                    
                                'd26
                                Case cpCreatorNames
                                    m_oFWTable.AddCellAsHTML _
                                        PersonDescriptionToDisplay(var.lID), typCellStyle
                            End Select
                            m_oFWTable.EndRow
                        Next

                    End If

                    'Adding HTML to output
                    GetUpdatedCatalogueDisplayPage = GetUpdatedCatalogueDisplayPage & _
                                                    m_ofwhtmlReference.Paragraph(m_oFWTable.GetHTML)
                    'Setting fresh table, as the page contains a second
                    Set m_oFWTable = New FWTable
                    
                End If
            
            End If
            
            'Display the heading for the list of catalogues affected by the bulk update
            typCellStyle = m_ofwhtmlReference.GetCellStyle(ECellStyle.CELLColumnHeadingStyle)
            m_oFWTable.AddHCellAsHTML m_ofwhtmlReference.FontWithStyle(m_oFWAppData.GetDisplayInformation( _
                                        ESharedResourceFile.srfCatalogueReference), _
                                        m_ofwhtmlReference.GetFontStyle(EFontStyle.FONTColumnHeadingStyle)) _
                                        , typCellStyle
            
            'Show the list of catalogue references affected in the update
            For Each ofwRecord In t_colCataloguesAffected
                m_oFWTable.StartRow
                
                typCellStyle = m_ofwhtmlReference.GetCellStyle(ECellStyle.CELLReadOnlyDataStyle)
                m_oFWTable.AddCellAsHTML m_ofwhtmlReference.FontWithStyle( _
                                        GenerateCatalogueReferenceFromID( _
                                        ofwRecord.GetValueForField(gk_sDBCOLCatalogueID), _
                                        ofwRecord.GetValueForField(gk_sDBCOLCatalogueLevel)), _
                                        m_ofwhtmlReference.GetFontStyle(EFontStyle.FONTReadOnlyDataStyle)) _
                                        , typCellStyle
            Next
            
            'Add the table content to the rest of the text
            GetUpdatedCatalogueDisplayPage = GetUpdatedCatalogueDisplayPage & m_oFWTable.GetHTML _
                                            & gk_sHTMLLineBreak
        End If
    End If
    
    'Build a link with the editset ID as parameter
    sEditsetParam = gk_sLinkDisplayEditSetDetails & "?" & gk_sParamEditSetID & "=" & t_lEditSetID
    
    'Add the done button for the page
    GetUpdatedCatalogueDisplayPage = GetUpdatedCatalogueDisplayPage & _
                                    m_ofwhtmlReference.Hyperlink(sEditsetParam, _
                                    m_ofwhtmlReference.Image(mk_sNameButtonOk, gk_sImagesPath & gk_sImageDone, _
                                    m_ofwRefData.Search(gToolTip, ETOOLTIP.ttButtonDone)), _
                                    gk_sMainFrame, mk_sNameButtonOk, _
                                    m_ofwRefData.Search(gToolTip, ETOOLTIP.ttButtonDone))
                                    
    
End Function


Private Sub CheckPermissions(t_eFunctionality As ESecurityCheck, _
                        t_sSource As String, _
                        t_sDescription As String)
'Subroutine that checks and raises an error if the user has no rights to
' the passed function
Dim eCheckRights As ESecurityCheckResult
Dim sSource As String

    sSource = "VWCatalogueBulkupdate.CheckPermissions"
    
    m_oboEditorialPermission.EditsetID = m_lEditsetID
    m_oboEditorialPermission.UserID = m_lLoggedUserID
        
    eCheckRights = m_oboEditorialPermission.Check(t_eFunctionality)
    If eCheckRights <> _
            ESecurityCheckResult.ESRSecurityCheckSuccessful Then
    
        RaiseError sSource, _
            GetErrorEnumForSecurityCheckResult(eCheckRights), _
            GetDescriptionForSecurityCheckError(GetErrorEnumForSecurityCheckResult(eCheckRights))
    End If
        

End Sub


Private Sub DisplayOneLineHeadingForContent(ByVal t_vHeading As Variant)
'
'Description : Displays the heading on the left side of a new row
'
' d8   ksk      27Oct1999       Created

    On Error GoTo ErrorHandler
   
    Dim sSource As String
    
    sSource = "VWCatalogueBulkUpdate.DisplayOneLineHeadingForContent"
    
    m_oFWTable.StartRow
    
    'Show the heading
    m_oFWTable.AddCellAsHTML _
        m_ofwhtmlReference.FontWithStyle( _
            m_ofwhtmlReference.EncodeAsHTML(CStr(t_vHeading)), m_typFontStyleLeftHeading), m_typCellStyleLeftHeading
    
    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub


Private Sub DisplayOneLineEditForContent(ByVal t_vDetails As Variant, _
                                          t_lPaddingCells As Long, _
                                          t_lSpanningCells As Long)
'
'Description : Displays the content on the right side on the current row
'              The displayed data generally is a html element like textbox
'              This function is used when catalogue entry is being edited
'
' d8   ksk      27Oct1999       Created
' d14  tmj      13Mar2000       Changed so takes cell parameters, making function more
'                               flexible

    On Error GoTo ErrorHandler
   
    Dim sSource As String
    
    sSource = "VWCatalogueBulkUpdate.DisplayOneLineEditForContent"

'    '4 because there are max four cells while editing on the right side
'    m_typCellStyleRightDetails.lCOLSPAN = 4
'
'    m_oFWTable.AddCellAsHTML CStr(t_vDetails), m_typCellStyleRightDetails
    
    '''''''''''''''
    Dim lCount As Long
    lCount = 0
    Do While lCount < t_lPaddingCells
        m_oFWTable.AddCellAsHTML "", m_typCellStyleRightDetails
        lCount = lCount + 1
    Loop
    
    m_typCellStyleRightDetails.lCOLSPAN = t_lSpanningCells
    m_oFWTable.AddCellAsHTML CStr(t_vDetails), m_typCellStyleRightDetails
    '''''''''''''''
    
    'End the current row
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
' d8   ksk      27Oct1999       Created

    On Error GoTo ErrorHandler
   
    Dim sSource As String
    
    sSource = "VWCatalogueBulkUpdate.DisplayOneLineEditForContent"

    m_oFWTable.StartRow
    
    'Display the heading
    m_oFWTable.AddCellAsHTML _
        m_ofwhtmlReference.FontWithStyle( _
            m_ofwhtmlReference.EncodeAsHTML(CStr(t_vHeading)), m_typFontStyleLeftHeading), m_typCellStyleLeftHeading
    
    'call the function to display the contents
    DisplayOneLineEditAndDisplayForContent _
         t_vDisplayDetails, t_vRemoveButton, t_vHiddenField, t_vTextBox, t_vAddButton

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
'
' d8   ksk      27Oct1999       Created
' d14  tmj      13Mar2000       Moved description so it is before the textbox

    On Error GoTo ErrorHandler
   
    Dim sSource As String
    
    sSource = "VWCatalogueBulkUpdate.DisplayOneLineEditForContent"
         
    'Check if text box is to be displayed
    If t_vTextBox <> "" Then
        'Display the content - ie details of the link item
        m_typCellStyleRightDetails.lCOLSPAN = 1
        m_typCellStyleRightDetails.sWidth = "40%"
        m_oFWTable.AddCellAsHTML _
            m_ofwhtmlReference.FontWithStyle( _
                m_ofwhtmlReference.EncodeAsHTML(CStr(t_vDisplayDetails)), _
                m_typFontStyleRightDetails), m_typCellStyleRightDetails
                
        'Display the text box for optional description
        m_typCellStyleRightDetails.lCOLSPAN = 1
        m_typCellStyleRightDetails.sWidth = "20%"
        m_oFWTable.AddCellAsHTML _
            m_ofwhtmlReference.FontWithStyle( _
                CStr(t_vTextBox) & CStr(t_vHiddenField), m_typFontStyleRightDetails), _
                m_typCellStyleRightDetails
                    
    Else
        'No text box
        
        'Display the content - ie details of the link item
        'The hidden field is attached to the content
        m_typCellStyleRightDetails.lCOLSPAN = 2
        m_typCellStyleRightDetails.sWidth = "40%"
        m_oFWTable.AddCellAsHTML _
            m_ofwhtmlReference.FontWithStyle( _
            m_ofwhtmlReference.EncodeAsHTML(CStr(t_vDisplayDetails)) & _
            CStr(t_vHiddenField), m_typFontStyleRightDetails), m_typCellStyleRightDetails
    End If
    
    'Add link button  (Only for related and separated material)
    m_typCellStyleRightDetails.lCOLSPAN = 1
    m_typCellStyleRightDetails.sWidth = "10%"
    If t_vAddButton = "" Then
        m_oFWTable.AddCellAsHTML _
            m_ofwhtmlReference.FontWithStyle(gk_sHTMLNonBreakSpace, m_typFontStyleRightDetails), m_typCellStyleRightDetails
    Else
        m_oFWTable.AddCellAsHTML _
            m_ofwhtmlReference.FontWithStyle(CStr(t_vAddButton), m_typFontStyleRightDetails), m_typCellStyleRightDetails
    End If
    
    'Remove link Button
    m_typCellStyleRightDetails.lCOLSPAN = 1
    m_typCellStyleRightDetails.sWidth = "10%"
    m_oFWTable.AddCellAsHTML _
        m_ofwhtmlReference.FontWithStyle(CStr(t_vRemoveButton), m_typFontStyleRightDetails), _
        m_typCellStyleRightDetails
    
    m_oFWTable.EndRow
    
    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub




Private Function StoreFormLinkItemsIntoCollections()
'
'Description : Stores the details of the link items into both the collections
'
' d8   ksk      27Oct1999       Created

On Error GoTo ErrorHandler

    Dim sSource As String
    
    sSource = "VWCatalogueBulkUpdate.StoreFormLinkItemsInCollection"
    
    
    Set m_colFirstSet = ParseAndFillCollectionWithValue(mk_sNameHiddenIDFieldFirstLink, _
                                        mk_sNameDescriptionOptionalFirstLink)
    
    'Check to see if the second collection is required if so save it too.
    If m_bBothTheCollectionsRequired Then
        
        Set m_colSecondSet = ParseAndFillCollectionWithValue(mk_sNameHiddenIDFieldSecondLink, _
                                            mk_sNameDescriptionOptionalSecondLink)
    
    End If
    
    Exit Function
ErrorHandler:
    ConvertError sSource
End Function

Private Function ParseAndFillCollectionWithValue(t_sIDFieldName As String, _
                    t_sDescriptionFieldName As String) As Collection
'Parses the query string for a list of ID (link) and VALUE (link description)
' fields starting with the names t_sIDFieldName and t_sDescriptionFieldName
' d8   ksk      27Oct1999       Created
    Dim iMainCounter As Integer
    Dim typGenericLinkItem As GenericLinkRecord
    
    iMainCounter = 1
    
    Set ParseAndFillCollectionWithValue = New Collection
    'Process all the links
    Do While Not IsEmpty(m_aspRequest.Form(t_sIDFieldName & iMainCounter))
        
        'Get the id
        typGenericLinkItem.lID = CLng(m_aspRequest.Form(t_sIDFieldName & iMainCounter))
        
        If Not t_sDescriptionFieldName = "" Then
            'Description
            typGenericLinkItem.sOptionalDescription = m_aspRequest.Form(t_sDescriptionFieldName & iMainCounter)
        End If
        
        'Add the id and description record to the collection
        ParseAndFillCollectionWithValue.Add typGenericLinkItem, gk_sPrefixForLinkTerm & typGenericLinkItem.lID
        
        iMainCounter = iMainCounter + 1 'Next id
    Loop
    
                    
End Function





Private Function HeaderHTML() As String
' d8   ksk      27Oct1999       Created
' d22  gdb      15Feb2001       Added s extra header - NOW REMOVED

    'Display the header of the page
    HeaderHTML = m_ofwhtmlReference.FontWithStyle( _
                m_oFWAppData.GetDisplayInformation(EEditorialResourceFile.erfBulkEdit), _
                    m_ofwhtmlReference.GetFontStyle(EFontStyle.FONTPageHeadingStyle)) _
                    & gk_sHTMLLineBreak
    HeaderHTML = HeaderHTML & m_ofwhtmlReference.FontWithStyle(m_oFWAppData.GetDisplayInformation( _
                                    EEditorialResourceFile.erfBulkEditValueEntryHelp), _
                                    m_ofwhtmlReference.GetFontStyle( _
                                                        EFontStyle.FONTHelpTextStyle)) & _
                                    gk_sHTMLLineBreak

   

End Function




Private Sub UpdateBusinessObjectWithParsedValues(t_bUseLinkdataFields As Boolean)
'This method updates the business object with the local values
' t_bUseLinkdataFields - when true the collections need to be assigned
' d8   ksk      27Oct1999       Created
' d13  wdp      24Feb2000       Transiently store the levels affected too.

    m_oboCatalogueBulkUpdate.EditsetID = m_lEditsetID
    m_oboCatalogueBulkUpdate.BulkUpdateField = m_eSelectedFieldID
    m_oboCatalogueBulkUpdate.BulkUpdateValue1 = m_vUpdateValue1
    m_oboCatalogueBulkUpdate.BulkUpdateValue2 = m_vUpdateValue2
    m_oboCatalogueBulkUpdate.SetLevels m_bSelectedLevels()
    If t_bUseLinkdataFields Then
        Set m_oboCatalogueBulkUpdate.BulkUpdateValue3 = m_colFirstSet
        Set m_oboCatalogueBulkUpdate.BulkUpdateValue4 = m_colSecondSet
    End If

End Sub

Private Sub UpdateLocalObjectWithBusinessobjectValues(t_bUseLinkdataFields As Boolean)
'This method updates the business object with the local values
' t_bUseLinkdataFields - when true the collections need to be assigned
' d8   ksk      27Oct1999       Created
' d13  wdp      24Feb2000       Transiently store the levels affected too.

    m_lEditsetID = m_oboCatalogueBulkUpdate.EditsetID
    m_eSelectedFieldID = m_oboCatalogueBulkUpdate.BulkUpdateField
    m_vUpdateValue1 = m_oboCatalogueBulkUpdate.BulkUpdateValue1
    m_vUpdateValue2 = m_oboCatalogueBulkUpdate.BulkUpdateValue2
    m_oboCatalogueBulkUpdate.GetLevels m_bSelectedLevels()
    If t_bUseLinkdataFields Then
        Set m_colFirstSet = m_oboCatalogueBulkUpdate.BulkUpdateValue3
        Set m_colSecondSet = m_oboCatalogueBulkUpdate.BulkUpdateValue4
    End If

End Sub


Private Function CorporateBodyDescriptionToDisplay(t_lCorporateBodyID As Long) As String
'Returns the corporate body desciption that needs to be displayed when the link is selected
' d8   ksk      27Oct1999       Created
    
    Dim ofwReferenceData As FWReferenceData

    Set ofwReferenceData = New FWReferenceData
    
    CorporateBodyDescriptionToDisplay = ofwReferenceData.Search(EGroup.gCorporateBody, _
                                                    t_lCorporateBodyID, _
                                                    gk_sDBCOLCorporateBodyReferenceID, _
                                                    gk_sDBCOLCorporateBodyName)
                            
End Function

Private Function PersonDescriptionToDisplay(t_lPersonID As Long) As String
'Returns the Persons desciption that needs to be displayed when the link is selected
' d8   ksk      27Oct1999       Created
    
    Dim ofwReferenceData As FWReferenceData

    Set ofwReferenceData = New FWReferenceData
    
    PersonDescriptionToDisplay = VariantStringAsString( _
                            ofwReferenceData.Search(EGroup.gPerson, _
                                            t_lPersonID, _
                                            gk_sDBCOLPersonReferenceID, _
                                            gk_sDBCOLPerson_Title))
    
    PersonDescriptionToDisplay = PersonDescriptionToDisplay & _
                                 VariantStringAsString( _
                                 ofwReferenceData.Search(EGroup.gPerson, _
                                            t_lPersonID, _
                                            gk_sDBCOLPersonReferenceID, _
                                            gk_sDBCOLForename_Text))
    
    PersonDescriptionToDisplay = PersonDescriptionToDisplay & _
                                 VariantStringAsString( _
                                 ofwReferenceData.Search(EGroup.gPerson, _
                                            t_lPersonID, _
                                            gk_sDBCOLPersonReferenceID, _
                                            gk_sDBCOLSurname_Text))
    

End Function

Private Function LeafletDescriptionToDisplay(t_lLeafletID As Long) As String
'Returns the Leaflet title that needs to be displayed when the link is selected
' d9    TMJ     14Mar2000   Created
    
    Dim oboLeafletLoader As BOLeafletLoader
    Dim sSource As String
    Dim typ_Data As LeafletDetails
    
    sSource = "VWCataloguebulkUpdate.LeafletDescriptionToDisplay"
    
    Set oboLeafletLoader = New BOLeafletLoader
    
    oboLeafletLoader.LoadData (t_lLeafletID)
    
    typ_Data = oboLeafletLoader.Data

    LeafletDescriptionToDisplay = typ_Data.sTitle
    
ErrorHandler:
    Set oboLeafletLoader = Nothing
End Function

Private Function IsIDPresentInCollection(t_lID As Long, t_colData As Collection) As Boolean
On Error GoTo Error_Handler
Dim lValue As Long
    
    IsIDPresentInCollection = True
    lValue = t_colData.Item(gk_sPrefixForLinkTerm & CStr(t_lID)).lID
    
    Exit Function
    
Error_Handler:
    IsIDPresentInCollection = False
End Function

Private Function SFADescriptionToDisplay(t_lSFAID As Long) As String
'Returns the string of SFA description to be displayed
' d8   ksk      27Oct1999       Created

    Dim oboSFADetail As BOSFALoader
    Dim typSFARecord As SFADetails

    Set oboSFADetail = New BOSFALoader

    oboSFADetail.LoadData (t_lSFAID)
    
    typSFARecord = oboSFADetail.Data
    
    SFADescriptionToDisplay = VariantStringAsString(typSFARecord.sTitle) & _
                                ", " & VariantStringAsString(typSFARecord.sSubTitle) & _
                                ", " & VariantStringAsString(typSFARecord.sEdition) & _
                                ", " & typSFARecord.sAuthor
    
    If typSFARecord.bInProLibaray Then
    
        SFADescriptionToDisplay = SFADescriptionToDisplay & m_oFWAppData.GetDisplayInformation( _
                    ESharedResourceFile.srfTabDisplayContentDetailsMessageCopyAtPROLibrary)

    End If
    
    If typSFARecord.bInProReference Then
    
        SFADescriptionToDisplay = SFADescriptionToDisplay & m_oFWAppData.GetDisplayInformation( _
                    ESharedResourceFile.srfTabDisplayContentDetailsMessageAvailableAtReadingRoom)

    End If
    

End Function

Private Sub UpdateStartAndEndDates(ByVal t_sCoveringDates As String)
' Desc: Determines if the start and end dates should be overwritten, based
'       on the covering date. Works out what the new start and end date
'       should be, based on the covering date, updating the relevant
'       properties
'
' Parm: t_sCoveringDates - The covering date
'
' Rtrn: None
'
' Amendment history:
'   d21  AManic  13-Feb-2001     Created
On Error GoTo ErrorHandler
    
    ' Declare locals
    Dim sSource         As String
    Dim dStartDate      As Date
    Dim dEndDate        As Date
    Dim oBOCoveringDate As BOCoveringDate
    Dim bResult         As Boolean
    
    ' Define locals
    Set oBOCoveringDate = New BOCoveringDate
    
    ' Define error string
    sSource = "VWCatalogueBulkUpdate.UpdateStartAndEndDates"
        
    ' Update the start / end dates based on the covering date. If the
    ' covering date can not be converted then use the start / end dates
    ' specified by the user
    If (oBOCoveringDate.GetStartAndEndDate(t_sCoveringDates, _
            dStartDate, dEndDate) = True) Then
                                          
        ' Use the start and end dates generated from the covering date
        m_vUpdateValue2 = ConvertUserDate(dStartDate) & _
                          ConvertUserDate(dEndDate)
        
    Else
        
        m_vUpdateValue2 = Null
        
    End If

    ' Clean up
    Set oBOCoveringDate = Nothing
    
    Exit Sub
ErrorHandler:

    ' Clean up
    Set oBOCoveringDate = Nothing
    
    ConvertError sSource
End Sub


