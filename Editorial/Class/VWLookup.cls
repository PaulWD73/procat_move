VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "VWLookup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    VWLookup.cls
' System:       PROCAT, PRO
' Copyright:    (C) Quidnunc Limited
'
' Description:  Displays the Lookup maintanence page to add, modify and delete
'               the records from the lookup table
'
' Amendment history:
'   d1   aga      06Jul1999     Created
'   d2   aga      16Jul1999     Added ShowAddModifyStatus()
'   d3   ksk      16Jul1999     Integrated
'   d4   aga      20Jul1999     Changes made according to the BOLookup interface
'                               Changes to Add, Delete, Modify methods and to ShowAddModifyStatus
'   d5   aga      10Jul1999     Removed "Remove image Button" for Closure Status & Integration
'   d6   ksk      03Sep1999     Removed Physical description from the lookup
'   d7   aga      05Oct1999     Added using FWAppData for diplaying Resource contants ...
'                               Added ShowWarningPage method to display a warning message
'                               Modified to allow deletion of lookups
'   d8   aga      26Oct1999     Taken d6 and added undeletion.
'                               Added using FWAppData for diplaying Resource contants ...
'                               Added ShowWarningPage method to display a warning message
'                               Added permissions
'   d9   aga      28Oct1999     Removed ShowWarningPage
'   d10  AGA      02Nov1999     Modified CheckPermissions
'   d11  KSK      03Nov1999     Changed Checkpermissions method to used the parameter source
'                               value while raising error
'   d10  AGA      17Dec1999     Added ISO_Code field for Language lookup
'   d11  AGA      24Dec1999     Fixing the defects & Changing the cell styles
'   d12  AGA      12Jan2000     Fixed minor defects
'   d13  tmj      06Mar2000     changed widths of columns
'   d14  tmj      28Mar2000     Fixed cosmetic defect in ShowLookupHomepage
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'Interfaces
Implements ObjectControl


Private m_eLookupType As ELOOKUPTYPE

'Private variables to hold the Lookup table type, Lookup ID & Lookup Value

Private m_iLookupType As Integer    ' Lookup type
Private m_vLookupID As Variant          ' ID of the lookup record
Private m_sLookupValue As String        ' Value of the Lookup


'private variables to hold table and field names
Private m_sLookupTableName As String
Private m_sLookupIDField As String
Private m_sLookupValueField As String
Private m_sCurrencyIndicatorField As String

'Private Variable to hold the Field headings of the Lookup table
' to display to the user
'   Note: Currency Indicator is at present not to be displayed to the user
'         In case of displaying it to user in future, add a variable for that here

Private m_sLookupIDHeading As String
Private m_sLookupValueHeading As String
Public m_lLookupValueFieldLength As Long 'Var to hold the field length in the database
'd2
Private m_eSaveReturnCode As ELOOKUPSAVERETURNCODE   ' to check the return code of the BOLookup.Save

'Private variables to hold object variables

Private m_oObjectContext As ObjectContext
Private m_aspResponse As Response
Private m_aspRequest As Request
Private m_aspServer As Server

'd8  'For displaying resource strings
Private m_oFWAppData As FWApplicationData

Private m_ofwRefData As FWReferenceData

'Constants for Parameters for values passes to and from asp form & querystring variables
Private Const ksParamModifyID As String = "MID"
Private Const ksParamModifyValue As String = "MVal"
Private Const ksParamRemoveID As String = "RID"
Private Const ksParamRemoveValue As String = "RVal"
Private Const ksNameAddID As String = "AddID"
Private Const ksNameAddValue As String = "AddVal"
Private Const ksParamLookupType As String = "LookupType"
Private Const ksNameModifyID As String = "MID"
Private Const ksNameModifyValue As String = "MVal"

'd10
Private Const ksISOCode As String = "ISOCode"

'd8  Permissions object
Private m_oboEditorialPermission As BOEditorialPermissions
Private m_lLoggedUserID As Long



'Object control interfaces

Private Sub ObjectControl_Activate()
    Initialise
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    Terminate
End Sub
Private Sub Initialise()
' Description:
'
' Initialises the objects
' Amendment history:
'
' d1   aga        06Jul1999    Created
' d8   aga        26Oct1999    Added initialization of FWAppData, Permissions
'
    Dim sSource As String
    sSource = "VWLookup.Initialise"
    On Error GoTo ErrorHandler
    
    Set m_oObjectContext = GetObjectContext()
    If m_oObjectContext Is Nothing Then
        RaiseError sSource, peOCUnableToGetObjectContext, "Objectcontext is nothing"
    Else
        
        Set m_aspResponse = m_oObjectContext.Item(gk_sResponseObject)
        Set m_aspRequest = m_oObjectContext.Item(gk_sRequestObject)
        Set m_aspServer = m_oObjectContext.Item(gk_sServerObject)
        Set m_oFWAppData = New FWApplicationData
        Set m_oboEditorialPermission = New BOEditorialPermissions
        Set m_ofwRefData = New FWReferenceData
                
    End If
    
    m_eLookupType = m_aspRequest(ksParamLookupType)
    
    m_lLoggedUserID = GetDataFromSession(dsUserID, m_oObjectContext.Item(gk_sSessionObject))
    If m_lLoggedUserID = 0 Then     'Not a valid user
        RaiseError sSource, EPROCATINTERNALERROR.peDBInvalidUser, "Invalid User"
    End If

    Exit Sub
    
ErrorHandler:
    ShowError sSource
End Sub
Private Sub Terminate()
' Description:
'
' Releases all the objects
' Amendment history:
'
' d1   aga        06Jul1999    Created
' d8   aga        26Oct1999    Added releasing of FWAppData, permissions object
'
    'Release the objects
    Set m_oObjectContext = Nothing
    Set m_aspResponse = Nothing
    Set m_aspRequest = Nothing
    Set m_aspServer = Nothing
    Set m_oFWAppData = Nothing
    Set m_oboEditorialPermission = Nothing
    Set m_ofwRefData = Nothing

    
End Sub

Public Sub AddLookupAndShowRecords()
' Description:
'
'to add a record to the Lookup table and to show the modified table
'It uses the BOLookup.Save to add the record
'
' Amendment history:
'
' d1   aga      12Jul1999   Created
' d2   aga      20Jul1999   Changed code to call save with out parameters & to set the
'                               BOLookup properties
' d8   aga      26Oct1999   Added security
'
    Dim sSource As String
    Dim oBOLookup As BOLookup

    sSource = "VWLookup.AddLookupAndShowRecords"

    On Error GoTo ErrorHandler
    'Check the permissions
    CheckPermissions ESecurityCheck.EUserCanMaintainLookUps, sSource, ""
    
    Set oBOLookup = New BOLookup
    
    'To display in the ShowAddModifyStatus method to display the status to the user
    m_sLookupValue = CStr(m_aspRequest.Form(ksNameAddValue))
    m_vLookupID = CVar(m_aspRequest.Form(ksNameAddID))
   
    'Get the TableID, Lookup Value, LookupID from the ASP page using Response object
    'and set the properties of the BOLookup and call save
    oBOLookup.LookupTableID = CInt(Trim(m_aspRequest(ksParamLookupType)))
    oBOLookup.LookupID = CVar(Trim(m_aspRequest.Form(ksNameAddID)))
    oBOLookup.LookupValue = CStr(Trim(m_aspRequest.Form(ksNameAddValue)))
    If Not Trim(m_aspRequest(ksISOCode)) = "" Then
        oBOLookup.LanguageISOCode = Trim(m_aspRequest(ksISOCode))
    End If
    m_eSaveReturnCode = oBOLookup.Save

    'Display the homepage
    m_aspResponse.Redirect gk_sLookupDisplay & "?" & ksParamLookupType & "=" & oBOLookup.LookupTableID

    
    Set oBOLookup = Nothing
    
    'transaction successfull ... set the transaction complete
    m_oObjectContext.SetComplete


    Exit Sub

ErrorHandler:

    'transaction failed ... Abort it
    m_oObjectContext.SetAbort
    
    ShowError sSource

End Sub

Public Sub DeleteLookupAndShowRecords()
' Description:
'
'to delete a record from a lookup table, and shows the modified table
'  It uses the BOLookup.Delete to delete a record
'
' Amendment history:
'
' d1   aga      12Jul1999   Created
' d2   aga      20Jul199    Changed code to call save with out parameters & to set the
'                               BOLookup properties
' d8   aga      26Oct1999   Added security
'

    Dim sSource As String
    Dim oBOLookup As BOLookup

    sSource = "VWLookup.DeleteLookupAndShowRecords"

    On Error GoTo ErrorHandler
    'Check the permissions
    CheckPermissions ESecurityCheck.EUserCanMaintainLookUps, sSource, ""
    
    Set oBOLookup = New BOLookup
    
    'Set the BOLookup properties and call the BOLookup.Delete
    oBOLookup.LookupTableID = CInt(m_aspRequest(ksParamLookupType))
    oBOLookup.LookupID = m_aspRequest(ksParamRemoveID)
    Call oBOLookup.Delete
    

    'Display lookup homepate
    m_aspResponse.Redirect gk_sLookupDisplay & "?" & ksParamLookupType & "=" & oBOLookup.LookupTableID

    Set oBOLookup = Nothing
    
    'transaction successfull ... set the transaction complete
    m_oObjectContext.SetComplete

    Exit Sub

ErrorHandler:

    'transaction failed ... Abort it
    m_oObjectContext.SetAbort

    ShowError sSource


End Sub
Public Sub UnDeleteLookupAndShowRecords()
' Description:
'
'   to delete a record from a lookup table, and shows the modified table
'   It uses the BOLookup.Delete to delete a record
'
' Amendment history:
'
'   d8  aga  26-Oct-1999    Created
'

    Dim sSource As String
    Dim oBOLookup As BOLookup

    sSource = "VWLookup.UnDeleteLookupAndShowRecords"

    On Error GoTo ErrorHandler
    
    
    Set oBOLookup = New BOLookup
    
    'Set the BOLookup properties and call the BOLookup.Delete
    oBOLookup.LookupTableID = CInt(m_aspRequest(ksParamLookupType))
    oBOLookup.LookupID = m_aspRequest(ksParamRemoveID)
    Call oBOLookup.UnDelete
    
    'Display lookup homepate
    m_aspResponse.Redirect gk_sLookupDisplay & "?" & ksParamLookupType & "=" & oBOLookup.LookupTableID

    Set oBOLookup = Nothing
    
    'transaction successfull ... set the transaction complete
    m_oObjectContext.SetComplete

    Exit Sub

ErrorHandler:

    'transaction failed ... Abort it
    m_oObjectContext.SetAbort

    ShowError sSource

End Sub


Public Sub ModifyLookupAndShowRecords()
' Description:
'
'to modify a record from a lookup table, and shows the modified table
'  It uses the BOLookup.Save to modify a record
'
' Amendment history:
'
' d1   aga      12Jul1999       Created
' d2   aga      20Jul1999       Changed code to call save with out parameters & to set the
'                               BOLookup properties
' d8   aga      26Oct1999       Added security
'


    Dim sSource As String
    Dim oBOLookup As BOLookup

    sSource = "VWLookup.ModifyLookupAndShowRecords"

    On Error GoTo ErrorHandler
    'Check the permissions
    CheckPermissions ESecurityCheck.EUserCanMaintainLookUps, sSource, ""
    
    Set oBOLookup = New BOLookup

    'To display in the ShowAddModifyStatus method to display the status to the user
    m_sLookupValue = CStr(m_aspRequest(ksNameModifyValue))
    m_vLookupID = CVar(m_aspRequest.Form(ksNameModifyID))

    'set the BOLookup properties and call Save
    oBOLookup.LookupTableID = CInt(Trim(m_aspRequest(ksParamLookupType)))
    oBOLookup.LookupID = Trim(m_aspRequest(ksNameModifyID))
    oBOLookup.LookupValue = CStr(Trim(m_aspRequest(ksNameModifyValue)))
    
    If Not Trim(m_aspRequest(ksISOCode)) = "" Then
        oBOLookup.LanguageISOCode = Trim(m_aspRequest(ksISOCode))
    End If
    
    m_eSaveReturnCode = oBOLookup.Save

    'Display lookup homepage
    m_aspResponse.Redirect gk_sLookupDisplay & "?" & ksParamLookupType & "=" & oBOLookup.LookupTableID

    Set oBOLookup = Nothing
    
    'transaction successfull ... set the transaction complete
    m_oObjectContext.SetComplete

    Exit Sub

ErrorHandler:
    'transaction failed ... Abort it
    m_oObjectContext.SetAbort
        
    ShowError sSource

End Sub

Public Sub ShowModificationPage()
' Description:
'
'shows the output in html form and enables the user
'   to modify a particular lookup
'
' Amendment history:
'
' d1   aga      13Jul1999   Created
' d8   aga      26Oct1999   Added security
'
    On Error GoTo ErrorHandler
    
    
    'Local variables
    Dim sSource As String    ' to store error string
    Dim sHTML As String      ' to store html tags and data
    Dim ofwHTML As FWHTML
    Dim ofwTable As FWTable
    
    Dim typCellStyle As CellStyle
    Dim typTableStyle As TableStyle
    
    'Temporary variables to hold the ID & Value pair of a record in the record list
    Dim sFieldID As String
    Dim sFieldValue As String
    
    Dim oBOLookup As BOLookup
    
    sSource = "VWLookup.ShowModificationPage"
    
    'Check the permissions
    CheckPermissions ESecurityCheck.EUserCanMaintainLookUps, sSource, ""
    
    'Get the data from ASP and assign to Tabletype, Lookup ID & Values
    m_iLookupType = CInt(m_aspRequest(ksParamLookupType))
    m_vLookupID = CVar(m_aspRequest(ksNameModifyID))
    m_sLookupValue = CStr(m_aspRequest(ksNameModifyValue))
        
    Set ofwHTML = New FWHTML
    
   
    'Map the corresponding table and fields names to the member variables
    ' to be used for displaying to the user
    MapTableAndFields (m_eLookupType)
    
    sHTML = ""
    sHTML = sHTML & ofwHTML.Bold(ofwHTML.Font(ShowLookupHeading, COLOURBlack, 4))
    sHTML = sHTML & gk_sHTMLLineBreak
    
    sHTML = sHTML & ofwHTML.Paragraph(m_oFWAppData.GetDisplayInformation(ELookupMaintainResource.LMModify) _
        & " " & m_sLookupValueHeading)
    
    Set oBOLookup = New BOLookup
    
    sHTML = sHTML & ofwHTML.FormStart(gk_sLookupUpdate & "?" & ksParamLookupType & "=" _
        & m_iLookupType, gk_sMainFrame, "frmModify", MPost, True)
        
    
    'show the values in a tabular format
    
    Set ofwTable = New FWTable
    
    typTableStyle = ofwHTML.GetTableStyle(TABCustomStyle)
    typTableStyle.sWidth = "100%"
    typTableStyle.lCELLSPACING = 0
    ofwTable.TableStyle = typTableStyle
            
    ofwTable.AddBlankRow
    
    'Display Fields names & textbox for adding lookup value
    
    If m_eLookupType = ELOOKUPTYPE.lutClosureStatus Then
            'Text box for ID (for closure lookup, user has to enter the ID also)
        ofwTable.StartRow
            typCellStyle = ofwHTML.GetCellStyle(ECellStyle.CELLReadWriteLabelStyle)
            typCellStyle.sWidth = "5%"
            ofwTable.AddCellAsHTML m_sLookupIDHeading, typCellStyle
            typCellStyle.sWidth = "20%"
            ofwTable.AddCellAsHTML ofwHTML.HiddenField(ksNameModifyID, CStr(m_vLookupID)) & _
                ofwHTML.Font(CStr(m_vLookupID), COLOURBlack, 3), typCellStyle
            typCellStyle.sWidth = "60%"
            ofwTable.AddCellAsHTML " ", typCellStyle
        ofwTable.EndRow
        ofwTable.StartRow
            typCellStyle = ofwHTML.GetCellStyle(ECellStyle.CELLReadWriteLabelStyle)
            typCellStyle.sWidth = "5%"
            ofwTable.AddCellAsHTML m_sLookupValueHeading, typCellStyle
            typCellStyle.sWidth = "20%"
            
            ofwTable.AddCellAsHTML ofwHTML.InputField(EInputType.ITText, ksNameModifyValue, 50, _
                m_lLookupValueFieldLength, m_sLookupValue, , True), typCellStyle
            
            ofwTable.AddCellAsHTML " ", typCellStyle
        ofwTable.StartRow
            typCellStyle.sWidth = "5%"
            ofwTable.AddCellAsHTML " "
            typCellStyle.sWidth = "20%"
            ofwTable.AddCellAsHTML ofwHTML.SubmitImage(gk_sImagesPath & gk_sImageOK, "frmSubmitOK", "somevalue", "") & _
                gk_sHTMLNonBreakSpace & gk_sHTMLNonBreakSpace & gk_sHTMLNonBreakSpace & _
                ofwHTML.Hyperlink(gk_sLookupDisplay & "?" & ksParamLookupType & "=" & _
                CStr(m_eLookupType), ofwHTML.Image("name", gk_sImagesPath & gk_sImageCancel, "Cancel", 0, , , ""), _
                gk_sMainFrame), typCellStyle
            typCellStyle.sWidth = "60%"
            ofwTable.AddCellAsHTML " ", typCellStyle
        ofwTable.EndRow
    Else
        
        typCellStyle = ofwHTML.GetCellStyle(ECellStyle.CELLReadWriteLabelStyle)
        
        ofwTable.StartRow
            typCellStyle.sWidth = "20%"
            ofwTable.AddCellAsHTML m_sLookupValueHeading, typCellStyle
            
            typCellStyle = ofwHTML.GetCellStyle(ECellStyle.CELLReadWriteDataStyle)
            typCellStyle.sWidth = "20%"
            ofwTable.AddCellAsHTML ofwHTML.InputField(EInputType.ITText, ksNameModifyValue, 50, _
                m_lLookupValueFieldLength, m_sLookupValue, , True), typCellStyle
            
            typCellStyle.sWidth = "45%"
            ofwTable.AddCellAsHTML " ", typCellStyle
        ofwTable.EndRow
        'ISO_Code field for Language
        If m_eLookupType = ELOOKUPTYPE.lutLanguage Then
            ofwTable.StartRow
            typCellStyle = ofwHTML.GetCellStyle(ECellStyle.CELLReadWriteLabelStyle)
            typCellStyle.sWidth = "20%"
            ofwTable.AddCellAsHTML "ISO Code", typCellStyle
            
            typCellStyle = ofwHTML.GetCellStyle(ECellStyle.CELLReadWriteDataStyle)
            typCellStyle.sWidth = "20%"
            ofwTable.AddCellAsHTML ofwHTML.InputField(EInputType.ITText, ksISOCode, 50, _
                m_lLookupValueFieldLength, CStr(m_aspRequest(ksISOCode)), , True), typCellStyle
            
            typCellStyle.sWidth = "45%"
            ofwTable.AddCellAsHTML " ", typCellStyle
        ofwTable.EndRow
        
        End If
        ofwTable.StartRow
            typCellStyle = ofwHTML.GetCellStyle(ECellStyle.CELLReadOnlyLabelStyle)
            typCellStyle.sWidth = "5%"
            ofwTable.AddCellAsHTML ofwHTML.HiddenField(ksNameModifyID, CInt(m_vLookupID))
            typCellStyle.sWidth = "20%"
            
            ofwTable.AddCellAsHTML ofwHTML.SubmitImage(gk_sImagesPath & gk_sImageOK, "frmSubmitOK", "somevalue", "") & _
                ofwHTML.ImageButton(gk_sImagesPath & gk_sImageCancel, , gk_sLookupDisplay & "?" & ksParamLookupType & "=" & _
                CStr(m_eLookupType), gk_sMainFrame, m_ofwRefData.Search(gToolTip, ETOOLTIP.ttButtonCancelHelp)) _
                , typCellStyle
            ofwTable.AddCellAsHTML " ", typCellStyle
        ofwTable.EndRow
         
    End If
   
    
    sHTML = sHTML & ofwTable.GetHTML
    
    sHTML = sHTML & gk_sHTMLLineBreak & ofwHTML.EndTag("form")
    
    'Add context sensitive help button
    sHTML = sHTML & HTMLToAddHelpButton(gk_sHelpLinkToMaintLookup)
    
    m_aspResponse.Write (sHTML)
    
    'transaction successfull ... set the transaction complete
    
    m_oObjectContext.SetComplete
    
    Set ofwHTML = Nothing
    Set ofwTable = Nothing
    Set oBOLookup = Nothing
    Exit Sub
ErrorHandler:

    'transaction failed ... Abort it
    m_oObjectContext.SetAbort
    ShowError sSource


End Sub

Public Sub ShowLookupHomePage()
' Description:
'
'shows the lookup entries of a lookup table and enables the user to add, modify
' and remove
'
' Amendment history:
'
' d1   aga        07Jul1999     Created
' d2   aga        20Jul1999     Changed code to call save with out parameters & to set the
'                               BOLookup properties
' d3   aga        10Jul1999     Removed "Remove image Button" for Closure Status
' d8   aga        26Oct1999     Added permissions
' d12  AGA        12Jan2000     Adding done button
' d13  tmj        06Mar2000     changed widths of columns
' d14  tmj        28Mar2000     Added EncodeAsHTML so tags are not inadvertantly displayed
'
    On Error GoTo ErrorHandler
    
    
    'Local variables
    Dim sSource As String    ' to store error string
    Dim sHTML As String      ' to store html tags and data
    Dim ofwHTML As FWHTML
    Dim ofwTable As FWTable
    
    Dim vRowColour As Variant
    
    Dim typRowStyle As RowStyle
    Dim typCellStyle As CellStyle
    Dim typTableStyle As TableStyle
    
    Dim oFWRecordList As FWRecordList
    'Local variables to hold the ID & Value pair of a record in the record list
    Dim sFieldID As String
    Dim sFieldValue As String
    
    Dim oBOLookup As BOLookup
    
    sSource = "VWLookup.ShowLookupHomePage"
    
    'Check the permissions
    CheckPermissions ESecurityCheck.EUserCanMaintainLookUps, sSource, ""
    
    Set ofwHTML = New FWHTML
    
   
    m_iLookupType = m_aspRequest(ksParamLookupType)
   
    'Map the corresponding table and fields names to the member variables
    ' to be used for displaying to the user
    MapTableAndFields (m_eLookupType)
    
    sHTML = ""
    sHTML = sHTML & ofwHTML.Bold(ofwHTML.Font(ShowLookupHeading, COLOURBlack, 4))
    sHTML = sHTML & gk_sHTMLLineBreak
    
    sHTML = sHTML & ofwHTML.Paragraph(m_oFWAppData.GetDisplayInformation(ELookupMaintainResource.LMAddNew) & " " & m_sLookupValueHeading)
     
    
    Set oBOLookup = New BOLookup
    
    'Get the records from the table
    oBOLookup.LookupTableID = m_eLookupType
    
    'd8 pass true parameter to return all records
    Set oFWRecordList = oBOLookup.LookupList(True)
    
    
    sHTML = sHTML & ofwHTML.FormStart(gk_sLookupAdd & "?" & ksParamLookupType & "=" & m_iLookupType, gk_sMainFrame, "frmAdd", MPost, True)

    
    'show the values in a tabular format
    
    Set ofwTable = New FWTable
    
    typTableStyle.eStyle = TABCustomStyle
    typTableStyle.sWidth = "100%"
    typTableStyle.lCELLSPACING = 2
    ofwTable.TableStyle = typTableStyle
            
    
    typRowStyle = ofwHTML.GetRowStyle(RowBackgroundPink)
    
    'Display textbox for adding lookup value
    typCellStyle = ofwHTML.GetCellStyle(ECellStyle.CELLReadWriteDataStyle)
    
    If m_eLookupType = ELOOKUPTYPE.lutClosureStatus Then
        ofwTable.StartRow typRowStyle
            typCellStyle = ofwHTML.GetCellStyle(ECellStyle.CELLReadWriteLabelStyle)
            typCellStyle.sWidth = "5%"
            ofwTable.AddCellAsHTML m_sLookupIDHeading, typCellStyle
            'Text box for ID (for closure lookup, user has to enter the ID also)
            typCellStyle = ofwHTML.GetCellStyle(ECellStyle.CELLReadOnlyDataStyle)
            typCellStyle.sWidth = "30%"
            ofwTable.AddCellAsHTML ofwHTML.InputField(ITText, ksNameAddID, 1, 1, "", EValidCheck.VCNone, True), typCellStyle
        ofwTable.EndRow
        ofwTable.StartRow typRowStyle
            typCellStyle = ofwHTML.GetCellStyle(ECellStyle.CELLReadWriteLabelStyle)
            typCellStyle.sWidth = "5%"
            ofwTable.AddCellAsHTML m_sLookupValueHeading, typCellStyle
            typCellStyle = ofwHTML.GetCellStyle(ECellStyle.CELLReadWriteDataStyle)
            typCellStyle.sWidth = "30%"
            ofwTable.AddCellAsHTML ofwHTML.InputField(ITText, ksNameAddValue, 50, _
                m_lLookupValueFieldLength, "", EValidCheck.VCNone, True), typCellStyle
        ofwTable.EndRow
        ofwTable.StartRow typRowStyle
            typCellStyle = ofwHTML.GetCellStyle(ECellStyle.CELLReadWriteLabelStyle)
            typCellStyle.sWidth = "5%"
            ofwTable.AddCellAsHTML " ", typCellStyle
            typCellStyle.sWidth = "30%"
            ofwTable.AddCellAsHTML ofwHTML.SubmitImage(gk_sImagesPath & gk_sImageAdd, "frmSubmitAdd", "somevalue"), typCellStyle
        ofwTable.EndRow
    Else
        
        ofwTable.StartRow typRowStyle
            typCellStyle = ofwHTML.GetCellStyle(ECellStyle.CELLReadWriteLabelStyle)
            typCellStyle.sWidth = "5%"
            ofwTable.AddCellAsHTML m_sLookupValueHeading, typCellStyle
            typCellStyle = ofwHTML.GetCellStyle(ECellStyle.CELLReadWriteDataStyle)
            typCellStyle.sWidth = "30%"
            ofwTable.AddCellAsHTML ofwHTML.InputField(ITText, ksNameAddValue, 50, m_lLookupValueFieldLength, "", EValidCheck.VCNone, True), typCellStyle
        ofwTable.EndRow
        'For language lookup add a ISO_CODE field
        If m_eLookupType = ELOOKUPTYPE.lutLanguage Then
            ofwTable.StartRow typRowStyle
                typCellStyle = ofwHTML.GetCellStyle(ECellStyle.CELLReadWriteLabelStyle)
                typCellStyle.sWidth = "5%"
                ofwTable.AddCellAsHTML "ISO Code", typCellStyle
                typCellStyle = ofwHTML.GetCellStyle(ECellStyle.CELLReadWriteDataStyle)
                typCellStyle.sWidth = "30%"
                ofwTable.AddCellAsHTML ofwHTML.InputField(ITText, ksISOCode, 50, m_lLookupValueFieldLength, "", EValidCheck.VCNone, True), typCellStyle
            ofwTable.EndRow
        End If
        ofwTable.StartRow typRowStyle
            typCellStyle = ofwHTML.GetCellStyle(ECellStyle.CELLReadOnlyLabelStyle)
            typCellStyle.sWidth = "5%"
            ofwTable.AddCellAsHTML " ", typCellStyle
            ofwTable.AddCellAsHTML ofwHTML.SubmitImage(gk_sImagesPath & gk_sImageAdd, "frmSubmitAdd", "somevalue"), typCellStyle
        ofwTable.EndRow
    End If
    
    
    sHTML = sHTML & ofwTable.GetHTML
    sHTML = sHTML & ofwHTML.Paragraph(ofwHTML.Font(ShowAddModifyStatus, COLOURRed))
    
    If oFWRecordList.RecordCount <> 0 Then
        sHTML = sHTML & ofwHTML.Paragraph(m_oFWAppData.GetDisplayInformation( _
            ELookupMaintainResource.LMListOfExisting) & " " & m_sLookupValueHeading & " lookups: ")
    Else
        sHTML = sHTML & ofwHTML.Paragraph(m_oFWAppData.GetDisplayInformation(ELookupMaintainResource.LMNoRecords))
    End If
    
    ofwTable.Reset
    
    typTableStyle.eStyle = TABCustomStyle
    typTableStyle.sWidth = "100%"
    typTableStyle.lCELLSPACING = 2
    ofwTable.TableStyle = typTableStyle
    
    typCellStyle = ofwHTML.GetCellStyle(ECellStyle.CELLColumnHeadingStyle)
    
    'Set the red back ground for the heading of table
    typRowStyle = ofwHTML.GetRowStyle(RowBackgroundRed)
    
    
    'Print the heading before the first row
    
    ofwTable.StartRow typRowStyle
    If m_eLookupType = ELOOKUPTYPE.lutClosureStatus Then
        typCellStyle.sWidth = "5%"
        ofwTable.AddCellAsHTML m_sLookupIDHeading, typCellStyle
    End If
    'For language lookup add the ISO code field
    If m_eLookupType = ELOOKUPTYPE.lutLanguage Then
        typCellStyle.sWidth = "35%"
        ofwTable.AddCellAsHTML "ISO Code", typCellStyle
        typCellStyle.sWidth = "35%"
        ofwTable.AddCellAsHTML m_sLookupValueHeading, typCellStyle
    Else
        typCellStyle.sWidth = "70%"
        ofwTable.AddCellAsHTML m_sLookupValueHeading, typCellStyle
    End If
    typCellStyle = ofwHTML.GetCellStyle(ECellStyle.CELLReadOnlyLabelStyle)
    ofwTable.AddCellAsHTML "", typCellStyle
    ofwTable.AddCellAsHTML "", typCellStyle
    ofwTable.EndRow
    
    While Not oFWRecordList.IsAtEnd
        
        
        sFieldID = CStr(oFWRecordList.GetFieldValue(m_sLookupIDField))
        sFieldValue = CStr(oFWRecordList.GetFieldValue(m_sLookupValueField))
        
        ofwTable.StartRow
        
        typCellStyle = ofwHTML.GetCellStyle(ECellStyle.CELLReadOnlyDataStyle)

        If m_eLookupType = ELOOKUPTYPE.lutClosureStatus Then
            typCellStyle.sWidth = "5%"
            ofwTable.AddCellAsHTML ofwHTML.EncodeAsHTML(sFieldID), typCellStyle
        End If
        
        'For language lookup add a ISO_CODE field
        If m_eLookupType = ELOOKUPTYPE.lutLanguage Then
            typCellStyle.sWidth = "35%"
            ofwTable.AddCellAsHTML ofwHTML.EncodeAsHTML( _
                         VariantStringAsString(oFWRecordList.GetFieldValue(gk_sDBCOLLanugageISOCode))), typCellStyle
            typCellStyle.sWidth = "35%"
            ofwTable.AddCellAsHTML ofwHTML.EncodeAsHTML(sFieldValue), typCellStyle
        Else
            typCellStyle.sWidth = "70%"
            ofwTable.AddCellAsHTML ofwHTML.EncodeAsHTML(sFieldValue), typCellStyle
        End If
        'Creating Hyperlinks to the Modify & Remove images & corresponding querystring values of a record
        typCellStyle = ofwHTML.GetCellStyle(ECellStyle.CELLReadOnlyLabelStyle)
        typCellStyle.sWidth = "25%"

        Select Case m_eLookupType
            Case ELOOKUPTYPE.lutClosureStatus
                'show the edit image button
                ofwTable.AddCellAsHTML ofwHTML.ImageButton(gk_sImagesPath & gk_sImageEdit, "Edit", gk_sLookupModify & "?" & ksParamLookupType & "=" & _
                    CStr(m_eLookupType) & "&" & ksParamModifyID & "=" & sFieldID & "&" & ksParamModifyValue & "=" & _
                    m_aspServer.URLEncode(sFieldValue)), typCellStyle
            Case ELOOKUPTYPE.lutLanguage
                If oFWRecordList.GetFieldValue(m_sCurrencyIndicatorField) <> 0 Then
                    'show the edit  & Remove image  buttons
                    ofwTable.AddCellAsHTML ofwHTML.ImageButton(gk_sImagesPath & gk_sImageEdit, "Edit", gk_sLookupModify & "?" & ksParamLookupType & "=" & _
                        CStr(m_eLookupType) & "&" & ksParamModifyID & "=" & sFieldID & "&" & ksParamModifyValue & "=" & _
                        m_aspServer.URLEncode(sFieldValue) & "&" & ksISOCode & "=" & m_aspServer.URLEncode(VariantStringAsString(oFWRecordList.GetFieldValue(gk_sDBCOLLanugageISOCode)))) _
                        & ofwHTML.ImageButton(gk_sImagesPath & gk_sImageRemove, "Remove", gk_sLookupRemove & "?" & ksParamLookupType & "=" & _
                        CStr(m_eLookupType) & "&" & ksParamRemoveID & "=" & sFieldID & "&" & ksParamRemoveValue & "=" & _
                        m_aspServer.URLEncode(sFieldValue)), typCellStyle
                Else
                    ''show the edit  & Undelete image  buttons
                    ofwTable.AddCellAsHTML ofwHTML.ImageButton(gk_sImagesPath & gk_sImageEdit, "Edit", gk_sLookupModify & "?" & ksParamLookupType & "=" & _
                        CStr(m_eLookupType) & "&" & ksParamModifyID & "=" & sFieldID & "&" & ksParamModifyValue & "=" & _
                        m_aspServer.URLEncode(sFieldValue) & "&" & ksISOCode & "=" & m_aspServer.URLEncode(VariantStringAsString(oFWRecordList.GetFieldValue(gk_sDBCOLLanugageISOCode)))) _
                        & ofwHTML.ImageButton(gk_sImagesPath & gk_sImageUndelete, "Undelete", gk_sLookupUndelete & "?" & ksParamLookupType & "=" & _
                        CStr(m_eLookupType) & "&" & ksParamRemoveID & "=" & sFieldID & "&" & ksParamRemoveValue & "=" & _
                        m_aspServer.URLEncode(sFieldValue)), typCellStyle
                End If
            Case Else
                If oFWRecordList.GetFieldValue(m_sCurrencyIndicatorField) <> 0 Then
                    'show the edit  & Remove image  buttons
                    ofwTable.AddCellAsHTML ofwHTML.ImageButton(gk_sImagesPath & gk_sImageEdit, "Edit", gk_sLookupModify & "?" & ksParamLookupType & "=" & _
                        CStr(m_eLookupType) & "&" & ksParamModifyID & "=" & sFieldID & "&" & ksParamModifyValue & "=" & _
                        m_aspServer.URLEncode(sFieldValue)) _
                        & ofwHTML.ImageButton(gk_sImagesPath & gk_sImageRemove, "Remove", gk_sLookupRemove & "?" & ksParamLookupType & "=" & _
                        CStr(m_eLookupType) & "&" & ksParamRemoveID & "=" & sFieldID & "&" & ksParamRemoveValue & "=" & _
                        m_aspServer.URLEncode(sFieldValue)), typCellStyle
                Else
                    ''show the edit  & Undelete image  buttons
                    ofwTable.AddCellAsHTML ofwHTML.ImageButton(gk_sImagesPath & gk_sImageEdit, "Edit", gk_sLookupModify & "?" & ksParamLookupType & "=" & _
                        CStr(m_eLookupType) & "&" & ksParamModifyID & "=" & sFieldID & "&" & ksParamModifyValue & "=" & _
                        m_aspServer.URLEncode(sFieldValue)) _
                        & ofwHTML.ImageButton(gk_sImagesPath & gk_sImageUndelete, "Undelete", gk_sLookupUndelete & "?" & ksParamLookupType & "=" & _
                        CStr(m_eLookupType) & "&" & ksParamRemoveID & "=" & sFieldID & "&" & ksParamRemoveValue & "=" & _
                        m_aspServer.URLEncode(sFieldValue)), typCellStyle
                End If
        End Select
        
        
        oFWRecordList.MoveNext
        ofwTable.EndRow
    Wend
    
    sHTML = sHTML & ofwTable.GetHTML
    
    sHTML = sHTML & gk_sHTMLLineBreak & ofwHTML.EndTag("form")
    
    'd12 Done button
    sHTML = sHTML & gk_sHTMLLineBreak & ofwHTML.ImageButton(gk_sImagesPath & gk_sImageDone, , _
         gk_sMaintenanceHome, gk_sMainFrame, "Go to maintenance home")
    'Add context sensitive help button
    sHTML = sHTML & HTMLToAddHelpButton(gk_sHelpLinkToMaintLookup)
    m_aspResponse.Write sHTML
    
    'transaction successfull ... set the transaction complete
    
    m_oObjectContext.SetComplete
    
    Set ofwHTML = Nothing
    Set ofwTable = Nothing
    Set oFWRecordList = Nothing
    Set oBOLookup = Nothing
    Exit Sub
ErrorHandler:

    'transaction failed ... Abort it
    m_oObjectContext.SetAbort
    ShowError sSource
End Sub

Private Function ShowLookupHeading() As String
' Description:
'
'returns the message based on the user's selection of lookup and type of operation
'
' Amendment history:
'
' d1   aga        9Jul1999     Created
' d2   aga       22Jul1999     Moved the string constants to the Resource file
'
    Dim sMessage As String
    Dim sSource As String
    
    sSource = "VWLookup.ShowLookupHeading"
    
    On Error GoTo ErrorHandler:
    
    'Select the message based on the lookup table type
    Select Case m_eLookupType
    
        Case ELOOKUPTYPE.lutLanguage
            sMessage = m_oFWAppData.GetDisplayInformation(ELookupMaintainResource.LMLanguageLookup)
        Case ELOOKUPTYPE.lutLegalStatus
            sMessage = m_oFWAppData.GetDisplayInformation(ELookupMaintainResource.LMLegalStatusLookup)
        Case ELOOKUPTYPE.lutPhysicalForm
            sMessage = m_oFWAppData.GetDisplayInformation(ELookupMaintainResource.LMPhysicalFormLookup)
        Case ELOOKUPTYPE.lutClosureStatus
            sMessage = m_oFWAppData.GetDisplayInformation(ELookupMaintainResource.LMClosureStatusLookup)
        Case ELOOKUPTYPE.lutAccessCondition
            sMessage = m_oFWAppData.GetDisplayInformation(ELookupMaintainResource.LMAccessConditionLookup)
        Case Else
            sMessage = m_oFWAppData.GetDisplayInformation(ELookupMaintainResource.LMInvalidLookup)
            RaiseError sSource, PEGeneralInternalError, sMessage
    End Select
        
    ShowLookupHeading = sMessage
     
    Exit Function
    
ErrorHandler:
    
    ConvertError sSource
    
End Function

Private Function ShowAddModifyStatus() As String
' Description:
'
'returns the message based on the BOLookup.Save Return code. This message will be
'useful to inform the user about the result of the operation.
'
' Amendment history:
'
' d1   aga        16Jul1999     Created
' d2   aga        20Jul1999     Changed the messages according to the changes in BOLookup.Save
    
    Dim sSource As String
    Dim sErrorMessage As String
    
    sSource = "VWLookup.ShowAddModifyStatus"
    On Error GoTo ErrorHandler
    
        Select Case m_eSaveReturnCode
           Case ELOOKUPSAVERETURNCODE.LMSaveDuplicateRecordExists
                sErrorMessage = "'" & m_sLookupValue & "' " & m_oFWAppData.GetDisplayInformation(ELookupMaintainResource.LMAlreadyExists)
            Case ELOOKUPSAVERETURNCODE.LMSaveDuplicateRecordExistsAsDeleted
                sErrorMessage = "'" & m_sLookupValue & "' " & m_oFWAppData.GetDisplayInformation(ELookupMaintainResource.LMDuplicateRecordExistAsDeleted)
            Case ELOOKUPSAVERETURNCODE.LMSaveNoLookupExistsWithID
                sErrorMessage = m_oFWAppData.GetDisplayInformation(ELookupMaintainResource.LMNoLookupExistsWithThis) & " '" & CStr(m_vLookupID) & "'"
            Case ELOOKUPSAVERETURNCODE.LMSaveLookupIDAlreadyExists
                sErrorMessage = m_oFWAppData.GetDisplayInformation(ELookupMaintainResource.LMCode) & " '" & CStr(m_vLookupID) & "' " & m_oFWAppData.GetDisplayInformation(ELookupMaintainResource.LMAlreadyExists)
            
            Case Else
                sErrorMessage = ""
        End Select
            
        ShowAddModifyStatus = sErrorMessage
    Exit Function

ErrorHandler:
    ConvertError sSource
End Function



Private Sub MapTableAndFields(t_eLookup As ELOOKUPTYPE)
' Description:
'
' maps the table name and the field names to the member variables which are
'  used in identifying the table and field names based on the selection
'
' Amendment history:
'
' d1   aga        08Jul1999     Created
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String         ' to store the error string
    
    sSource = "VWLookup.MapTableAndFields(" & t_eLookup & ")"
    
    'mapping table name and fields names to the member variables based on
    ' the lookup table type selected
    
    Select Case t_eLookup
    
        Case ELOOKUPTYPE.lutLanguage
        
            m_sLookupTableName = gk_sDBTABLELanguage
            
            m_sLookupIDField = gk_sDBCOLLanguageID
            m_sLookupValueField = gk_sDBCOLLanguageDescription
            m_sCurrencyIndicatorField = gk_sDBCOLCurrencyIndicator
            
            m_lLookupValueFieldLength = gk_iDBCOLLenLanguage
            m_sLookupValueHeading = m_oFWAppData.GetDisplayInformation(ELookupMaintainResource.LMLanguage)
            
        Case ELOOKUPTYPE.lutLegalStatus
        
            m_sLookupTableName = gk_sDBTABLELegalStatus
            
            m_sLookupIDField = gk_sDBCOLLegalStatusID
            m_sLookupValueField = gk_sDBCOLLegalStatusDescription
            m_sCurrencyIndicatorField = gk_sDBCOLCurrencyIndicator
            
            m_lLookupValueFieldLength = gk_iDBCOLLenLegalStatus
            m_sLookupValueHeading = m_oFWAppData.GetDisplayInformation(ELookupMaintainResource.LMLegalStatus)
            
            
        Case ELOOKUPTYPE.lutPhysicalForm
            
            m_sLookupTableName = gk_sDBTABLEPhysicalForm
            
            m_sLookupIDField = gk_sDBCOLPhysicalDescriptionFormID
            m_sLookupValueField = gk_sDBCOLPhysicalDescriptionFormDescription
            m_sCurrencyIndicatorField = gk_sDBCOLCurrencyIndicator
            
            m_lLookupValueFieldLength = gk_iDBCOLLenPhysicalForm
            m_sLookupValueHeading = m_oFWAppData.GetDisplayInformation(ELookupMaintainResource.LMPhysicalForm)
            
        Case ELOOKUPTYPE.lutClosureStatus
            
            m_sLookupTableName = gk_sDBTABLEClosureStatus
            
            m_sLookupIDField = gk_sDBCOLClosureStatus
            m_sLookupValueField = gk_sDBCOLClosureStatusDescription
            
            m_lLookupValueFieldLength = gk_iDBCOLLenClosureStatusDescription
            m_sLookupIDHeading = m_oFWAppData.GetDisplayInformation(ELookupMaintainResource.LMCode)
            m_sLookupValueHeading = m_oFWAppData.GetDisplayInformation(ELookupMaintainResource.LMClosureStatus)
            
        Case ELOOKUPTYPE.lutAccessCondition
            
            m_sLookupTableName = gk_sDBTABLEAccessCondition
            
            m_sLookupIDField = gk_sDBCOLAccessConditionID
            m_sLookupValueField = gk_sDBCOLAccessConditionDescription
            m_sCurrencyIndicatorField = gk_sDBCOLCurrencyIndicator
            
            m_lLookupValueFieldLength = gk_iDBCOLLenAccessCondition
            m_sLookupValueHeading = m_oFWAppData.GetDisplayInformation(ELookupMaintainResource.LMAccessCondition)
            
        Case Else
            RaiseError sSource, PEGeneralInternalError, "Invalid Lookup Table type : " & t_eLookup
            
    End Select
    
    Exit Sub
    
ErrorHandler:
  
        ConvertError sSource
End Sub




Private Sub CheckPermissions(t_eFunctionality As ESecurityCheck, _
                         t_sSource As String, _
                        t_sDescription As String)
'Subroutine that checks and raises an error if the user has no rights to
' the passed function
'
'd8    aga     26Oct1999     Copied from VWCatalogueBulkUpdate and modified
'd10   AGA     02Nov1999     Modified due to change in Check method
'd11   KSK     03Nov1999     Changed Checkpermissions method to used the parameter source
'                            value while raising error
    Dim eCheckRights As ESecurityCheckResult
    
    m_oboEditorialPermission.UserID = m_lLoggedUserID
    eCheckRights = m_oboEditorialPermission.Check(t_eFunctionality)
    
    If eCheckRights <> _
            ESecurityCheckResult.ESRSecurityCheckSuccessful Then
    
        RaiseError t_sSource, _
            GetErrorEnumForSecurityCheckResult(eCheckRights), _
            GetDescriptionForSecurityCheckError(GetErrorEnumForSecurityCheckResult(eCheckRights))
    End If
        
End Sub



