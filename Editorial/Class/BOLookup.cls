VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BOLookup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    BOLookup.cls
' System:       PROCAT, PRO
' Copyright:    (C) Quidnunc Limited
'
' Description:  Business Object class for Lookup Maintanence
'
' Amendment history:
'   d1      aga     07Jul1999   Created.
'   d2      aga     16Jul1999   Added return code from save method
'   d3      ksk     16Jul1999   Integrated
'   d4      aga     20Jul1999   Added properties & removed parameters of Save method, delete method and
'                               LookupList method and implemented business logic in the Save method itself
'                               instead of Stored Procedure
'                               Added enum to save method's return status
'   d5      aga     06Jul1999   Changes to eliminate currency indicator from closure status
'   d6      ksk     03Sep1999   Removed Physical condition as lookup
'   d7      aga     04Oct1999   Modified to remove currency indicator
'   d8      aga     26Oct1999   Taken d6 and modified to undelete a lookup
'   d9      aga     04Nov1999   Refreshing the reference data after
'                               adding/deleting/undeleting a lookup
'   d10     ksk     17Nov1999   Added closure type as a lookup type
'   d11     aga     17Nov1999   Added ISO_Code field to Language
'   d12     daj     15Jun2000   Tidied up code to always make ISO code lower case, and
'                               make the updating of the value case sensitive (ie user can now change
'                               case of the value)
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'Interfaces
Implements ObjectControl

'Enumerations for lookup table type
Public Enum ELOOKUPTYPE
    ' Add entry here for new lookup table
    lutLanguage = 1
    lutLegalStatus = 2
    lutPhysicalForm = 4
    lutClosureStatus = 5
    lutAccessCondition = 6
    lutClosureType = 7
End Enum

'Enumeration to save method's return status
Public Enum ELOOKUPSAVERETURNCODE
    LMSaveSuccessful = 1
    LMSaveDuplicateRecordExists = 2
    LMSaveDuplicateRecordExistsAsDeleted = 3
    LMSaveNoLookupExistsWithID = 4
    LMSaveLookupIDAlreadyExists = 5
End Enum



'private variables (Get/Let)
Private m_eLookupTableID As ELOOKUPTYPE 'LookupTableID
Private m_vLookupID As Variant          ' ID of the lookup record
Private m_sLookupValue As String        ' Value of the Lookup
Private m_vISOCode As Variant        ' Value of the ISOCode for language

'private variables to hold the field names and table name
Private m_sLookupTableName As String
Private m_sLookupIDField As String
Private m_sLookupValueField As String
Private m_sCurrencyIndicatorField As String

Private m_iLookupValueFieldLength As Integer 'Var to hold the field length in the database

'Constants to be used for Query type
Private Const mk_iSelect = 1            'For Select Query
Private Const mk_iDelete = 2            'For Update Query to mark the record for deletion
Private Const mk_iInsert = 3            'For Inserting new Lookup
Private Const mk_iUpdateID = 4            'For Updating record if ID exists already
Private Const mk_iFindID = 5            'For finding the ID is already existing
Private Const mk_iFindValue = 6         'For finding the Value is already existing
Private Const mk_iUpdateValue = 7            'For Updating record if value exists already
Private Const mk_iUnDelete = 8      'd8 For undeleting a lookup
Private Const mk_iFindCode = 9         'For finding the ISO code of a language

'd8. This is added to add an optional parameter to LookupList to return all records
'Because, others are already using this to return undeleted records
Private m_bSelectAllRecordIncludingMarkedAsDeleted As Boolean
'Private variables to hold object variables

Private m_oObjectContext As ObjectContext

'Properties
Public Property Let LookupTableID(ByVal t_eLookupTableType As ELOOKUPTYPE)
    m_eLookupTableID = t_eLookupTableType
End Property
Public Property Get LookupTableID() As ELOOKUPTYPE
    LookupTableID = m_eLookupTableID
End Property
Public Property Let LookupID(ByVal t_vLookupID As Variant)
    If t_vLookupID <> "" Then
        If m_eLookupTableID = ELOOKUPTYPE.lutClosureStatus Then
            m_vLookupID = Left(CStr(t_vLookupID), 1)
        Else
            m_vLookupID = t_vLookupID
        End If
    End If
    
End Property
Public Property Get LookupID() As Variant
    LookupID = m_vLookupID
End Property
Public Property Let LookupValue(ByVal t_sLookupValue As String)
    'Replace any single quotes with two single quotes for SQL purposes
    m_sLookupValue = ReplaceString(t_sLookupValue, "'", "''")
End Property
Public Property Get LookupValue() As String
    LookupValue = m_sLookupValue
End Property
Public Property Let LanguageISOCode(ByVal t_vLanguageISOCode As Variant)
'   d12     daj     15Jun2000   Tidied up code to always make ISO code lower case, and
'                               make the updating of the value case sensitive (ie user can now change
'                               case of the value)
    'Replace any single quotes with two single quotes for SQL purposes
    If Not IsNull(t_vLanguageISOCode) Then
        m_vISOCode = CVar(ReplaceString(StrConv(CStr(t_vLanguageISOCode), vbLowerCase), "'", "''"))
    End If
End Property
Public Property Get LanguageISOCode() As Variant
    LanguageISOCode = m_vISOCode
End Property

'Interfaces
Private Sub ObjectControl_Activate()
    Initialise
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    Terminate
End Sub

Private Sub Class_Initialize()
    Call Initialise
End Sub

Private Sub Class_Terminate()
    Call Terminate
End Sub
'Class methods
Private Sub Initialise()

' Description:
'
' Initialises the objects
' Amendment history:
'
' d1   aga        13Jul1999     Created
'
    Dim sSource As String
    sSource = "BOLookup.Initialise"
    On Error GoTo ErrorHandler
    
    Set m_oObjectContext = GetObjectContext()
    If m_oObjectContext Is Nothing Then
        RaiseError sSource, peOCUnableToGetObjectContext, "Objectcontext is nothing"
    End If
    
    Exit Sub
    
ErrorHandler:
    ConvertError sSource
End Sub
Private Sub Terminate()
' Description:
'
' Releases all the objects
' Amendment history:
'
' d1   aga        13Jul1999     Created
'
    'Release the objects
    Set m_oObjectContext = Nothing

End Sub

Public Function LookupList(Optional t_bDisplayAllRecords As Boolean = False) As FWRecordList
' Description:
'
'Gets the record list from the database based Lookup table type
'
' Amendment history:
'
' d1   aga        08Jul1999     Created
' d2   aga        20Jul1999     Changed the interface(Removed the parameter and use the properties)
' d8   aga        26Oct1999    Added optional parameter to return all records / undeleted records
'

    Dim sSource As String
    Dim ofwDBHelper As FWDBHelper
    Dim sSQLString As String

    sSource = "BOLookup.LookupList()"

    On Error GoTo ErrorHandler:

    Set ofwDBHelper = New FWDBHelper

    'd8 set the property
    If t_bDisplayAllRecords = True Then
        m_bSelectAllRecordIncludingMarkedAsDeleted = True
    Else
        m_bSelectAllRecordIncludingMarkedAsDeleted = False
    End If
    'Get the sql query string
    sSQLString = GenerateTheQuery(m_eLookupTableID)

    Set LookupList = ofwDBHelper.FindBySql(sSQLString)

    'release the objects
    Set ofwDBHelper = Nothing
    
    Exit Function
    
ErrorHandler:
    
    ConvertError sSource

End Function

Public Sub Delete()
' Description:
'
'Deletes the lookup value from the lookup table.
'  Or sets the currency_indicator to "true" if lookup value already exists
'
' Amendment history:
'
' d1    aga     09Jul1999       Created
' d2    aga     20Jul1999       Changed the interface(Removed the parameter and use the properties)
' d9    aga     04Nov1999       Refreshing the reference data after deleting a lookup
'
'
    Dim sSource As String
    Dim ofwDBHelper As FWDBHelper
    Dim sSQLString As String
    Dim ofwRefdata As FWReferenceData
    
    sSource = "BOLookup.Delete"
    On Error GoTo ErrorHandler
    
    
    sSQLString = GenerateTheQuery(m_eLookupTableID, mk_iDelete)
    
    Set ofwDBHelper = New FWDBHelper
    
    'Executes the Query string returned by GenerateTheQuery
    ofwDBHelper.Execute sSQLString
    
    'Now refresh the reference data
    Set ofwRefdata = New FWReferenceData
    
    Select Case m_eLookupTableID
        Case ELOOKUPTYPE.lutAccessCondition
            ofwRefdata.Refresh (EGroup.gAccessCondition)
        Case ELOOKUPTYPE.lutAccessCondition
            ofwRefdata.Refresh (EGroup.gClosureStatus)
        Case ELOOKUPTYPE.lutLanguage
            ofwRefdata.Refresh (EGroup.gLanguage)
        Case ELOOKUPTYPE.lutLegalStatus
            ofwRefdata.Refresh (EGroup.gLegalStatus)
        Case ELOOKUPTYPE.lutPhysicalForm
            ofwRefdata.Refresh (EGroup.gPhysicalForm)
    End Select
    
    'release the objects
    Set ofwDBHelper = Nothing
    Set ofwRefdata = Nothing
    Exit Sub
    
ErrorHandler:
    
    ConvertError sSource
End Sub

Public Sub UnDelete()
' Description:
'
'UnDeletes the lookup value
'
' Amendment history:
'
' d8    aga     26Oct1999       Created
' d9    aga     04Nov1999       Refreshing the reference data after undeleting a lookup
'
    Dim sSource As String
    Dim ofwDBHelper As FWDBHelper
    Dim sSQLString As String
    Dim ofwRefdata As FWReferenceData
    
    sSource = "BOLookup.UnDelete"
    On Error GoTo ErrorHandler
    
    
    sSQLString = GenerateTheQuery(m_eLookupTableID, mk_iUnDelete)
    
    Set ofwDBHelper = New FWDBHelper
    
    'Executes the Query string returned by GenerateTheQuery
    ofwDBHelper.Execute sSQLString
    
    Set ofwRefdata = New FWReferenceData
    
    Select Case m_eLookupTableID
        Case ELOOKUPTYPE.lutAccessCondition
            ofwRefdata.Refresh (EGroup.gAccessCondition)
        Case ELOOKUPTYPE.lutAccessCondition
            ofwRefdata.Refresh (EGroup.gClosureStatus)
        Case ELOOKUPTYPE.lutLanguage
            ofwRefdata.Refresh (EGroup.gLanguage)
        Case ELOOKUPTYPE.lutLegalStatus
            ofwRefdata.Refresh (EGroup.gLegalStatus)
        Case ELOOKUPTYPE.lutPhysicalForm
            ofwRefdata.Refresh (EGroup.gPhysicalForm)
    End Select
    
    'release the objects
    Set ofwDBHelper = Nothing
    Set ofwRefdata = Nothing

    Exit Sub
    
ErrorHandler:

    ConvertError sSource
End Sub



Public Function Save() As ELOOKUPSAVERETURNCODE

' Description:
'
'Adds or modifies a lookup based on the Lookup Value and/or Lookup ID properties
'
'Checks if the ID already exists
'  checks if the currency status is true
'    if both are true returns a number based on the message
'   if status is false sets the status to true
'else inserts the record
'
'
' Amendment history:
'
' d1    aga     09Jul1999       Created
' d2    aga     16Jul1999       Added code for returning a recordlist which is used
'                               hold the recordset returned by stored procedure.
' d3    aga     20Jul1999       Changed the interface and implementation,
'                               Remove the stored procedure and added code in the class
'                               Changed the return type to Enum type
' d4    aga     06Jul1999       Changes to eliminate currency indicator from closure status
' d9    aga     04Nov1999       Refreshing the reference data after adding a lookup
'   d12     daj     15Jun2000   Tidied up code to always make ISO code lower case, and
'                               make the updating of the value case sensitive (ie user can now change
'                               case of the value)

    Dim sSource As String
    Dim ofwDBHelper As FWDBHelper
    Dim sSQLString As String
    Dim eSaveReturnCode As ELOOKUPSAVERETURNCODE
    
    Dim cSingleCharID As String
    Dim oFWRecordList As FWRecordList
    Dim ofwRefdata As FWReferenceData
    
    sSource = "BOLookup.Save"
    On Error GoTo ErrorHandler
    
    GetValueFieldLength (m_eLookupTableID)
    
    
    Set oFWRecordList = New FWRecordList
    Set ofwDBHelper = New FWDBHelper
              
    eSaveReturnCode = ELOOKUPSAVERETURNCODE.LMSaveSuccessful  'Initialise to successful return
    
    
    'IF NOT closure_status table
    If Not m_eLookupTableID = lutClosureStatus Then
    
        'check IF id is set
        If Not IsEmpty(m_vLookupID) Then
        
            'check IF the ID is already present in the database
            sSQLString = GenerateTheQuery(m_eLookupTableID, mk_iFindID)
            Set oFWRecordList = ofwDBHelper.FindBySql(sSQLString)
    
            If (oFWRecordList.RecordCount > 0) Then
            
                'Find if the value already exists
                sSQLString = GenerateTheQuery(m_eLookupTableID, mk_iFindValue)
                Set oFWRecordList = ofwDBHelper.FindBySql(sSQLString)
                
                If oFWRecordList.RecordCount > 0 Then
                    ' d12 - Record already exist with that value, so throw error
                    RaiseError sSource, peUSERCannotAddDuplicateLookupValue, "There is already a lookup term with that value"
                End If
            
                ' Is this a language?
                If m_eLookupTableID = lutLanguage Then
                    ' Check to see if the ISO code exists as well
                    sSQLString = GenerateTheQuery(m_eLookupTableID, mk_iFindCode)
                    Set oFWRecordList = ofwDBHelper.FindBySql(sSQLString)
                    If oFWRecordList.RecordCount > 0 Then
                        RaiseError sSource, peUSERCannotAddDuplicateLookupValue, "There is already a lookup term with that ISO code"
                    End If
                End If
                
                'Update the record with the new value (SET THE STATUS TRUE)
                sSQLString = GenerateTheQuery(m_eLookupTableID, mk_iUpdateID)
                ofwDBHelper.Execute (sSQLString)
    
            Else
        
                eSaveReturnCode = ELOOKUPSAVERETURNCODE.LMSaveNoLookupExistsWithID
            End If
    
        Else
            'check IF Value already exists
            sSQLString = GenerateTheQuery(m_eLookupTableID, mk_iFindValue)
            Set oFWRecordList = ofwDBHelper.FindBySql(sSQLString)
            If (oFWRecordList.RecordCount > 0) Then
                ' d12 - Already a value like that, so throw an error
                eSaveReturnCode = ELOOKUPSAVERETURNCODE.LMSaveDuplicateRecordExists
                RaiseError sSource, peUSERCannotAddDuplicateLookupValue, "There is already a lookup term with that value"
            End If
            
            ' Is this a language?
            If m_eLookupTableID = lutLanguage Then
                ' Check to see if the ISO code is valid
                If m_vISOCode = "" Then
                    RaiseError sSource, peMandatoryValuesMissing, "You must supply an ISO code"
                End If
                
                ' Check to see if the ISO code exists as well
                sSQLString = GenerateTheQuery(m_eLookupTableID, mk_iFindCode)
                Set oFWRecordList = ofwDBHelper.FindBySql(sSQLString)
                If oFWRecordList.RecordCount > 0 Then
                    RaiseError sSource, peUSERCannotAddDuplicateLookupValue, "There is already a lookup term with that ISO code"
                End If
            End If
            
            'If passed al that, then insert the record
            sSQLString = GenerateTheQuery(m_eLookupTableID, mk_iInsert)
            ofwDBHelper.Execute (sSQLString)
        End If
    
    'If closure status table
    Else
        'check IF the ID is present
        sSQLString = GenerateTheQuery(m_eLookupTableID, mk_iFindID)
        Set oFWRecordList = ofwDBHelper.FindBySql(sSQLString)
        
        
        If (oFWRecordList.RecordCount > 0) Then
            'Find if the value already exists
            sSQLString = GenerateTheQuery(m_eLookupTableID, mk_iFindValue)
            Set oFWRecordList = ofwDBHelper.FindBySql(sSQLString)
            
            If oFWRecordList.RecordCount > 0 Then
                'Record already exist with that value
                eSaveReturnCode = ELOOKUPSAVERETURNCODE.LMSaveDuplicateRecordExists
            Else
                'Update the record with the new value
                sSQLString = GenerateTheQuery(m_eLookupTableID, mk_iUpdateID)
                ofwDBHelper.Execute (sSQLString)
                
            End If

        Else
            'Find if the value already exists
            sSQLString = GenerateTheQuery(m_eLookupTableID, mk_iFindValue)
            Set oFWRecordList = ofwDBHelper.FindBySql(sSQLString)
            If oFWRecordList.RecordCount > 0 Then
                
                    eSaveReturnCode = ELOOKUPSAVERETURNCODE.LMSaveDuplicateRecordExists
                
            Else
                'Insert the record
                sSQLString = GenerateTheQuery(m_eLookupTableID, mk_iInsert)
                ofwDBHelper.Execute (sSQLString)
            End If
        End If
    
    End If
        
    'Now refresh the reference data
    Set ofwRefdata = New FWReferenceData
    
    Select Case m_eLookupTableID
        Case ELOOKUPTYPE.lutAccessCondition
            ofwRefdata.Refresh (EGroup.gAccessCondition)
        Case ELOOKUPTYPE.lutAccessCondition
            ofwRefdata.Refresh (EGroup.gClosureStatus)
        Case ELOOKUPTYPE.lutLanguage
            ofwRefdata.Refresh (EGroup.gLanguage)
        Case ELOOKUPTYPE.lutLegalStatus
            ofwRefdata.Refresh (EGroup.gLegalStatus)
        Case ELOOKUPTYPE.lutPhysicalForm
            ofwRefdata.Refresh (EGroup.gPhysicalForm)
    End Select
    
    'release the objects
    Set oFWRecordList = Nothing
    Set ofwDBHelper = Nothing
    Set ofwRefdata = Nothing
    
    Save = eSaveReturnCode
    
    Exit Function
    
ErrorHandler:
    
    ConvertError sSource

End Function

Private Function GenerateTheQuery(t_eLookup As ELOOKUPTYPE, Optional t_iQueryType As Integer = mk_iSelect) As String
' Description:
'
' maps the table name and the field names to the member variables
'  and generates the query and returns the query as a string
'
' Amendment history:
'
' d1   aga      08Jul1999       Created
' d2   aga      20Jul1999       Added queries for insert, findID, findValue, UpdateID, UpdateValue
' d3   aga      06Jul1999       Changes made to eliminate currency indicator from closure status table
' d8   aga      26Oct1999       Added undelete query, modified Select query
' d11  aga      17Nov1999       Added ISO_Code field to Language
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String         ' to store the error string
    Dim sSQLString As String
    
    sSource = "BOLookup.GenerateTheQuery(" & t_eLookup & t_iQueryType & ")"
    
    'mapping table and field names based on the users selection i.e., t_eLookup
    Select Case t_eLookup
    
        Case ELOOKUPTYPE.lutLanguage
        
            m_sLookupTableName = gk_sDBTABLELanguage
            m_sLookupIDField = gk_sDBCOLLanguageID
            m_sLookupValueField = gk_sDBCOLLanguageDescription
            m_sCurrencyIndicatorField = gk_sDBCOLCurrencyIndicator
            m_iLookupValueFieldLength = gk_iDBCOLLenLanguage
            
        Case ELOOKUPTYPE.lutLegalStatus
        
            m_sLookupTableName = gk_sDBTABLELegalStatus
            m_sLookupIDField = gk_sDBCOLLegalStatusID
            m_sLookupValueField = gk_sDBCOLLegalStatusDescription
            m_sCurrencyIndicatorField = gk_sDBCOLCurrencyIndicator
            m_iLookupValueFieldLength = gk_iDBCOLLenLegalStatus
            
            
        Case ELOOKUPTYPE.lutPhysicalForm
            
            m_sLookupTableName = gk_sDBTABLEPhysicalForm
            m_sLookupIDField = gk_sDBCOLPhysicalDescriptionFormID
            m_sLookupValueField = gk_sDBCOLPhysicalDescriptionFormDescription
            m_sCurrencyIndicatorField = gk_sDBCOLCurrencyIndicator
            m_iLookupValueFieldLength = gk_iDBCOLLenPhysicalForm
            
        Case ELOOKUPTYPE.lutClosureStatus
            
            m_sLookupTableName = gk_sDBTABLEClosureStatus
            m_sLookupIDField = gk_sDBCOLClosureStatus
            m_sLookupValueField = gk_sDBCOLClosureStatusDescription
            m_iLookupValueFieldLength = gk_iDBCOLLenClosureStatusDescription
            
        Case ELOOKUPTYPE.lutAccessCondition
            
            m_sLookupTableName = gk_sDBTABLEAccessCondition
            m_sLookupIDField = gk_sDBCOLAccessConditionID
            m_sLookupValueField = gk_sDBCOLAccessConditionDescription
            m_sCurrencyIndicatorField = gk_sDBCOLCurrencyIndicator
            m_iLookupValueFieldLength = gk_iDBCOLLenAccessCondition
            
        Case ELOOKUPTYPE.lutClosureType
            
            m_sLookupTableName = gk_sDBTABLEClosureType
            m_sLookupIDField = gk_sDBCOLClosureType
            m_sLookupValueField = gk_sDBCOLClosureTypeDescription
            m_sCurrencyIndicatorField = gk_sDBCOLCurrencyIndicator
            m_iLookupValueFieldLength = gk_iDBCOLLenAccessCondition
            
        Case Else
            RaiseError sSource, PEGeneralInternalError, "Invalid Lookup Table type : " & t_eLookup
            
    End Select
        
    ' Gererate the query string
    Select Case t_iQueryType
        Case mk_iSelect
        
            'Generate the Select query to retrieve those records whose currency indicator is not false
            If t_eLookup = ELOOKUPTYPE.lutClosureStatus Or _
               t_eLookup = ELOOKUPTYPE.lutClosureType Then
                sSQLString = "SELECT * FROM " & m_sLookupTableName & _
                    " ORDER BY " & m_sLookupValueField
            Else
                
                If m_bSelectAllRecordIncludingMarkedAsDeleted = True Then
                    'All records including marked as deleted records
                    sSQLString = "SELECT * FROM " & m_sLookupTableName & _
                        " ORDER BY " & m_sLookupValueField
                Else
                    'Undeleted records only
                    sSQLString = "SELECT * FROM " & m_sLookupTableName & " WHERE " & m_sCurrencyIndicatorField & "<> 0" & _
                    " ORDER BY " & m_sLookupValueField
                End If
                
            End If
                 
        Case mk_iDelete
        
            'Delete (Update) query
            ' Marks the record for deletion i.e., set the currency indicator to False
            If t_eLookup = ELOOKUPTYPE.lutClosureStatus Then
                'disabled deletion (marking)
                'sSQLString = "UPDATE " & m_sLookupTableName & _
                '     " SET " & m_sCurrencyIndicatorField & "=0 WHERE " & _
                '     m_sLookupIDField & "=" & Quotes(CStr(m_vLookupID))
            Else
                sSQLString = "UPDATE " & m_sLookupTableName & _
                     " SET " & m_sCurrencyIndicatorField & "=0 WHERE " & _
                     m_sLookupIDField & "=" & CInt(m_vLookupID)
            End If
        
        Case mk_iUnDelete       'd8
        
            'UnDelete (Update) query
            ' Marks the record for undeletion i.e., set the currency indicator to true
            If t_eLookup = ELOOKUPTYPE.lutClosureStatus Then
                'No deletion for closure status
            Else
                sSQLString = "UPDATE " & m_sLookupTableName & _
                     " SET " & m_sCurrencyIndicatorField & "=1 WHERE " & _
                     m_sLookupIDField & "=" & CInt(m_vLookupID)
            End If
            
        Case mk_iInsert
            
            'Generate query to insert a new lookup
            If t_eLookup = ELOOKUPTYPE.lutClosureStatus Then
                sSQLString = "INSERT INTO " & m_sLookupTableName & _
                    " ( " & m_sLookupIDField & "," & m_sLookupValueField & _
                     ") " & " VALUES " & _
                    " ( " & Quotes(Left(CStr(m_vLookupID), 1)) & "," & _
                    Quotes(RTrim(Left(Trim(m_sLookupValue), m_iLookupValueFieldLength))) & ")"
            Else
                If t_eLookup = ELOOKUPTYPE.lutLanguage Then
                    If Not IsMissing(m_vISOCode) Then
                        sSQLString = "INSERT INTO " & m_sLookupTableName & _
                            " ( " & m_sLookupValueField & "," & m_sCurrencyIndicatorField & ", " & gk_sDBCOLLanugageISOCode & ") " & _
                            " VALUES " & " ( " & Quotes(RTrim(Left(Trim(m_sLookupValue), m_iLookupValueFieldLength))) & ",1 ," _
                            & Quotes(VariantStringAsString(m_vISOCode)) & ")"
                    Else
                        sSQLString = "INSERT INTO " & m_sLookupTableName & _
                            " ( " & m_sLookupValueField & "," & m_sCurrencyIndicatorField & ", " & gk_sDBCOLLanugageISOCode & ") " & _
                            " VALUES " & " ( " & Quotes(RTrim(Left(Trim(m_sLookupValue), m_iLookupValueFieldLength))) & ",1 " _
                            & ", Null " & ")"
                    End If
                        
                Else
                    sSQLString = "INSERT INTO " & m_sLookupTableName & _
                        " ( " & m_sLookupValueField & "," & m_sCurrencyIndicatorField & ") " & _
                        " VALUES " & " ( " & Quotes(RTrim(Left(Trim(m_sLookupValue), m_iLookupValueFieldLength))) & ",1 )"
                End If
            End If
        Case mk_iUpdateID
            'Generate the query to update (modify) an existing record
            'And also set the currency indicator to true
            If t_eLookup = ELOOKUPTYPE.lutClosureStatus Then
                sSQLString = "UPDATE " & m_sLookupTableName & _
                     " SET " & m_sLookupValueField & "=" & _
                     Quotes(RTrim(Left(Trim(m_sLookupValue), m_iLookupValueFieldLength))) & " WHERE " & _
                     m_sLookupIDField & "=" & Quotes(CStr(m_vLookupID))
            Else
                If t_eLookup = ELOOKUPTYPE.lutLanguage Then
                    If Not IsMissing(m_vISOCode) Then
                        sSQLString = "UPDATE " & m_sLookupTableName & _
                        " SET " & m_sCurrencyIndicatorField & "=1 , " & m_sLookupValueField & "=" & _
                        Quotes(RTrim(Left(Trim(m_sLookupValue), m_iLookupValueFieldLength))) _
                        & ", " & gk_sDBCOLLanugageISOCode & "=" & Quotes(VariantStringAsString(m_vISOCode)) & " WHERE " & _
                        m_sLookupIDField & "=" & CInt(m_vLookupID)
                    Else
                        sSQLString = "UPDATE " & m_sLookupTableName & _
                        " SET " & m_sCurrencyIndicatorField & "=1 , " & m_sLookupValueField & "=" & _
                        Quotes(RTrim(Left(Trim(m_sLookupValue), m_iLookupValueFieldLength))) & " WHERE " & _
                        m_sLookupIDField & "=" & CInt(m_vLookupID)
                    End If
                        
                Else
                    sSQLString = "UPDATE " & m_sLookupTableName & _
                        " SET " & m_sCurrencyIndicatorField & "=1 , " & m_sLookupValueField & "=" & _
                        Quotes(RTrim(Left(Trim(m_sLookupValue), m_iLookupValueFieldLength))) & " WHERE " & _
                        m_sLookupIDField & "=" & CInt(m_vLookupID)
                End If
                
            End If
        Case mk_iFindID
        
            'Generate the Select query to find the ID is already existing
            If t_eLookup = ELOOKUPTYPE.lutClosureStatus Then
                sSQLString = "SELECT * FROM " & m_sLookupTableName & " WHERE " & m_sLookupIDField & _
                    "=" & Quotes(CStr(m_vLookupID))
            Else
                sSQLString = "SELECT * FROM " & m_sLookupTableName & " WHERE " & m_sLookupIDField & _
                    "=" & CInt(m_vLookupID)
            End If
            
                 
        Case mk_iFindValue
        
            'Generate the Select query to find the Value is already existing
            If t_eLookup = ELOOKUPTYPE.lutClosureStatus Then
                sSQLString = "SELECT " & m_sLookupIDField & "," & _
                    m_sLookupValueField & _
                    " FROM " & m_sLookupTableName & " WHERE LOWER(" & m_sLookupValueField & _
                    ")=" & Quotes(StrConv((Trim(m_sLookupValue)), vbLowerCase)) & _
                    " AND NOT " & m_sLookupIDField & " = " & Quotes(CStr(m_vLookupID))
            Else
                sSQLString = "SELECT " & m_sLookupIDField & "," & _
                    m_sLookupValueField & "," & m_sCurrencyIndicatorField & _
                    " FROM " & m_sLookupTableName & " WHERE LOWER(" & m_sLookupValueField & _
                    ")=" & Quotes(StrConv(Trim(m_sLookupValue), vbLowerCase)) & _
                    " AND NOT " & m_sLookupIDField & " = " & Quotes(CStr(m_vLookupID))
            End If
        Case mk_iUpdateValue
            'Generate the query to update (modify) an existing record
            'with the given value and also set the currency indicator to true
            If t_eLookup = ELOOKUPTYPE.lutClosureStatus Then
                'disabled
                'sSQLString = "UPDATE " & m_sLookupTableName & _
                 '    " SET " & m_sCurrencyIndicatorField & "=1 " & " WHERE LOWER(" & _
                 '    m_sLookupValueField & ")=" & _
                 '    Quotes(StrConv(RTrim(Left(Trim(m_sLookupValue), m_iLookupValueFieldLength)), vbLowerCase))
            Else
                If t_eLookup = ELOOKUPTYPE.lutLanguage Then
                    If Not IsMissing(m_vISOCode) Then
                        sSQLString = "UPDATE " & m_sLookupTableName & _
                            " SET " & m_sCurrencyIndicatorField & "=1 " & _
                             ", " & gk_sDBCOLLanugageISOCode & "=" & Quotes(VariantStringAsString(m_vISOCode)) & _
                            " WHERE LOWER(" & m_sLookupValueField & ")=" & _
                            Quotes(StrConv(RTrim(Left(Trim(m_sLookupValue), m_iLookupValueFieldLength)), vbLowerCase))
                           
                    Else
                        sSQLString = "UPDATE " & m_sLookupTableName & _
                        " SET " & m_sCurrencyIndicatorField & "=1 " & _
                        " WHERE LOWER(" & m_sLookupValueField & ")=" & _
                        Quotes(StrConv(RTrim(Left(Trim(m_sLookupValue), m_iLookupValueFieldLength)), vbLowerCase))
                    End If
                        
                Else
                    sSQLString = "UPDATE " & m_sLookupTableName & _
                        " SET " & m_sCurrencyIndicatorField & "=1 " & _
                        " WHERE LOWER(" & m_sLookupValueField & ")=" & _
                        Quotes(StrConv(RTrim(Left(Trim(m_sLookupValue), m_iLookupValueFieldLength)), vbLowerCase))
                End If
                
            End If
            
        Case mk_iFindCode
            ' d12 - Find an existing ISO code
            sSQLString = "SELECT * FROM " & m_sLookupTableName & " WHERE " & _
                gk_sDBCOLLanugageISOCode & "=" & Quotes(VariantStringAsString(m_vISOCode)) & _
                " AND NOT " & m_sLookupIDField & " = " & Quotes(CStr(m_vLookupID))
        Case Else
            RaiseError sSource, PEGeneralInternalError, "Invalid Query type " & t_iQueryType
    End Select
    
    GenerateTheQuery = sSQLString
    
    Exit Function
    
ErrorHandler:

        ConvertError sSource
End Function

Private Sub GetValueFieldLength(t_eLookup As ELOOKUPTYPE)
' Description:
'
'To assign the field length of the Lookup Value field
'
' Amendment history:
'
' d1   aga        13Jul1999     Created

    Dim sSource As String
    
    sSource = "BOLookup.GetVAlueFieldLength (" & t_eLookup & ")"
    On Error GoTo ErrorHandler
    Select Case t_eLookup
    
        Case ELOOKUPTYPE.lutLanguage
        
            m_iLookupValueFieldLength = gk_iDBCOLLenLanguage
            
        Case ELOOKUPTYPE.lutLegalStatus
        
            m_iLookupValueFieldLength = gk_iDBCOLLenLegalStatus
            
            
        Case ELOOKUPTYPE.lutPhysicalForm
            
            m_iLookupValueFieldLength = gk_iDBCOLLenPhysicalForm
            
        Case ELOOKUPTYPE.lutClosureStatus
            
            m_iLookupValueFieldLength = gk_iDBCOLLenClosureStatusDescription
            
        Case ELOOKUPTYPE.lutAccessCondition
            
            m_iLookupValueFieldLength = gk_iDBCOLLenAccessCondition
            
        Case Else
            RaiseError sSource, PEGeneralInternalError, "Invalid Lookup Table type : " & t_eLookup
            
    End Select
    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub





