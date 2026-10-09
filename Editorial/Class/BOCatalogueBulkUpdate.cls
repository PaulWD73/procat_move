VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BOCatalogueBulkUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    BOCatalogueBulkUpdate.cls
' System:       PROCAT, PRO
' Copyright:    (C) Quidnunc Limited
'
' Description:
'
' Business object to bulk update the properties of the a catalogue
'
' Amendment history:
'
' d1    ksk     07Ju11999       Created.
' d2    ksk     14Jul1999       Bulkupdate to change the editset_entry_type Id
'                               flag to modified. Added constant for the entry
'                               type values.
'                               Changed All methods that save strings from using
'                               Quotes method to QuoteValueWithQuotes method
'                               totake care of the double quote occuring in a value
' d3    ksk     22Jul1999       Added datatype and length checks for the
' d4    ksk     04Aug1999       Changed usage of catalogue element matrix from local to
'                               private variable
' d5    ksk     13Aug1999       Changes made to reflect a alphanumeric ID for closure status
' d6    ksk     17Aug1999       Changes to bulkupdate to list the catalogues reflected by
'                               the update
' d7    ksk     03Sep1999       Changes made after the physical condition field was
'                               changed to a text field from lookup
' d8    ksk     25Oct1999       Changes made after the editset entry type status field
'                               was made as a Enumeration and changes for closure status
'                               and access condition change.
' d9    ksk     18Nov1999       Added the extra join that was missing to link the extenstion
'                               and main table for the list of records affected
' d10   ksk     19Nov1999       Changed the Piece and Item levels language update to use
'                               join to extension table
' d11   ksk     22Nov1999       Changed the condition in the update for imported matrial.
'                               Used the wrong constant
' d12   ksk     22Nov1999       Added a max constant to be used by all the Findbysql
'                               To return the affected records
' d13   ksk     09Dec1999       Fixed a bug in raising error for length violation in
'                               Checkifpropertyset
' d14   ksk     29Dec1999       Changed the usage of open_date in Piece and Item level
' d15   wdp     24Feb2000       Added transient storage of levels to update
' d16   tmj     28Feb2000       Wrapped date values in VariantDateAsDDMMYYYY so can be
'                               understood by the Database
' d17   tmj     14Mar2000       Modified so leaflets can be added to Publication notes and
'                               Unpublished finding aids
' d18   tmj     27Mar2000       Ensured BulkUpdateDivision selected only one copy of each
'                               division
' d19   alm     12Feb2001       Updates the start / end dates automatically
'                               when a covering date is updated
' d20   alm     02Oct2001       Added support for accesioned, orderable, LCI and
'                               extract as transfer properties + creator, deposit and
'                               originals link items
' d21   fjd     05Feb2003       Renamed LCI flag to AR
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'TO DO
'EAD editing for Title, Scope and Content and arrangement

Private Const mk_iMaxNoOfRecords As Integer = 32000


Private m_ofwUpdate As FWDBHelper
Private m_oBOCatalogueElementMatrix As BOCatalogueElementMatrix
Private m_oBOCatalogueDBHelper As BOCatalogueDBHelper

Private m_lEditsetID As Long
Private m_eUpdateField As ECATALOGUEPROPERTIES

'Stores the primary value for all fields. Unknown description for the link data items
' and for closure type + closure code this stores the closure type value
Private m_vUpdateValue1 As Variant

'Used by only closure type + closure code fields.
' For closure type + closure code this stores the closure code value
Private m_vUpdateValue2 As Variant

'Used by link data item and closure type + closure code fields.
' For closure type + closure code this stores the closure code value
' For Link data item this store the ID, Description collection
Private m_colUpdateValue3 As Collection

'Used by only link data item fields to store the second collection list
' This store the ID, Description collection
Private m_colUpdateValue4 As Collection

'Bit map to indicate which levels we are updating (for transient storage)
Private m_lLevels As Long

Private Const mk_lLettercodeVal As Long = 1
Private Const mk_lDivisionVal As Long = 2
Private Const mk_lClassVal As Long = 4
Private Const mk_lHeaderVal As Long = 8
Private Const mk_lSubheaderVal As Long = 16
Private Const mk_lPieceVal As Long = 32
Private Const mk_lItemVal As Long = 64


Implements PROCATTransientInterface




Public Function BulkUpdateLevels(t_bSelectedLevels() As Boolean) As Collection
' Method : BulkUpdateLevels
'  t_bSelectedLevels takes in a boolean array with the number of levels as the number of rows
'                    if the value corresponding to the index(which stands for the level is true
'                    then this level was selected for bulk updation.
'
' d6    ksk     changed the function to return a recordlist that will be a list of catalogues
'               affected by the change.
'
On Error GoTo ErrHandler
    Dim sSource As String
    Dim ofwAffectedCatalogues As FWRecordList
    Dim i As Integer
    
    sSource = "VWBulkUpdate.BulkUpdate"
    If Not CheckIfPropertiesSet Then
        Exit Function
    End If
    
    

    For i = 1 To gk_iNumberOfLevels
        If t_bSelectedLevels(i) Then
            Select Case i
                Case ECATALOGUELEVEL.ctLetterCode
                    Set ofwAffectedCatalogues = BulkUpdateLettercode
                Case ECATALOGUELEVEL.ctDivision
                    Set ofwAffectedCatalogues = BulkUpdateDivision
                Case ECATALOGUELEVEL.ctClass
                    Set ofwAffectedCatalogues = BulkUpdateClass
                Case ECATALOGUELEVEL.ctHeaders
                    Set ofwAffectedCatalogues = BulkUpdateHeader
                Case ECATALOGUELEVEL.ctSubHeaders
                    Set ofwAffectedCatalogues = BulkUpdateSubheader
                Case ECATALOGUELEVEL.ctPiece
                    Set ofwAffectedCatalogues = BulkUpdatePiece
                Case ECATALOGUELEVEL.ctItem
                    Set ofwAffectedCatalogues = BulkUpdateItem
            End Select
            'Add the returned list of catalogues to the existing one
            ' to return to the client
            If Not ofwAffectedCatalogues Is Nothing Then
                If BulkUpdateLevels Is Nothing Then
                    Set BulkUpdateLevels = New Collection
                End If
                'Loop adding all the records to return
                While Not ofwAffectedCatalogues.IsAtEnd
                    BulkUpdateLevels.Add ofwAffectedCatalogues.GetRecord
                    ofwAffectedCatalogues.MoveNext
                Wend
            End If
        
        End If
        
    Next i
    
    Exit Function
ErrHandler:
    ConvertError sSource
End Function


Private Function BulkUpdateLettercode() As FWRecordList
' Bulkupdates the value of the set property of the Lettercode
' This procedure generates the update query based on the field to be updated and
'    executes the query.
' d2 Change Quotes to QuoteValueWithQuotes and updated entry type ID field
' d5    ksk     Changes for closure status ID being alphanumeric
' d6    ksk     changed the function to return a recordlist that will be a list of catalogues
'               affected by the change.
' d7    ksk     changed the physical condition field to be a text field insted of lookup
' d9    ksk     Added the extra join that was missing to link the extenstion
'               and main table for the list of records affected
' d11   ksk     Changed the condition in the update for imported matrial. Used the
'               wrong constant
' d12   ksk     Added a max constant to be used by all the Findbysql
'               To return the affected records
' d19   alm     Updates the start / end dates automatically when a covering date is updated
' d20   alm     Supports Creator, Deposit and Originals
'
On Error GoTo ErrHandler

    Dim sSql          As String
    Dim sStartDateSql As String 'd19
    Dim sEndDateSql   As String 'd19
    Dim bCanUpdate    As Boolean
    Dim sSource       As String
    
    sStartDateSql = vbNullString
    sEndDateSql = vbNullString
    
    sSource = "BOCatalogueBulkUpdate.BulkUpdateLettercode"
    
    'Check if the set property is available for tthe lettercode level
    bCanUpdate = m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(ctLetterCode, m_eUpdateField)
        
    
    'If the current property is available for Lettercode then generate the corresponding SQL
    '  to update the field in the corresponding table.
    If bCanUpdate Then
        Select Case m_eUpdateField
            'Main table
            'SQL to update both the TITLE fields
            Case ECATALOGUEPROPERTIES.cpTitle
                
                
                sSql = "update " & gk_sDBTABLEEditsetLettercode & " Set " & _
                                    gk_sDBCOLLetterCodeTitle & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetLettercodeExtension & _
                        " Where " & gk_sDBTABLEEditsetLettercode & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetLettercodeExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetLettercode & "." & gk_sDBCOLLetterCodeID & _
                                    " = " & gk_sDBTABLEEditsetLettercodeExtension & "." & _
                                    gk_sDBCOLLetterCodeID & _
                        " AND " & gk_sDBTABLEEditsetLettercodeExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                        
                sSql = sSql & "  UPDATE " & gk_sDBTABLEEditsetLettercodeTitle & _
                            " Set " & gk_sDBCOLTitle & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetLettercodeExtension & _
                        " Where " & gk_sDBTABLEEditsetLettercodeTitle & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetLettercodeExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetLettercodeTitle & "." & gk_sDBCOLLetterCodeID & _
                                    " = " & gk_sDBTABLEEditsetLettercodeExtension & "." & _
                                    gk_sDBCOLLetterCodeID & _
                        " AND " & gk_sDBTABLEEditsetLettercodeExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            
            
            'SQL to update the CLOSURE STATUS ID Link
            Case ECATALOGUEPROPERTIES.cpClosureStatus
                
                sSql = "update " & gk_sDBTABLEEditsetLettercode & " Set " & _
                            gk_sDBCOLClosureStatus & " = " & StringToNullOrString( _
                                            m_vUpdateValue1) & _
                        " FROM " & gk_sDBTABLEEditsetLettercodeExtension & _
                        " Where " & gk_sDBTABLEEditsetLettercode & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetLettercodeExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetLettercode & "." & gk_sDBCOLLetterCodeID & _
                                    " = " & gk_sDBTABLEEditsetLettercodeExtension & "." & _
                                    gk_sDBCOLLetterCodeID & _
                        " AND " & gk_sDBTABLEEditsetLettercodeExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            
            'Extension table
            'SQL to update the FORMER REFERENCE (Department) field
            Case ECATALOGUEPROPERTIES.cpFormerDepartmentReference
                
                sSql = "update " & gk_sDBTABLEEditsetLettercodeExtension & " Set " & _
                                    gk_sDBCOLFormerDepartmentReference & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'SQL to update the FORMER PRO REFERENCE  field
            Case ECATALOGUEPROPERTIES.cpFormerPROReference
                
                sSql = "update " & gk_sDBTABLEEditsetLettercodeExtension & " Set " & _
                                    gk_sDBCOLFormerPROReference & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'SQL to update the COVERING DATES field
            Case ECATALOGUEPROPERTIES.cpCoveringDates
                
                sSql = "update " & gk_sDBTABLEEditsetLettercodeExtension & " Set " & _
                                    gk_sDBCOLCoveringDate & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                    
                ' d19 We also need to update the start and end dates
                If Not IsNull(m_vUpdateValue2) Then
                    
                    sStartDateSql = "update " & gk_sDBTABLEEditsetLettercodeExtension & " Set " & _
                                    gk_sDBCOLCoveringFromDate & " = " & _
                                    StringToNullOrString(Left(m_vUpdateValue2, 8)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                                        
                    sEndDateSql = "update " & gk_sDBTABLEEditsetLettercodeExtension & " Set " & _
                                    gk_sDBCOLCoveringToDate & " = " & _
                                    StringToNullOrString(Right(m_vUpdateValue2, 8)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                                    
                End If
            
            'SQL to update the PHYSICAL CONDITION ID Link
            Case ECATALOGUEPROPERTIES.cpPhysicalCondition
                
                sSql = "update " & gk_sDBTABLEEditsetLettercodeExtension & " Set " & _
                                    gk_sDBCOLPhysicalCondition & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'SQL to update the PHYSICAL DESCRIPTION RECORD QUANTITY field
            Case ECATALOGUEPROPERTIES.cpPhysicalDescription_RecordQuantity
                
                sSql = "update " & gk_sDBTABLEEditsetLettercodeExtension & " Set " & _
                                    gk_sDBCOLPhysicalDescriptionRecordQuantity & " = " & _
                                    NullToSQLNull(StringToNullOrLong(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                        
            'SQL to update the PHYSICAL DESCRIPTION FORM ID Link
            Case ECATALOGUEPROPERTIES.cpPhysicalDescription_Form
                
                sSql = "update " & gk_sDBTABLEEditsetLettercodeExtension & " Set " & _
                                    gk_sDBCOLPhysicalDescriptionFormID & " = " & _
                                    NullToSQLNull(StringToNullOrLong(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
    
            'SQL to update the PHYSICAL DESCRIPTION DIMENSION field
            Case ECATALOGUEPROPERTIES.cpPhysicalDescription_Dimension
                
                
                sSql = "update " & gk_sDBTABLEEditsetLettercodeExtension & " Set " & _
                                    gk_sDBCOLPhysicalDescriptionDimension & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'SQL to update the MAP DESIGNATION field
            Case ECATALOGUEPROPERTIES.cpMapDesignation
                
                sSql = "update " & gk_sDBTABLEEditsetLettercodeExtension & " Set " & _
                                    gk_sDBCOLMapDesignation & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'SQL to update the MAP SCALE NUMBER field
            Case ECATALOGUEPROPERTIES.cpMapscaleNumber
                
                sSql = "update " & gk_sDBTABLEEditsetLettercodeExtension & " Set " & _
                                    gk_sDBCOLMapScaleNumber & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'SQL to update the ACCESS CONDITION ID Link
            Case ECATALOGUEPROPERTIES.cpAccessCondition
                
                sSql = "update " & gk_sDBTABLEEditsetLettercodeExtension & " Set " & _
                                    gk_sDBCOLAccessConditionID & " = " & _
                                    NullToSQLNull(StringToNullOrLong(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)


            'SQL to update the RESTRICTIONS ON USE field
            Case ECATALOGUEPROPERTIES.cpRestrictionsOnUse
                
                sSql = "update " & gk_sDBTABLEEditsetLettercodeExtension & " Set " & _
                                    gk_sDBCOLRestrictionsOnUse & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'SQL to update the ACCUMULATION DATES field
            Case ECATALOGUEPROPERTIES.cpAccumulationDates
                
                sSql = "update " & gk_sDBTABLEEditsetLettercodeExtension & " Set " & _
                                    gk_sDBCOLAccumulationdates & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                        
            'SQL to update the RECORD OPENING DATE field
            Case ECATALOGUEPROPERTIES.cpRecordOpeningDate
                
                sSql = "update " & gk_sDBTABLEEditsetLettercodeExtension & " Set " & _
                                    gk_sDBCOLRecordOpeningDate & " = " & _
                                    QuoteValueWithQuotes(VariantDateAsDDMMMYYYY(StringToNullOrDate(m_vUpdateValue1))) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'SQL to update the LANGUAGE ID Link
            Case ECATALOGUEPROPERTIES.cpLanguage
                
                sSql = "update " & gk_sDBTABLEEditsetLettercodeExtension & " Set " & _
                                    gk_sDBCOLLanguageID & " = " & _
                                    NullToSQLNull(StringToNullOrLong(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'SQL to update the LEGAL STATUS ID Link
            Case ECATALOGUEPROPERTIES.cpLegalStatus
                
                sSql = "update " & gk_sDBTABLEEditsetLettercodeExtension & " Set " & _
                                    gk_sDBCOLLegalStatusID & " = " & _
                                    NullToSQLNull(StringToNullOrLong(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'SQL to update the CLOSURE CODE value and CLOSURE TYPE lookup
            Case ECATALOGUEPROPERTIES.cpClosureCode, ECATALOGUEPROPERTIES.cpClosureType
            
                sSql = " update " & gk_sDBTABLEEditsetLettercodeExtension & " Set " & _
                                    gk_sDBCOLClosureType & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly) & vbNewLine
                
                sSql = sSql & " update " & gk_sDBTABLEEditsetLettercodeExtension & " Set " & _
                                    gk_sDBCOLClosureCode & " = " & _
                                    NullToSQLNull(StringToNullOrLong(m_vUpdateValue2)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly) & vbNewLine
            
            
            'Seperate tables, one property per table
            'SQL to update the seperate SCOPE AND CONTENT field
            Case ECATALOGUEPROPERTIES.cpScopeAndContent
                sSql = "update " & gk_sDBTABLEEditsetLettercodeScopeAndContent & " Set " & _
                                    gk_sDBCOLScopeAndContent & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetLettercodeExtension & _
                        " Where " & gk_sDBTABLEEditsetLettercodeScopeAndContent & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetLettercodeExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetLettercodeScopeAndContent & "." & gk_sDBCOLLetterCodeID & _
                                    " = " & gk_sDBTABLEEditsetLettercodeExtension & "." & _
                                    gk_sDBCOLLetterCodeID & _
                        " AND " & gk_sDBTABLEEditsetLettercodeExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'SQL to update the seperate ARRANGAMENT field
            Case ECATALOGUEPROPERTIES.cpArrangement
                
                sSql = "update " & gk_sDBTABLEEditsetLettercodeArrangement & " Set " & _
                                    gk_sDBCOLArrangement & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetLettercodeExtension & _
                        " Where " & gk_sDBTABLEEditsetLettercodeArrangement & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetLettercodeExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetLettercodeArrangement & "." & gk_sDBCOLLetterCodeID & _
                                    " = " & gk_sDBTABLEEditsetLettercodeExtension & "." & _
                                    gk_sDBCOLLetterCodeID & _
                        " AND " & gk_sDBTABLEEditsetLettercodeExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            
            'SQL to update the seperate ADMINISTRATIVE OR BIOGRAPHICAL HISTORY field
            Case ECATALOGUEPROPERTIES.cpAdministrative_BiographicalHistory
                
                
                sSql = "update " & gk_sDBTABLEEditsetLettercodeAdministrativeHistory & " Set " & _
                                    gk_sDBCOLAdministrativeHistory & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetLettercodeExtension & _
                        " Where " & gk_sDBTABLEEditsetLettercodeAdministrativeHistory & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetLettercodeExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetLettercodeAdministrativeHistory & "." & gk_sDBCOLLetterCodeID & _
                                    " = " & gk_sDBTABLEEditsetLettercodeExtension & "." & _
                                    gk_sDBCOLLetterCodeID & _
                        " AND " & gk_sDBTABLEEditsetLettercodeExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            
            'SQL to update the CUSTODIAL HISTORY field
            Case ECATALOGUEPROPERTIES.cpCustodialHistory
                
                sSql = "update " & gk_sDBTABLEEditsetLettercodeCustodialHistory & " Set " & _
                                    gk_sDBCOLCustodialHistory & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetLettercodeExtension & _
                        " Where " & gk_sDBTABLEEditsetLettercodeCustodialHistory & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetLettercodeExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetLettercodeCustodialHistory & "." & gk_sDBCOLLetterCodeID & _
                                    " = " & gk_sDBTABLEEditsetLettercodeExtension & "." & _
                                    gk_sDBCOLLetterCodeID & _
                        " AND " & gk_sDBTABLEEditsetLettercodeExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            
            
            'SQL to update the APPRAISAL DESTRUCTION INFORMATION field
            Case ECATALOGUEPROPERTIES.cpAppraisal_DestructionInformation
                
                sSql = "update " & gk_sDBTABLEEditsetLettercodeAppraisalDestructionInformation & _
                            " Set " & gk_sDBCOLAppraisalDestructionInformation & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetLettercodeExtension & _
                        " Where " & gk_sDBTABLEEditsetLettercodeAppraisalDestructionInformation & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetLettercodeExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetLettercodeAppraisalDestructionInformation & "." & gk_sDBCOLLetterCodeID & _
                                    " = " & gk_sDBTABLEEditsetLettercodeExtension & "." & _
                                    gk_sDBCOLLetterCodeID & _
                        " AND " & gk_sDBTABLEEditsetLettercodeExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            
            'SQL to update the ACCRUALS field
            Case ECATALOGUEPROPERTIES.cpAccruals
                
                sSql = "update " & gk_sDBTABLEEditsetLettercodeAccruals & " Set " & _
                                    gk_sDBCOLAccruals & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetLettercodeExtension & _
                        " Where " & gk_sDBTABLEEditsetLettercodeAccruals & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetLettercodeExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetLettercodeAccruals & "." & gk_sDBCOLLetterCodeID & _
                                    " = " & gk_sDBTABLEEditsetLettercodeExtension & "." & _
                                    gk_sDBCOLLetterCodeID & _
                        " AND " & gk_sDBTABLEEditsetLettercodeExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'PC TABLES
            'SQL to update the IMMEDIATE SOURCE OF ACQUISITION description field ONLY
            Case ECATALOGUEPROPERTIES.cpImmediateSourceOfAcquisition
                    
                UpdateLinkDataForImmediateSource ECATALOGUELEVEL.ctLetterCode
            
            'SQL to update the COPIES INFORMATION description field ONLY
            Case ECATALOGUEPROPERTIES.cpCopiesInformation
                
                UpdateLinkDataForCopiesInfo ECATALOGUELEVEL.ctLetterCode
            
            'SQL to update the PUBLICATION NOTE description field ONLY
            Case ECATALOGUEPROPERTIES.cpPublicationNote
            
                UpdateLinkDataForPublicationNote ECATALOGUELEVEL.ctLetterCode
                
            'SQL to update the UNPUBLISHED FINDING AIDS description field ONLY
            Case ECATALOGUEPROPERTIES.cpUnpublishedFindingAids
                
                UpdateLinkDataForUnpublishedFindingAids ECATALOGUELEVEL.ctLetterCode
        
            'd20 SQL to update the CREATOR NAMES description field ONLY
            Case ECATALOGUEPROPERTIES.cpCreatorNames
                
                UpdateLinkDataForCreatorNames ECATALOGUELEVEL.ctLetterCode
                
            'd20 SQL to update the PLACE OF DEPOSIT description field ONLY
            Case ECATALOGUEPROPERTIES.cpPlaceOfDeposit
                
                UpdateLinkDataForPlaceOfDeposit ECATALOGUELEVEL.ctLetterCode
                
            'd20 SQL to update the LOCATION OF ORIGINALS description field ONLY
            Case ECATALOGUEPROPERTIES.cpLocationOfOriginals
                
                UpdateLinkDataForLocationOfOriginals ECATALOGUELEVEL.ctLetterCode
        
            'd20 SQL to update the CREATOR NAMES description field ONLY
            Case ECATALOGUEPROPERTIES.cpCreatorNames
                
                UpdateLinkDataForCreatorNames ECATALOGUELEVEL.ctLetterCode
                
            'd20 SQL to update the PLACE OF DEPOSIT description field ONLY
            Case ECATALOGUEPROPERTIES.cpPlaceOfDeposit
                
                UpdateLinkDataForPlaceOfDeposit ECATALOGUELEVEL.ctLetterCode
                
            'd20 SQL to update the LOCATION OF ORIGINALS description field ONLY
            Case ECATALOGUEPROPERTIES.cpLocationOfOriginals
                
                UpdateLinkDataForLocationOfOriginals ECATALOGUELEVEL.ctLetterCode
        
        End Select
        
        If Not (m_eUpdateField = ECATALOGUEPROPERTIES.cpImmediateSourceOfAcquisition Or _
                m_eUpdateField = ECATALOGUEPROPERTIES.cpCopiesInformation Or _
                m_eUpdateField = ECATALOGUEPROPERTIES.cpPublicationNote Or _
                m_eUpdateField = ECATALOGUEPROPERTIES.cpUnpublishedFindingAids Or _
                m_eUpdateField = ECATALOGUEPROPERTIES.cpCreatorNames Or _
                m_eUpdateField = ECATALOGUEPROPERTIES.cpPlaceOfDeposit Or _
                m_eUpdateField = ECATALOGUEPROPERTIES.cpLocationOfOriginals) Then
            
            m_ofwUpdate.Execute sSql
            
            ' d19 - Update the start AND end dates if possible
            If Not (sStartDateSql = vbNullString And sEndDateSql = vbNullString) Then
                m_ofwUpdate.Execute (sStartDateSql)
                m_ofwUpdate.Execute (sEndDateSql)
            End If
            
        End If
        
        'update the editset entry type to modified if not change for all catalogues the level
        sSql = "update " & gk_sDBTABLEEditsetLettercodeExtension & " Set " & _
                gk_sDBCOLEntryTypeID & " = CASE " & gk_sDBCOLEntryTypeID & _
                    " WHEN " & CStr(ECATALOGUEENTRYSTATUS.cesPROCATCreated) & " THEN " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesPROCATCreatedModified) & _
                    " WHEN " & CStr(ECATALOGUEENTRYSTATUS.cesAddedFromLive) & " THEN " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesAddedFromLiveModified) & _
                    " WHEN " & CStr(ECATALOGUEENTRYSTATUS.cesImported) & " THEN " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesImportedModified) & _
                    " ELSE " & gk_sDBCOLEntryTypeID & _
                    " END " & Chr(vbKeyReturn) & _
                " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID)
    
        m_ofwUpdate.Execute sSql
    
        
        sSql = "SELECT " & gk_sDBTABLEEditsetLettercode & "." & gk_sDBCOLLetterCodeID & " as " & gk_sDBCOLCatalogueID & _
                ", 1 as " & gk_sDBCOLCatalogueLevel & _
                " from " & gk_sDBTABLEEditsetLettercode & "(nolock) " & _
                ", " & gk_sDBTABLEEditsetLettercodeExtension & "(nolock) " & _
                " Where " & gk_sDBTABLEEditsetLettercode & "." & gk_sDBCOLEditSetID & _
                                " = " & CStr(m_lEditsetID) & _
                " AND " & gk_sDBTABLEEditsetLettercodeExtension & "." & gk_sDBCOLLetterCodeID & _
                                " = " & gk_sDBTABLEEditsetLettercode & "." & gk_sDBCOLLetterCodeID & _
                " AND " & gk_sDBTABLEEditsetLettercodeExtension & "." & gk_sDBCOLEditSetID & _
                                " = " & CStr(m_lEditsetID) & _
                " AND " & gk_sDBTABLEEditsetLettercodeExtension & "." & _
                gk_sDBCOLEditSetEntryTypeID & " <> " & CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
        
        Set BulkUpdateLettercode = m_ofwUpdate.FindBySql(sSql, mk_iMaxNoOfRecords)
    
    End If

    Exit Function

ErrHandler:
    ConvertError sSource
End Function




Private Function BulkUpdateDivision() As FWRecordList
' Bulkupdates the value of the set property of the Division
' This procedure generates the update query based on the field to be updated and
'    executes the query.
' d2 Change Quotes to QuoteValueWithQuotes and updated entry type ID field
' d5    ksk     Changes for closure status ID being alphanumeric
'
' d6    ksk     changed the function to return a recordlist that will be a list of catalogues
'               affected by the change.
' d7    ksk     changed the physical condition field to be a text field insted of lookup
' d9    ksk     Added the extra join that was missing to link the extenstion
'               and main table for the list of records affected
' d11   ksk     Changed the condition in the update for imported matrial. Used the
'               wrong constant
' d12   ksk     Added a max constant to be used by all the Findbysql
'               To return the affected records
' d18   tmj     Ensured selected from Primary key of EditsetID and DivisionID
' d19   alm     12Feb2001       Updates the start / end dates automatically
'                               when a covering date is updated
' d20   alm     Supports Creator, Deposit and Originals
'
On Error GoTo ErrHandler
    
    Dim sSql          As String
    Dim sStartDateSql As String 'd19
    Dim sEndDateSql   As String 'd19
    Dim bCanUpdate    As Boolean
    Dim sSource       As String
    
    sSource = "BOCatalogueBulkUpdate.BulkUpdateDivision"
    
    'Check if the set property is available for tthe Division level
    bCanUpdate = m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(ctDivision, m_eUpdateField)
    
    'If the current property is available for Division then generate the corresponding SQL
    '  to update the field in the corresponding table.
    If bCanUpdate Then
        Select Case m_eUpdateField
            'Main table
            'SQL to update both the TITLE fields
            Case ECATALOGUEPROPERTIES.cpTitle
                
                sSql = "update " & gk_sDBTABLEEditsetDivision & " Set " & _
                                    gk_sDBCOLDivisionTitle & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetDivisionExtension & _
                        " Where " & gk_sDBTABLEEditsetDivision & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetDivisionExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetDivision & "." & gk_sDBCOLDivisionID & _
                                    " = " & gk_sDBTABLEEditsetDivisionExtension & "." & _
                                    gk_sDBCOLDivisionID & _
                        " AND " & gk_sDBTABLEEditsetDivisionExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                        
                        
                sSql = sSql & "  UPDATE " & gk_sDBTABLEEditsetDivisionTitle & " Set " & _
                                    gk_sDBCOLTitle & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetDivisionExtension & _
                        " Where " & gk_sDBTABLEEditsetDivisionTitle & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetDivisionExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetDivisionTitle & "." & gk_sDBCOLDivisionID & _
                                    " = " & gk_sDBTABLEEditsetDivisionExtension & "." & _
                                    gk_sDBCOLDivisionID & _
                        " AND " & gk_sDBTABLEEditsetDivisionExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'SQL to update the CLOSURE STATUS ID Link
            Case ECATALOGUEPROPERTIES.cpClosureStatus
                
                sSql = "update " & gk_sDBTABLEEditsetDivision & " Set " & _
                                    gk_sDBCOLClosureStatus & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetDivisionExtension & _
                        " Where " & gk_sDBTABLEEditsetDivision & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetDivisionExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetDivision & "." & gk_sDBCOLDivisionID & _
                                    " = " & gk_sDBTABLEEditsetDivisionExtension & "." & _
                                    gk_sDBCOLDivisionID & _
                        " AND " & gk_sDBTABLEEditsetDivisionExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'Extension table
            'SQL to update the FORMER REFERENCE (Department) field
            Case ECATALOGUEPROPERTIES.cpFormerDepartmentReference
                
                sSql = "update " & gk_sDBTABLEEditsetDivisionExtension & " Set " & _
                                    gk_sDBCOLFormerDepartmentReference & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'SQL to update the FORMER PRO REFERENCE  field
            Case ECATALOGUEPROPERTIES.cpFormerPROReference
            
                sSql = "update " & gk_sDBTABLEEditsetDivisionExtension & " Set " & _
                                    gk_sDBCOLFormerPROReference & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                    
            'SQL to update the COVERING DATES field
            Case ECATALOGUEPROPERTIES.cpCoveringDates
            
                sSql = "update " & gk_sDBTABLEEditsetDivisionExtension & " Set " & _
                                    gk_sDBCOLCoveringDate & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                    
                ' d19 We also need to update the start and end dates
                If Not IsNull(m_vUpdateValue2) Then
                    
                    sStartDateSql = "update " & gk_sDBTABLEEditsetDivisionExtension & " Set " & _
                                    gk_sDBCOLCoveringFromDate & " = " & _
                                    StringToNullOrString(Left(m_vUpdateValue2, 8)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                                        
                    sEndDateSql = "update " & gk_sDBTABLEEditsetDivisionExtension & " Set " & _
                                    gk_sDBCOLCoveringToDate & " = " & _
                                    StringToNullOrString(Right(m_vUpdateValue2, 8)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                                    
                End If
                                    
            'SQL to update the PHYSICAL CONDITION ID Link
            Case ECATALOGUEPROPERTIES.cpPhysicalCondition
            
                sSql = "update " & gk_sDBTABLEEditsetDivisionExtension & " Set " & _
                                    gk_sDBCOLPhysicalCondition & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                    
            'SQL to update the PHYSICAL DESCRIPTION RECORD QUANTITY field
            Case ECATALOGUEPROPERTIES.cpPhysicalDescription_RecordQuantity
            
                sSql = "update " & gk_sDBTABLEEditsetDivisionExtension & " Set " & _
                                    gk_sDBCOLPhysicalDescriptionRecordQuantity & " = " & _
                                    NullToSQLNull(StringToNullOrLong(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                    
            'SQL to update the PHYSICAL DESCRIPTION FORM ID Link
            Case ECATALOGUEPROPERTIES.cpPhysicalDescription_Form
            
                sSql = "update " & gk_sDBTABLEEditsetDivisionExtension & " Set " & _
                                    gk_sDBCOLPhysicalDescriptionFormID & " = " & _
                                    NullToSQLNull(StringToNullOrLong(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                    
            'SQL to update the PHYSICAL DESCRIPTION DIMENSION field
            Case ECATALOGUEPROPERTIES.cpPhysicalDescription_Dimension
            
                sSql = "update " & gk_sDBTABLEEditsetDivisionExtension & " Set " & _
                                    gk_sDBCOLPhysicalDescriptionDimension & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                        
            'SQL to update the MAP DESIGNATION field
            Case ECATALOGUEPROPERTIES.cpMapDesignation
            
                sSql = "update " & gk_sDBTABLEEditsetDivisionExtension & " Set " & _
                                    gk_sDBCOLMapDesignation & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                    
            'SQL to update the MAP SCALE NUMBER field
            Case ECATALOGUEPROPERTIES.cpMapscaleNumber
            
                sSql = "update " & gk_sDBTABLEEditsetDivisionExtension & " Set " & _
                                    gk_sDBCOLMapScaleNumber & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                    
            'SQL to update the ACCESS CONDITION ID Link
            Case ECATALOGUEPROPERTIES.cpAccessCondition
            
                sSql = "update " & gk_sDBTABLEEditsetDivisionExtension & " Set " & _
                                    gk_sDBCOLAccessConditionID & " = " & _
                                    NullToSQLNull(StringToNullOrLong(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                    
            'SQL to update the RESTRICTIONS ON USE field
            Case ECATALOGUEPROPERTIES.cpRestrictionsOnUse
            
                sSql = "update " & gk_sDBTABLEEditsetDivisionExtension & " Set " & _
                                    gk_sDBCOLRestrictionsOnUse & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                    
            'SQL to update the ACCUMULATION DATES field
            Case ECATALOGUEPROPERTIES.cpAccumulationDates
            
                sSql = "update " & gk_sDBTABLEEditsetDivisionExtension & " Set " & _
                                    gk_sDBCOLAccumulationdates & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                    
            'SQL to update the RECORD OPENING DATE field
            Case ECATALOGUEPROPERTIES.cpRecordOpeningDate
            
                sSql = "update " & gk_sDBTABLEEditsetDivisionExtension & " Set " & _
                                    gk_sDBCOLRecordOpeningDate & " = " & _
                                    QuoteValueWithQuotes(VariantDateAsDDMMMYYYY(StringToNullOrDate(m_vUpdateValue1))) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                    
            'SQL to update the LANGUAGE ID Link
            Case ECATALOGUEPROPERTIES.cpLanguage
            
                sSql = "update " & gk_sDBTABLEEditsetDivisionExtension & " Set " & _
                                    gk_sDBCOLLanguageID & " = " & _
                                    NullToSQLNull(StringToNullOrLong(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                        
            'SQL to update the LEGAL STATUS ID Link
            Case ECATALOGUEPROPERTIES.cpLegalStatus
            
                sSql = "update " & gk_sDBTABLEEditsetDivisionExtension & " Set " & _
                                    gk_sDBCOLLegalStatusID & " = " & _
                                    NullToSQLNull(StringToNullOrLong(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'SQL to update the CLOSURE CODE value and CLOSURE TYPE lookup
            Case ECATALOGUEPROPERTIES.cpClosureCode, ECATALOGUEPROPERTIES.cpClosureType
                
                sSql = " update " & gk_sDBTABLEEditsetDivisionExtension & " Set " & _
                                    gk_sDBCOLClosureType & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly) & vbNewLine
                                    
                sSql = sSql & " update " & gk_sDBTABLEEditsetDivisionExtension & " Set " & _
                                    gk_sDBCOLClosureCode & " = " & _
                                    NullToSQLNull(StringToNullOrLong(m_vUpdateValue2)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                
            
            'Seperate tables, one property per table
            'SQL to update the seperate SCOPE AND CONTENT field
            Case ECATALOGUEPROPERTIES.cpScopeAndContent
                
                sSql = "update " & gk_sDBTABLEEditsetDivisionScopeAndContent & " Set " & _
                                    gk_sDBCOLScopeAndContent & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetDivisionExtension & _
                        " Where " & gk_sDBTABLEEditsetDivisionScopeAndContent & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetDivisionExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetDivisionScopeAndContent & "." & gk_sDBCOLDivisionID & _
                                    " = " & gk_sDBTABLEEditsetDivisionExtension & "." & _
                                    gk_sDBCOLDivisionID & _
                        " AND " & gk_sDBTABLEEditsetDivisionExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            
            'SQL to update the seperate ARRANGAMENT field
            Case ECATALOGUEPROPERTIES.cpArrangement
                
                sSql = "update " & gk_sDBTABLEEditsetDivisionArrangement & " Set " & _
                                    gk_sDBCOLArrangement & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetDivisionExtension & _
                        " Where " & gk_sDBTABLEEditsetDivisionArrangement & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetDivisionExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetDivisionArrangement & "." & gk_sDBCOLDivisionID & _
                                    " = " & gk_sDBTABLEEditsetDivisionExtension & "." & _
                                    gk_sDBCOLDivisionID & _
                        " AND " & gk_sDBTABLEEditsetDivisionExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'SQL to update the seperate ADMINISTRATIVE OR BIOGRAPHICAL HISTORY field
            Case ECATALOGUEPROPERTIES.cpAdministrative_BiographicalHistory
                
                sSql = "update " & gk_sDBTABLEEditsetDivisionAdministrativeHistory & " Set " & _
                                    gk_sDBCOLAdministrativeHistory & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetDivisionExtension & _
                        " Where " & gk_sDBTABLEEditsetDivisionAdministrativeHistory & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetDivisionExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetDivisionAdministrativeHistory & "." & gk_sDBCOLDivisionID & _
                                    " = " & gk_sDBTABLEEditsetDivisionExtension & "." & _
                                    gk_sDBCOLDivisionID & _
                        " AND " & gk_sDBTABLEEditsetDivisionExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'SQL to update the CUSTODIAL HISTORY field
            Case ECATALOGUEPROPERTIES.cpCustodialHistory
                
                sSql = "update " & gk_sDBTABLEEditsetDivisionCustodialHistory & " Set " & _
                                    gk_sDBCOLCustodialHistory & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetDivisionExtension & _
                        " Where " & gk_sDBTABLEEditsetDivisionCustodialHistory & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetDivisionExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetDivisionCustodialHistory & "." & gk_sDBCOLDivisionID & _
                                    " = " & gk_sDBTABLEEditsetDivisionExtension & "." & _
                                    gk_sDBCOLDivisionID & _
                        " AND " & gk_sDBTABLEEditsetDivisionExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'SQL to update the APPRAISAL DESTRUCTION INFORMATION field
            Case ECATALOGUEPROPERTIES.cpAppraisal_DestructionInformation
                
                sSql = "update " & gk_sDBTABLEEditsetDivisionAppraisalDestructionInformation & _
                            " Set " & gk_sDBCOLAppraisalDestructionInformation & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetDivisionExtension & _
                        " Where " & gk_sDBTABLEEditsetDivisionAppraisalDestructionInformation & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetDivisionExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetDivisionAppraisalDestructionInformation & "." & gk_sDBCOLDivisionID & _
                                    " = " & gk_sDBTABLEEditsetDivisionExtension & "." & _
                                    gk_sDBCOLDivisionID & _
                        " AND " & gk_sDBTABLEEditsetDivisionExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'SQL to update the ACCRUALS field
            Case ECATALOGUEPROPERTIES.cpAccruals
                
                sSql = "update " & gk_sDBTABLEEditsetDivisionAccruals & " Set " & _
                                    gk_sDBCOLAccruals & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetDivisionExtension & _
                        " Where " & gk_sDBTABLEEditsetDivisionAccruals & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetDivisionExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetDivisionAccruals & "." & gk_sDBCOLDivisionID & _
                                    " = " & gk_sDBTABLEEditsetDivisionExtension & "." & _
                                    gk_sDBCOLDivisionID & _
                        " AND " & gk_sDBTABLEEditsetDivisionExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
        
            'PC TABLES
            'SQL to update the IMMEDIATE SOURCE OF ACQUISITION description field ONLY
            Case ECATALOGUEPROPERTIES.cpImmediateSourceOfAcquisition
                    
                UpdateLinkDataForImmediateSource ECATALOGUELEVEL.ctDivision
            
            'SQL to update the COPIES INFORMATION description field ONLY
            Case ECATALOGUEPROPERTIES.cpCopiesInformation
                
                UpdateLinkDataForCopiesInfo ECATALOGUELEVEL.ctDivision
            
            'SQL to update the PUBLICATION NOTE description field ONLY
            Case ECATALOGUEPROPERTIES.cpPublicationNote
            
                UpdateLinkDataForPublicationNote ECATALOGUELEVEL.ctDivision
                
            'SQL to update the UNPUBLISHED FINDING AIDS description field ONLY
            Case ECATALOGUEPROPERTIES.cpUnpublishedFindingAids
                
                UpdateLinkDataForUnpublishedFindingAids ECATALOGUELEVEL.ctDivision
        
            'd20 SQL to update the CREATOR NAMES description field ONLY
            Case ECATALOGUEPROPERTIES.cpCreatorNames
                
                UpdateLinkDataForCreatorNames ECATALOGUELEVEL.ctDivision
                
            'd20 SQL to update the PLACE OF DEPOSIT description field ONLY
            Case ECATALOGUEPROPERTIES.cpPlaceOfDeposit
                
                UpdateLinkDataForPlaceOfDeposit ECATALOGUELEVEL.ctDivision
                
            'd20 SQL to update the LOCATION OF ORIGINALS description field ONLY
            Case ECATALOGUEPROPERTIES.cpLocationOfOriginals
                
                UpdateLinkDataForLocationOfOriginals ECATALOGUELEVEL.ctDivision
        
        End Select
        
        
        If Not (m_eUpdateField = ECATALOGUEPROPERTIES.cpImmediateSourceOfAcquisition Or _
            m_eUpdateField = ECATALOGUEPROPERTIES.cpCopiesInformation Or _
            m_eUpdateField = ECATALOGUEPROPERTIES.cpPublicationNote Or _
            m_eUpdateField = ECATALOGUEPROPERTIES.cpUnpublishedFindingAids Or _
            m_eUpdateField = ECATALOGUEPROPERTIES.cpCreatorNames Or _
            m_eUpdateField = ECATALOGUEPROPERTIES.cpPlaceOfDeposit Or _
            m_eUpdateField = ECATALOGUEPROPERTIES.cpLocationOfOriginals) Then
            
            m_ofwUpdate.Execute sSql
            
            ' d19 - Update the start AND end dates if possible
            If Not (sStartDateSql = vbNullString And sEndDateSql = vbNullString) Then
                m_ofwUpdate.Execute (sStartDateSql)
                m_ofwUpdate.Execute (sEndDateSql)
            End If
            
        End If
    
        'update the editset entry type to modified if not change for all catalogues the level
        sSql = "update " & gk_sDBTABLEEditsetDivisionExtension & " Set " & _
                gk_sDBCOLEntryTypeID & " = CASE " & gk_sDBCOLEntryTypeID & _
                    " WHEN " & CStr(ECATALOGUEENTRYSTATUS.cesPROCATCreated) & " THEN " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesPROCATCreatedModified) & _
                    " WHEN " & CStr(ECATALOGUEENTRYSTATUS.cesAddedFromLive) & " THEN " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesAddedFromLiveModified) & _
                    " WHEN " & CStr(ECATALOGUEENTRYSTATUS.cesImported) & " THEN " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesImportedModified) & _
                    " ELSE " & gk_sDBCOLEntryTypeID & _
                    " END " & _
                " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID)
    
        m_ofwUpdate.Execute sSql
    
        sSql = "Select " & gk_sDBTABLEEditsetDivision & "." & gk_sDBCOLDivisionID & " as " & gk_sDBCOLCatalogueID & _
                ", 2 as " & gk_sDBCOLCatalogueLevel & _
                " from " & gk_sDBTABLEEditsetDivision & "(nolock) " & _
                ", " & gk_sDBTABLEEditsetDivisionExtension & "(nolock) " & _
                " Where " & gk_sDBTABLEEditsetDivision & "." & gk_sDBCOLEditSetID & _
                                " = " & CStr(m_lEditsetID) & _
                " AND " & gk_sDBTABLEEditsetDivisionExtension & "." & gk_sDBCOLEditSetID & _
                                " = " & CStr(m_lEditsetID) & _
                " AND " & gk_sDBTABLEEditsetDivisionExtension & "." & gk_sDBCOLDivisionID & _
                                " = " & gk_sDBTABLEEditsetDivision & "." & gk_sDBCOLDivisionID & _
                " AND " & gk_sDBTABLEEditsetDivisionExtension & "." & gk_sDBCOLEditSetID & _
                                " = " & gk_sDBTABLEEditsetDivision & "." & gk_sDBCOLEditSetID & _
                " AND " & gk_sDBTABLEEditsetDivisionExtension & "." & _
                gk_sDBCOLEditSetEntryTypeID & " <> " & CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
        Set BulkUpdateDivision = m_ofwUpdate.FindBySql(sSql, mk_iMaxNoOfRecords)
    
    
    End If
    Exit Function

ErrHandler:
    ConvertError sSource
End Function




Private Function BulkUpdateClass() As FWRecordList
' Bulkupdates the value of the set property of the Class
' This procedure generates the update query based on the field to be updated and
'    executes the query.
' d2 Change Quotes to QuoteValueWithQuotes and updated entry type ID field
' d5    ksk     Changes for closure status ID being alphanumeric
'
' d6    ksk     changed the function to return a recordlist that will be a list of catalogues
'               affected by the change.
' d7    ksk     changed the physical condition field to be a text field insted of lookup
' d9    ksk     Added the extra join that was missing to link the extenstion
'               and main table for the list of records affected
' d11   ksk     Changed the condition in the update for imported matrial. Used the
'               wrong constant
' d12   ksk     Added a max constant to be used by all the Findbysql
'               To return the affected records
' d19   alm     12Feb2001       Updates the start / end dates automatically
'                               when a covering date is updated
' d20   alm     Supports Creator, Deposit and Originals
'
On Error GoTo ErrHandler

    Dim sSql          As String
    Dim sStartDateSql As String 'd19
    Dim sEndDateSql   As String 'd19
    Dim bCanUpdate    As Boolean
    Dim sSource       As String
    
    sSource = "BOCatalogueBulkUpdate.BulkUpdateClass"
    
    'Check if the set property is available for tthe Class level
    bCanUpdate = m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(ctClass, m_eUpdateField)
    
    'If the current property is available for Class then generate the corresponding SQL
    '  to update the field in the corresponding table.
    If bCanUpdate Then
        Select Case m_eUpdateField
            'Main table
            'SQL to update both the TITLE fields
            Case ECATALOGUEPROPERTIES.cpTitle
                sSql = "update " & gk_sDBTABLEEditsetClass & " Set " & _
                                    gk_sDBCOLClassTitle & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetClassExtension & _
                        " Where " & gk_sDBTABLEEditsetClass & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetClassExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetClass & "." & gk_sDBCOLClassID & _
                                    " = " & gk_sDBTABLEEditsetClassExtension & "." & _
                                    gk_sDBCOLClassID & _
                        " AND " & gk_sDBTABLEEditsetClassExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                        
                        
                sSql = sSql & "  UPDATE " & gk_sDBTABLEEditsetClassTitle & " Set " & _
                                    gk_sDBCOLTitle & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetClassExtension & _
                        " Where " & gk_sDBTABLEEditsetClassTitle & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetClassExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetClassTitle & "." & gk_sDBCOLClassID & _
                                    " = " & gk_sDBTABLEEditsetClassExtension & "." & _
                                    gk_sDBCOLClassID & _
                        " AND " & gk_sDBTABLEEditsetClassExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            
            'SQL to update the CLOSURE STATUS ID Link
            Case ECATALOGUEPROPERTIES.cpClosureStatus
                
                sSql = "update " & gk_sDBTABLEEditsetClass & " Set " & _
                                    gk_sDBCOLClosureStatus & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetClassExtension & _
                        " Where " & gk_sDBTABLEEditsetClass & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetClassExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetClass & "." & gk_sDBCOLClassID & _
                                    " = " & gk_sDBTABLEEditsetClassExtension & "." & _
                                    gk_sDBCOLClassID & _
                        " AND " & gk_sDBTABLEEditsetClassExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'SQL to update the COVERING DATES field
            Case ECATALOGUEPROPERTIES.cpCoveringDates
                
                sSql = "update " & gk_sDBTABLEEditsetClass & " Set " & _
                                    gk_sDBCOLDateText & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetClassExtension & _
                        " Where " & gk_sDBTABLEEditsetClass & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetClassExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetClass & "." & gk_sDBCOLClassID & _
                                    " = " & gk_sDBTABLEEditsetClassExtension & "." & _
                                    gk_sDBCOLClassID & _
                        " AND " & gk_sDBTABLEEditsetClassExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
                ' d19 We also need to update the start and end dates
                If Not IsNull(m_vUpdateValue2) Then
                    
                    sStartDateSql = "update " & gk_sDBTABLEEditsetClass & " Set " & _
                                    gk_sDBCOLFirstDate & " = " & _
                                    StringToNullOrString(Left(m_vUpdateValue2, 8)) & _
                        " FROM " & gk_sDBTABLEEditsetClassExtension & _
                        " Where " & gk_sDBTABLEEditsetClass & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetClassExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetClass & "." & gk_sDBCOLClassID & _
                                    " = " & gk_sDBTABLEEditsetClassExtension & "." & _
                                    gk_sDBCOLClassID & _
                        " AND " & gk_sDBTABLEEditsetClassExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                                        
                    sEndDateSql = "update " & gk_sDBTABLEEditsetClass & " Set " & _
                                    gk_sDBCOLLastDate & " = " & _
                                    StringToNullOrString(Right(m_vUpdateValue2, 8)) & _
                        " FROM " & gk_sDBTABLEEditsetClassExtension & _
                        " Where " & gk_sDBTABLEEditsetClass & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetClassExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetClass & "." & gk_sDBCOLClassID & _
                                    " = " & gk_sDBTABLEEditsetClassExtension & "." & _
                                    gk_sDBCOLClassID & _
                        " AND " & gk_sDBTABLEEditsetClassExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                                        
                End If
                            
                            
            'Extension table
            'SQL to update the FORMER REFERENCE (Department) field
            Case ECATALOGUEPROPERTIES.cpFormerDepartmentReference
                
                sSql = "update " & gk_sDBTABLEEditsetClassExtension & " Set " & _
                                    gk_sDBCOLFormerDepartmentReference & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'SQL to update the FORMER PRO REFERENCE  field
            Case ECATALOGUEPROPERTIES.cpFormerPROReference
                
                sSql = "update " & gk_sDBTABLEEditsetClassExtension & " Set " & _
                                    gk_sDBCOLFormerPROReference & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            
            'SQL to update the PHYSICAL CONDITION ID Link
            Case ECATALOGUEPROPERTIES.cpPhysicalCondition
                
                sSql = "update " & gk_sDBTABLEEditsetClassExtension & " Set " & _
                                    gk_sDBCOLPhysicalCondition & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'SQL to update the PHYSICAL DESCRIPTION RECORD QUANTITY field
            Case ECATALOGUEPROPERTIES.cpPhysicalDescription_RecordQuantity
                
                sSql = "update " & gk_sDBTABLEEditsetClassExtension & " Set " & _
                                    gk_sDBCOLPhysicalDescriptionRecordQuantity & " = " & _
                                    NullToSQLNull(StringToNullOrLong(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'SQL to update the PHYSICAL DESCRIPTION FORM ID Link
            Case ECATALOGUEPROPERTIES.cpPhysicalDescription_Form
                
                sSql = "update " & gk_sDBTABLEEditsetClassExtension & " Set " & _
                                    gk_sDBCOLPhysicalDescriptionFormID & " = " & _
                                    NullToSQLNull(StringToNullOrLong(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'SQL to update the PHYSICAL DESCRIPTION DIMENSION field
            Case ECATALOGUEPROPERTIES.cpPhysicalDescription_Dimension
                
                sSql = "update " & gk_sDBTABLEEditsetClassExtension & " Set " & _
                                    gk_sDBCOLPhysicalDescriptionDimension & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'SQL to update the MAP DESIGNATION field
            Case ECATALOGUEPROPERTIES.cpMapDesignation
                
                sSql = "update " & gk_sDBTABLEEditsetClassExtension & " Set " & _
                                    gk_sDBCOLMapDesignation & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'SQL to update the MAP SCALE NUMBER field
            Case ECATALOGUEPROPERTIES.cpMapscaleNumber
                
                sSql = "update " & gk_sDBTABLEEditsetClassExtension & " Set " & _
                                    gk_sDBCOLMapScaleNumber & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'SQL to update the ACCESS CONDITION ID Link
            Case ECATALOGUEPROPERTIES.cpAccessCondition
                
                sSql = "update " & gk_sDBTABLEEditsetClassExtension & " Set " & _
                                    gk_sDBCOLAccessConditionID & " = " & _
                                    NullToSQLNull(StringToNullOrLong(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'SQL to update the RESTRICTIONS ON USE field
            Case ECATALOGUEPROPERTIES.cpRestrictionsOnUse
                
                sSql = "update " & gk_sDBTABLEEditsetClassExtension & " Set " & _
                                    gk_sDBCOLRestrictionsOnUse & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'SQL to update the ACCUMULATION DATES field
            Case ECATALOGUEPROPERTIES.cpAccumulationDates
                
                sSql = "update " & gk_sDBTABLEEditsetClassExtension & " Set " & _
                                    gk_sDBCOLAccumulationdates & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                        
            'SQL to update the RECORD OPENING DATE field
            Case ECATALOGUEPROPERTIES.cpRecordOpeningDate
                
                sSql = "update " & gk_sDBTABLEEditsetClassExtension & " Set " & _
                                    gk_sDBCOLRecordOpeningDate & " = " & _
                                    QuoteValueWithQuotes(VariantDateAsDDMMMYYYY(StringToNullOrDate(m_vUpdateValue1))) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'SQL to update the LANGUAGE ID Link
            Case ECATALOGUEPROPERTIES.cpLanguage
                
                sSql = "update " & gk_sDBTABLEEditsetClassExtension & " Set " & _
                                    gk_sDBCOLLanguageID & " = " & _
                                    NullToSQLNull(StringToNullOrLong(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            
            'SQL to update the LEGAL STATUS ID Link
            Case ECATALOGUEPROPERTIES.cpLegalStatus
                
                sSql = "update " & gk_sDBTABLEEditsetClassExtension & " Set " & _
                                    gk_sDBCOLLegalStatusID & " = " & _
                                    NullToSQLNull(StringToNullOrLong(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'SQL to update the CLOSURE CODE value and CLOSURE TYPE lookup
            Case ECATALOGUEPROPERTIES.cpClosureCode, ECATALOGUEPROPERTIES.cpClosureType
                
                sSql = " update " & gk_sDBTABLEEditsetClassExtension & " Set " & _
                                    gk_sDBCOLClosureType & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly) & vbNewLine
                
                sSql = sSql & " update " & gk_sDBTABLEEditsetClassExtension & " Set " & _
                                    gk_sDBCOLClosureCode & " = " & _
                                    NullToSQLNull(StringToNullOrLong(m_vUpdateValue2)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            
            'Seperate tables, one property per table
            'SQL to update the seperate SCOPE AND CONTENT field
            Case ECATALOGUEPROPERTIES.cpScopeAndContent
                
                sSql = "update " & gk_sDBTABLEEditsetClassScopeAndContent & " Set " & _
                                    gk_sDBCOLScopeAndContent & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetClassExtension & _
                        " Where " & gk_sDBTABLEEditsetClassScopeAndContent & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetClassExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetClassScopeAndContent & "." & gk_sDBCOLClassID & _
                                    " = " & gk_sDBTABLEEditsetClassExtension & "." & _
                                    gk_sDBCOLClassID & _
                        " AND " & gk_sDBTABLEEditsetClassExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'SQL to update the seperate ARRANGAMENT field
            Case ECATALOGUEPROPERTIES.cpArrangement
                
                sSql = "update " & gk_sDBTABLEEditsetClassArrangement & " Set " & _
                                    gk_sDBCOLArrangement & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetClassExtension & _
                        " Where " & gk_sDBTABLEEditsetClassArrangement & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetClassExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetClassArrangement & "." & gk_sDBCOLClassID & _
                                    " = " & gk_sDBTABLEEditsetClassExtension & "." & _
                                    gk_sDBCOLClassID & _
                        " AND " & gk_sDBTABLEEditsetClassExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'SQL to update the seperate ADMINISTRATIVE OR BIOGRAPHICAL HISTORY field
            Case ECATALOGUEPROPERTIES.cpAdministrative_BiographicalHistory
                
                sSql = "update " & gk_sDBTABLEEditsetClassAdministrativeHistory & " Set " & _
                                    gk_sDBCOLAdministrativeHistory & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetClassExtension & _
                        " Where " & gk_sDBTABLEEditsetClassAdministrativeHistory & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetClassExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetClassAdministrativeHistory & "." & gk_sDBCOLClassID & _
                                    " = " & gk_sDBTABLEEditsetClassExtension & "." & _
                                    gk_sDBCOLClassID & _
                        " AND " & gk_sDBTABLEEditsetClassExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'SQL to update the CUSTODIAL HISTORY field
            Case ECATALOGUEPROPERTIES.cpCustodialHistory
                
                sSql = "update " & gk_sDBTABLEEditsetClassCustodialHistory & " Set " & _
                                    gk_sDBCOLCustodialHistory & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetClassExtension & _
                        " Where " & gk_sDBTABLEEditsetClassCustodialHistory & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetClassExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetClassCustodialHistory & "." & gk_sDBCOLClassID & _
                                    " = " & gk_sDBTABLEEditsetClassExtension & "." & _
                                    gk_sDBCOLClassID & _
                        " AND " & gk_sDBTABLEEditsetClassExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'SQL to update the APPRAISAL DESTRUCTION INFORMATION field
            Case ECATALOGUEPROPERTIES.cpAppraisal_DestructionInformation
                
                sSql = "update " & gk_sDBTABLEEditsetClassAppraisalDestructionInformation & _
                        " Set " & gk_sDBCOLAppraisalDestructionInformation & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetClassExtension & _
                        " Where " & gk_sDBTABLEEditsetClassAppraisalDestructionInformation & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetClassExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetClassAppraisalDestructionInformation & "." & gk_sDBCOLClassID & _
                                    " = " & gk_sDBTABLEEditsetClassExtension & "." & _
                                    gk_sDBCOLClassID & _
                        " AND " & gk_sDBTABLEEditsetClassExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'SQL to update the ACCRUALS field
            Case ECATALOGUEPROPERTIES.cpAccruals
                
                sSql = "update " & gk_sDBTABLEEditsetClassAccruals & " Set " & _
                                    gk_sDBCOLAccruals & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetClassExtension & _
                        " Where " & gk_sDBTABLEEditsetClassAccruals & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetClassExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetClassAccruals & "." & gk_sDBCOLClassID & _
                                    " = " & gk_sDBTABLEEditsetClassExtension & "." & _
                                    gk_sDBCOLClassID & _
                        " AND " & gk_sDBTABLEEditsetClassExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
        
            'PC TABLES
            'SQL to update the IMMEDIATE SOURCE OF ACQUISITION description field ONLY
            Case ECATALOGUEPROPERTIES.cpImmediateSourceOfAcquisition
                    
                UpdateLinkDataForImmediateSource ECATALOGUELEVEL.ctClass
            
            'SQL to update the COPIES INFORMATION description field ONLY
            Case ECATALOGUEPROPERTIES.cpCopiesInformation
                
                UpdateLinkDataForCopiesInfo ECATALOGUELEVEL.ctClass
            
            'SQL to update the PUBLICATION NOTE description field ONLY
            Case ECATALOGUEPROPERTIES.cpPublicationNote
            
                UpdateLinkDataForPublicationNote ECATALOGUELEVEL.ctClass
                
            'SQL to update the UNPUBLISHED FINDING AIDS description field ONLY
            Case ECATALOGUEPROPERTIES.cpUnpublishedFindingAids
                
                UpdateLinkDataForUnpublishedFindingAids ECATALOGUELEVEL.ctClass
        
            'd20 SQL to update the CREATOR NAMES description field ONLY
            Case ECATALOGUEPROPERTIES.cpCreatorNames
                
                UpdateLinkDataForCreatorNames ECATALOGUELEVEL.ctClass
                
            'd20 SQL to update the PLACE OF DEPOSIT description field ONLY
            Case ECATALOGUEPROPERTIES.cpPlaceOfDeposit
                
                UpdateLinkDataForPlaceOfDeposit ECATALOGUELEVEL.ctClass
                
            'd20 SQL to update the LOCATION OF ORIGINALS description field ONLY
            Case ECATALOGUEPROPERTIES.cpLocationOfOriginals
                
                UpdateLinkDataForLocationOfOriginals ECATALOGUELEVEL.ctClass
        
        End Select
        
        If Not (m_eUpdateField = ECATALOGUEPROPERTIES.cpImmediateSourceOfAcquisition Or _
            m_eUpdateField = ECATALOGUEPROPERTIES.cpCopiesInformation Or _
            m_eUpdateField = ECATALOGUEPROPERTIES.cpPublicationNote Or _
            m_eUpdateField = ECATALOGUEPROPERTIES.cpUnpublishedFindingAids Or _
            m_eUpdateField = ECATALOGUEPROPERTIES.cpCreatorNames Or _
            m_eUpdateField = ECATALOGUEPROPERTIES.cpPlaceOfDeposit Or _
            m_eUpdateField = ECATALOGUEPROPERTIES.cpLocationOfOriginals) Then
            
            m_ofwUpdate.Execute sSql
            
            ' d19 - Update the start AND end dates if possible
            If Not (sStartDateSql = vbNullString And sEndDateSql = vbNullString) Then
                m_ofwUpdate.Execute (sStartDateSql)
                m_ofwUpdate.Execute (sEndDateSql)
            End If
            
        End If
    
        'update the editset entry type to modified if not change for all catalogues the level
        sSql = "update " & gk_sDBTABLEEditsetClassExtension & " Set " & _
                gk_sDBCOLEntryTypeID & " = CASE " & gk_sDBCOLEntryTypeID & _
                    " WHEN " & CStr(ECATALOGUEENTRYSTATUS.cesPROCATCreated) & " THEN " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesPROCATCreatedModified) & _
                    " WHEN " & CStr(ECATALOGUEENTRYSTATUS.cesAddedFromLive) & " THEN " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesAddedFromLiveModified) & _
                    " WHEN " & CStr(ECATALOGUEENTRYSTATUS.cesImported) & " THEN " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesImportedModified) & _
                    " ELSE " & gk_sDBCOLEntryTypeID & _
                    " END " & _
                " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID)
    
        m_ofwUpdate.Execute sSql
    
    
        sSql = "Select " & gk_sDBTABLEEditsetClass & "." & gk_sDBCOLClassID & " as " & gk_sDBCOLCatalogueID & _
                ", 3 as " & gk_sDBCOLCatalogueLevel & _
                " from " & gk_sDBTABLEEditsetClass & "(nolock) " & _
                ", " & gk_sDBTABLEEditsetClassExtension & "(nolock) " & _
                " Where " & gk_sDBTABLEEditsetClass & "." & gk_sDBCOLEditSetID & _
                                " = " & CStr(m_lEditsetID) & _
                " AND " & gk_sDBTABLEEditsetClassExtension & "." & gk_sDBCOLClassID & _
                                " = " & gk_sDBTABLEEditsetClass & "." & gk_sDBCOLClassID & _
                " AND " & gk_sDBTABLEEditsetClassExtension & "." & gk_sDBCOLEditSetID & _
                                " = " & CStr(m_lEditsetID) & _
                " AND " & gk_sDBTABLEEditsetClassExtension & "." & _
                gk_sDBCOLEditSetEntryTypeID & " <> " & CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
        
        Set BulkUpdateClass = m_ofwUpdate.FindBySql(sSql, mk_iMaxNoOfRecords)
    
    End If
    Exit Function

ErrHandler:
    ConvertError sSource
End Function





Private Function BulkUpdateHeader() As FWRecordList
' Bulkupdates the value of the set property of the Header
' This procedure generates the update query based on the field to be updated and
'    executes the query.
' d2 Change Quotes to QuoteValueWithQuotes and updated entry type ID field
' d5    ksk     Changes for closure status ID being alphanumeric
'
' d6    ksk     changed the function to return a recordlist that will be a list of catalogues
'               affected by the change.
' d7    ksk     changed the physical condition field to be a text field insted of lookup
' d9    ksk     Added the extra join that was missing to link the extenstion
'               and main table for the list of records affected
' d11   ksk     Changed the condition in the update for imported matrial. Used the
'               wrong constant
' d12   ksk     Added a max constant to be used by all the Findbysql
'               To return the affected records
' d19   alm     12Feb2001       Updates the start / end dates automatically
'                               when a covering date is updated
' d20   alm     Supports Creator, Deposit and Originals
'
On Error GoTo ErrHandler
    
    Dim sSql          As String
    Dim sStartDateSql As String 'd19
    Dim sEndDateSql   As String 'd19
    Dim bCanUpdate    As Boolean
    Dim sSource       As String
    
    sSource = "BOCatalogueBulkUpdate.BulkUpdateHeader"
    
    'Check if the set property is available for tthe Header level
    bCanUpdate = m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(ctHeaders, m_eUpdateField)
    
    'If the current property is available for Header then generate the corresponding SQL
    '  to update the field in the corresponding table.
    If bCanUpdate Then
        Select Case m_eUpdateField
            'Main table
            'SQL to update both the TITLE fields
            Case ECATALOGUEPROPERTIES.cpTitle
                
                sSql = "update " & gk_sDBTABLEEditsetHeader & " Set " & _
                                    gk_sDBCOLHeaderTitle & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetHeaderExtension & _
                        " Where " & gk_sDBTABLEEditsetHeader & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetHeaderExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetHeader & "." & gk_sDBCOLHeaderID & _
                                    " = " & gk_sDBTABLEEditsetHeaderExtension & "." & _
                                    gk_sDBCOLHeaderID & _
                        " AND " & gk_sDBTABLEEditsetHeaderExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                
                sSql = sSql & "  UPDATE " & gk_sDBTABLEEditsetHeaderTitle & " Set " & _
                                    gk_sDBCOLTitle & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetHeaderExtension & _
                        " Where " & gk_sDBTABLEEditsetHeaderTitle & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetHeaderExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetHeaderTitle & "." & gk_sDBCOLHeaderID & _
                                    " = " & gk_sDBTABLEEditsetHeaderExtension & "." & _
                                    gk_sDBCOLHeaderID & _
                        " AND " & gk_sDBTABLEEditsetHeaderExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'SQL to update the CLOSURE STATUS ID Link
            Case ECATALOGUEPROPERTIES.cpClosureStatus
                
                sSql = "update " & gk_sDBTABLEEditsetHeader & " Set " & _
                                    gk_sDBCOLClosureStatus & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetHeaderExtension & _
                        " Where " & gk_sDBTABLEEditsetHeader & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetHeaderExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetHeader & "." & gk_sDBCOLHeaderID & _
                                    " = " & gk_sDBTABLEEditsetHeaderExtension & "." & _
                                    gk_sDBCOLHeaderID & _
                        " AND " & gk_sDBTABLEEditsetHeaderExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'SQL to update the COVERING DATES field
            Case ECATALOGUEPROPERTIES.cpCoveringDates
            
                sSql = "update " & gk_sDBTABLEEditsetHeader & " Set " & _
                                    gk_sDBCOLDateText & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetHeaderExtension & _
                        " Where " & gk_sDBTABLEEditsetHeader & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetHeaderExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetHeader & "." & gk_sDBCOLHeaderID & _
                                    " = " & gk_sDBTABLEEditsetHeaderExtension & "." & _
                                    gk_sDBCOLHeaderID & _
                        " AND " & gk_sDBTABLEEditsetHeaderExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                    
                ' d19 We also need to update the start and end dates
                If Not IsNull(m_vUpdateValue2) Then
                    
                    sStartDateSql = "update " & gk_sDBTABLEEditsetHeaderExtension & " Set " & _
                                    gk_sDBCOLCoveringFromDate & " = " & _
                                    StringToNullOrString(Left(m_vUpdateValue2, 8)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)

                    sEndDateSql = "update " & gk_sDBTABLEEditsetHeaderExtension & " Set " & _
                                    gk_sDBCOLCoveringToDate & " = " & _
                                    StringToNullOrString(Right(m_vUpdateValue2, 8)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                                    
                End If
            
            'Extension table
            'SQL to update the FORMER REFERENCE (Department) field
            Case ECATALOGUEPROPERTIES.cpFormerDepartmentReference
                
                sSql = "update " & gk_sDBTABLEEditsetHeaderExtension & " Set " & _
                                    gk_sDBCOLFormerDepartmentReference & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'SQL to update the FORMER PRO REFERENCE  field
            Case ECATALOGUEPROPERTIES.cpFormerPROReference
            
                sSql = "update " & gk_sDBTABLEEditsetHeaderExtension & " Set " & _
                                    gk_sDBCOLFormerPROReference & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                    
                                    
            'SQL to update the PHYSICAL CONDITION ID Link
            Case ECATALOGUEPROPERTIES.cpPhysicalCondition
            
                sSql = "update " & gk_sDBTABLEEditsetHeaderExtension & " Set " & _
                                    gk_sDBCOLPhysicalCondition & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                        
            'SQL to update the PHYSICAL DESCRIPTION RECORD QUANTITY field
            Case ECATALOGUEPROPERTIES.cpPhysicalDescription_RecordQuantity
            
                sSql = "update " & gk_sDBTABLEEditsetHeaderExtension & " Set " & _
                                    gk_sDBCOLPhysicalDescriptionRecordQuantity & " = " & _
                                    NullToSQLNull(StringToNullOrLong(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                    
            'SQL to update the PHYSICAL DESCRIPTION FORM ID Link
            Case ECATALOGUEPROPERTIES.cpPhysicalDescription_Form
            
                sSql = "update " & gk_sDBTABLEEditsetHeaderExtension & " Set " & _
                                    gk_sDBCOLPhysicalDescriptionFormID & " = " & _
                                    NullToSQLNull(StringToNullOrLong(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                    
            'SQL to update the PHYSICAL DESCRIPTION DIMENSION field
            Case ECATALOGUEPROPERTIES.cpPhysicalDescription_Dimension
            
                sSql = "update " & gk_sDBTABLEEditsetHeaderExtension & " Set " & _
                                    gk_sDBCOLPhysicalDescriptionDimension & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                    
            'SQL to update the MAP DESIGNATION field
            Case ECATALOGUEPROPERTIES.cpMapDesignation
            
                sSql = "update " & gk_sDBTABLEEditsetHeaderExtension & " Set " & _
                                    gk_sDBCOLMapDesignation & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                    
            'SQL to update the MAP SCALE NUMBER field
            Case ECATALOGUEPROPERTIES.cpMapscaleNumber
            
                sSql = "update " & gk_sDBTABLEEditsetHeaderExtension & " Set " & _
                                    gk_sDBCOLMapScaleNumber & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                    
            'SQL to update the ACCESS CONDITION ID Link
            Case ECATALOGUEPROPERTIES.cpAccessCondition
            
                sSql = "update " & gk_sDBTABLEEditsetHeaderExtension & " Set " & _
                                    gk_sDBCOLAccessConditionID & " = " & _
                                    NullToSQLNull(StringToNullOrLong(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                    
            'SQL to update the RESTRICTIONS ON USE field
            Case ECATALOGUEPROPERTIES.cpRestrictionsOnUse
            
                sSql = "update " & gk_sDBTABLEEditsetHeaderExtension & " Set " & _
                                    gk_sDBCOLRestrictionsOnUse & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                    
            'SQL to update the ACCUMULATION DATES field
            Case ECATALOGUEPROPERTIES.cpAccumulationDates
            
                sSql = "update " & gk_sDBTABLEEditsetHeaderExtension & " Set " & _
                                    gk_sDBCOLAccumulationdates & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                    
            'SQL to update the RECORD OPENING DATE field
            Case ECATALOGUEPROPERTIES.cpRecordOpeningDate
            
                sSql = "update " & gk_sDBTABLEEditsetHeaderExtension & " Set " & _
                                    gk_sDBCOLRecordOpeningDate & " = " & _
                                    QuoteValueWithQuotes(VariantDateAsDDMMMYYYY(StringToNullOrDate(m_vUpdateValue1))) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                    
            'SQL to update the LANGUAGE ID Link
            Case ECATALOGUEPROPERTIES.cpLanguage
            
                sSql = "update " & gk_sDBTABLEEditsetHeaderExtension & " Set " & _
                                    gk_sDBCOLLanguageID & " = " & _
                                    NullToSQLNull(StringToNullOrLong(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                    
            'SQL to update the LEGAL STATUS ID Link
            Case ECATALOGUEPROPERTIES.cpLegalStatus
            
                sSql = "update " & gk_sDBTABLEEditsetHeaderExtension & " Set " & _
                                    gk_sDBCOLLegalStatusID & " = " & _
                                    NullToSQLNull(StringToNullOrLong(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                    
            'SQL to update the CLOSURE CODE value and CLOSURE TYPE lookup
            Case ECATALOGUEPROPERTIES.cpClosureCode, ECATALOGUEPROPERTIES.cpClosureType
                
                sSql = " update " & gk_sDBTABLEEditsetHeaderExtension & " Set " & _
                                    gk_sDBCOLClosureType & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly) & vbNewLine
                
                sSql = sSql & " update " & gk_sDBTABLEEditsetHeaderExtension & " Set " & _
                                    gk_sDBCOLClosureCode & " = " & _
                                    NullToSQLNull(StringToNullOrLong(m_vUpdateValue2)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
                                    
            
            'Seperate tables, one property per table
            'SQL to update the seperate SCOPE AND CONTENT field
            Case ECATALOGUEPROPERTIES.cpScopeAndContent
                
                sSql = "update " & gk_sDBTABLEEditsetHeaderScopeAndContent & " Set " & _
                                    gk_sDBCOLScopeAndContent & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetHeaderExtension & _
                        " Where " & gk_sDBTABLEEditsetHeaderScopeAndContent & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetHeaderExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetHeaderScopeAndContent & "." & gk_sDBCOLHeaderID & _
                                    " = " & gk_sDBTABLEEditsetHeaderExtension & "." & _
                                    gk_sDBCOLHeaderID & _
                        " AND " & gk_sDBTABLEEditsetHeaderExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'SQL to update the seperate ARRANGAMENT field
            Case ECATALOGUEPROPERTIES.cpArrangement
                
                sSql = "update " & gk_sDBTABLEEditsetHeaderArrangement & " Set " & _
                                    gk_sDBCOLArrangement & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetHeaderExtension & _
                        " Where " & gk_sDBTABLEEditsetHeaderArrangement & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetHeaderExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetHeaderArrangement & "." & gk_sDBCOLHeaderID & _
                                    " = " & gk_sDBTABLEEditsetHeaderExtension & "." & _
                                    gk_sDBCOLHeaderID & _
                        " AND " & gk_sDBTABLEEditsetHeaderExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'SQL to update the seperate ADMINISTRATIVE OR BIOGRAPHICAL HISTORY field
            Case ECATALOGUEPROPERTIES.cpAdministrative_BiographicalHistory
                
                sSql = "update " & gk_sDBTABLEEditsetHeaderAdministrativeHistory & " Set " & _
                                    gk_sDBCOLAdministrativeHistory & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetHeaderExtension & _
                        " Where " & gk_sDBTABLEEditsetHeaderAdministrativeHistory & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetHeaderExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetHeaderAdministrativeHistory & "." & gk_sDBCOLHeaderID & _
                                    " = " & gk_sDBTABLEEditsetHeaderExtension & "." & _
                                    gk_sDBCOLHeaderID & _
                        " AND " & gk_sDBTABLEEditsetHeaderExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'SQL to update the CUSTODIAL HISTORY field
            Case ECATALOGUEPROPERTIES.cpCustodialHistory
                
                sSql = "update " & gk_sDBTABLEEditsetHeaderCustodialHistory & " Set " & _
                                    gk_sDBCOLCustodialHistory & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetHeaderExtension & _
                        " Where " & gk_sDBTABLEEditsetHeaderCustodialHistory & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetHeaderExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetHeaderCustodialHistory & "." & gk_sDBCOLHeaderID & _
                                    " = " & gk_sDBTABLEEditsetHeaderExtension & "." & _
                                    gk_sDBCOLHeaderID & _
                        " AND " & gk_sDBTABLEEditsetHeaderExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'SQL to update the APPRAISAL DESTRUCTION INFORMATION field
            Case ECATALOGUEPROPERTIES.cpAppraisal_DestructionInformation
                
                sSql = "update " & gk_sDBTABLEEditsetHeaderAppraisalDestructionInformation & _
                        " Set " & gk_sDBCOLAppraisalDestructionInformation & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetHeaderExtension & _
                        " Where " & gk_sDBTABLEEditsetHeaderAppraisalDestructionInformation & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetHeaderExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetHeaderAppraisalDestructionInformation & "." & gk_sDBCOLHeaderID & _
                                    " = " & gk_sDBTABLEEditsetHeaderExtension & "." & _
                                    gk_sDBCOLHeaderID & _
                        " AND " & gk_sDBTABLEEditsetHeaderExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'SQL to update the ACCRUALS field
            Case ECATALOGUEPROPERTIES.cpAccruals
                
                sSql = "update " & gk_sDBTABLEEditsetHeaderAccruals & " Set " & _
                                    gk_sDBCOLAccruals & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetHeaderExtension & _
                        " Where " & gk_sDBTABLEEditsetHeaderAccruals & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetHeaderExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetHeaderAccruals & "." & gk_sDBCOLHeaderID & _
                                    " = " & gk_sDBTABLEEditsetHeaderExtension & "." & _
                                    gk_sDBCOLHeaderID & _
                        " AND " & gk_sDBTABLEEditsetHeaderExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
        
            'PC TABLES
            'SQL to update the IMMEDIATE SOURCE OF ACQUISITION description field ONLY
            Case ECATALOGUEPROPERTIES.cpImmediateSourceOfAcquisition
                    
                UpdateLinkDataForImmediateSource ECATALOGUELEVEL.ctHeaders
            
            'SQL to update the COPIES INFORMATION description field ONLY
            Case ECATALOGUEPROPERTIES.cpCopiesInformation
                
                UpdateLinkDataForCopiesInfo ECATALOGUELEVEL.ctHeaders
            
            'SQL to update the PUBLICATION NOTE description field ONLY
            Case ECATALOGUEPROPERTIES.cpPublicationNote
            
                UpdateLinkDataForPublicationNote ECATALOGUELEVEL.ctHeaders
                
            'SQL to update the UNPUBLISHED FINDING AIDS description field ONLY
            Case ECATALOGUEPROPERTIES.cpUnpublishedFindingAids
                
                UpdateLinkDataForUnpublishedFindingAids ECATALOGUELEVEL.ctHeaders
        
            'd20 SQL to update the CREATOR NAMES description field ONLY
            Case ECATALOGUEPROPERTIES.cpCreatorNames
                
                UpdateLinkDataForCreatorNames ECATALOGUELEVEL.ctHeaders
                
            'd20 SQL to update the PLACE OF DEPOSIT description field ONLY
            Case ECATALOGUEPROPERTIES.cpPlaceOfDeposit
                
                UpdateLinkDataForPlaceOfDeposit ECATALOGUELEVEL.ctHeaders
                
            'd20 SQL to update the LOCATION OF ORIGINALS description field ONLY
            Case ECATALOGUEPROPERTIES.cpLocationOfOriginals
                
                UpdateLinkDataForLocationOfOriginals ECATALOGUELEVEL.ctHeaders
        
        End Select
        
        If Not (m_eUpdateField = ECATALOGUEPROPERTIES.cpImmediateSourceOfAcquisition Or _
            m_eUpdateField = ECATALOGUEPROPERTIES.cpCopiesInformation Or _
            m_eUpdateField = ECATALOGUEPROPERTIES.cpPublicationNote Or _
            m_eUpdateField = ECATALOGUEPROPERTIES.cpUnpublishedFindingAids Or _
            m_eUpdateField = ECATALOGUEPROPERTIES.cpCreatorNames Or _
            m_eUpdateField = ECATALOGUEPROPERTIES.cpPlaceOfDeposit Or _
            m_eUpdateField = ECATALOGUEPROPERTIES.cpLocationOfOriginals) Then
            
            m_ofwUpdate.Execute sSql
            
            ' d19 - Update the start AND end dates if possible
            If Not (sStartDateSql = vbNullString And sEndDateSql = vbNullString) Then
                m_ofwUpdate.Execute (sStartDateSql)
                m_ofwUpdate.Execute (sEndDateSql)
            End If
        End If
    
        'update the editset entry type to modified if not change for all catalogues the level
        sSql = "update " & gk_sDBTABLEEditsetHeaderExtension & " Set " & _
                gk_sDBCOLEntryTypeID & " = CASE " & gk_sDBCOLEntryTypeID & _
                    " WHEN " & CStr(ECATALOGUEENTRYSTATUS.cesPROCATCreated) & " THEN " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesPROCATCreatedModified) & _
                    " WHEN " & CStr(ECATALOGUEENTRYSTATUS.cesAddedFromLive) & " THEN " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesAddedFromLiveModified) & _
                    " WHEN " & CStr(ECATALOGUEENTRYSTATUS.cesImported) & " THEN " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesImportedModified) & _
                    " ELSE " & gk_sDBCOLEntryTypeID & _
                    " END " & _
                " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID)
    
        m_ofwUpdate.Execute sSql
    
        sSql = "Select " & gk_sDBTABLEEditsetHeader & "." & gk_sDBCOLHeaderID & " as " & gk_sDBCOLCatalogueID & _
                ", 4 as " & gk_sDBCOLCatalogueLevel & _
                " from " & gk_sDBTABLEEditsetHeader & "(nolock) " & _
                ", " & gk_sDBTABLEEditsetHeaderExtension & "(nolock) " & _
                " Where " & gk_sDBTABLEEditsetHeader & "." & gk_sDBCOLEditSetID & _
                                " = " & CStr(m_lEditsetID) & _
                " AND " & gk_sDBTABLEEditsetHeaderExtension & "." & gk_sDBCOLHeaderID & _
                                " = " & gk_sDBTABLEEditsetHeader & "." & gk_sDBCOLHeaderID & _
                " AND " & gk_sDBTABLEEditsetHeaderExtension & "." & gk_sDBCOLEditSetID & _
                                " = " & CStr(m_lEditsetID) & _
                " AND " & gk_sDBTABLEEditsetHeaderExtension & "." & _
                gk_sDBCOLEditSetEntryTypeID & " <> " & CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
        
        Set BulkUpdateHeader = m_ofwUpdate.FindBySql(sSql, mk_iMaxNoOfRecords)
    End If
    
    Exit Function

ErrHandler:
    ConvertError sSource
End Function





Private Function BulkUpdateSubheader() As FWRecordList
' Bulkupdates the value of the set property of the Subheader
' This procedure generates the update query based on the field to be updated and
'    executes the query.
' d2 Change Quotes to QuoteValueWithQuotes and updated entry type ID field
' d5    ksk     Changes for closure status ID being alphanumeric
' d6    ksk     changed the function to return a recordlist that will be a list of catalogues
'               affected by the change.
' d7    ksk     changed the physical condition field to be a text field insted of lookup
' d9    ksk     Added the extra join that was missing to link the extenstion
'               and main table for the list of records affected
' d11   ksk     Changed the condition in the update for imported matrial. Used the
'               wrong constant
' d12   ksk     Added a max constant to be used by all the Findbysql
'               To return the affected records
' d19   alm     12Feb2001       Updates the start / end dates automatically
'                               when a covering date is updated
' d20   alm     Supports Creator, Deposit and Originals
'
On Error GoTo ErrHandler

    Dim sSql          As String
    Dim sStartDateSql As String 'd19
    Dim sEndDateSql   As String 'd19
    Dim bCanUpdate    As Boolean
    Dim sSource       As String
    
    sSource = "BOCatalogueBulkUpdate.BulkUpdateSubheader"
    
    'Check if the set property is available for tthe Subheader level
    bCanUpdate = m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(ctSubHeaders, m_eUpdateField)
    
    'If the current property is available for Subheader then generate the corresponding SQL
    '  to update the field in the corresponding table.
    If bCanUpdate Then
        Select Case m_eUpdateField
            'Main table
            'SQL to update both the TITLE fields
            Case ECATALOGUEPROPERTIES.cpTitle
                sSql = "update " & gk_sDBTABLEEditsetSubheader & " Set " & _
                                    gk_sDBCOLSubheaderTitle & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetSubheaderExtension & _
                        " Where " & gk_sDBTABLEEditsetSubheader & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetSubheaderExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetSubheader & "." & gk_sDBCOLSubHeaderID & _
                                    " = " & gk_sDBTABLEEditsetSubheaderExtension & "." & _
                                    gk_sDBCOLSubHeaderID & _
                        " AND " & gk_sDBTABLEEditsetSubheaderExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                        
                sSql = sSql & "  UPDATE " & gk_sDBTABLEEditsetSubheaderTitle & " Set " & _
                                    gk_sDBCOLTitle & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetSubheaderExtension & _
                        " Where " & gk_sDBTABLEEditsetSubheaderTitle & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetSubheaderExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetSubheaderTitle & "." & gk_sDBCOLSubHeaderID & _
                                    " = " & gk_sDBTABLEEditsetSubheaderExtension & "." & _
                                    gk_sDBCOLSubHeaderID & _
                        " AND " & gk_sDBTABLEEditsetSubheaderExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'SQL to update the CLOSURE STATUS ID Link
            Case ECATALOGUEPROPERTIES.cpClosureStatus
                sSql = "update " & gk_sDBTABLEEditsetSubheader & " Set " & _
                                    gk_sDBCOLClosureStatus & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetSubheaderExtension & _
                        " Where " & gk_sDBTABLEEditsetSubheader & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetSubheaderExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetSubheader & "." & gk_sDBCOLSubHeaderID & _
                                    " = " & gk_sDBTABLEEditsetSubheaderExtension & "." & _
                                    gk_sDBCOLSubHeaderID & _
                        " AND " & gk_sDBTABLEEditsetSubheaderExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'SQL to update the COVERING DATES field
            Case ECATALOGUEPROPERTIES.cpCoveringDates
            
                sSql = "update " & gk_sDBTABLEEditsetSubheader & " Set " & _
                                    gk_sDBCOLDateText & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetSubheaderExtension & _
                        " Where " & gk_sDBTABLEEditsetSubheader & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetSubheaderExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetSubheader & "." & gk_sDBCOLSubHeaderID & _
                                    " = " & gk_sDBTABLEEditsetSubheaderExtension & "." & _
                                    gk_sDBCOLSubHeaderID & _
                        " AND " & gk_sDBTABLEEditsetSubheaderExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                                                 
                ' d19 We also need to update the start and end dates
                If Not IsNull(m_vUpdateValue2) Then
                    
                    sStartDateSql = "update " & gk_sDBTABLEEditsetSubheaderExtension & " Set " & _
                                    gk_sDBCOLCoveringFromDate & " = " & _
                                    StringToNullOrString(Left(m_vUpdateValue2, 8)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                                        
                    sEndDateSql = "update " & gk_sDBTABLEEditsetSubheaderExtension & " Set " & _
                                    gk_sDBCOLCoveringToDate & " = " & _
                                    StringToNullOrString(Right(m_vUpdateValue2, 8)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                                    
                End If
                
            'Extension table
            'SQL to update the FORMER REFERENCE (Department) field
            Case ECATALOGUEPROPERTIES.cpFormerDepartmentReference
                
                sSql = "update " & gk_sDBTABLEEditsetSubheaderExtension & " Set " & _
                                    gk_sDBCOLFormerDepartmentReference & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'SQL to update the FORMER PRO REFERENCE  field
            Case ECATALOGUEPROPERTIES.cpFormerPROReference
            
                sSql = "update " & gk_sDBTABLEEditsetSubheaderExtension & " Set " & _
                                    gk_sDBCOLFormerPROReference & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                    
            'SQL to update the PHYSICAL CONDITION ID Link
            Case ECATALOGUEPROPERTIES.cpPhysicalCondition
            
                sSql = "update " & gk_sDBTABLEEditsetSubheaderExtension & " Set " & _
                                    gk_sDBCOLPhysicalCondition & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                    
            'SQL to update the PHYSICAL DESCRIPTION RECORD QUANTITY field
            Case ECATALOGUEPROPERTIES.cpPhysicalDescription_RecordQuantity
            
                sSql = "update " & gk_sDBTABLEEditsetSubheaderExtension & " Set " & _
                                    gk_sDBCOLPhysicalDescriptionRecordQuantity & " = " & _
                                    NullToSQLNull(StringToNullOrLong(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                        
            'SQL to update the PHYSICAL DESCRIPTION FORM ID Link
            Case ECATALOGUEPROPERTIES.cpPhysicalDescription_Form
            
                sSql = "update " & gk_sDBTABLEEditsetSubheaderExtension & " Set " & _
                                    gk_sDBCOLPhysicalDescriptionFormID & " = " & _
                                    NullToSQLNull(StringToNullOrLong(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                    
            'SQL to update the PHYSICAL DESCRIPTION DIMENSION field
            Case ECATALOGUEPROPERTIES.cpPhysicalDescription_Dimension
            
                sSql = "update " & gk_sDBTABLEEditsetSubheaderExtension & " Set " & _
                                    gk_sDBCOLPhysicalDescriptionDimension & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                    
            'SQL to update the MAP DESIGNATION field
            Case ECATALOGUEPROPERTIES.cpMapDesignation
            
                sSql = "update " & gk_sDBTABLEEditsetSubheaderExtension & " Set " & _
                                    gk_sDBCOLMapDesignation & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                    
            'SQL to update the MAP SCALE NUMBER field
            Case ECATALOGUEPROPERTIES.cpMapscaleNumber
            
                sSql = "update " & gk_sDBTABLEEditsetSubheaderExtension & " Set " & _
                                    gk_sDBCOLMapScaleNumber & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                    
            'SQL to update the ACCESS CONDITION ID Link
            Case ECATALOGUEPROPERTIES.cpAccessCondition
            
                sSql = "update " & gk_sDBTABLEEditsetSubheaderExtension & " Set " & _
                                    gk_sDBCOLAccessConditionID & " = " & _
                                    NullToSQLNull(StringToNullOrLong(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                    
            'SQL to update the RESTRICTIONS ON USE field
            Case ECATALOGUEPROPERTIES.cpRestrictionsOnUse
            
                sSql = "update " & gk_sDBTABLEEditsetSubheaderExtension & " Set " & _
                                    gk_sDBCOLRestrictionsOnUse & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                    
            'SQL to update the ACCUMULATION DATES field
            Case ECATALOGUEPROPERTIES.cpAccumulationDates
            
                sSql = "update " & gk_sDBTABLEEditsetSubheaderExtension & " Set " & _
                                    gk_sDBCOLAccumulationdates & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                    
            'SQL to update the RECORD OPENING DATE field
            Case ECATALOGUEPROPERTIES.cpRecordOpeningDate
            
                sSql = "update " & gk_sDBTABLEEditsetSubheaderExtension & " Set " & _
                                    gk_sDBCOLRecordOpeningDate & " = " & _
                                    QuoteValueWithQuotes(VariantDateAsDDMMMYYYY(StringToNullOrDate(m_vUpdateValue1))) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                    
            'SQL to update the LANGUAGE ID Link
            Case ECATALOGUEPROPERTIES.cpLanguage
            
                sSql = "update " & gk_sDBTABLEEditsetSubheaderExtension & " Set " & _
                                    gk_sDBCOLLanguageID & " = " & _
                                    NullToSQLNull(StringToNullOrLong(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                    
            'SQL to update the LEGAL STATUS ID Link
            Case ECATALOGUEPROPERTIES.cpLegalStatus
            
                sSql = "update " & gk_sDBTABLEEditsetSubheaderExtension & " Set " & _
                                    gk_sDBCOLLegalStatusID & " = " & _
                                    NullToSQLNull(StringToNullOrLong(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'SQL to update the CLOSURE CODE value and CLOSURE TYPE lookup
            Case ECATALOGUEPROPERTIES.cpClosureCode, ECATALOGUEPROPERTIES.cpClosureType
                
                sSql = " update " & gk_sDBTABLEEditsetSubheaderExtension & " Set " & _
                                    gk_sDBCOLClosureType & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly) & vbNewLine
                
                sSql = sSql & " update " & gk_sDBTABLEEditsetSubheaderExtension & " Set " & _
                                    gk_sDBCOLClosureCode & " = " & _
                                    NullToSQLNull(StringToNullOrLong(m_vUpdateValue2)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            
            
            
            'Seperate tables, one property per table
            'SQL to update the seperate SCOPE AND CONTENT field
            Case ECATALOGUEPROPERTIES.cpScopeAndContent
                
                sSql = "update " & gk_sDBTABLEEditsetSubheaderScopeAndContent & " Set " & _
                                    gk_sDBCOLScopeAndContent & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetSubheaderExtension & _
                        " Where " & gk_sDBTABLEEditsetSubheaderScopeAndContent & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetSubheaderExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetSubheaderScopeAndContent & "." & gk_sDBCOLSubHeaderID & _
                                    " = " & gk_sDBTABLEEditsetSubheaderExtension & "." & _
                                    gk_sDBCOLSubHeaderID & _
                        " AND " & gk_sDBTABLEEditsetSubheaderExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'SQL to update the seperate ARRANGAMENT field
            Case ECATALOGUEPROPERTIES.cpArrangement
                
                sSql = "update " & gk_sDBTABLEEditsetSubheaderArrangement & " Set " & _
                                    gk_sDBCOLArrangement & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetSubheaderExtension & _
                        " Where " & gk_sDBTABLEEditsetSubheaderArrangement & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetSubheaderExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetSubheaderArrangement & "." & gk_sDBCOLSubHeaderID & _
                                    " = " & gk_sDBTABLEEditsetSubheaderExtension & "." & _
                                    gk_sDBCOLSubHeaderID & _
                        " AND " & gk_sDBTABLEEditsetSubheaderExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'SQL to update the seperate ADMINISTRATIVE OR BIOGRAPHICAL HISTORY field
            Case ECATALOGUEPROPERTIES.cpAdministrative_BiographicalHistory
                
                sSql = "update " & gk_sDBTABLEEditsetSubheaderAdministrativeHistory & " Set " & _
                                    gk_sDBCOLAdministrativeHistory & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetSubheaderExtension & _
                        " Where " & gk_sDBTABLEEditsetSubheaderAdministrativeHistory & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetSubheaderExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetSubheaderAdministrativeHistory & "." & gk_sDBCOLSubHeaderID & _
                                    " = " & gk_sDBTABLEEditsetSubheaderExtension & "." & _
                                    gk_sDBCOLSubHeaderID & _
                        " AND " & gk_sDBTABLEEditsetSubheaderExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'SQL to update the CUSTODIAL HISTORY field
            Case ECATALOGUEPROPERTIES.cpCustodialHistory
                
                sSql = "update " & gk_sDBTABLEEditsetSubheaderCustodialHistory & " Set " & _
                                    gk_sDBCOLCustodialHistory & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetSubheaderExtension & _
                        " Where " & gk_sDBTABLEEditsetSubheaderCustodialHistory & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetSubheaderExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetSubheaderCustodialHistory & "." & gk_sDBCOLSubHeaderID & _
                                    " = " & gk_sDBTABLEEditsetSubheaderExtension & "." & _
                                    gk_sDBCOLSubHeaderID & _
                        " AND " & gk_sDBTABLEEditsetSubheaderExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            
            'SQL to update the APPRAISAL DESTRUCTION INFORMATION field
            Case ECATALOGUEPROPERTIES.cpAppraisal_DestructionInformation
                
                sSql = "update " & gk_sDBTABLEEditsetSubheaderAppraisalDestructionInformation & _
                        " Set " & gk_sDBCOLAppraisalDestructionInformation & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetSubheaderExtension & _
                        " Where " & gk_sDBTABLEEditsetSubheaderAppraisalDestructionInformation & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetSubheaderExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetSubheaderAppraisalDestructionInformation & "." & gk_sDBCOLSubHeaderID & _
                                    " = " & gk_sDBTABLEEditsetSubheaderExtension & "." & _
                                    gk_sDBCOLSubHeaderID & _
                        " AND " & gk_sDBTABLEEditsetSubheaderExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'SQL to update the ACCRUALS field
            Case ECATALOGUEPROPERTIES.cpAccruals
                
                sSql = "update " & gk_sDBTABLEEditsetSubheaderAccruals & " Set " & _
                                    gk_sDBCOLAccruals & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetSubheaderExtension & _
                        " Where " & gk_sDBTABLEEditsetSubheaderAccruals & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetSubheaderExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetSubheaderAccruals & "." & gk_sDBCOLSubHeaderID & _
                                    " = " & gk_sDBTABLEEditsetSubheaderExtension & "." & _
                                    gk_sDBCOLSubHeaderID & _
                        " AND " & gk_sDBTABLEEditsetSubheaderExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
        
            'PC TABLES
            'SQL to update the IMMEDIATE SOURCE OF ACQUISITION description field ONLY
            Case ECATALOGUEPROPERTIES.cpImmediateSourceOfAcquisition
                    
                UpdateLinkDataForImmediateSource ECATALOGUELEVEL.ctSubHeaders
            
            'SQL to update the COPIES INFORMATION description field ONLY
            Case ECATALOGUEPROPERTIES.cpCopiesInformation
                
                UpdateLinkDataForCopiesInfo ECATALOGUELEVEL.ctSubHeaders
            
            'SQL to update the PUBLICATION NOTE description field ONLY
            Case ECATALOGUEPROPERTIES.cpPublicationNote
            
                UpdateLinkDataForPublicationNote ECATALOGUELEVEL.ctSubHeaders
                
            'SQL to update the UNPUBLISHED FINDING AIDS description field ONLY
            Case ECATALOGUEPROPERTIES.cpUnpublishedFindingAids
                
                UpdateLinkDataForUnpublishedFindingAids ECATALOGUELEVEL.ctSubHeaders
        
            'd20 SQL to update the CREATOR NAMES description field ONLY
            Case ECATALOGUEPROPERTIES.cpCreatorNames
                
                UpdateLinkDataForCreatorNames ECATALOGUELEVEL.ctSubHeaders
                
            'd20 SQL to update the PLACE OF DEPOSIT description field ONLY
            Case ECATALOGUEPROPERTIES.cpPlaceOfDeposit
                
                UpdateLinkDataForPlaceOfDeposit ECATALOGUELEVEL.ctSubHeaders
                
            'd20 SQL to update the LOCATION OF ORIGINALS description field ONLY
            Case ECATALOGUEPROPERTIES.cpLocationOfOriginals
                
                UpdateLinkDataForLocationOfOriginals ECATALOGUELEVEL.ctSubHeaders
        
        End Select
        
        If Not (m_eUpdateField = ECATALOGUEPROPERTIES.cpImmediateSourceOfAcquisition Or _
            m_eUpdateField = ECATALOGUEPROPERTIES.cpCopiesInformation Or _
            m_eUpdateField = ECATALOGUEPROPERTIES.cpPublicationNote Or _
            m_eUpdateField = ECATALOGUEPROPERTIES.cpUnpublishedFindingAids Or _
            m_eUpdateField = ECATALOGUEPROPERTIES.cpCreatorNames Or _
            m_eUpdateField = ECATALOGUEPROPERTIES.cpPlaceOfDeposit Or _
            m_eUpdateField = ECATALOGUEPROPERTIES.cpLocationOfOriginals) Then
            
            m_ofwUpdate.Execute sSql
            
            ' d19 - Update the start AND end dates if possible
            If Not (sStartDateSql = vbNullString And sEndDateSql = vbNullString) Then
                m_ofwUpdate.Execute (sStartDateSql)
                m_ofwUpdate.Execute (sEndDateSql)
            End If
            
        End If
        
        'update the editset entry type to modified if not change for all catalogues the level
        sSql = "update " & gk_sDBTABLEEditsetSubheaderExtension & " Set " & _
                gk_sDBCOLEntryTypeID & " = CASE " & gk_sDBCOLEntryTypeID & _
                    " WHEN " & CStr(ECATALOGUEENTRYSTATUS.cesPROCATCreated) & " THEN " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesPROCATCreatedModified) & _
                    " WHEN " & CStr(ECATALOGUEENTRYSTATUS.cesAddedFromLive) & " THEN " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesAddedFromLiveModified) & _
                    " WHEN " & CStr(ECATALOGUEENTRYSTATUS.cesImported) & " THEN " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesImportedModified) & _
                    " ELSE " & gk_sDBCOLEntryTypeID & _
                    " END " & _
                " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID)
    
        m_ofwUpdate.Execute sSql
    
    
        sSql = "Select " & gk_sDBTABLEEditsetSubheader & "." & gk_sDBCOLSubHeaderID & " as " & gk_sDBCOLCatalogueID & _
                ", 5 as " & gk_sDBCOLCatalogueLevel & _
                " from " & gk_sDBTABLEEditsetSubheader & " (nolock) " & _
                ", " & gk_sDBTABLEEditsetSubheaderExtension & "(nolock) " & _
                " Where " & gk_sDBTABLEEditsetSubheader & "." & gk_sDBCOLEditSetID & _
                                " = " & CStr(m_lEditsetID) & _
                " AND " & gk_sDBTABLEEditsetSubheaderExtension & "." & gk_sDBCOLSubHeaderID & _
                                " = " & gk_sDBTABLEEditsetSubheader & "." & gk_sDBCOLSubHeaderID & _
                " AND " & gk_sDBTABLEEditsetSubheaderExtension & "." & gk_sDBCOLEditSetID & _
                                " = " & CStr(m_lEditsetID) & _
                " AND " & gk_sDBTABLEEditsetSubheaderExtension & "." & _
                gk_sDBCOLEditSetEntryTypeID & " <> " & CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
        
        Set BulkUpdateSubheader = m_ofwUpdate.FindBySql(sSql, mk_iMaxNoOfRecords)
    End If
    Exit Function

ErrHandler:
    ConvertError sSource
End Function





Private Function BulkUpdatePiece() As FWRecordList
' Bulkupdates the value of the set property of the Piece
' This procedure generates the update query based on the field to be updated and
'    executes the query.
' d2 Change Quotes to QuoteValueWithQuotes and updated entry type ID field
' d5    ksk     Changes for closure status ID being alphanumeric
' d6    ksk     changed the function to return a recordlist that will be a list of catalogues
'               affected by the change.
' d7    ksk     changed the physical condition field to be a text field insted of lookup
' d9    ksk     Added the extra join that was missing to link the extenstion
'               and main table for the list of records affected
' d10   ksk     Changed the Piece and Item levels language update to use
'               join to extension table
' d11   ksk     Changed the condition in the update for imported matrial. Used the
'               wrong constant
' d12   ksk     Added a max constant to be used by all the Findbysql
'               To return the affected records
' d14   ksk     29Dec1999       Changed the usage of open_date in Piece and Item level
' d19   alm     12Feb2001       Updates the start / end dates automatically
'                               when a covering date is updated
' d20   alm     02Oct2001       Added support for accesioned, orderable, LCI and extract as transfer properties
'                               and Creator, Deposit and Originals
'
On Error GoTo ErrHandler

    Dim sSql          As String
    Dim sStartDateSql As String 'd19
    Dim sEndDateSql   As String 'd19
    Dim bCanUpdate    As Boolean
    Dim sSource       As String
    
    sSource = "BOCatalogueBulkUpdate.BulkUpdatePiece"
    
    'Check if the set property is available for tthe Piece level
    bCanUpdate = m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(ctPiece, m_eUpdateField)
    
    'If the current property is available for Piece then generate the corresponding SQL
    '  to update the field in the corresponding table.
    If bCanUpdate Then
        Select Case m_eUpdateField
            'Main table
            'SQL to update both the TITLE fields
            Case ECATALOGUEPROPERTIES.cpTitle
                
                sSql = "update " & gk_sDBTABLEEditsetPiece & " Set " & _
                                    gk_sDBCOLPieceTitle & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetPieceExtension & _
                        " Where " & gk_sDBTABLEEditsetPiece & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetPieceExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetPiece & "." & gk_sDBCOLPieceID & _
                                    " = " & gk_sDBTABLEEditsetPieceExtension & "." & _
                                    gk_sDBCOLPieceID & _
                        " AND " & gk_sDBTABLEEditsetPieceExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                
                sSql = sSql & "  UPDATE " & gk_sDBTABLEEditsetPieceTitle & " Set " & _
                                    gk_sDBCOLTitle & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetPieceExtension & _
                        " Where " & gk_sDBTABLEEditsetPieceTitle & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetPieceExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetPieceTitle & "." & gk_sDBCOLPieceID & _
                                    " = " & gk_sDBTABLEEditsetPieceExtension & "." & _
                                    gk_sDBCOLPieceID & _
                        " AND " & gk_sDBTABLEEditsetPieceExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'SQL to update the CLOSURE STATUS ID Link
            Case ECATALOGUEPROPERTIES.cpClosureStatus
                sSql = "update " & gk_sDBTABLEEditsetPiece & " Set " & _
                                    gk_sDBCOLClosureStatus & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetPieceExtension & _
                        " Where " & gk_sDBTABLEEditsetPiece & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetPieceExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetPiece & "." & gk_sDBCOLPieceID & _
                                    " = " & gk_sDBTABLEEditsetPieceExtension & "." & _
                                    gk_sDBCOLPieceID & _
                        " AND " & gk_sDBTABLEEditsetPieceExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                
            
            
            'SQL to update the CLOSURE CODE value and CLOSURE TYPE lookup
            Case ECATALOGUEPROPERTIES.cpClosureCode, ECATALOGUEPROPERTIES.cpClosureType
            
                sSql = " update " & gk_sDBTABLEEditsetPiece & " Set " & _
                                    gk_sDBCOLClosureType & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetPieceExtension & _
                        " Where " & gk_sDBTABLEEditsetPiece & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetPieceExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetPiece & "." & gk_sDBCOLPieceID & _
                                    " = " & gk_sDBTABLEEditsetPieceExtension & "." & _
                                    gk_sDBCOLPieceID & _
                        " AND " & gk_sDBTABLEEditsetPieceExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly) & vbNewLine
                
                sSql = sSql & " update " & gk_sDBTABLEEditsetPiece & " Set " & _
                                    gk_sDBCOLClosureCode & " = " & _
                                    NullToSQLNull(StringToNullOrLong(m_vUpdateValue2)) & _
                        " FROM " & gk_sDBTABLEEditsetPieceExtension & _
                        " Where " & gk_sDBTABLEEditsetPiece & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetPieceExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetPiece & "." & gk_sDBCOLPieceID & _
                                    " = " & gk_sDBTABLEEditsetPieceExtension & "." & _
                                    gk_sDBCOLPieceID & _
                        " AND " & gk_sDBTABLEEditsetPieceExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
                
            
            'SQL to update the COVERING DATES field
            Case ECATALOGUEPROPERTIES.cpCoveringDates
            
                sSql = "update " & gk_sDBTABLEEditsetPiece & " Set " & _
                                    gk_sDBCOLDateText & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetPieceExtension & _
                        " Where " & gk_sDBTABLEEditsetPiece & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetPieceExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetPiece & "." & gk_sDBCOLPieceID & _
                                    " = " & gk_sDBTABLEEditsetPieceExtension & "." & _
                                    gk_sDBCOLPieceID & _
                        " AND " & gk_sDBTABLEEditsetPieceExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
                ' d19 We also need to update the start and end dates
                If Not IsNull(m_vUpdateValue2) Then
                    
                    sStartDateSql = "update " & gk_sDBTABLEEditsetPiece & " Set " & _
                                    gk_sDBCOLFirstDate & " = " & _
                                    StringToNullOrString(Left(m_vUpdateValue2, 8)) & _
                        " FROM " & gk_sDBTABLEEditsetPieceExtension & _
                        " Where " & gk_sDBTABLEEditsetPiece & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetPieceExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetPiece & "." & gk_sDBCOLPieceID & _
                                    " = " & gk_sDBTABLEEditsetPieceExtension & "." & _
                                    gk_sDBCOLPieceID & _
                        " AND " & gk_sDBTABLEEditsetPieceExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                                        
                    sEndDateSql = "update " & gk_sDBTABLEEditsetPiece & " Set " & _
                                    gk_sDBCOLLastDate & " = " & _
                                    StringToNullOrString(Right(m_vUpdateValue2, 8)) & _
                        " FROM " & gk_sDBTABLEEditsetPieceExtension & _
                        " Where " & gk_sDBTABLEEditsetPiece & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetPieceExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetPiece & "." & gk_sDBCOLPieceID & _
                                    " = " & gk_sDBTABLEEditsetPieceExtension & "." & _
                                    gk_sDBCOLPieceID & _
                        " AND " & gk_sDBTABLEEditsetPieceExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                                    
                End If
                
            'd14 changed to update the live table instead of the extenstion table
            'SQL to update the RECORD OPENING DATE field
            Case ECATALOGUEPROPERTIES.cpRecordOpeningDate
            
                sSql = "update " & gk_sDBTABLEEditsetPiece & " Set " & _
                                    gk_sDBCOLOpenDate & " = " & _
                                    QuoteValueWithQuotes(VariantDateAsDDMMMYYYY(StringToNullOrDate(m_vUpdateValue1))) & _
                        " FROM " & gk_sDBTABLEEditsetPieceExtension & _
                        " Where " & gk_sDBTABLEEditsetPiece & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetPieceExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetPiece & "." & gk_sDBCOLPieceID & _
                                    " = " & gk_sDBTABLEEditsetPieceExtension & "." & _
                                    gk_sDBCOLPieceID & _
                        " AND " & gk_sDBTABLEEditsetPieceExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            
            'Extension table
            'SQL to update the FORMER REFERENCE (Department) field
            Case ECATALOGUEPROPERTIES.cpFormerDepartmentReference
                
                sSql = "update " & gk_sDBTABLEEditsetPieceExtension & " Set " & _
                                    gk_sDBCOLFormerDepartmentReference & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'SQL to update the FORMER PRO REFERENCE  field
            Case ECATALOGUEPROPERTIES.cpFormerPROReference
            
                sSql = "update " & gk_sDBTABLEEditsetPieceExtension & " Set " & _
                                    gk_sDBCOLFormerPROReference & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                    
                                    
            'SQL to update the PHYSICAL CONDITION ID Link
            Case ECATALOGUEPROPERTIES.cpPhysicalCondition
            
                sSql = "update " & gk_sDBTABLEEditsetPieceExtension & " Set " & _
                                    gk_sDBCOLPhysicalCondition & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                    
            'SQL to update the PHYSICAL DESCRIPTION RECORD QUANTITY field
            Case ECATALOGUEPROPERTIES.cpPhysicalDescription_RecordQuantity
            
                sSql = "update " & gk_sDBTABLEEditsetPieceExtension & " Set " & _
                                    gk_sDBCOLPhysicalDescriptionRecordQuantity & " = " & _
                                    NullToSQLNull(StringToNullOrLong(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                    
            'SQL to update the PHYSICAL DESCRIPTION FORM ID Link
            Case ECATALOGUEPROPERTIES.cpPhysicalDescription_Form
            
                sSql = "update " & gk_sDBTABLEEditsetPieceExtension & " Set " & _
                                    gk_sDBCOLPhysicalDescriptionFormID & " = " & _
                                    NullToSQLNull(StringToNullOrLong(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                    
            'SQL to update the PHYSICAL DESCRIPTION DIMENSION field
            Case ECATALOGUEPROPERTIES.cpPhysicalDescription_Dimension
            
                sSql = "update " & gk_sDBTABLEEditsetPieceExtension & " Set " & _
                                    gk_sDBCOLPhysicalDescriptionDimension & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                    
            'SQL to update the MAP DESIGNATION field
            Case ECATALOGUEPROPERTIES.cpMapDesignation
            
                sSql = "update " & gk_sDBTABLEEditsetPieceExtension & " Set " & _
                                    gk_sDBCOLMapDesignation & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                        
            'SQL to update the MAP SCALE NUMBER field
            Case ECATALOGUEPROPERTIES.cpMapscaleNumber
            
                sSql = "update " & gk_sDBTABLEEditsetPieceExtension & " Set " & _
                                    gk_sDBCOLMapScaleNumber & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                    
            'SQL to update the ACCESS CONDITION ID Link
            Case ECATALOGUEPROPERTIES.cpAccessCondition
            
                sSql = "update " & gk_sDBTABLEEditsetPieceExtension & " Set " & _
                                    gk_sDBCOLAccessConditionID & " = " & _
                                    NullToSQLNull(StringToNullOrLong(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                    
            'SQL to update the RESTRICTIONS ON USE field
            Case ECATALOGUEPROPERTIES.cpRestrictionsOnUse
            
                sSql = "update " & gk_sDBTABLEEditsetPieceExtension & " Set " & _
                                    gk_sDBCOLRestrictionsOnUse & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                    
            'SQL to update the ACCUMULATION DATES field
            Case ECATALOGUEPROPERTIES.cpAccumulationDates
            
                sSql = "update " & gk_sDBTABLEEditsetPieceExtension & " Set " & _
                                    gk_sDBCOLAccumulationdates & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                    
                                    
            'SQL to update the LANGUAGE ID Link
            Case ECATALOGUEPROPERTIES.cpLanguage
            
                sSql = "update " & gk_sDBTABLEEditsetPiece & " Set " & _
                                    gk_sDBCOLLanguageID & " = " & _
                                    NullToSQLNull(StringToNullOrLong(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetPieceExtension & _
                        " Where " & gk_sDBTABLEEditsetPiece & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetPieceExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetPiece & "." & gk_sDBCOLPieceID & _
                                    " = " & gk_sDBTABLEEditsetPieceExtension & "." & _
                                    gk_sDBCOLPieceID & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                    
            'SQL to update the LEGAL STATUS ID Link
            Case ECATALOGUEPROPERTIES.cpLegalStatus
            
                sSql = "update " & gk_sDBTABLEEditsetPieceExtension & " Set " & _
                                    gk_sDBCOLLegalStatusID & " = " & _
                                    NullToSQLNull(StringToNullOrLong(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                    
            
            'Seperate tables, one property per table
            'SQL to update the seperate SCOPE AND CONTENT field
            Case ECATALOGUEPROPERTIES.cpScopeAndContent
                
                sSql = "update " & gk_sDBTABLEEditsetPieceScopeAndContent & " Set " & _
                                    gk_sDBCOLScopeAndContent & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetPieceExtension & _
                        " Where " & gk_sDBTABLEEditsetPieceScopeAndContent & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetPieceExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetPieceScopeAndContent & "." & gk_sDBCOLPieceID & _
                                    " = " & gk_sDBTABLEEditsetPieceExtension & "." & _
                                    gk_sDBCOLPieceID & _
                        " AND " & gk_sDBTABLEEditsetPieceExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                
            
            'SQL to update the seperate ARRANGAMENT field
            Case ECATALOGUEPROPERTIES.cpArrangement
                
                sSql = "update " & gk_sDBTABLEEditsetPieceArrangement & " Set " & _
                                    gk_sDBCOLArrangement & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetPieceExtension & _
                        " Where " & gk_sDBTABLEEditsetPieceArrangement & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetPieceExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetPieceArrangement & "." & gk_sDBCOLPieceID & _
                                    " = " & gk_sDBTABLEEditsetPieceExtension & "." & _
                                    gk_sDBCOLPieceID & _
                        " AND " & gk_sDBTABLEEditsetPieceExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                
            'SQL to update the seperate ADMINISTRATIVE OR BIOGRAPHICAL HISTORY field
            Case ECATALOGUEPROPERTIES.cpAdministrative_BiographicalHistory
                
                sSql = "update " & gk_sDBTABLEEditsetPieceAdministrativeHistory & " Set " & _
                                    gk_sDBCOLAdministrativeHistory & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetPieceExtension & _
                        " Where " & gk_sDBTABLEEditsetPieceAdministrativeHistory & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetPieceExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetPieceAdministrativeHistory & "." & gk_sDBCOLPieceID & _
                                    " = " & gk_sDBTABLEEditsetPieceExtension & "." & _
                                    gk_sDBCOLPieceID & _
                        " AND " & gk_sDBTABLEEditsetPieceExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'SQL to update the CUSTODIAL HISTORY field
            Case ECATALOGUEPROPERTIES.cpCustodialHistory
                
                sSql = "update " & gk_sDBTABLEEditsetPieceCustodialHistory & " Set " & _
                                    gk_sDBCOLCustodialHistory & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetPieceExtension & _
                        " Where " & gk_sDBTABLEEditsetPieceCustodialHistory & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetPieceExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetPieceCustodialHistory & "." & gk_sDBCOLPieceID & _
                                    " = " & gk_sDBTABLEEditsetPieceExtension & "." & _
                                    gk_sDBCOLPieceID & _
                        " AND " & gk_sDBTABLEEditsetPieceExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'SQL to update the APPRAISAL DESTRUCTION INFORMATION field
            Case ECATALOGUEPROPERTIES.cpAppraisal_DestructionInformation
                
                sSql = "update " & gk_sDBTABLEEditsetPieceAppraisalDestructionInformation & _
                        " Set " & gk_sDBCOLAppraisalDestructionInformation & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetPieceExtension & _
                        " Where " & gk_sDBTABLEEditsetPieceAppraisalDestructionInformation & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetPieceExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetPieceAppraisalDestructionInformation & "." & gk_sDBCOLPieceID & _
                                    " = " & gk_sDBTABLEEditsetPieceExtension & "." & _
                                    gk_sDBCOLPieceID & _
                        " AND " & gk_sDBTABLEEditsetPieceExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'SQL to update the ACCRUALS field
            Case ECATALOGUEPROPERTIES.cpAccruals
                
                sSql = "update " & gk_sDBTABLEEditsetPieceAccruals & " Set " & _
                                    gk_sDBCOLAccruals & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetPieceExtension & _
                        " Where " & gk_sDBTABLEEditsetPieceAccruals & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetPieceExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetPieceAccruals & "." & gk_sDBCOLPieceID & _
                                    " = " & gk_sDBTABLEEditsetPieceExtension & "." & _
                                    gk_sDBCOLPieceID & _
                        " AND " & gk_sDBTABLEEditsetPieceExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'd20 Update record status fields (accession)
            Case ECATALOGUEPROPERTIES.cpAccessionedAsCopy
                
                ' Turn the flag on or off
                If m_vUpdateValue1 = True Then
                    m_oBOCatalogueDBHelper.BulkUpdateRecordStatus m_lEditsetID, ctPiece, gk_lFLAGAccessionAsCopy, True
                Else
                    m_oBOCatalogueDBHelper.BulkUpdateRecordStatus m_lEditsetID, ctPiece, gk_lFLAGAccessionAsCopy, False
                End If
            
            'd20 Update record status fields (orderable)
            Case ECATALOGUEPROPERTIES.cpOrderable
                
                ' Turn the flag on or off
                If m_vUpdateValue1 = True Then
                    m_oBOCatalogueDBHelper.BulkUpdateRecordStatus m_lEditsetID, ctPiece, gk_lFLAGOrderable, True
                    m_oBOCatalogueDBHelper.BulkUpdateRecordStatus m_lEditsetID, ctPiece, gk_lFLAGNotOrderable, False
                Else
                    m_oBOCatalogueDBHelper.BulkUpdateRecordStatus m_lEditsetID, ctPiece, gk_lFLAGOrderable, False
                    m_oBOCatalogueDBHelper.BulkUpdateRecordStatus m_lEditsetID, ctPiece, gk_lFLAGNotOrderable, True
                End If
            
            'd20 Update record status fields (LCI)
            Case ECATALOGUEPROPERTIES.cpAR
                
                ' Turn the flag on or off
                If m_vUpdateValue1 = True Then
                    m_oBOCatalogueDBHelper.BulkUpdateRecordStatus m_lEditsetID, ctPiece, gk_lFLAGAR, True
                Else
                    m_oBOCatalogueDBHelper.BulkUpdateRecordStatus m_lEditsetID, ctPiece, gk_lFLAGAR, False
                End If
                
            'd20 Update record status fields (extract at transfer)
            Case ECATALOGUEPROPERTIES.cpExtractAtTransfer
                                    
                ' Turn the flag on or off
                If m_vUpdateValue1 = True Then
                    m_oBOCatalogueDBHelper.BulkUpdateRecordStatus m_lEditsetID, ctPiece, gk_lFLAGExtractAtTransfer, True
                Else
                    m_oBOCatalogueDBHelper.BulkUpdateRecordStatus m_lEditsetID, ctPiece, gk_lFLAGExtractAtTransfer, False
                End If
                     
            'PC TABLES
            'SQL to update the IMMEDIATE SOURCE OF ACQUISITION description field ONLY
            Case ECATALOGUEPROPERTIES.cpImmediateSourceOfAcquisition
                    
                UpdateLinkDataForImmediateSource ECATALOGUELEVEL.ctPiece
            
            'SQL to update the COPIES INFORMATION description field ONLY
            Case ECATALOGUEPROPERTIES.cpCopiesInformation
                
                UpdateLinkDataForCopiesInfo ECATALOGUELEVEL.ctPiece
            
            'SQL to update the PUBLICATION NOTE description field ONLY
            Case ECATALOGUEPROPERTIES.cpPublicationNote
            
                UpdateLinkDataForPublicationNote ECATALOGUELEVEL.ctPiece
                
            'SQL to update the UNPUBLISHED FINDING AIDS description field ONLY
            Case ECATALOGUEPROPERTIES.cpUnpublishedFindingAids
                
                UpdateLinkDataForUnpublishedFindingAids ECATALOGUELEVEL.ctPiece
        
            'd20 SQL to update the CREATOR NAMES description field ONLY
            Case ECATALOGUEPROPERTIES.cpCreatorNames
                
                UpdateLinkDataForCreatorNames ECATALOGUELEVEL.ctPiece
                
            'd20 SQL to update the PLACE OF DEPOSIT description field ONLY
            Case ECATALOGUEPROPERTIES.cpPlaceOfDeposit
                
                UpdateLinkDataForPlaceOfDeposit ECATALOGUELEVEL.ctPiece
                
            'd20 SQL to update the LOCATION OF ORIGINALS description field ONLY
            Case ECATALOGUEPROPERTIES.cpLocationOfOriginals
                
                UpdateLinkDataForLocationOfOriginals ECATALOGUELEVEL.ctPiece
        
        End Select
        
        If Not (m_eUpdateField = ECATALOGUEPROPERTIES.cpImmediateSourceOfAcquisition Or _
            m_eUpdateField = ECATALOGUEPROPERTIES.cpCopiesInformation Or _
            m_eUpdateField = ECATALOGUEPROPERTIES.cpPublicationNote Or _
            m_eUpdateField = ECATALOGUEPROPERTIES.cpUnpublishedFindingAids Or _
            m_eUpdateField = ECATALOGUEPROPERTIES.cpAccessionedAsCopy Or _
            m_eUpdateField = ECATALOGUEPROPERTIES.cpOrderable Or _
            m_eUpdateField = ECATALOGUEPROPERTIES.cpAR Or _
            m_eUpdateField = ECATALOGUEPROPERTIES.cpExtractAtTransfer Or _
            m_eUpdateField = ECATALOGUEPROPERTIES.cpCreatorNames Or _
            m_eUpdateField = ECATALOGUEPROPERTIES.cpPlaceOfDeposit Or _
            m_eUpdateField = ECATALOGUEPROPERTIES.cpLocationOfOriginals) Then
            
            m_ofwUpdate.Execute sSql
            
            ' d19 - Update the start AND end dates if possible
            If Not (sStartDateSql = vbNullString And sEndDateSql = vbNullString) Then
                m_ofwUpdate.Execute (sStartDateSql)
                m_ofwUpdate.Execute (sEndDateSql)
            End If
                            
        End If
        
        'update the editset entry type to modified if not change for all catalogues the level
        sSql = "update " & gk_sDBTABLEEditsetPieceExtension & " Set " & _
                gk_sDBCOLEntryTypeID & " = CASE " & gk_sDBCOLEntryTypeID & _
                    " WHEN " & CStr(ECATALOGUEENTRYSTATUS.cesPROCATCreated) & " THEN " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesPROCATCreatedModified) & _
                    " WHEN " & CStr(ECATALOGUEENTRYSTATUS.cesAddedFromLive) & " THEN " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesAddedFromLiveModified) & _
                    " WHEN " & CStr(ECATALOGUEENTRYSTATUS.cesImported) & " THEN " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesImportedModified) & _
                    " ELSE " & gk_sDBCOLEntryTypeID & _
                    " END " & _
                " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID)
        m_ofwUpdate.Execute sSql
    
        sSql = "Select " & gk_sDBTABLEEditsetPiece & "." & gk_sDBCOLPieceID & " as " & gk_sDBCOLCatalogueID & _
                ", 6 as " & gk_sDBCOLCatalogueLevel & _
                " from " & gk_sDBTABLEEditsetPiece & "(nolock) " & _
                ", " & gk_sDBTABLEEditsetPieceExtension & "(nolock) " & _
                " Where " & gk_sDBTABLEEditsetPiece & "." & gk_sDBCOLEditSetID & _
                                " = " & CStr(m_lEditsetID) & _
                " AND " & gk_sDBTABLEEditsetPieceExtension & "." & gk_sDBCOLPieceID & _
                                " = " & gk_sDBTABLEEditsetPiece & "." & gk_sDBCOLPieceID & _
                " AND " & gk_sDBTABLEEditsetPieceExtension & "." & gk_sDBCOLEditSetID & _
                                " = " & CStr(m_lEditsetID) & _
                " AND " & gk_sDBTABLEEditsetPieceExtension & "." & _
                gk_sDBCOLEditSetEntryTypeID & " <> " & CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
        
        Set BulkUpdatePiece = m_ofwUpdate.FindBySql(sSql, mk_iMaxNoOfRecords)
    
    End If
    Exit Function

ErrHandler:
    ConvertError sSource
End Function





Private Function BulkUpdateItem() As FWRecordList
' Bulkupdates the value of the set property of the Item
' This procedure generates the update query based on the field to be updated and
'    executes the query.
' d2 Change Quotes to QuoteValueWithQuotes and updated entry type ID field
' d5    ksk     Changes for closure status ID being alphanumeric
' d6    ksk     changed the function to return a recordlist that will be a list of catalogues
'               affected by the change.
' d7    ksk     changed the physical condition field to be a text field insted of lookup
' d9    ksk     Added the extra join that was missing to link the extenstion
'               and main table for the list of records affected
' d10   ksk     Changed the Piece and Item levels language update to use
'               join to extension table
' d11   ksk     Changed the condition in the update for imported matrial. Used the
'               wrong constant
' d12   ksk     Added a max constant to be used by all the Findbysql
'               To return the affected records
' d14   ksk     29Dec1999       Changed the usage of open_date in Piece and Item level
' d19   alm     12Feb2001       Updates the start / end dates automatically
'                               when a covering date is updated
' d20   alm     02Oct2001       Added support for accesioned, orderable, LCI and extract as transfer properties
'                               and Creator, Deposit and Originals
'
On Error GoTo ErrHandler

    Dim sSql          As String
    Dim sStartDateSql As String 'd19
    Dim sEndDateSql   As String 'd19
    Dim bCanUpdate    As Boolean
    Dim sSource       As String
    
    sSource = "BOCatalogueBulkUpdate.BulkUpdateItem"
    
    'Check if the set property is available for tthe Item level
    bCanUpdate = m_oBOCatalogueElementMatrix.LevelHasCatalogueProperty(ctItem, m_eUpdateField)
    
    'If the current property is available for Item then generate the corresponding SQL
    '  to update the field in the corresponding table.
    If bCanUpdate Then
        Select Case m_eUpdateField
            'Main table
            'SQL to update both the TITLE fields
            Case ECATALOGUEPROPERTIES.cpTitle
            
                sSql = "update " & gk_sDBTABLEEditsetItem & " Set " & _
                                    gk_sDBCOLItemTitle & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetItemExtension & _
                        " Where " & gk_sDBTABLEEditsetItem & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetItemExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetItem & "." & gk_sDBCOLItemID & _
                                    " = " & gk_sDBTABLEEditsetItemExtension & "." & _
                                    gk_sDBCOLItemID & _
                        " AND " & gk_sDBTABLEEditsetItemExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                
                sSql = sSql & "  UPDATE " & gk_sDBTABLEEditsetItemTitle & " Set " & _
                                    gk_sDBCOLTitle & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetItemExtension & _
                        " Where " & gk_sDBTABLEEditsetItemTitle & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetItemExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetItemTitle & "." & gk_sDBCOLItemID & _
                                    " = " & gk_sDBTABLEEditsetItemExtension & "." & _
                                    gk_sDBCOLItemID & _
                        " AND " & gk_sDBTABLEEditsetItemExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'SQL to update the CLOSURE CODE value and CLOSURE TYPE lookup
            Case ECATALOGUEPROPERTIES.cpClosureCode, ECATALOGUEPROPERTIES.cpClosureType
            
                sSql = " update " & gk_sDBTABLEEditsetItem & " Set " & _
                                    gk_sDBCOLClosureType & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetItemExtension & _
                        " Where " & gk_sDBTABLEEditsetItem & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetItemExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetItem & "." & gk_sDBCOLItemID & _
                                    " = " & gk_sDBTABLEEditsetItemExtension & "." & _
                                    gk_sDBCOLItemID & _
                        " AND " & gk_sDBTABLEEditsetItemExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                
                sSql = sSql & " update " & gk_sDBTABLEEditsetItem & " Set " & _
                                    gk_sDBCOLClosureCode & " = " & _
                                    NullToSQLNull(StringToNullOrLong(m_vUpdateValue2)) & _
                        " FROM " & gk_sDBTABLEEditsetItemExtension & _
                        " Where " & gk_sDBTABLEEditsetItem & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetItemExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetItem & "." & gk_sDBCOLItemID & _
                                    " = " & gk_sDBTABLEEditsetItemExtension & "." & _
                                    gk_sDBCOLItemID & _
                        " AND " & gk_sDBTABLEEditsetItemExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly) & vbNewLine
            
            'SQL to update the CLOSURE TYPE lookup
            Case ECATALOGUEPROPERTIES.cpClosureStatus
            
                sSql = "update " & gk_sDBTABLEEditsetItem & " Set " & _
                                    gk_sDBCOLClosureStatus & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetItemExtension & _
                        " Where " & gk_sDBTABLEEditsetItem & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetItemExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetItem & "." & gk_sDBCOLItemID & _
                                    " = " & gk_sDBTABLEEditsetItemExtension & "." & _
                                    gk_sDBCOLItemID & _
                        " AND " & gk_sDBTABLEEditsetItemExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'SQL to update the COVERING DATES field
            Case ECATALOGUEPROPERTIES.cpCoveringDates
            
                sSql = "update " & gk_sDBTABLEEditsetItem & " Set " & _
                                    gk_sDBCOLDateText & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetItemExtension & _
                        " Where " & gk_sDBTABLEEditsetItem & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetItemExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetItem & "." & gk_sDBCOLItemID & _
                                    " = " & gk_sDBTABLEEditsetItemExtension & "." & _
                                    gk_sDBCOLItemID & _
                        " AND " & gk_sDBTABLEEditsetItemExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                    
                ' d19 We also need to update the start and end dates
                If Not IsNull(m_vUpdateValue2) Then
                    
                    sStartDateSql = "update " & gk_sDBTABLEEditsetItem & " Set " & _
                                    gk_sDBCOLFirstDate & " = " & _
                                    StringToNullOrString(Left(m_vUpdateValue2, 8)) & _
                        " FROM " & gk_sDBTABLEEditsetItemExtension & _
                        " Where " & gk_sDBTABLEEditsetItem & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetItemExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetItem & "." & gk_sDBCOLItemID & _
                                    " = " & gk_sDBTABLEEditsetItemExtension & "." & _
                                    gk_sDBCOLItemID & _
                        " AND " & gk_sDBTABLEEditsetItemExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                                        
                    sEndDateSql = "update " & gk_sDBTABLEEditsetItem & " Set " & _
                                    gk_sDBCOLLastDate & " = " & _
                                    StringToNullOrString(Right(m_vUpdateValue2, 8)) & _
                        " FROM " & gk_sDBTABLEEditsetItemExtension & _
                        " Where " & gk_sDBTABLEEditsetItem & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetItemExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetItem & "." & gk_sDBCOLItemID & _
                                    " = " & gk_sDBTABLEEditsetItemExtension & "." & _
                                    gk_sDBCOLItemID & _
                        " AND " & gk_sDBTABLEEditsetItemExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                                    
                End If
            
                            
            'd14 changed to update the live table instead of the extenstion table
            'SQL to update the RECORD OPENING DATE field
            Case ECATALOGUEPROPERTIES.cpRecordOpeningDate
            
                sSql = "update " & gk_sDBTABLEEditsetItem & " Set " & _
                                    gk_sDBCOLOpenDate & " = " & _
                                    QuoteValueWithQuotes(VariantDateAsDDMMMYYYY(StringToNullOrDate(m_vUpdateValue1))) & _
                        " FROM " & gk_sDBTABLEEditsetItemExtension & _
                        " Where " & gk_sDBTABLEEditsetItem & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetItemExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetItem & "." & gk_sDBCOLItemID & _
                                    " = " & gk_sDBTABLEEditsetItemExtension & "." & _
                                    gk_sDBCOLItemID & _
                        " AND " & gk_sDBTABLEEditsetItemExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'Extension table
            'SQL to update the FORMER REFERENCE (Department) field
            Case ECATALOGUEPROPERTIES.cpFormerDepartmentReference
                
                sSql = "update " & gk_sDBTABLEEditsetItemExtension & " Set " & _
                                    gk_sDBCOLFormerDepartmentReference & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'SQL to update the FORMER PRO REFERENCE  field
            Case ECATALOGUEPROPERTIES.cpFormerPROReference
            
                sSql = "update " & gk_sDBTABLEEditsetItemExtension & " Set " & _
                                    gk_sDBCOLFormerPROReference & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                    
                                    
            'SQL to update the PHYSICAL CONDITION ID Link
            Case ECATALOGUEPROPERTIES.cpPhysicalCondition
            
                sSql = "update " & gk_sDBTABLEEditsetItemExtension & " Set " & _
                                    gk_sDBCOLPhysicalCondition & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                    
            'SQL to update the PHYSICAL DESCRIPTION RECORD QUANTITY field
            Case ECATALOGUEPROPERTIES.cpPhysicalDescription_RecordQuantity
            
                sSql = "update " & gk_sDBTABLEEditsetItemExtension & " Set " & _
                                    gk_sDBCOLPhysicalDescriptionRecordQuantity & " = " & _
                                    NullToSQLNull(StringToNullOrLong(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                    
            'SQL to update the PHYSICAL DESCRIPTION FORM ID Link
            Case ECATALOGUEPROPERTIES.cpPhysicalDescription_Form
            
                sSql = "update " & gk_sDBTABLEEditsetItemExtension & " Set " & _
                                    gk_sDBCOLPhysicalDescriptionFormID & " = " & _
                                    NullToSQLNull(StringToNullOrLong(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                    
            'SQL to update the PHYSICAL DESCRIPTION DIMENSION field
            Case ECATALOGUEPROPERTIES.cpPhysicalDescription_Dimension
            
                sSql = "update " & gk_sDBTABLEEditsetItemExtension & " Set " & _
                                    gk_sDBCOLPhysicalDescriptionDimension & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                    
            'SQL to update the MAP DESIGNATION field
            Case ECATALOGUEPROPERTIES.cpMapDesignation
            
                sSql = "update " & gk_sDBTABLEEditsetItemExtension & " Set " & _
                                    gk_sDBCOLMapDesignation & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                    
            'SQL to update the MAP SCALE NUMBER field
            Case ECATALOGUEPROPERTIES.cpMapscaleNumber
            
                sSql = "update " & gk_sDBTABLEEditsetItemExtension & " Set " & _
                                    gk_sDBCOLMapScaleNumber & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                    
            'SQL to update the ACCESS CONDITION ID Link
            Case ECATALOGUEPROPERTIES.cpAccessCondition
            
                sSql = "update " & gk_sDBTABLEEditsetItemExtension & " Set " & _
                                    gk_sDBCOLAccessConditionID & " = " & _
                                    NullToSQLNull(StringToNullOrLong(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                    
            'SQL to update the RESTRICTIONS ON USE field
            Case ECATALOGUEPROPERTIES.cpRestrictionsOnUse
            
                sSql = "update " & gk_sDBTABLEEditsetItemExtension & " Set " & _
                                    gk_sDBCOLRestrictionsOnUse & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                    
            'SQL to update the ACCUMULATION DATES field
            Case ECATALOGUEPROPERTIES.cpAccumulationDates
            
                sSql = "update " & gk_sDBTABLEEditsetItemExtension & " Set " & _
                                    gk_sDBCOLAccumulationdates & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                    
                                    
            'SQL to update the LANGUAGE ID Link
            Case ECATALOGUEPROPERTIES.cpLanguage
            
                sSql = "update " & gk_sDBTABLEEditsetItem & " Set " & _
                                    gk_sDBCOLLanguageID & " = " & _
                                    NullToSQLNull(StringToNullOrLong(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetItemExtension & _
                        " Where " & gk_sDBTABLEEditsetItem & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetItemExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetItem & "." & gk_sDBCOLItemID & _
                                    " = " & gk_sDBTABLEEditsetItemExtension & "." & _
                                    gk_sDBCOLItemID & _
                        " AND " & gk_sDBTABLEEditsetItemExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
                                    
            'SQL to update the LEGAL STATUS ID Link
            Case ECATALOGUEPROPERTIES.cpLegalStatus
            
                sSql = "update " & gk_sDBTABLEEditsetItemExtension & " Set " & _
                                    gk_sDBCOLLegalStatusID & " = " & _
                                    NullToSQLNull(StringToNullOrLong(m_vUpdateValue1)) & _
                        " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'Seperate tables, one property per table
            'SQL to update the seperate SCOPE AND CONTENT field
            Case ECATALOGUEPROPERTIES.cpScopeAndContent
                
                sSql = "update " & gk_sDBTABLEEditsetItemScopeAndContent & " Set " & _
                                    gk_sDBCOLScopeAndContent & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetItemExtension & _
                        " Where " & gk_sDBTABLEEditsetItemScopeAndContent & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetItemExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetItemScopeAndContent & "." & gk_sDBCOLItemID & _
                                    " = " & gk_sDBTABLEEditsetItemExtension & "." & _
                                    gk_sDBCOLItemID & _
                        " AND " & gk_sDBTABLEEditsetItemExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'SQL to update the seperate ARRANGAMENT field
            Case ECATALOGUEPROPERTIES.cpArrangement
                
                sSql = "update " & gk_sDBTABLEEditsetItemArrangement & " Set " & _
                                    gk_sDBCOLArrangement & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetItemExtension & _
                        " Where " & gk_sDBTABLEEditsetItemArrangement & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetItemExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetItemArrangement & "." & gk_sDBCOLItemID & _
                                    " = " & gk_sDBTABLEEditsetItemExtension & "." & _
                                    gk_sDBCOLItemID & _
                        " AND " & gk_sDBTABLEEditsetItemExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'SQL to update the seperate ADMINISTRATIVE OR BIOGRAPHICAL HISTORY field
            Case ECATALOGUEPROPERTIES.cpAdministrative_BiographicalHistory
                
                sSql = "update " & gk_sDBTABLEEditsetItemAdministrativeHistory & " Set " & _
                                    gk_sDBCOLAdministrativeHistory & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetItemExtension & _
                        " Where " & gk_sDBTABLEEditsetItemAdministrativeHistory & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetItemExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetItemAdministrativeHistory & "." & gk_sDBCOLItemID & _
                                    " = " & gk_sDBTABLEEditsetItemExtension & "." & _
                                    gk_sDBCOLItemID & _
                        " AND " & gk_sDBTABLEEditsetItemExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'SQL to update the CUSTODIAL HISTORY field
            Case ECATALOGUEPROPERTIES.cpCustodialHistory
                
                sSql = "update " & gk_sDBTABLEEditsetItemCustodialHistory & " Set " & _
                                    gk_sDBCOLCustodialHistory & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetItemExtension & _
                        " Where " & gk_sDBTABLEEditsetItemCustodialHistory & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetItemExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetItemCustodialHistory & "." & gk_sDBCOLItemID & _
                                    " = " & gk_sDBTABLEEditsetItemExtension & "." & _
                                    gk_sDBCOLItemID & _
                        " AND " & gk_sDBTABLEEditsetItemExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'SQL to update the APPRAISAL DESTRUCTION INFORMATION field
            Case ECATALOGUEPROPERTIES.cpAppraisal_DestructionInformation
                
                sSql = "update " & gk_sDBTABLEEditsetItemAppraisalDestructionInformation & _
                        " Set " & gk_sDBCOLAppraisalDestructionInformation & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetItemExtension & _
                        " Where " & gk_sDBTABLEEditsetItemAppraisalDestructionInformation & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetItemExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetItemAppraisalDestructionInformation & "." & gk_sDBCOLItemID & _
                                    " = " & gk_sDBTABLEEditsetItemExtension & "." & _
                                    gk_sDBCOLItemID & _
                        " AND " & gk_sDBTABLEEditsetItemExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'SQL to update the ACCRUALS field
            Case ECATALOGUEPROPERTIES.cpAccruals
                
                sSql = "update " & gk_sDBTABLEEditsetItemAccruals & " Set " & _
                                    gk_sDBCOLAccruals & " = " & _
                                    QuoteValueWithQuotes(StringToNullOrString(m_vUpdateValue1)) & _
                        " FROM " & gk_sDBTABLEEditsetItemExtension & _
                        " Where " & gk_sDBTABLEEditsetItemAccruals & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetItemExtension & "." & gk_sDBCOLEditSetID & _
                                    " = " & CStr(m_lEditsetID) & _
                        " AND " & gk_sDBTABLEEditsetItemAccruals & "." & gk_sDBCOLItemID & _
                                    " = " & gk_sDBTABLEEditsetItemExtension & "." & _
                                    gk_sDBCOLItemID & _
                        " AND " & gk_sDBTABLEEditsetItemExtension & "." & _
                                    gk_sDBCOLEditSetEntryTypeID & " <> " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
            
            'd20 Update record status fields (accession)
            Case ECATALOGUEPROPERTIES.cpAccessionedAsCopy
                
                ' Turn the flag on or off
                If m_vUpdateValue1 = True Then
                    m_oBOCatalogueDBHelper.BulkUpdateRecordStatus m_lEditsetID, ctItem, gk_lFLAGAccessionAsCopy, True
                Else
                    m_oBOCatalogueDBHelper.BulkUpdateRecordStatus m_lEditsetID, ctItem, gk_lFLAGAccessionAsCopy, False
                End If
            
            'd20 Update record status fields (orderable)
            Case ECATALOGUEPROPERTIES.cpOrderable
                
                ' Turn the flag on or off
                If m_vUpdateValue1 = True Then
                    m_oBOCatalogueDBHelper.BulkUpdateRecordStatus m_lEditsetID, ctItem, gk_lFLAGOrderable, True
                    m_oBOCatalogueDBHelper.BulkUpdateRecordStatus m_lEditsetID, ctItem, gk_lFLAGNotOrderable, False
                Else
                    m_oBOCatalogueDBHelper.BulkUpdateRecordStatus m_lEditsetID, ctItem, gk_lFLAGOrderable, False
                    m_oBOCatalogueDBHelper.BulkUpdateRecordStatus m_lEditsetID, ctItem, gk_lFLAGNotOrderable, True
                End If
            
            'd20 Update record status fields (LCI)
            Case ECATALOGUEPROPERTIES.cpAR
                
                ' Turn the flag on or off
                If m_vUpdateValue1 = True Then
                    m_oBOCatalogueDBHelper.BulkUpdateRecordStatus m_lEditsetID, ctItem, gk_lFLAGAR, True
                Else
                    m_oBOCatalogueDBHelper.BulkUpdateRecordStatus m_lEditsetID, ctItem, gk_lFLAGAR, False
                End If
                
            'd20 Update record status fields (extract at transfer)
            Case ECATALOGUEPROPERTIES.cpExtractAtTransfer
                                    
                ' Turn the flag on or off
                If m_vUpdateValue1 = True Then
                    m_oBOCatalogueDBHelper.BulkUpdateRecordStatus m_lEditsetID, ctItem, gk_lFLAGExtractAtTransfer, True
                Else
                    m_oBOCatalogueDBHelper.BulkUpdateRecordStatus m_lEditsetID, ctItem, gk_lFLAGExtractAtTransfer, False
                End If
                
            'PC TABLES
            'SQL to update the IMMEDIATE SOURCE OF ACQUISITION description field ONLY
            Case ECATALOGUEPROPERTIES.cpImmediateSourceOfAcquisition
                    
                UpdateLinkDataForImmediateSource ECATALOGUELEVEL.ctItem
            
            'SQL to update the COPIES INFORMATION description field ONLY
            Case ECATALOGUEPROPERTIES.cpCopiesInformation
                
                UpdateLinkDataForCopiesInfo ECATALOGUELEVEL.ctItem
            
            'SQL to update the PUBLICATION NOTE description field ONLY
            Case ECATALOGUEPROPERTIES.cpPublicationNote
            
                UpdateLinkDataForPublicationNote ECATALOGUELEVEL.ctItem
                
            'SQL to update the UNPUBLISHED FINDING AIDS description field ONLY
            Case ECATALOGUEPROPERTIES.cpUnpublishedFindingAids
                
                UpdateLinkDataForUnpublishedFindingAids ECATALOGUELEVEL.ctItem
        
            'd20 SQL to update the CREATOR NAMES description field ONLY
            Case ECATALOGUEPROPERTIES.cpCreatorNames
                
                UpdateLinkDataForCreatorNames ECATALOGUELEVEL.ctItem
                
            'd20 SQL to update the PLACE OF DEPOSIT description field ONLY
            Case ECATALOGUEPROPERTIES.cpPlaceOfDeposit
                
                UpdateLinkDataForPlaceOfDeposit ECATALOGUELEVEL.ctItem
                
            'd20 SQL to update the LOCATION OF ORIGINALS description field ONLY
            Case ECATALOGUEPROPERTIES.cpLocationOfOriginals
                
                UpdateLinkDataForLocationOfOriginals ECATALOGUELEVEL.ctItem
        
        End Select
        
        If Not (m_eUpdateField = ECATALOGUEPROPERTIES.cpImmediateSourceOfAcquisition Or _
            m_eUpdateField = ECATALOGUEPROPERTIES.cpCopiesInformation Or _
            m_eUpdateField = ECATALOGUEPROPERTIES.cpPublicationNote Or _
            m_eUpdateField = ECATALOGUEPROPERTIES.cpUnpublishedFindingAids Or _
            m_eUpdateField = ECATALOGUEPROPERTIES.cpAccessionedAsCopy Or _
            m_eUpdateField = ECATALOGUEPROPERTIES.cpOrderable Or _
            m_eUpdateField = ECATALOGUEPROPERTIES.cpAR Or _
            m_eUpdateField = ECATALOGUEPROPERTIES.cpExtractAtTransfer Or _
            m_eUpdateField = ECATALOGUEPROPERTIES.cpCreatorNames Or _
            m_eUpdateField = ECATALOGUEPROPERTIES.cpPlaceOfDeposit Or _
            m_eUpdateField = ECATALOGUEPROPERTIES.cpLocationOfOriginals) Then
            
            m_ofwUpdate.Execute sSql
            
            ' d19 - Update the start AND end dates if possible
            If Not (sStartDateSql = vbNullString And sEndDateSql = vbNullString) Then
                m_ofwUpdate.Execute (sStartDateSql)
                m_ofwUpdate.Execute (sEndDateSql)
            End If
                            
        End If
        
        'update the editset entry type to modified if not change for all catalogues the level
        sSql = "update " & gk_sDBTABLEEditsetItemExtension & " Set " & _
                gk_sDBCOLEntryTypeID & " = CASE " & gk_sDBCOLEntryTypeID & _
                    " WHEN " & CStr(ECATALOGUEENTRYSTATUS.cesPROCATCreated) & " THEN " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesPROCATCreatedModified) & _
                    " WHEN " & CStr(ECATALOGUEENTRYSTATUS.cesAddedFromLive) & " THEN " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesAddedFromLiveModified) & _
                    " WHEN " & CStr(ECATALOGUEENTRYSTATUS.cesImported) & " THEN " & _
                                    CStr(ECATALOGUEENTRYSTATUS.cesImportedModified) & _
                    " ELSE " & gk_sDBCOLEntryTypeID & _
                    " END " & _
                " Where " & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID)
    
        m_ofwUpdate.Execute sSql
    
        sSql = "Select " & gk_sDBTABLEEditsetItem & "." & gk_sDBCOLItemID & " as " & gk_sDBCOLCatalogueID & _
                ", 7 as " & gk_sDBCOLCatalogueLevel & _
                " from " & gk_sDBTABLEEditsetItem & "(nolock) " & _
                ", " & gk_sDBTABLEEditsetItemExtension & "(nolock) " & _
                " Where " & gk_sDBTABLEEditsetItem & "." & gk_sDBCOLEditSetID & _
                                " = " & CStr(m_lEditsetID) & _
                " AND " & gk_sDBTABLEEditsetItemExtension & "." & gk_sDBCOLItemID & _
                                " = " & gk_sDBTABLEEditsetItem & "." & gk_sDBCOLItemID & _
                " AND " & gk_sDBTABLEEditsetItemExtension & "." & gk_sDBCOLEditSetID & _
                                " = " & CStr(m_lEditsetID) & _
                " AND " & gk_sDBTABLEEditsetItemExtension & "." & _
                gk_sDBCOLEditSetEntryTypeID & " <> " & CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
        
        Set BulkUpdateItem = m_ofwUpdate.FindBySql(sSql, mk_iMaxNoOfRecords)
    End If
    Exit Function

ErrHandler:
    ConvertError sSource
End Function

Private Function UpdateLinkDataForCreatorNames(t_eLevel As ECATALOGUELEVEL)
' d20   alm     02Oct2001       Created (based on UpdateLinkDataForImmediateSource)
    Dim lPropertyID As Long
    
    lPropertyID = ECATALOGUEPROPERTIES.cpCreatorNames
    
    'Remove and add the description in corporate body table
    ReplaceAForSFALinks t_eLevel, lPropertyID, EAFTermType.AFCorporateBody, m_colUpdateValue3, False
    
    'Remove and add the description in person table
    ReplaceAForSFALinks t_eLevel, lPropertyID, EAFTermType.AFPerson, m_colUpdateValue4, False
    
End Function

Private Function UpdateLinkDataForPlaceOfDeposit(t_eLevel As ECATALOGUELEVEL)
' d20   alm     02Oct2001       Created (based on UpdateLinkDataForCopiesInfo)
    Dim lPropertyID As Long
    
    lPropertyID = ECATALOGUEPROPERTIES.cpPlaceOfDeposit
    
    'Remove and add the description in corporate body table
    ReplaceAForSFALinks t_eLevel, lPropertyID, EAFTermType.AFCorporateBody, m_colUpdateValue3, False

End Function

Private Function UpdateLinkDataForLocationOfOriginals(t_eLevel As ECATALOGUELEVEL)
' d20   alm     02Oct2001       Created (based on UpdateLinkDataForCopiesInfo)
    Dim lPropertyID As Long
    
    lPropertyID = ECATALOGUEPROPERTIES.cpLocationOfOriginals
    
    'Remove and add the description in corporate body table
    ReplaceAForSFALinks t_eLevel, lPropertyID, EAFTermType.AFCorporateBody, m_colUpdateValue3, False

End Function

Private Function UpdateLinkDataForImmediateSource(t_eLevel As ECATALOGUELEVEL)

    Dim lPropertyID As Long
    
    lPropertyID = ECATALOGUEPROPERTIES.cpImmediateSourceOfAcquisition
    
    'Removed unknown description and add then new description
    ReplaceUnknownDescription t_eLevel, lPropertyID, m_vUpdateValue1
    
    'Remove and add the description in corporate body table
    ReplaceAForSFALinks t_eLevel, lPropertyID, EAFTermType.AFCorporateBody, m_colUpdateValue3, False
    
    'Remove and add the description in person table
    ReplaceAForSFALinks t_eLevel, lPropertyID, EAFTermType.AFPerson, m_colUpdateValue4, False
    
End Function


Private Function UpdateLinkDataForCopiesInfo(t_eLevel As ECATALOGUELEVEL)
    
    Dim lPropertyID As Long
    
    lPropertyID = ECATALOGUEPROPERTIES.cpCopiesInformation
    
    'Removed unknown description and add then new description
    ReplaceUnknownDescription t_eLevel, lPropertyID, m_vUpdateValue1
    
    'Remove and add the description in corporate body table
    ReplaceAForSFALinks t_eLevel, lPropertyID, EAFTermType.AFCorporateBody, m_colUpdateValue3, False

End Function

Private Function UpdateLinkDataForPublicationNote(t_eLevel As ECATALOGUELEVEL)
' d15   tmj     15Mar2000   Added ReplaceLeaflets call

    Dim lPropertyID As Long
    
    lPropertyID = ECATALOGUEPROPERTIES.cpPublicationNote
    
    'Removed unknown description and add then new description
    ReplaceUnknownDescription t_eLevel, lPropertyID, m_vUpdateValue1
    
    'Remove and add the description in corporate body table
    ReplaceAForSFALinks t_eLevel, lPropertyID, EAFTermType.AFCorporateBody, m_colUpdateValue3, True

    'Remove and add the description in leaflet table
    ReplaceLeaflets t_eLevel, lPropertyID, m_colUpdateValue4

End Function



Private Function UpdateLinkDataForUnpublishedFindingAids(t_eLevel As ECATALOGUELEVEL)
' d15   tmj     15Mar2000   Added ReplaceLeaflets call
        
    Dim lPropertyID As Long
    
    lPropertyID = ECATALOGUEPROPERTIES.cpUnpublishedFindingAids
    
    'Removed unknown description and add then new description
    ReplaceUnknownDescription t_eLevel, lPropertyID, m_vUpdateValue1
    
    'Remove and add the description in corporate body table
    ReplaceAForSFALinks t_eLevel, lPropertyID, EAFTermType.AFCorporateBody, m_colUpdateValue3, True

    'Remove and add the description in leaflet table
    ReplaceLeaflets t_eLevel, lPropertyID, m_colUpdateValue4

End Function

Private Function ReplaceLeaflets(t_eLevel As ECATALOGUELEVEL, _
                        t_lPropertyID As Long, t_colLeaflets As Collection)
' This removes the optional leaflet and comment and adds the new leaflet back in
' for the property and level and adds the new description
'
'Version History
'
' d15   TMJ     15Mar2000   Created

    Dim sSource As String
    Dim vColItem As Variant
    Dim sSql As String
    Dim sColumnName As String
    Dim sTablename As String
    Dim sExtensionTableName As String

    On Error GoTo ErrorHandler
    
    sSource = "BOCatalogueBulkUpdate.ReplaceLeaflets"

    sColumnName = GetColumnName(ECATALOGUECOLUMNTYPE.cctIDColumn, t_eLevel)
    sTablename = GetTableName(ECATALOGUETABLETYPE.cttMainTable, ESectionType.SECeditorial, t_eLevel)
    sExtensionTableName = GetTableName(ECATALOGUETABLETYPE.cttExtensionTable, ESectionType.SECeditorial, t_eLevel)

    sSql = "DELETE FROM " & gk_sDBTABLEEditsetPCLeaflet & _
           " FROM " & sExtensionTableName & _
           " WHERE " & gk_sDBTABLEEditsetPCLeaflet & "." & gk_sDBCOLEditSetID & _
           "=" & CStr(m_lEditsetID) & _
           " AND " & gk_sDBTABLEEditsetPCLeaflet & "." & gk_sDBCOLCatalogueLevel & _
           "=" & CStr(t_eLevel) & _
           " AND " & gk_sDBTABLEEditsetPCLeaflet & "." & gk_sDBCOLPropertyID & _
           "=" & CStr(t_lPropertyID)

    sSql = sSql & " AND " & sExtensionTableName & "." & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID)
    sSql = sSql & " AND " & sExtensionTableName & "." & sColumnName & " = " & gk_sDBTABLEEditsetPCLeaflet & "." & gk_sDBCOLCatalogueID
    sSql = sSql & " AND " & sExtensionTableName & "." & gk_sDBCOLEditSetEntryTypeID & " <> " & CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)

    m_ofwUpdate.Execute sSql

    For Each vColItem In t_colLeaflets
        sSql = " INSERT INTO " & gk_sDBTABLEEditsetPCLeaflet & _
                         " ( " & gk_sDBCOLLeafletReferenceID & " ,  " & _
                                 gk_sDBCOLEditSetID & " ,  " & _
                                 gk_sDBCOLCatalogueID & " ,  " & _
                                 gk_sDBCOLCatalogueLevel & " ,  " & _
                                 gk_sDBCOLPropertyID & " ,  " & _
                                 gk_sDBCOLDescription & ")"
                    
        sSql = sSql & " SELECT " & CStr(vColItem.lID) & ", " _
                                 & CStr(m_lEditsetID) & ", " _
                                 & sTablename & "." & sColumnName & ", " _
                                 & CStr(t_eLevel) & ", " _
                                 & CStr(t_lPropertyID) & ", " _
                                 & QuoteValueWithQuotes _
                                    (StringToNullOrString( _
                                      vColItem.sOptionalDescription))
                                      
        sSql = sSql & " FROM " & sTablename & " (nolock), " & _
                                 sExtensionTableName & " (nolock) "
                                 
        sSql = sSql & " WHERE " & sTablename & "." & gk_sDBCOLEditSetID & "=" _
                                & CStr(m_lEditsetID)
                                
        sSql = sSql & " AND " & sExtensionTableName & "." & sColumnName & " = " & sTablename & "." & sColumnName
        sSql = sSql & " AND " & sExtensionTableName & "." & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID)
        sSql = sSql & " AND " & sExtensionTableName & "." & gk_sDBCOLEditSetEntryTypeID & " <> " & CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
        
        m_ofwUpdate.Execute sSql
        
    Next

    Exit Function

ErrorHandler:
    ConvertError sSource
End Function

Private Function ReplaceUnknownDescription(t_eLevel As ECATALOGUELEVEL, _
                        t_lPropertyID As Long, t_vUnknownDescription As Variant)
'This removes the optional unknown description and adds the new description back in
' for the property and level and adds the new description

On Error GoTo ErrHandler
    
    Dim sColumnName As String
    Dim sTablename As String
    Dim sExtensionTableName As String
    Dim iLevelVal As Integer
    Dim sSource As String
    Dim sSql As String
    
    
    sSource = Name & ".ReplaceUnknownDescription"
    
    iLevelVal = GetCatalogueLevelCode(t_eLevel)
    sColumnName = GetColumnName(ECATALOGUECOLUMNTYPE.cctIDColumn, t_eLevel)
    sTablename = GetTableName(ECATALOGUETABLETYPE.cttMainTable, ESectionType.SECeditorial, t_eLevel)
    sExtensionTableName = GetTableName(ECATALOGUETABLETYPE.cttExtensionTable, ESectionType.SECeditorial, t_eLevel)
    
    'Remove the current values of unknown description
    sSql = "DELETE FROM " & gk_sDBTABLEEditsetPCLinkDataElement
    sSql = sSql & " FROM " & sExtensionTableName
    sSql = sSql & " WHERE " & gk_sDBTABLEEditsetPCLinkDataElement & "." & gk_sDBCOLCatalogueLevel & " = " & CStr(iLevelVal)
    sSql = sSql & " AND " & gk_sDBTABLEEditsetPCLinkDataElement & "." & gk_sDBCOLPropertyID & " = " & CStr(t_lPropertyID)
    sSql = sSql & " AND " & gk_sDBTABLEEditsetPCLinkDataElement & "." & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID)
    
    sSql = sSql & " AND " & sExtensionTableName & "." & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID)
    sSql = sSql & " AND " & sExtensionTableName & "." & sColumnName & " = " & gk_sDBTABLEEditsetPCLinkDataElement & "." & gk_sDBCOLCatalogueID
    sSql = sSql & " AND " & sExtensionTableName & "." & gk_sDBCOLEditSetEntryTypeID & " <> " & CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
    
    m_ofwUpdate.Execute sSql
    
    'Add the records only if the description value is set
    If t_vUnknownDescription <> "" Then
        'The main table is used to get the list of all catalogues to insert into
        sSql = " INSERT INTO " & gk_sDBTABLEEditsetPCLinkDataElement & _
                    " ( " & gk_sDBCOLEditSetID & " ,  " & _
                    gk_sDBCOLPropertyID & " ,  " & _
                    gk_sDBCOLCatalogueLevel & " ,  " & _
                    gk_sDBCOLCatalogueID & " ,  " & _
                    gk_sDBCOLDescription & ")"
                    
        sSql = sSql & " SELECT " & (m_lEditsetID) & ", " & _
                        CStr(t_lPropertyID) & ", " & _
                        iLevelVal & ", " & _
                        sTablename & "." & sColumnName & ", " & _
                        QuoteValueWithQuotes(t_vUnknownDescription) & _
                    " FROM " & sTablename & " (nolock), " & sExtensionTableName & " (nolock) " & _
                    " WHERE " & sTablename & "." & gk_sDBCOLEditSetID & "=" & CStr(m_lEditsetID)
                    
        sSql = sSql & " AND " & sExtensionTableName & "." & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID)
        sSql = sSql & " AND " & sExtensionTableName & "." & sColumnName & " = " & sTablename & "." & sColumnName
        sSql = sSql & " AND " & sExtensionTableName & "." & gk_sDBCOLEditSetEntryTypeID & " <> " & CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
        
        m_ofwUpdate.Execute sSql
        
    End If
    Exit Function

ErrHandler:
    ConvertError sSource
End Function


Private Function ReplaceAForSFALinks(t_eLevel As ECATALOGUELEVEL, _
                                            t_lPropertyID As Long, _
                                            t_eAFType As EAFTermType, _
                                            t_colLinkDataItem As Collection, _
                                            t_bSFAUpdation As Boolean)
' This method removes the current existing links from the editsets link term type table
' and adds in the new set of links and description for all the catalogues in a given level
'   t_eLevel - The level which has to be replaced by the new set of link data items
'   t_lPropertyID - The link data elements porperty ID (In this case copies info,
'                       publication note, unpublished finding aids and immediate source)
'   t_eAFType - The term type that is being removed and added. Can be ES_PC_Person or ES_PC_Corp_Body
'   t_colLinkDataItem - The collection of the new (link, description data) pairs
'   t_bSFAUpdation - if true the updation id for ES_PC_SFA else the two AF tables
'
On Error GoTo ErrHandler
    Dim sColumnName As String
    Dim sTablename As String
    Dim sExtensionTableName As String
    Dim sLinkDataItemTableToUpdate As String
    Dim sLinkDataItemColumnToUpdate As String
    Dim iLevelVal As Integer
    Dim sSource As String
    Dim sSql As String
    Dim typLinkDataItem As Variant
    
    sSource = Name & ".ReplaceUnknownDescription"
    
    If t_bSFAUpdation Then
        sLinkDataItemTableToUpdate = gk_sDBTABLEEditsetPCSFA
        sLinkDataItemColumnToUpdate = gk_sDBCOLSFAReferenceID
    Else
        If t_eAFType = EAFTermType.AFAllTypes Then
            RaiseError sSource, EPROCATINTERNALERROR.peBulkUpdate, "Parameter t_eAFType value has to be within 1 - 4"
        End If
        
        sLinkDataItemTableToUpdate = GetEditsetLinkDataItemTableName(t_eAFType)
        sLinkDataItemColumnToUpdate = GetLinkDataItemColumnName(t_eAFType)
    End If
    
    iLevelVal = GetCatalogueLevelCode(t_eLevel)
    sColumnName = GetColumnName(ECATALOGUECOLUMNTYPE.cctIDColumn, t_eLevel)
    sTablename = GetTableName(ECATALOGUETABLETYPE.cttMainTable, ESectionType.SECeditorial, t_eLevel)
    sExtensionTableName = GetTableName(ECATALOGUETABLETYPE.cttExtensionTable, ESectionType.SECeditorial, t_eLevel)
    
    'Remove the current values of unknown description
    sSql = "DELETE FROM " & sLinkDataItemTableToUpdate
    sSql = sSql & " FROM " & sExtensionTableName
    sSql = sSql & " WHERE " & sLinkDataItemTableToUpdate & "." & gk_sDBCOLCatalogueLevel & " = " & CStr(iLevelVal)
    sSql = sSql & " AND " & sLinkDataItemTableToUpdate & "." & gk_sDBCOLPropertyID & " = " & CStr(t_lPropertyID)
    sSql = sSql & " AND " & sLinkDataItemTableToUpdate & "." & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID)
    
    sSql = sSql & " AND " & sExtensionTableName & "." & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID)
    sSql = sSql & " AND " & sExtensionTableName & "." & sColumnName & " = " & sLinkDataItemTableToUpdate & "." & gk_sDBCOLCatalogueID
    sSql = sSql & " AND " & sExtensionTableName & "." & gk_sDBCOLEditSetEntryTypeID & " <> " & CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
    
    m_ofwUpdate.Execute sSql
    
    'Add the records only if the description value is set
    For Each typLinkDataItem In t_colLinkDataItem

        'The main table is used to get the list of all catalogues to insert into
        sSql = " INSERT INTO " & sLinkDataItemTableToUpdate & _
                    " ( " & gk_sDBCOLEditSetID & " ,  " & _
                    gk_sDBCOLPropertyID & " ,  " & _
                    gk_sDBCOLCatalogueLevel & " ,  " & _
                    gk_sDBCOLCatalogueID & " ,  " & _
                    sLinkDataItemColumnToUpdate & " ,  " & _
                    gk_sDBCOLDescription & ")"
                    
        sSql = sSql & " SELECT " & CStr(m_lEditsetID) & ", " & _
                        CStr(t_lPropertyID) & ", " & _
                        iLevelVal & ", " & _
                        sTablename & "." & sColumnName & ", " & _
                        CStr(typLinkDataItem.lID) & ", " & _
                         QuoteValueWithQuotes(StringToNullOrString(typLinkDataItem.sOptionalDescription)) & _
                    " FROM " & sTablename & " (nolock), " & sExtensionTableName & " (nolock) " & _
                    " WHERE " & sTablename & "." & gk_sDBCOLEditSetID & "=" & CStr(m_lEditsetID)
        
        sSql = sSql & " AND " & sExtensionTableName & "." & sColumnName & " = " & sTablename & "." & sColumnName
        sSql = sSql & " AND " & sExtensionTableName & "." & gk_sDBCOLEditSetID & " = " & CStr(m_lEditsetID)
        sSql = sSql & " AND " & sExtensionTableName & "." & gk_sDBCOLEditSetEntryTypeID & " <> " & CStr(ECATALOGUEENTRYSTATUS.cesReadeOnly)
        
        m_ofwUpdate.Execute sSql
    Next
    
    Exit Function

ErrHandler:
    ConvertError sSource
End Function




Private Function CheckIfPropertiesSet() As Boolean
' d4    ksk     04Aug1999       Changed usage of catalogue element matrix from local to
'                               private variable
' d5    ksk     13Aug1999       Changes for closure status ID being alphanumeric
' d13   ksk     09Dec1999       Fixed a bug in raising error for length violation in Checkifpropertyset
' d20   alm     02Oct2001       Added support for accesioned, orderable, LCI and extract as transfer properties
'
    Dim sSource As String
    Dim sErrorDescription As String
    Dim lLength As Integer
    Dim eDataType As EValidCheck
    
    sSource = "BOCatalogueBulkUpdate.CheckIfPropertySet"
    
    
    'If closure type then the length check is for closure code field
    If m_eUpdateField = ECATALOGUEPROPERTIES.cpClosureCode Or _
            m_eUpdateField = ECATALOGUEPROPERTIES.cpClosureType Then
        eDataType = m_oBOCatalogueElementMatrix.CataloguePropertyDataType(ECATALOGUEPROPERTIES.cpClosureCode)
        lLength = m_oBOCatalogueElementMatrix.CataloguePropertyLength(ECATALOGUEPROPERTIES.cpClosureCode)
    Else
        eDataType = m_oBOCatalogueElementMatrix.CataloguePropertyDataType(m_eUpdateField)
        lLength = m_oBOCatalogueElementMatrix.CataloguePropertyLength(m_eUpdateField)
    End If
    
    'Make sure that the properties are set
    If m_eUpdateField < 1 Then
        RaiseError sSource, peBulkUpdate, "Invalid Catalogue Property (field) to update"
    End If
    
    If Len(m_vUpdateValue1) > lLength Then
        RaiseError sSource, EPROCATINTERNALERROR.peIncorrectLength, _
                            m_oBOCatalogueElementMatrix.DescriptionFromEnumeration(m_eUpdateField) _
                            & " property has more than the allowed text length of " & _
                            CStr(lLength)
    End If
    
    'Validate the field type values
    Select Case m_eUpdateField
    Case ECATALOGUEPROPERTIES.cpMapscaleNumber
        If Trim(m_vUpdateValue1) = "" Or Not IsNumeric(m_vUpdateValue1) Then
            RaiseError sSource, EPROCATINTERNALERROR.peIncorrectDataType, _
                            m_oBOCatalogueElementMatrix.DescriptionFromEnumeration(m_eUpdateField) _
                            & " property requires a long number"
        End If
    Case ECATALOGUEPROPERTIES.cpRecordOpeningDate
        If Trim(m_vUpdateValue1) = "" Or _
        Not IsDate(ConvertDateString(StringToNullOrDate(m_vUpdateValue1))) Then
            RaiseError sSource, EPROCATINTERNALERROR.peIncorrectDataType, _
                            m_oBOCatalogueElementMatrix.DescriptionFromEnumeration(m_eUpdateField) _
                            & " property  requires a date (dd/mm/yyyy) format"
        End If
    Case ECATALOGUEPROPERTIES.cpClosureStatus
        If Trim(m_vUpdateValue1) = "" Then
            RaiseError sSource, EPROCATINTERNALERROR.peIncorrectDataType, _
                            m_oBOCatalogueElementMatrix.DescriptionFromEnumeration(m_eUpdateField) _
                            & " property needs a the Lookup ID"
        End If
    Case ECATALOGUEPROPERTIES.cpClosureCode, ECATALOGUEPROPERTIES.cpClosureType
        
        If Not IsNumeric(m_vUpdateValue2) And m_vUpdateValue2 <> "" Then
            RaiseError sSource, EPROCATINTERNALERROR.peIncorrectDataType, _
                            m_oBOCatalogueElementMatrix.DescriptionFromEnumeration(m_eUpdateField) _
                            & " The bulkupdate closure code value has to be numeric"
        End If
        If Not IsValidClosureTypeAndCode(m_vUpdateValue1, m_vUpdateValue2, sErrorDescription) Then
            RaiseError sSource, EPROCATINTERNALERROR.peInvalidCombinationOfClosureTypeAndCode, _
                            m_oBOCatalogueElementMatrix.DescriptionFromEnumeration(m_eUpdateField) _
                            & sErrorDescription
        End If
        
    Case Else
        'Check if ID value is empty 'd20
        If (m_eUpdateField = cpAccessCondition Or m_eUpdateField = cpLanguage _
            Or m_eUpdateField = cpLegalStatus Or m_eUpdateField = cpPhysicalDescription_Form _
            Or m_eUpdateField = cpPhysicalDescription_Form _
            Or m_eUpdateField = cpAccessionedAsCopy Or m_eUpdateField = cpOrderable _
            Or m_eUpdateField = cpAR Or m_eUpdateField = cpExtractAtTransfer) And _
            Trim(m_vUpdateValue1) = "" Then
            RaiseError sSource, EPROCATINTERNALERROR.peIncorrectDataType, _
                            m_oBOCatalogueElementMatrix.DescriptionFromEnumeration(m_eUpdateField) _
                            & " property needs a the Lookup ID"
        End If
        'Make sure that the link fields have the ID passes as numeric value
        If m_eUpdateField = cpAccessCondition Or m_eUpdateField = cpLanguage _
            Or m_eUpdateField = cpLegalStatus _
            Or m_eUpdateField = cpPhysicalDescription_Form Or _
            m_eUpdateField = cpPhysicalDescription_Form Then
            
            If CLng(Trim(m_vUpdateValue1)) = 0 Then
                RaiseError sSource, EPROCATINTERNALERROR.peIncorrectDataType, _
                        m_oBOCatalogueElementMatrix.DescriptionFromEnumeration(m_eUpdateField) _
                            & " property has to be numeric"
            End If
        
        End If
        'd20 Make sure checkbox fields are true of false
        If (m_eUpdateField = cpAccessionedAsCopy Or m_eUpdateField = cpOrderable _
            Or m_eUpdateField = cpAR Or m_eUpdateField = cpExtractAtTransfer) Then
            
            If Not (m_vUpdateValue1 = True Or m_vUpdateValue1 = False) Then
                RaiseError sSource, EPROCATINTERNALERROR.peIncorrectDataType, _
                        m_oBOCatalogueElementMatrix.DescriptionFromEnumeration(m_eUpdateField) _
                            & " property has to be true or false"
            End If
            
        End If
    End Select
        
    CheckIfPropertiesSet = True
End Function


Private Sub Initialise()
    Set m_ofwUpdate = New FWDBHelper
    Set m_oBOCatalogueElementMatrix = New BOCatalogueElementMatrix
    Set m_oBOCatalogueDBHelper = New BOCatalogueDBHelper
    Set m_colUpdateValue3 = Nothing
    Set m_colUpdateValue4 = Nothing
    m_lLevels = 0
End Sub

Private Sub Terminate()
    Set m_ofwUpdate = Nothing
    Set m_oBOCatalogueElementMatrix = Nothing
    Set m_oBOCatalogueDBHelper = Nothing
    Set m_colUpdateValue3 = Nothing
    Set m_colUpdateValue4 = Nothing
End Sub




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'Properties
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Public Property Let EditsetID(t_lEditSetID As Long)
    m_lEditsetID = t_lEditSetID
End Property

Public Property Get EditsetID() As Long
    EditsetID = m_lEditsetID
End Property

Public Property Let BulkUpdateField(t_eBulkUpdateField As ECATALOGUEPROPERTIES)
    m_eUpdateField = t_eBulkUpdateField
End Property

Public Property Get BulkUpdateField() As ECATALOGUEPROPERTIES
    BulkUpdateField = m_eUpdateField
End Property




'Property used by All bulk update fields. The following are the options of usage
' 1) Regular fields for value
' 2) Closure Type + Closure Code field for storing Closure Type value
' 3) Link data item fields for storing the unknown description (ie description without a link)
'
Public Property Let BulkUpdateValue1(t_vBulkUpdateValue As Variant)
    m_vUpdateValue1 = t_vBulkUpdateValue
End Property

Public Property Get BulkUpdateValue1() As Variant
    BulkUpdateValue1 = m_vUpdateValue1
End Property

'Property used to store only closure code field value when bulkupdating closure type + closure code option.
Public Property Let BulkUpdateValue2(t_vBulkUpdateValue As Variant)
    m_vUpdateValue2 = t_vBulkUpdateValue
End Property

Public Property Get BulkUpdateValue2() As Variant
    BulkUpdateValue2 = m_vUpdateValue2
End Property

'Property to store the ID, Value (Description) pairs for links
' 1) Used by All Link data items (Immediate source, copies info, publication notes, unpublished finding)
'   Immediate source stores only the links to corporate body
Public Property Set BulkUpdateValue3(t_colBulkUpdateValue As Collection)
    Set m_colUpdateValue3 = t_colBulkUpdateValue
End Property

Public Property Get BulkUpdateValue3() As Collection
    Set BulkUpdateValue3 = m_colUpdateValue3
End Property

'Additional Property to store the ID, Value (Description) pairs for links
' 1) Used by only Immediate source stores only the links to person links
Public Property Set BulkUpdateValue4(t_colBulkUpdateValue As Collection)
    Set m_colUpdateValue4 = t_colBulkUpdateValue
End Property

Public Property Get BulkUpdateValue4() As Collection
    Set BulkUpdateValue4 = m_colUpdateValue4
End Property



Public Property Get Name() As String
    Name = "BOCatalogueBulkUpdate"
End Property


Private Sub Class_Initialize()
    Call Initialise
End Sub

Private Sub Class_Terminate()
    Call Terminate
End Sub




Public Sub PROCATTransientInterface_DeleteState(Optional t_sInstanceReference As String)
'
'Description: Removes a stored state
'
' d8    ksk     25Oct1999   created
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String
    
    Dim sInstanceReference As String
    Dim ofwTransientManager As FWTransientManager
    
    sSource = Name & ".PROCATTransientInterface_DeleteState"
    
    'Initialise object
    Set ofwTransientManager = New FWTransientManager
    
    'Create the instance reference
    sInstanceReference = CreateDelimitedString(Name, t_sInstanceReference)

    ofwTransientManager.DeleteData GetObjectContext.Item(gk_sSessionObject).SessionID, sInstanceReference

    LogMessage "Delete " & Name & " state" & sInstanceReference, EProfile.pComplete

    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub

Public Sub PROCATTransientInterface_GetState(Optional t_sInstanceReference As String)
'
'Description: Gets the the current state of the bulkupdate values selected
'
' d8    ksk     25Oct1999   created
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String
    
    Dim sInstanceReference As String
    Dim sSerialisedString As String
    Dim ofwTransientManager As FWTransientManager
    
    sSource = Name & ".PROCATTransientInterface_GetState"
    
    'Initialise object
    Set ofwTransientManager = New FWTransientManager
    
    'Create the instance reference
    sInstanceReference = CreateDelimitedString(Name, t_sInstanceReference)
    
    'Read from the transient store
    sSerialisedString = _
        ofwTransientManager.GetData(GetObjectContext.Item(gk_sSessionObject).SessionID, sInstanceReference)

    If Trim(sSerialisedString) = "" Then
        LogMessage Name & " state not found " & sInstanceReference, EProfile.pProfileInformation
    Else
        
        m_lEditsetID = CLng(GetDataFromSerialisedString(sSerialisedString))
        m_eUpdateField = CInt(GetDataFromSerialisedString(sSerialisedString))
        'String values
        m_vUpdateValue1 = GetDataFromSerialisedString(sSerialisedString)
        m_vUpdateValue2 = GetDataFromSerialisedString(sSerialisedString)
        m_lLevels = GetDataFromSerialisedString(sSerialisedString)
        
        'Collections
        Set m_colUpdateValue3 = m_oBOCatalogueDBHelper.ExtractLinkItemsFromSerialIsedString(sSerialisedString)
        Set m_colUpdateValue4 = m_oBOCatalogueDBHelper.ExtractLinkItemsFromSerialIsedString(sSerialisedString)

        LogMessage "Read " & Me.Name & " state" & sInstanceReference, EProfile.pComplete
    End If
    
    Set ofwTransientManager = Nothing
    
    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub


Public Sub PROCATTransientInterface_SetState(Optional t_sInstanceReference As String)
'
'Description: Saves the the current state of the bulkupdate values selected
'
' d8    ksk     25Oct1999   created
'
    On Error GoTo ErrorHandler
    
    Dim sSource As String
    
    Dim sInstanceReference As String
    Dim sSerialisedString As String
    Dim ofwTransientManager As FWTransientManager
    
    sSource = Name & ".PROCATTransientInterface_SetState"
    
    'Initialise object
    Set ofwTransientManager = New FWTransientManager
    
    'Create the instance reference
    sInstanceReference = CreateDelimitedString(Name, t_sInstanceReference)

    'Simple Strings
    sSerialisedString = ""
    
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, m_lEditsetID)
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, m_eUpdateField)
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, VariantStringAsString(m_vUpdateValue1))
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, VariantStringAsString(m_vUpdateValue2))
    sSerialisedString = AddDataToSerialisedString(sSerialisedString, m_lLevels)
    
    'Collections
    sSerialisedString = m_oBOCatalogueDBHelper.SerialiseLinkItem(sSerialisedString, m_colUpdateValue3)
    sSerialisedString = m_oBOCatalogueDBHelper.SerialiseLinkItem(sSerialisedString, m_colUpdateValue4)

    'Store in the transient manager
    ofwTransientManager.AddData _
        GetObjectContext.Item(gk_sSessionObject).SessionID, sInstanceReference, sSerialisedString

    LogMessage "Stored " & Me.Name & sInstanceReference, EProfile.pComplete

    Set ofwTransientManager = Nothing
    
    Exit Sub
    
ErrorHandler:
    ConvertError sSource
End Sub


Private Function GetEditsetLinkDataItemTableName(t_eLinkDataItemType As EAFTermType) As String
'This function returns the edit set table name of the AD term type for the value passed in t_eLinkDataItemType
'
' d8    ksk     25Oct1999   created

    Select Case t_eLinkDataItemType
    Case EAFTermType.AFPerson
        GetEditsetLinkDataItemTableName = gk_sDBTABLEEditSetPCPerson
    Case EAFTermType.AFPlace
        GetEditsetLinkDataItemTableName = gk_sDBTABLEEditsetPCPlace
    Case EAFTermType.AFSubject
        GetEditsetLinkDataItemTableName = gk_sDBTABLEEditsetPCSubject
    Case EAFTermType.AFCorporateBody
        GetEditsetLinkDataItemTableName = gk_sDBTABLEEditSetPCCorporateBody
    End Select
    
End Function

Public Function GetLinkDataItemColumnName(t_eLinkDataItemType As EAFTermType) As String
'This function returns the column name of the index term ID filed in the AF term type table
'
' d8    ksk     25Oct1999   created

    Select Case t_eLinkDataItemType
    Case EAFTermType.AFPerson
        GetLinkDataItemColumnName = gk_sDBCOLPersonReferenceID
    Case EAFTermType.AFPlace
        GetLinkDataItemColumnName = gk_sDBCOLPlaceReferenceID
    Case EAFTermType.AFSubject
        GetLinkDataItemColumnName = gk_sDBCOLSubjectReferenceID
    Case EAFTermType.AFCorporateBody
        GetLinkDataItemColumnName = gk_sDBCOLCorporateBodyReferenceID
    End Select

End Function

Public Sub SetLevels(t_bLevels() As Boolean)
' Sets the m_lLevels bit map which indicates which levels to update
'
' d15   WDP     24Feb2000       Created
    On Error GoTo ErrorHandler
    Dim sSource As String
    sSource = "BOCatalogueBulkUpdate.SetLevels"
    
    'Reset the member value
    m_lLevels = 0
    
    Dim i As Integer
    For i = 1 To 7
        If t_bLevels(i) Then
            m_lLevels = m_lLevels + (2 ^ i)
        End If
    Next
    Exit Sub
ErrorHandler:
    ConvertError sSource
End Sub

Public Sub GetLevels(ByRef t_bLevels() As Boolean)
' Sets the m_lLevels bit map which indicates which levels to update
'
' d15   WDP     24Feb2000       Created
    On Error GoTo ErrorHandler
    Dim sSource As String
    sSource = "BOCatalogueBulkUpdate.GetLevels"
    
    Dim i As Integer
    For i = 1 To 7
        If m_lLevels And (2 ^ i) Then
            t_bLevels(i) = True
        Else
            t_bLevels(i) = False
        End If
    Next
    Exit Sub
ErrorHandler:
    ConvertError sSource

End Sub
