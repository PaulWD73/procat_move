Attribute VB_Name = "StyleConstants"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    StyleConstants.bas
' System:       PROCAT, PRO
' Copyright:    (C) Quidnunc Limited
'
' Description:  Constants for style of display
'
' Amendment history:
'   d1      ndk     17Aug1999       Created
'   d2      ndk     31Aug1999       Added new style for helptext, added code for bold and italic properties
'   d3      ndk     31Aug1999       Pageheading style changed
'   d4      dtm     10feb2000       Colours changed to ones that will display on 256-colour setup
'   d5      wdp     15Mar2000       Changed colours again! used web safe ones
'                                   see http://www.visibone.com/colorlab/
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Colour Constants

Public Const gk_sCOLOURWhite As String = "#FFFFFF"
Public Const gk_sCOLOURBlack As String = "#000000"
Public Const gk_sCOLOURRed As String = "#CC0000"
Public Const gk_sCOLOURPALEGREEN As String = "#CCFFCC"
Public Const gk_sCOLOURPink As String = "#FFCCFF"
Public Const gk_sCOLOURGRAY As String = "#CCCCCC"
Public Const gk_sCOLOURDarkPink As String = "#FF99FF"


'Font Face constants

Public Const gk_sFONTNone As String = "" 'When default font is to be considered
Public Const gk_sFONTArial As String = "Arial"
Public Const gk_sFONTTimesNewRoman As String = "Times New Roman"

'Custom Styles

Public Const gk_sFONTNormal As String = gk_sFONTNone
Public Const gk_iStyleFontSizeNormal As Integer = 3
Public Const gk_sStyleForeColourNormal As String = gk_sCOLOURBlack
Public Const gk_sStyleCellColourNormal As String = gk_sCOLOURWhite
Public Const gk_bStyleIsBoldNormal As Boolean = False
Public Const gk_bStyleIsItalicNormal As Boolean = False

Public Const gk_sFONTColumnHeading As String = gk_sFONTNone
Public Const gk_iStyleFontSizeColumnHeading As Integer = 3
Public Const gk_sStyleForeColourColumnHeading As String = gk_sCOLOURWhite
Public Const gk_sStyleCellColourColumnHeading As String = gk_sCOLOURRed
Public Const gk_bStyleIsBoldColumnHeading As Boolean = True
Public Const gk_bStyleIsItalicColumnHeading As Boolean = False

Public Const gk_sFONTColumnData As String = gk_sFONTNone
Public Const gk_iStyleFontSizeColumnData As Integer = 3
Public Const gk_sStyleForeColourColumnData As String = gk_sCOLOURBlack
Public Const gk_sStyleCellColourColumnData As String = gk_sCOLOURPink
Public Const gk_bStyleIsBoldColumnData As Boolean = False
Public Const gk_bStyleIsItalicColumnData As Boolean = False

Public Const gk_sFONTReadOnlyLabel As String = gk_sFONTArial
Public Const gk_iStyleFontSizeReadOnlyLabel As Integer = 3
Public Const gk_sStyleForeColourReadOnlyLabel As String = gk_sCOLOURBlack
Public Const gk_sStyleCellColourReadOnlyLabel As String = gk_sCOLOURWhite
Public Const gk_bStyleIsBoldReadOnlyLabel As Boolean = False
Public Const gk_bStyleIsItalicReadOnlyLabel As Boolean = False

Public Const gk_sFONTReadWriteLabel As String = gk_sFONTNone
Public Const gk_iStyleFontSizeReadWriteLabel As Integer = 3
Public Const gk_sStyleForeColourReadWriteLabel As String = gk_sCOLOURRed
Public Const gk_sStyleCellColourReadWriteLabel As String = gk_sCOLOURWhite
Public Const gk_bStyleIsBoldReadWriteLabel As Boolean = False
Public Const gk_bStyleIsItalicReadWriteLabel As Boolean = False

Public Const gk_sFONTPageHeading As String = gk_sFONTNone
Public Const gk_iStyleFontSizePageHeading As Integer = 3
Public Const gk_sStyleForeColourPageHeading As String = gk_sCOLOURBlack
Public Const gk_sStyleCellColourPageHeading As String = gk_sCOLOURWhite
Public Const gk_bStyleIsBoldPageHeading As Boolean = True
Public Const gk_bStyleIsItalicPageHeading As Boolean = False

Public Const gk_sFONTReadOnlyData As String = gk_sFONTNone
Public Const gk_iStyleFontSizeReadOnlyData As Integer = 3
Public Const gk_sStyleForeColourReadOnlyData As String = gk_sCOLOURBlack
Public Const gk_sStyleCellColourReadOnlyData As String = gk_sCOLOURPink
Public Const gk_bStyleIsBoldReadOnlyData As Boolean = False
Public Const gk_bStyleIsItalicReadOnlyData As Boolean = False

Public Const gk_sFONTReadWriteData As String = gk_sFONTNone
Public Const gk_iStyleFontSizeReadWriteData As Integer = 3
Public Const gk_sStyleForeColourReadWriteData As String = gk_sCOLOURBlack
Public Const gk_sStyleCellColourReadWriteData As String = gk_sCOLOURPink
Public Const gk_bStyleIsBoldReadWriteData As Boolean = False
Public Const gk_bStyleIsItalicReadWriteData As Boolean = False

Public Const gk_sFONTGeneralLabel As String = gk_sFONTNone
Public Const gk_iStyleFontSizeGeneralLabel As Integer = 3
Public Const gk_sStyleForeColourGeneralLabel As String = gk_sCOLOURBlack
Public Const gk_sStyleCellColourGeneralLabel As String = gk_sCOLOURWhite
Public Const gk_bStyleIsBoldGeneralLabel As Boolean = False
Public Const gk_bStyleIsItalicGeneralLabel As Boolean = False

Public Const gk_sFONTGeneralContent As String = gk_sFONTNone
Public Const gk_iStyleFontSizeGeneralContent As Integer = 3
Public Const gk_sStyleForeColourGeneralContent As String = gk_sCOLOURBlack
Public Const gk_sStyleCellColourGeneralContent As String = gk_sCOLOURPink
Public Const gk_bStyleIsBoldGeneralContent As Boolean = False
Public Const gk_bStyleIsItalicGeneralContent As Boolean = False

Public Const gk_sFONTReadyToApprove As String = gk_sFONTNone
Public Const gk_iStyleFontSizeReadyToApprove As Integer = 3
Public Const gk_sStyleForeColourReadyToApprove As String = gk_sCOLOURBlack
Public Const gk_sStyleCellColourReadyToApprove As String = gk_sCOLOURPALEGREEN
Public Const gk_bStyleIsBoldReadyToApprove As Boolean = False
Public Const gk_bStyleIsItalicReadyToApprove As Boolean = False

Public Const gk_sFONTWorkInProgress As String = gk_sFONTNone
Public Const gk_iStyleFontSizeWorkInProgress As Integer = 3
Public Const gk_sStyleForeColourWorkInProgress As String = gk_sCOLOURBlack
Public Const gk_sStyleCellColourWorkInProgress As String = gk_sCOLOURPink
Public Const gk_bStyleIsBoldWorkInProgress As Boolean = False
Public Const gk_bStyleIsItalicWorkInProgress As Boolean = False

Public Const gk_sFONTNotAtStage  As String = gk_sFONTNone
Public Const gk_iStyleFontSizeNotAtStage  As Integer = 3
Public Const gk_sStyleForeColourNotAtStage As String = gk_sCOLOURBlack
Public Const gk_sStyleCellColourNotAtStage As String = gk_sCOLOURGRAY
Public Const gk_bStyleIsBoldNotAtStage As Boolean = False
Public Const gk_bStyleIsItalicNotAtStage As Boolean = False

Public Const gk_sFONTHelpText As String = gk_sFONTNone
Public Const gk_iStyleFontSizeHelpText  As Integer = 3
Public Const gk_sStyleForeColourHelpText As String = gk_sCOLOURBlack
Public Const gk_sStyleCellColourHelpText As String = gk_sCOLOURWhite
Public Const gk_bStyleIsBoldHelpText As Boolean = False
Public Const gk_bStyleIsItalicHelpText As Boolean = True

Public Const gk_sStyleBackgroundHighlightRow As String = gk_sCOLOURDarkPink

