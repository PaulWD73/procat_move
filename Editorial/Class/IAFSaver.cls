VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IAFSaver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    IAFSaver.cls
' System:       PROCAT, PRO
' Copyright:    (C) Quidnunc Limited
'
' Description:  Interface to be implemented by authority file Saver
'               classes (BOSubjectSaver etc.)
'               Implemented in the same lines as Loader class send by WDP.
'
' Amendment history:
'   d1   Geedha Sivanadian    6-Aug-1999
'   d2   AGA                      14-Oct-1999    Modified SaveLinks, modifed save to return the ID
'                                                          Modified implementation of Save to return Long
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit


'Property to save the term data from vwclass
Public Property Let Data(t_vData As Variant)


End Property
Public Property Let SeeLinksData(t_vData As Collection)
'd2     AGA     14-Oct-1999     Created

End Property
Public Property Let SeeAlsoLinksData(t_vData As Collection)
'd2     AGA     14-Oct-1999     Created

End Property
'method to Save data into the object
Public Function Save(t_lID As Long) As Long
'd2     AGA     14-Oct-1999     Modified to return long

End Function

'method to get cross references for the term from the link table
Public Sub SaveLinks(t_lID As Long)

End Sub

