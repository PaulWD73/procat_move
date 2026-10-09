VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BOReleaseController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    BOReleaseController.cls
' System:       PROCAT, PRO
' Copyright:    (C) Quidnunc Limited
'
' Description:  A business object that encapsulates all the release logics.
'       The changes needed when
'       1 Editset is released/Abandoned
'       2 AF is released
'       3 Information Leaflet is released
'       This class will use methods in individual business object to move the object
'       to released state state. Since releasing an object will affect other objects,
'       because of inter-dependencies, this controlling class is used whenever the
'       object needs to release itself.
'
' Amendment history:
'   d1   ksk      02Aug1999      Created
'   d2   ksk      02Aug1999      Coded
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'The type of the object that needs to be released/abandoned
Public Enum ERELEASEOBJECT
    roEditSet
    roAuthorityFilePerson
    roAuthorityFilePlace
    roAuthorityFileSubject
    roAuthorityFileCorporateBody
    roInformationLeaflet
End Enum

Private m_eReleaseObject As ERELEASEOBJECT
Private m_lReleaseElementID As Long







'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'Properties
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Let ReleaseObject(t_eReleaseObject As ERELEASEOBJECT)
    m_eReleaseObject = t_eReleaseObject
End Property

Public Property Get ReleaseObject() As ERELEASEOBJECT
    ReleaseObject = m_eReleaseObject
End Property

Public Property Let ObjectID(t_lID As Long)
    m_lReleaseElementID = t_lID
End Property

Public Property Get ObjectID() As Long
    ObjectID = m_lReleaseElementID
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'Methods
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub Release()
'Release method is used to release the object (ID) to the live state depending on the conditions
    
    Select Case m_eReleaseObject
    Case ERELEASEOBJECT.roAuthorityFileCorporateBody
    
    Case ERELEASEOBJECT.roAuthorityFilePerson
    
    Case ERELEASEOBJECT.roAuthorityFilePlace
    
    Case ERELEASEOBJECT.roAuthorityFileSubject

    Case ERELEASEOBJECT.roEditSet
        'Call Insert related material links into the other catalogue
        'Call Insert Audit trial
        'Call Verity Indexing
        
    Case ERELEASEOBJECT.roInformationLeaflet
    
    End Select

End Sub

Public Sub Abandon()
'Abandoning method is currently only allowed for edit sets
Dim sSource As String
    sSource = "BOReleaseController.Abandon"
    
    If m_eReleaseObject <> ERELEASEOBJECT.roEditSet Then
        RaiseError sSource, EPROCATINTERNALERROR.PEInvalidReleaseMode, "Release mode can only be release"
    End If
    
    Select Case m_eReleaseObject
    Case ERELEASEOBJECT.roEditSet
    End Select

End Sub

