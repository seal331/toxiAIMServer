VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFeedbagObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_strName As String
Private m_lngGroupID As Long
Private m_lngItemID As Long
Private m_lngClassID As Long
Private m_oAttributes As clsTLVList

Public Property Get Name() As String
    Name = m_strName
End Property

Public Property Let Name(ByVal strName As String)
    m_strName = strName
End Property

Public Property Get GroupID() As Long
    GroupID = m_lngGroupID
End Property

Public Property Let GroupID(ByVal lngGroupID As Long)
    m_lngGroupID = lngGroupID
End Property

Public Property Get ItemID() As Long
    ItemID = m_lngItemID
End Property

Public Property Let ItemID(ByVal lngItemID As Long)
    m_lngItemID = lngItemID
End Property

Public Property Get ClassID() As Long
    ClassID = m_lngClassID
End Property

Public Property Let ClassID(ByVal lngClassID As Long)
    m_lngClassID = lngClassID
End Property

Public Property Get Attributes() As clsTLVList
    Set Attributes = m_oAttributes
End Property

Public Property Let Attributes(ByVal oAttributes As clsTLVList)
    m_oAttributes = oAttributes
End Property

Private Sub Class_Initialize()
    Set m_oAttributes = New clsTLVList
End Sub

Private Sub Class_Terminate()
    Set m_oAttributes = Nothing
End Sub
