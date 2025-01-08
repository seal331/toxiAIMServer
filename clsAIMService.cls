VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAIMService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_lngFoodgroup As Long
Private m_bytCookie() As Byte
Private m_intIndex As Integer

Public Property Get Foodgroup() As Long
    Foodgroup = m_lngFoodgroup
End Property

Public Property Let Foodgroup(ByVal lngFoodgroup As Long)
    m_lngFoodgroup = lngFoodgroup
End Property

Public Property Get Cookie() As Byte()
    Cookie = m_bytCookie
End Property

Public Sub SetCookie(ByRef bytCookie() As Byte)
    m_bytCookie = bytCookie
End Sub

Public Property Get Index() As Integer
    Index = m_intIndex
End Property

Public Property Let Index(ByVal intIndex As Integer)
    m_intIndex = intIndex
End Property
