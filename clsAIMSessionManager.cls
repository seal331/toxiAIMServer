VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAIMSessionManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_colAIMSessions As Collection

Public Function Add(ByVal strScreenName As String, ByVal strIPAddress As String, ByVal intAuthSocket As Integer, ByVal strChallenge As String) As clsAIMSession
    Dim oAIMSession As clsAIMSession
    Set oAIMSession = New clsAIMSession
    
    With oAIMSession
        .ScreenName = strScreenName
        .IPAddress = strIPAddress
        .AuthSocket = intAuthSocket
        .Challenge = strChallenge
    End With
    
    m_colAIMSessions.Add oAIMSession, TrimData(strScreenName)
    
    Set Add = oAIMSession
    Set oAIMSession = Nothing
End Function

Public Sub Remove(ByVal strScreenName As String)
    m_colAIMSessions.Remove TrimData(strScreenName)
End Sub

Public Property Get Item(vntIndexKey As Variant) As clsAIMSession
    On Error Resume Next
    Set Item = m_colAIMSessions(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = m_colAIMSessions.Count
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = m_colAIMSessions.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set m_colAIMSessions = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_colAIMSessions = Nothing
End Sub
