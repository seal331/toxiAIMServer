VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAIMSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_intIndex As Integer
Private m_dblID As Double
Private m_intWarningLevel As Integer
Private m_strScreenName As String
Private m_strFormattedScreenName As String
Private m_strPassword As String
Private m_strEmailAddress As String
Private m_strIPAddress As String
Private m_lngUserClass As Long
Private m_lngUserStatus As Long
Private m_dblSubscriptions As Double
Private m_dblParentalControls As Double
Private m_lngRegistrationStatus As Long
Private m_dtRegistrationTime As Date
Private m_dtSignOnTime As Date
Private m_dtIdleTime As Date
Private m_strChallenge As String
Private m_bytCookie() As Byte
Private m_intAuthSocket As Integer
Private m_blnIdle As Boolean
Private m_blnAuthorized As Boolean
Private m_blnSignedOn As Boolean
Private m_lngServiceID As Long
Private m_colServices As Collection

Public Property Get Index() As Integer
    Index = m_intIndex
End Property

Public Property Let Index(ByVal intIndex As Integer)
    m_intIndex = intIndex
End Property

Public Property Get ID() As Double
    ID = m_dblID
End Property

Public Property Let ID(ByVal dblID As Double)
    m_dblID = dblID
End Property

Public Property Get WarningLevel() As Integer
    WarningLevel = m_intWarningLevel
End Property

Public Property Let WarningLevel(ByVal intWarningLevel As Integer)
    m_intWarningLevel = intWarningLevel
End Property

Public Property Get ScreenName() As String
    ScreenName = m_strScreenName
End Property

Public Property Let ScreenName(ByVal strScreenName As String)
    m_strScreenName = strScreenName
End Property

Public Property Get FormattedScreenName() As String
    FormattedScreenName = m_strFormattedScreenName
End Property

Public Property Let FormattedScreenName(ByVal strFormattedScreenName As String)
    m_strFormattedScreenName = strFormattedScreenName
End Property

Public Property Get Password() As String
    Password = m_strPassword
End Property

Public Property Let Password(ByVal strPassword As String)
    m_strPassword = strPassword
End Property

Public Property Get EmailAddress() As String
    EmailAddress = m_strEmailAddress
End Property

Public Property Let EmailAddress(ByVal strEmailAddress As String)
    m_strEmailAddress = strEmailAddress
End Property

Public Property Get IPAddress() As String
    IPAddress = m_strIPAddress
End Property

Public Property Let IPAddress(ByVal strIPAddress As String)
    m_strIPAddress = strIPAddress
End Property

Public Property Get UserClass() As Long
    UserClass = m_lngUserClass
End Property

Public Property Let UserClass(ByVal lngUserClass As Long)
    m_lngUserClass = lngUserClass
End Property

Public Property Get UserStatus() As Long
    UserStatus = m_lngUserStatus
End Property

Public Property Let UserStatus(ByVal lngUserStatus As Long)
    m_lngUserStatus = lngUserStatus
End Property

Public Property Get Subscriptions() As Double
    Subscriptions = m_dblSubscriptions
End Property

Public Property Let Subscriptions(ByVal dblSubscriptions As Double)
    m_dblSubscriptions = dblSubscriptions
End Property

Public Property Get ParentalControls() As Double
    ParentalControls = m_dblParentalControls
End Property

Public Property Let ParentalControls(ByVal dblParentalControls As Double)
    m_dblParentalControls = dblParentalControls
End Property

Public Property Get RegistrationStatus() As Long
    RegistrationStatus = m_lngRegistrationStatus
End Property

Public Property Let RegistrationStatus(ByVal lngRegistrationStatus As Long)
    m_lngRegistrationStatus = lngRegistrationStatus
End Property

Public Property Get RegistrationTime() As Date
    RegistrationTime = m_dtRegistrationTime
End Property

Public Property Let RegistrationTime(ByVal dtRegistrationTime As Date)
    m_dtRegistrationTime = dtRegistrationTime
End Property

Public Property Get SignOnTime() As Date
    SignOnTime = m_dtSignOnTime
End Property

Public Property Let SignOnTime(ByVal dtSignOnTime As Date)
    m_dtSignOnTime = dtSignOnTime
End Property

Public Property Get IdleTime() As Date
    IdleTime = m_dtIdleTime
End Property

Public Property Let IdleTime(ByVal dtIdleTime As Date)
    m_dtIdleTime = dtIdleTime
End Property

Public Property Get Challenge() As String
    Challenge = m_strChallenge
End Property

Public Property Let Challenge(ByVal strChallenge As String)
    m_strChallenge = strChallenge
End Property

Public Property Get Cookie() As Byte()
    Cookie = m_bytCookie
End Property

Public Property Let Cookie(ByRef bytCookie() As Byte)
    m_bytCookie = bytCookie
End Property

Public Property Get AuthSocket() As Integer
    AuthSocket = m_intAuthSocket
End Property

Public Property Let AuthSocket(ByVal intAuthSocket As Integer)
    m_intAuthSocket = intAuthSocket
End Property

Public Property Get Idle() As Boolean
    Idle = m_blnIdle
End Property

Public Property Let Idle(ByVal blnIdle As Boolean)
    m_blnIdle = blnIdle
End Property

Public Property Get Authorized() As Boolean
    Authorized = m_blnAuthorized
End Property

Public Property Let Authorized(ByVal blnAuthorized As Boolean)
    m_blnAuthorized = blnAuthorized
End Property

Public Property Get SignedOn() As Boolean
    SignedOn = m_blnSignedOn
End Property

Public Property Let SignedOn(ByVal blnSignedOn As Boolean)
    m_blnSignedOn = blnSignedOn
End Property

Public Property Get Services() As Collection
    Set Services = m_colServices
End Property

Public Function AddService(ByVal lngFoodgroup As Long, ByRef bytCookie() As Byte) As clsAIMService
    Dim oAIMService As clsAIMService
    Set oAIMService = New clsAIMService
    
    oAIMService.Foodgroup = lngFoodgroup
    oAIMService.SetCookie bytCookie
    
    m_lngServiceID = m_lngServiceID + 1
    m_colServices.Add oAIMService, CStr(m_lngServiceID)
    
    Set AddService = oAIMService
    Set oAIMService = Nothing
End Function

Private Sub Class_Initialize()
    Set m_colServices = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_colServices = Nothing
End Sub
