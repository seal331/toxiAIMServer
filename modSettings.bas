Attribute VB_Name = "modSettings"
Option Explicit

' Declare API functions
Private Declare Function WritePrivateProfileString Lib "kernel32" _
Alias "WritePrivateProfileStringA" _
                        (ByVal lpApplicationName As String, _
                        ByVal lpKeyName As Any, _
                        ByVal lpString As Any, _
                        ByVal lpFileName As String) As Long

Private Declare Function GetPrivateProfileString Lib "kernel32" _
Alias "GetPrivateProfileStringA" _
                        (ByVal lpApplicationName As String, _
                        ByVal lpKeyName As Any, _
                        ByVal lpDefault As String, _
                        ByVal lpReturnedString As String, _
                        ByVal nSize As Long, _
                        ByVal lpFileName As String) As Long

' Define Settings types
Public Type ConnectionSettings
    ServerHost As String
    BucpPort As Long
    BosPort As Long
    AlertPort As Long
    AdminPort As Long
    DirectoryPort As Long
End Type

Public Type DatabaseSettings
    Driver As String
    Host As String
    Port As Long
    UserID As String
    Password As String
    Name As String
End Type

Public Type ErrorURLSettings
    UnregisteredAccount As String
    IncorrectPassword As String
    SuspendedAccount As String
    DeletedAccount As String
    PasswordChange As String
End Type

Public Type Settings
    Connection As ConnectionSettings
    Database As DatabaseSettings
    ErrorURLs As ErrorURLSettings
End Type

Public AppSettings As Settings

' Function to read a value from the INI file
Private Function ReadIniSetting(ByVal strSection As String, ByVal strKeyName As String, Optional ByVal strDefaultValue = "") As String
    Dim strRet As String
    
    strRet = String(255, Chr(0))
    ReadIniSetting = Left(strRet, GetPrivateProfileString(strSection, strKeyName, strDefaultValue, strRet, Len(strRet), App.Path & "\settings.ini"))
End Function

' Function to write a value to the INI file
Private Sub WriteIniSetting(ByVal strSection As String, ByVal strKeyName As String, ByVal strNewString As String)
    Call WritePrivateProfileString(strSection, strKeyName, strNewString, App.Path & "\settings.ini")
    
    If Err.Number <> 0 Then
        LogError "Settings", "Unable to write " & strKeyName & " of " & strSection & " to the settings file! (Error code: " & Err.Number & ")"
    End If
End Sub

' Function to load settings from the INI file into the AppSettings global
Public Sub LoadSettings()
    With AppSettings
        ' Load connection-related settings:
        .Connection.ServerHost = ReadIniSetting("Connection", "ServerHost", "127.0.0.1")
        .Connection.BucpPort = CLng(ReadIniSetting("Connection", "BucpPort", "5190"))
        .Connection.BosPort = CLng(ReadIniSetting("Connection", "BosPort", "5191"))
        .Connection.AlertPort = CLng(ReadIniSetting("Connection", "AlertPort", "5192"))
        .Connection.AdminPort = CLng(ReadIniSetting("Connection", "AdminPort", "5193"))
        .Connection.DirectoryPort = CLng(ReadIniSetting("Connection", "DirectoryPort", "5194"))
        
        ' Load database-related settings:
        .Database.Driver = ReadIniSetting("Database", "Driver")
        .Database.Host = ReadIniSetting("Database", "Host")
        .Database.Port = CLng(ReadIniSetting("Database", "Port", "0"))
        .Database.UserID = ReadIniSetting("Database", "UserID")
        .Database.Password = ReadIniSetting("Database", "Password")
        .Database.Name = ReadIniSetting("Database", "Name")
        
        ' Load error URL-related settings:
        .ErrorURLs.UnregisteredAccount = ReadIniSetting("ErrorURLs", "UnregisteredAccount", "http://www.aim.aol.com/errors/UNREGISTERED_SCREENNAME.html")
        .ErrorURLs.IncorrectPassword = ReadIniSetting("ErrorURLs", "IncorrectPassword", "http://www.aim.aol.com/errors/MISMATCH_PASSWD.html")
        .ErrorURLs.SuspendedAccount = ReadIniSetting("ErrorURLs", "SuspendedAccount", "http://www.aim.aol.com/errors/SUSPENDED.html")
        .ErrorURLs.DeletedAccount = ReadIniSetting("ErrorURLs", "DeletedAccount", "http://www.aim.aol.com/errors/DELETED_ACCT.html")
        .ErrorURLs.PasswordChange = ReadIniSetting("ErrorURLs", "PasswordChange", "http://aim.aol.com/password/change_password.adp")
    End With
End Sub

' Function to write settings from the AppSettings global to the INI file
Public Sub WriteSettings()
    With AppSettings
        ' Write connection-related settings
        WriteIniSetting "Connection", "ServerHost", .Connection.ServerHost
        WriteIniSetting "Connection", "BucpPort", CStr(.Connection.BucpPort)
        WriteIniSetting "Connection", "BosPort", CStr(.Connection.BosPort)
        WriteIniSetting "Connection", "AlertPort", CStr(.Connection.AlertPort)
        WriteIniSetting "Connection", "DirectoryPort", CStr(.Connection.DirectoryPort)
        
        ' Write database-related settings
        WriteIniSetting "Database", "Driver", .Database.Driver
        WriteIniSetting "Database", "Host", .Database.Host
        WriteIniSetting "Database", "Port", CStr(.Database.Port)
        WriteIniSetting "Database", "UserID", .Database.UserID
        WriteIniSetting "Database", "Password", .Database.Password
        WriteIniSetting "Database", "Name", .Database.Name
        
        ' Write error URL-related settings
        WriteIniSetting "ErrorURLs", "UnregisteredAccount", .ErrorURLs.UnregisteredAccount
        WriteIniSetting "ErrorURLs", "IncorrectPassword", .ErrorURLs.IncorrectPassword
        WriteIniSetting "ErrorURLs", "SuspendedAccount", .ErrorURLs.SuspendedAccount
        WriteIniSetting "ErrorURLs", "DeletedAccount", .ErrorURLs.DeletedAccount
        WriteIniSetting "ErrorURLs", "PasswordChange", .ErrorURLs.PasswordChange
    End With
End Sub
