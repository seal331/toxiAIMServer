Attribute VB_Name = "modSettings"
Option Explicit

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpString As Any, _
    ByVal lpFileName As String) As Long

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String) As Long

Public g_strServerHost             As String
Public g_lngBucpPort               As Long
Public g_lngBosPort                As Long
Public g_lngAlertPort              As Long
Public g_lngAdminPort              As Long
Public g_lngDirectoryPort          As Long
Public g_strDatabaseDriver         As String
Public g_strDatabaseHost           As String
Public g_lngDatabasePort           As Long
Public g_strDatabaseUserID         As String
Public g_strDatabasePassword       As String
Public g_strDatabaseName           As String
Public g_strUnregisteredAccountURL As String
Public g_strIncorrectPasswordURL   As String
Public g_strSuspendedAccountURL    As String
Public g_strDeletedAccountURL      As String
Public g_strPasswordChangeURL      As String

Private Function ReadINI(ByVal strSection As String, ByVal strKeyName As String, Optional ByVal strDefaultValue = "") As String
    Dim strRet As String
    
    strRet = String(255, Chr(0))
    ReadINI = Left(strRet, GetPrivateProfileString(strSection, strKeyName, strDefaultValue, strRet, Len(strRet), App.Path & "\settings.ini"))
End Function

Private Sub WriteINI(ByVal strSection As String, ByVal strKeyName As String, ByVal strNewString As String)
    Call WritePrivateProfileString(strSection, strKeyName, strNewString, App.Path & "\settings.ini")
    
    If Err.Number <> 0 Then
        LogError "Settings", "Unable to write " & strKeyName & " of " & strSection & " to the settings file! (Error code: " & Err.Number & ")"
    End If
End Sub

Public Sub LoadSettings()
    With frmMain
        ' Load connection-related settings:
        g_strServerHost = ReadINI("Connection", "ServerHost", "127.0.0.1")
        g_lngBucpPort = CLng(ReadINI("Connection", "BucpPort", "5190"))
        g_lngBosPort = CLng(ReadINI("Connection", "BosPort", "5191"))
        g_lngAlertPort = CLng(ReadINI("Connection", "AlertPort", "5192"))
        g_lngAdminPort = CLng(ReadINI("Connection", "AdminPort", "5193"))
        g_lngDirectoryPort = CLng(ReadINI("Connection", "DirectoryPort", "5194"))
        
        ' Load database-related settings:
        g_strDatabaseDriver = ReadINI("Database", "Driver")
        g_strDatabaseHost = ReadINI("Database", "Host")
        g_lngDatabasePort = CLng(ReadINI("Database", "Port", "0"))
        g_strDatabaseUserID = ReadINI("Database", "UserID")
        g_strDatabasePassword = ReadINI("Database", "Password")
        g_strDatabaseName = ReadINI("Database", "Name")
        
        ' Load error URL-related settings:
        g_strUnregisteredAccountURL = ReadINI("ErrorURLs", "UnregisteredAccount", "http://www.aim.aol.com/errors/UNREGISTERED_SCREENNAME.html")
        g_strIncorrectPasswordURL = ReadINI("ErrorURLs", "IncorrectPassword", "http://www.aim.aol.com/errors/MISMATCH_PASSWD.html")
        g_strSuspendedAccountURL = ReadINI("ErrorURLs", "SuspendedAccount", "http://www.aim.aol.com/errors/SUSPENDED.html")
        g_strDeletedAccountURL = ReadINI("ErrorURLs", "DeletedAccount", "http://www.aim.aol.com/errors/DELETED_ACCT.html")
        g_strPasswordChangeURL = ReadINI("ErrorURLs", "PasswordChange", "http://aim.aol.com/password/change_password.adp")
    
        ' Load them into the settings form aswell:
        .txtServerHost.Text = g_strServerHost
        .txtBucpServerPort.Text = g_lngBucpPort
        .txtBosServerPort.Text = g_lngBosPort
        .txtAlertServerPort.Text = g_lngAlertPort
        .txtAdminServerPort.Text = g_lngAdminPort
        .txtDirectoryServerPort.Text = g_lngDirectoryPort
        .cboDbDriver.Text = g_strDatabaseDriver
        .txtDbHost.Text = g_strDatabaseHost
        .txtDbPort.Text = g_lngDatabasePort
        .txtDbUserId.Text = g_strDatabaseUserID
        .txtDbPassword.Text = g_strDatabasePassword
        .txtDbName.Text = g_strDatabaseName
        .txtUnregisteredAcctUrl.Text = g_strUnregisteredAccountURL
        .txtIncorrectPasswdUrl.Text = g_strIncorrectPasswordURL
        .txtSuspendedAcctUrl.Text = g_strSuspendedAccountURL
        .txtDeletedAcctUrl.Text = g_strDeletedAccountURL
        .txtPasswdChangeUrl.Text = g_strPasswordChangeURL
    End With
End Sub

Public Sub WriteSettings()
    With frmMain
        ' Write connection-related settings
        WriteINI "Connection", "ServerHost", .txtServerHost.Text
        WriteINI "Connection", "BucpPort", .txtBucpServerPort.Text
        WriteINI "Connection", "BosPort", .txtBosServerPort.Text
        WriteINI "Connection", "AlertPort", .txtAlertServerPort.Text
        WriteINI "Connection", "DirectoryPort", .txtDirectoryServerPort.Text
        
        ' Write database-related settings
        WriteINI "Database", "Driver", .cboDbDriver.Text
        WriteINI "Database", "Host", .txtDbHost.Text
        WriteINI "Database", "Port", .txtDbPort.Text
        WriteINI "Database", "UserID", .txtDbUserId.Text
        WriteINI "Database", "Password", .txtDbPassword.Text
        WriteINI "Database", "Name", .txtDbName.Text
        
        ' Write error URL-related settings
        WriteINI "ErrorURLs", "UnregisteredAccount", .txtUnregisteredAcctUrl.Text
        WriteINI "ErrorURLs", "IncorrectPassword", .txtIncorrectPasswdUrl.Text
        WriteINI "ErrorURLs", "SuspendedAccount", .txtSuspendedAcctUrl.Text
        WriteINI "ErrorURLs", "DeletedAccount", .txtDeletedAcctUrl.Text
        WriteINI "ErrorURLs", "PasswordChange", .txtPasswdChangeUrl.Text
    End With
End Sub
