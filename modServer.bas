Attribute VB_Name = "modServer"
Public oAIMSessionManager As clsAIMSessionManager

Public Enum PasswordType
    PasswordTypeXor
    PasswordTypeWeakMD5
    PasswordTypeStrongMD5
End Enum

Public Enum LoginState
    LoginStateGood = 0
    LoginStateUnregistered = 1
    LoginStateIncorrectPassword = 2
    LoginStateSuspended = 3
    LoginStateDeleted = 4
End Enum

Public Sub Main()
    LoadSettings
    
    If Dir(App.Path & "\settings.ini") = "" Then
        WriteSettings
    End If
    
    Set oAIMSessionManager = New clsAIMSessionManager
    
    Load frmMain
    frmMain.Show
End Sub

Public Function TrimData(ByVal strData As String) As String
    TrimData = Replace(LCase(strData), " ", vbNullString)
End Function

Public Function GetUnixTimestamp(ByVal dt As Date) As Double
    GetUnixTimestamp = DateDiff("s", #1/1/1970#, dt)
End Function

Public Function ConvertUnixTimestamp(ByVal lngTimestamp As Double) As Date
    ConvertUnixTimestamp = DateAdd("s", lngTimestamp, #1/1/1970#)
End Function

Public Function CheckLogin(ByVal strScreenName As String, _
                           ByRef bytClientPassword() As Byte, _
                           ByVal intPasswordType As PasswordType, _
                           Optional ByVal strChallenge As String = vbNullString) As LoginState
    
    Dim RS As ADODB.Recordset
    Dim oMD5Hasher As clsMD5Hash
    Dim bytPassword() As Byte
    Dim bytServerPassword() As Byte
    Dim bytMD5Pass() As Byte
    
    ' Query the database for the user's password and status via their screen name.
    Set RS = ExecutePreparedQuery("SELECT `password`, `is_suspended`, `is_deleted` FROM `accounts` WHERE `screen_name` = ?", TrimData(strScreenName))
    
    ' Check if a record for the user was found
    If RS.EOF Then
        LogError "Server", "Unable to find user in database!"
        
        RS.Close
        Set RS = Nothing
        
        CheckLogin = LoginStateUnregistered
        Exit Function
    End If
    
    LogDebug "Server", "Found user in database!"
    
    ' Convert the password from the database to a byte array
    bytPassword = StringToBytes(RS.Fields("password"))
    
    Select Case intPasswordType
    
        ' Checking for XOR-based passwords, used by AIM 1.x - 3.5, is not yet implemented.
        Case PasswordTypeXor
            CheckLogin = LoginStateIncorrectPassword
            
        ' Check for MD5-based passwords used by AIM 3.5 up until 6.0, where they switched
        ' to UAS.
        '
        ' There exists 2 versions - a "weak" version used by clients pre-AIM 5.x, where
        ' the data before hashing consists of the challenge, the plaintext password,
        ' and the brand string.
        '
        ' However, in AIM 5.x, they switched to a "strong" version, discernible by TLV 0x4A,
        ' which is exactly the same, however the password is now hashed in a additional layer
        ' of MD5.
        Case PasswordTypeWeakMD5, PasswordTypeStrongMD5
            Set oMD5Hasher = New clsMD5Hash
                    
            If intPasswordType = PasswordTypeStrongMD5 Then
                bytMD5Pass = oMD5Hasher.HashBytes(bytPassword)
            End If
            
            bytServerPassword = oMD5Hasher.HashBytes(ConcatBytes( _
                StringToBytes(strChallenge), _
                IIf(intPasswordType = PasswordTypeStrongMD5, bytMD5Pass, bytPassword), _
                StringToBytes("AOL Instant Messenger (SM)") _
            ))
            
            LogDebug "Server", "Client-generated MD5 Password Hash: " & ByteArrayToHexString(bytClientPassword)
            LogDebug "Server", "Server-generated MD5 Password Hash: " & ByteArrayToHexString(bytServerPassword)
            
        Case Else
            LogError "Server", "Invalid password type - defaulting to incorrect password"
            
            CheckLogin = LoginStateIncorrectPassword
        
    End Select
    
    ' Compare both hashes to each other
    If IsBytesEqual(bytServerPassword, bytClientPassword) Then
        ' Ensure they aren't suspended or deleted
        If RS.Fields("is_suspended") = 1 Then
            CheckLogin = LoginStateSuspended
        ElseIf RS.Fields("is_deleted") = 1 Then
            CheckLogin = LoginStateDeleted
        Else
            CheckLogin = LoginStateGood
        End If
    Else
        CheckLogin = LoginStateIncorrectPassword
    End If
    
    RS.Close
    Set RS = Nothing
End Function

Public Sub SetupAccount(ByVal oAIMSession As clsAIMSession)
    Dim RS As ADODB.Recordset
    
    ' Query the account details
    Set RS = ExecutePreparedQuery("SELECT * FROM `accounts` WHERE `screen_name` = ?", TrimData(oAIMSession.ScreenName))
    
    If RS.EOF Then
        RS.Close
        Set RS = Nothing
        Exit Sub
    End If
    
    With oAIMSession
        ' Map basic properties
        .ID = RS.Fields("id")
        .FormattedScreenName = RS.Fields("format")
        .EmailAddress = RS.Fields("email")
        .Password = RS.Fields("password")
        .RegistrationStatus = RS.Fields("registration_status")
        .RegistrationTime = ConvertUnixTimestamp(RS.Fields("time_registered"))
        .SignOnTime = Now
        .WarningLevel = RS.Fields("evil_temporary")
        .Subscriptions = RS.Fields("subscriptions")
        .ParentalControls = RS.Fields("parental_controls")
            
        ' Set user class
        .UserClass = IIf(RS.Fields("is_confirmed") = 0, UserFlagsUnconfirmed, UserFlagsOscarFree)
            
        If RS.Fields("is_internal") = 1 Then
            .UserClass = .UserClass Or UserFlagsInternal Or UserFlagsAdministrator
        End If
            
        ' Update sign-on time in the database
        ExecutePreparedNonQuery "UPDATE `accounts` SET `time_login` = ? WHERE `id` = ?", GetUnixTimestamp(.SignOnTime), .ID
        
        ' Mark this session as authorized
        .Authorized = True
    End With
    
    RS.Close
    Set RS = Nothing
End Sub

' TODO(subpurple): pull from i.e. `feedbag` table in the database
Public Function GetFeedbagData(ByVal oAIMSession As clsAIMSession) As Byte()
    Dim oByteWriter As New clsByteBuffer
    
    With oByteWriter
        .WriteByte 0    ' Number of classes in the feedbag (always 0)
        .WriteByte 0    ' Number of items in the feedbag
        
        ' Add root group
        .WriteStringU16 vbNullString    ' The item's name as a UTF-8 string
        .WriteU16 0                     ' The item's group ID
        .WriteU16 0                     ' The item's ID
        .WriteU16 &H1                   ' The item's class (i.e. group)
        .WriteU16 0                     ' The number of attributes associated with the item (e.g. order)

        .WriteU32 GetUnixTimestamp(Now) ' Feedbag's last change time
        
        GetFeedbagData = .Buffer
    End With
End Function

Public Function FeedbagCheckIfNew(ByVal oAIMSession As clsAIMSession, ByVal dblFeedbagTimestamp As Double, ByVal lngFeedbagItems As Long) As Boolean
    FeedbagCheckIfNew = True
End Function

Public Function FeedbagAddItem(ByVal oAIMSession As clsAIMSession, ByVal strName As String, ByVal lngGroupID As Long, ByVal lngItemID As Long, ByVal lngClassID As Long, ByVal oAttributes As clsTLVList) As Long
    LogDebug "Server", oAIMSession.FormattedScreenName & " is adding feedbag item " & strName & " with ID " & DecimalToHex(lngItemID) & " via group ID " & DecimalToHex(lngGroupID) & " with attributes: " & ByteArrayToHexString(oAttributes.GetSerializedChain)

    FeedbagAddItem = 0
End Function

Public Function FeedbagUpdateItem(ByVal oAIMSession As clsAIMSession, ByVal strName As String, ByVal lngGroupID As Long, ByVal lngItemID As Long, ByVal lngClassID As Long, ByVal oAttributes As clsTLVList) As Long
    LogDebug "Server", oAIMSession.FormattedScreenName & " is updating feedbag item " & strName & " with ID " & DecimalToHex(lngItemID) & " via group ID " & DecimalToHex(lngGroupID) & " with attributes: " & ByteArrayToHexString(oAttributes.GetSerializedChain)

    FeedbagUpdateItem = 0
End Function

Public Function FeedbagDeleteItem(ByVal oAIMSession As clsAIMSession, ByVal strName As String, ByVal lngGroupID As Long, ByVal lngItemID As Long, ByVal lngClassID As Long, ByVal oAttributes As clsTLVList) As Long
    LogDebug "Server", oAIMSession.FormattedScreenName & " is deleting feedbag item " & strName & " with ID " & DecimalToHex(lngItemID) & " via group ID " & DecimalToHex(lngGroupID) & " with attributes: " & ByteArrayToHexString(oAttributes.GetSerializedChain)

    FeedbagRemoveItem = 0
End Function

