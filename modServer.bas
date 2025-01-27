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

' Converts a Date object into a Unix timestamp.
Public Function GetUnixTimestamp(ByVal dt As Date) As Double
    GetUnixTimestamp = DateDiff("s", #1/1/1970#, dt)
End Function

' Converts a Unix timestamp into a Date object.
Public Function ConvertUnixTimestamp(ByVal lngTimestamp As Double) As Date
    ConvertUnixTimestamp = DateAdd("s", lngTimestamp, #1/1/1970#)
End Function

Public Function RoastStringViaChars(ByVal strInput As String, ByVal strChars As String) As Byte()
    Dim bytOutput() As Byte
    Dim bytInput As Byte
    Dim bytChar As Byte
    Dim i As Long
    Dim lngCharsIndex As Long
    
    ReDim bytOutput(0 To Len(strInput) - 1)
    
    For i = 1 To Len(strInput)
        ' Get the byte value of the input character
        bytInput = Asc(Mid(strInput, i, 1))
        
        ' Get the byte value of the corresponding character in CHARS
        lngCharsIndex = ((i - 1) Mod Len(strChars)) + 1
        bytChar = Asc(Mid(strChars, lngCharsIndex, 1))
        
        ' XOR the values
        bytOutput(i - 1) = bytInput Xor bytChar
    Next i
    
    RoastStringViaChars = bytOutput
End Function

Public Function CheckLogin(ByVal strScreenName As String, _
                           ByRef bytClientPassword() As Byte, _
                           ByVal enuPasswordType As PasswordType, _
                           Optional ByVal strChallenge As String = vbNullString) As LoginState
    
    Dim rst As ADODB.Recordset
    Dim oMD5Hasher As clsMD5Hash
    Dim strPassword As String
    Dim bytPassword() As Byte
    Dim bytServerPassword() As Byte
    Dim bytMD5Pass() As Byte
    
    ' Query the database for the user's password and status via their screen name.
    Set rst = ExecutePreparedQuery("SELECT `password`, `is_suspended`, `is_deleted` FROM `users` WHERE `screen_name` = ?", TrimData(strScreenName))
    
    ' Check if a record for the user was found
    If rst.EOF Then
        LogError "Server", "Unable to find user in database!"
        
        rst.Close
        Set rst = Nothing
        
        CheckLogin = LoginStateUnregistered
        Exit Function
    End If
    
    LogDebug "Server", "Found user in database!"
    
    ' Get the password from the database and convert it to a byte array
    strPassword = rst.Fields("password")
    bytPassword = StringToBytes(strPassword)
    
    Select Case enuPasswordType
    
        ' Check for XOR-based passwords used prior to AIM 3.5.
        Case PasswordTypeXor
            ' TODO(subpurple):  The original Java client uses a different set of `CHARS`.
            ' We should check for those aswell in the future.
            bytServerPassword = RoastStringViaChars(strPassword, _
                Chr(&HF3) & Chr(&H26) & Chr(&H81) & Chr(&HC4) & _
                Chr(&H39) & Chr(&H86) & Chr(&HDB) & Chr(&H92) & _
                Chr(&H71) & Chr(&HA3) & Chr(&HB9) & Chr(&HE6) & _
                Chr(&H53) & Chr(&H7A) & Chr(&H95) & Chr(&H7C))
            
            LogDebug "Server", "Client-roasted password: " & BytesToHex(bytClientPassword)
            LogDebug "Server", "Server-roasted password: " & BytesToHex(bytServerPassword)
            
        ' Check for MD5-based passwords used by AIM 3.5 up until 6.0, when they switched
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
                    
            If enuPasswordType = PasswordTypeStrongMD5 Then
                bytMD5Pass = oMD5Hasher.HashBytes(bytPassword)
            End If
            
            bytServerPassword = oMD5Hasher.HashBytes(ConcatBytes( _
                StringToBytes(strChallenge), _
                IIf(enuPasswordType = PasswordTypeStrongMD5, bytMD5Pass, bytPassword), _
                StringToBytes("AOL Instant Messenger (SM)") _
            ))
            
            LogDebug "Server", "Client-generated MD5 Password Hash: " & BytesToHex(bytClientPassword)
            LogDebug "Server", "Server-generated MD5 Password Hash: " & BytesToHex(bytServerPassword)
            
        Case Else
            LogFatal "Server", "Invalid password type provided!"
        
    End Select
    
    ' Compare both hashes to each other
    If IsBytesEqual(bytServerPassword, bytClientPassword) Then
        ' Ensure they aren't suspended or deleted
        If rst.Fields("is_suspended") = 1 Then
            CheckLogin = LoginStateSuspended
        ElseIf rst.Fields("is_deleted") = 1 Then
            CheckLogin = LoginStateDeleted
        Else
            CheckLogin = LoginStateGood
        End If
    Else
        CheckLogin = LoginStateIncorrectPassword
    End If
    
    rst.Close
    Set rst = Nothing
End Function

Public Sub SetupAccount(ByVal oAIMUser As clsAIMSession)
    Dim rst As ADODB.Recordset
    
    ' Query the user details
    Set rst = ExecutePreparedQuery("SELECT * FROM `users` WHERE `screen_name` = ?", TrimData(oAIMUser.ScreenName))
    
    If rst.EOF Then
        rst.Close
        Set rst = Nothing
        Exit Sub
    End If
    
    With oAIMUser
        ' Map basic properties
        .ID = rst.Fields("id")
        .FormattedScreenName = rst.Fields("format")
        .EmailAddress = rst.Fields("email")
        .Password = rst.Fields("password")
        .RegistrationStatus = rst.Fields("registration_status")
        .RegistrationTime = ConvertUnixTimestamp(rst.Fields("time_registered"))
        .SignOnTime = Now
        .WarningLevel = rst.Fields("evil_temporary")
        .Subscriptions = rst.Fields("subscriptions")
        .ParentalControls = rst.Fields("parental_controls")
        
        ' Set user class
        If rst.Fields("is_confirmed") = 0 Then
            .UserClass = OSERVICE_USER_FLAG_DAMNED_TRANSIENT
        Else
            .UserClass = OSERVICE_USER_FLAG_OSCAR_FREE
        End If
        
        If rst.Fields("is_internal") = 1 Then
            .UserClass = _
                OSERVICE_USER_FLAG_OSCAR_PAY Or _
                OSERVICE_USER_FLAG_INTERNAL Or _
                OSERVICE_USER_FLAG_ADMINISTRATOR
        End If
        
        LogDebug "Server", .UserClass
        
        ' Update sign-on time in the database
        ExecutePreparedNonQuery "UPDATE `users` SET `time_login` = ? WHERE `id` = ?", GetUnixTimestamp(.SignOnTime), .ID
        
        ' Mark this session as authorized
        .Authorized = True
        
        ' Ensure feedbag is in order
        Call FeedbagEnsureRootGroupExists(oAIMUser)
    End With
    
    rst.Close
    Set rst = Nothing
End Sub

' Ensures the root group exists for the given account and creates it if not.
Public Sub FeedbagEnsureRootGroupExists(ByVal oAIMUser As clsAIMSession)
    Dim rst As ADODB.Recordset
    Set rst = ExecutePreparedQuery("SELECT EXISTS(SELECT * FROM `feedbag` WHERE `user_id` = ? LIMIT 1)", oAIMUser.ID)
    
    If rst.EOF Or rst.Fields(0).Value = 0 Then
        LogInformation "Server", "Initializing server-side feedbag..."
    
        ' Create master feedbag group
        ExecutePreparedNonQuery _
            "INSERT INTO `feedbag` (`user_id`, `name`, `group_id`, `item_id`, `class_id`) " & _
            "VALUES (?, '', 0, 0, 1)", oAIMUser.ID
            
        ' Update feedbag time and feedbag item count
        Call FeedbagUpdateDatabase(oAIMUser)
    End If
    
    rst.Close
    Set rst = Nothing
End Sub

' Retrieves all feedbag items for a given user.
Public Function FeedbagGetData(ByVal oAIMUser As clsAIMSession) As Collection
    Dim rst As ADODB.Recordset
    Dim oFeedbagItem As clsFeedbagItem
    Dim bFeedbagAttributes() As Byte
    Dim colFeedbagItems As New Collection
    
    Set rst = ExecuteQuery( _
        "SELECT `name`, `group_id`, `item_id`, `class_id`, `attributes` FROM `feedbag` " & _
        "WHERE `user_id` = " & oAIMUser.ID & " ORDER BY `group_id`, `item_id` ASC")
    
    Do Until rst.EOF
        Set oFeedbagItem = New clsFeedbagItem
        
        ' Check if attributes field has data and load it
        If rst.Fields("attributes").ActualSize > 0 Then
            bFeedbagAttributes = rst.Fields("attributes").GetChunk(rst.Fields("attributes").ActualSize)
        Else
            bFeedbagAttributes = GetEmptyBytes
        End If
        
        ' Populate the item with data
        With oFeedbagItem
            .Name = rst.Fields("name").Value
            .GroupID = rst.Fields("group_id").Value
            .ItemID = rst.Fields("item_id").Value
            .ClassID = rst.Fields("class_id").Value
            .SetAttributes bFeedbagAttributes
            
            LogDebug "Server", "Retrieved feedbag item " & .Name & " from database:"
            LogDebug "Server", "Group ID: 0x" & DecimalToHex(.GroupID)
            LogDebug "Server", "Item ID: 0x" & DecimalToHex(.ItemID)
            LogDebug "Server", "Class ID: 0x" & DecimalToHex(.ClassID)
            LogDebug "Server", "Attributes: " & BytesToHex(.Attributes)
        End With
        
        colFeedbagItems.Add oFeedbagItem
        
        rst.MoveNext
    Loop
    
    rst.Close
    
    Set rst = Nothing
    Set FeedbagGetData = colFeedbagItems
End Function

' Retrieves the feedbag timestamp for a given user.
Public Function FeedbagGetTime(ByVal oAIMUser As clsAIMSession) As Double
    Dim rst As ADODB.Recordset
    Set rst = ExecutePreparedQuery("SELECT `feedbag_time` FROM `users` WHERE `id` = ?", oAIMUser.ID)
    
    FeedbagGetTime = rst.Fields("feedbag_time").Value
    
    rst.Close
    Set rst = Nothing
End Function

' Checks if the user's feedbag is modified by passing the feedbag timestamp and total item count, and comparing it
' to the values from the database.
Public Function FeedbagIsModified(ByVal oAIMUser As clsAIMSession, ByVal dblFeedbagTimestamp As Double, ByVal lngFeedbagItems As Long) As Boolean
    Dim rst As ADODB.Recordset
    Set rst = ExecutePreparedQuery("SELECT `feedbag_time`, `feedbag_items` FROM `users` WHERE `id` = ?", oAIMUser.ID)
    
    If dblFeedbagTimestamp < rst.Fields("feedbag_time").Value Or lngFeedbagItems <> rst.Fields("feedbag_items").Value Then
        FeedbagIsModified = True
    Else
        FeedbagIsModified = False
    End If
    
    rst.Close
    Set rst = Nothing
End Function

' Updates the feedbag timestamp and total item count for a given user.
Public Sub FeedbagUpdateDatabase(ByVal oAIMUser As clsAIMSession)
    LogInformation "Database", "Updating feedbag time and items..."
    
    ExecutePreparedNonQuery _
        "UPDATE `users` SET `feedbag_time` = UNIX_TIMESTAMP(), `feedbag_items` = (" & _
            "SELECT COUNT(`id`) FROM `feedbag` WHERE `users`.`id` = `feedbag`.`user_id`" & _
        ") WHERE `id` = ?", _
        oAIMUser.ID
End Sub

' Adds a feedbag item for a given user.
Public Function FeedbagAddItem(ByVal oAIMUser As clsAIMSession, ByVal strName As String, ByVal lngGroupID As Long, ByVal lngItemID As Long, ByVal lngClassID As Long, ByRef bytAttributes() As Byte) As Long
    Dim rst As ADODB.Recordset
    
    LogInformation "Server", oAIMUser.FormattedScreenName & " is adding feedbag item " & strName
    LogInformation "Server", "Group ID: 0x" & DecimalToHex(lngGroupID)
    LogInformation "Server", "Item ID: 0x" & DecimalToHex(lngItemID)
    LogInformation "Server", "Class ID: 0x" & DecimalToHex(lngClassID)
    LogInformation "Server", "Attributes: " & BytesToHex(bytAttributes)
    
    ' Ensure it doesn't already exist:
    Set rst = ExecutePreparedQuery("SELECT EXISTS(SELECT * FROM `feedbag` WHERE `user_id` = ? AND `group_id` = ? AND `item_id` = ? LIMIT 1)", oAIMUser.ID, lngGroupID, lngItemID)
    
    If rst.Fields(0).Value <> 0 Then
        LogError "Server", "Feedbag item already exists!"
        
        ' Return a status code signifying it already exists
        FeedbagAddItem = FEEDBAG_STATUS_CODES_ALREADY_EXISTS
    Else
        ' Insert it database-side then:
        ExecutePreparedQuery _
            "INSERT INTO `feedbag` (`user_id`, `name`, `group_id`, `item_id`, `class_id`, `attributes`) " & _
            "VALUES (?, ?, ?, ?, ?, ?)", _
            oAIMUser.ID, strName, lngGroupID, lngItemID, lngClassID, bytAttributes
        
        ' Update feedbag time and total items from the user side
        Call FeedbagUpdateDatabase(oAIMUser)
        
        ' Return a status code signifying it was successful
        FeedbagAddItem = FEEDBAG_STATUS_CODES_SUCCESS
    End If
    
    rst.Close
    Set rst = Nothing
End Function

' Updates a feedbag item for a given user.
Public Function FeedbagUpdateItem(ByVal oAIMUser As clsAIMSession, ByVal strName As String, ByVal lngGroupID As Long, ByVal lngItemID As Long, ByVal lngClassID As Long, ByRef bytAttributes() As Byte) As Long
    Dim rst As ADODB.Recordset
        
    LogInformation "Server", oAIMUser.FormattedScreenName & " is updating feedbag item " & strName
    LogInformation "Server", "Group ID: 0x" & DecimalToHex(lngGroupID)
    LogInformation "Server", "Item ID: 0x" & DecimalToHex(lngItemID)
    LogInformation "Server", "Class ID: 0x" & DecimalToHex(lngClassID)
    LogInformation "Server", "Attributes: " & BytesToHex(bytAttributes)
    
    ' Ensure it exists:
    Set rst = ExecutePreparedQuery("SELECT EXISTS(SELECT * FROM `feedbag` WHERE `user_id` = ? AND `group_id` = ? AND `item_id` = ? LIMIT 1)", oAIMUser.ID, lngGroupID, lngItemID)
    
    If rst.EOF Or rst.Fields(0).Value = 0 Then
        LogError "Server", "Feedbag item does not exist."
        
        ' Return a status code signifying it doesn't exist
        FeedbagUpdateItem = FEEDBAG_STATUS_CODES_NOT_FOUND
    Else
        ' Update it database-side then:
        ExecutePreparedNonQuery _
            "UPDATE `feedbag` " & _
            "SET `name` = ?, `class_id` = ?, `attributes` = ? " & _
            "WHERE `user_id` = ? AND `group_id` = ? AND `item_id` = ?", _
            strName, lngClassID, bytAttributes, oAIMUser.ID, lngGroupID, lngItemID
        
        ' Update feedbag time and total items from the user side
        Call FeedbagUpdateDatabase(oAIMUser)
        
        ' Return a status code signifying it was successful
        FeedbagUpdateItem = FEEDBAG_STATUS_CODES_SUCCESS
    End If
    
    rst.Close
    Set rst = Nothing
End Function

' Deletes a feedbag item for a given user.
Public Function FeedbagDeleteItem(ByVal oAIMUser As clsAIMSession, ByVal strName As String, ByVal lngGroupID As Long, ByVal lngItemID As Long, ByVal lngClassID As Long, ByRef bytAttributes() As Byte) As Long
    Dim rst As ADODB.Recordset
    
    LogInformation "Server", oAIMUser.FormattedScreenName & " is deleting feedbag item " & strName
    LogInformation "Server", "Group ID: 0x" & DecimalToHex(lngGroupID)
    LogInformation "Server", "Item ID: 0x" & DecimalToHex(lngItemID)
    LogInformation "Server", "Class ID: 0x" & DecimalToHex(lngClassID)
    LogInformation "Server", "Attributes: " & BytesToHex(bytAttributes)
    
    ' Ensure it exists:
    Set rst = ExecutePreparedQuery("SELECT EXISTS(SELECT * FROM `feedbag` WHERE `user_id` = ? AND `group_id` = ? AND `item_id` = ? LIMIT 1)", oAIMUser.ID, lngGroupID, lngItemID)
    
    If rst.EOF Or rst.Fields(0).Value = 0 Then
        LogError "Server", "Feedbag item does not exist."
        
        ' Return a status code signifying it doesn't exist
        FeedbagDeleteItem = FEEDBAG_STATUS_CODES_NOT_FOUND
    Else
        ' Delete it database-side then:
        ExecutePreparedNonQuery _
            "DELETE FROM `feedbag` " & _
            "WHERE `user_id` = ? AND `group_id` = ? AND `item_id` = ?", _
            oAIMUser.ID, lngGroupID, lngItemID
        
        ' Update feedbag time and total items from the user side
        Call FeedbagUpdateDatabase(oAIMUser)
        
        ' Return a status code signifying it was successful
        FeedbagDeleteItem = FEEDBAG_STATUS_CODES_SUCCESS
    End If
    
    rst.Close
    Set rst = Nothing
End Function
