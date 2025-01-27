Attribute VB_Name = "modAIMPackets"
Public Function SnacError(ByVal lngCode As Long, Optional ByVal oTags As clsTLVList) As Byte()
    Dim oByteWriter As New clsByteBuffer
    
    With oByteWriter
        .WriteU16 lngCode
        
        If Not oTags Is Nothing Then
            .WriteBytes oTags.GetSerializedChain
        End If
        
        SnacError = .Buffer
    End With
End Function

Public Function LoginSuccessReply(ByVal strScreenName As String, _
                                  ByVal strEmail As String, _
                                  ByRef bytCookie() As Byte, _
                                  ByVal lngRegistrationStatus As Long, _
                                  ByVal strBOSAddress As String, _
                                  ByVal strChangePasswdURL As String) As Byte()
    Dim oTLVList As New clsTLVList
    
    With oTLVList
        .Add &H5, StringToBytes(strBOSAddress)
        .Add &H6, bytCookie
        .Add &H11, StringToBytes(strEmail)
        .Add &H13, Word(lngRegistrationStatus)
        .Add &H54, StringToBytes(strChangePasswdURL)
        .Add &H8E, SingleByte(0)
        .Add &H1, StringToBytes(strScreenName)
        
        LoginSuccessReply = .GetSerializedChain
    End With
End Function

Public Function LoginErrorReply(ByVal strScreenName As String, _
                               ByVal lngErrorCode As Long, _
                               ByVal strErrorURL As String) As Byte()
    Dim oTLVList As New clsTLVList
    
    With oTLVList
        .Add &H8, Word(lngErrorCode)
        .Add &H4, StringToBytes(strErrorURL)
        .Add &H1, StringToBytes(strScreenName)
    
        LoginErrorReply = .GetSerializedChain
    End With
End Function

Public Function ServiceHostOnline() As Byte()
    Dim oByteWriter As New clsByteBuffer
    
    With oByteWriter
        .WriteU16 &H1       ' OSERVICE
        .WriteU16 &H2       ' LOCATE
        .WriteU16 &H3       ' BUDDY
        .WriteU16 &H4       ' ICBM
        .WriteU16 &H6       ' INVITE
        .WriteU16 &H8       ' POPUP
        .WriteU16 &H9       ' BOS
        .WriteU16 &HA       ' USER_LOOKUP
        .WriteU16 &HB       ' STATS
        .WriteU16 &HC       ' TRANSLATE
        .WriteU16 &H13      ' FEEDBAG
        .WriteU16 &H15      ' ICQ
        .WriteU16 &H22      ' PLUGIN
        .WriteU16 &H24      '
        .WriteU16 &H25      ' MDIR
        
        ServiceHostOnline = .Buffer
    End With
End Function

Public Function ServiceHostVersions() As Byte()
    Dim oByteWriter As New clsByteBuffer
    
    With oByteWriter
        ' OSERVICE
        .WriteU16 &H1
        .WriteU16 4
        
        ' LOCATE
        .WriteU16 &H2
        .WriteU16 1
        
        ' BUDDY
        .WriteU16 &H3
        .WriteU16 1
        
        ' ICBM
        .WriteU16 &H4
        .WriteU16 1
        
        ' INVITE
        .WriteU16 &H6
        .WriteU16 1
        
        ' POPUP
        .WriteU16 &H8
        .WriteU16 1
        
        ' BOS
        .WriteU16 &H9
        .WriteU16 1
        
        ' USER_LOOKUP
        .WriteU16 &HA
        .WriteU16 1
        
        ' STATS
        .WriteU16 &HB
        .WriteU16 1
        
        ' TRANSLATE
        .WriteU16 &HC
        .WriteU16 1
        
        ' FEEDBAG
        .WriteU16 &H13
        .WriteU16 6
        
        ' ICQ
        .WriteU16 &H15
        .WriteU16 2
        
        ' PLUGIN
        .WriteU16 &H22
        .WriteU16 1
        
        '
        .WriteU16 &H24
        .WriteU16 1
        
        ' MDIR
        .WriteU16 &H25
        .WriteU16 1
        
        ServiceHostVersions = .Buffer
    End With
End Function

Public Function ServiceRateParamsReply() As Byte()
    Dim oByteWriter As New clsByteBuffer
    
    ' TODO(subpurple): in the future, this should pull from a `rate_classes`
    ' and `rate_groups` table via MySQL.
    With oByteWriter
        .WriteBytes HexToBytes("00 05 00 01 00 00 00 50 00 00 09 C4 00 00 07 D0")
        .WriteBytes HexToBytes("00 00 05 DC 00 00 03 20 00 00 0D 69 00 00 17 70")
        .WriteBytes HexToBytes("00 00 00 00 00 00 02 00 00 00 50 00 00 0B B8 00")
        .WriteBytes HexToBytes("00 07 D0 00 00 05 DC 00 00 03 E8 00 00 17 70 00")
        .WriteBytes HexToBytes("00 17 70 00 00 F9 0B 00 00 03 00 00 00 14 00 00")
        .WriteBytes HexToBytes("13 EC 00 00 13 88 00 00 0F A0 00 00 0B B8 00 00")
        .WriteBytes HexToBytes("11 47 00 00 17 70 00 00 5C D8 00 00 04 00 00 00")
        .WriteBytes HexToBytes("14 00 00 15 7C 00 00 14 B4 00 00 10 68 00 00 0B")
        .WriteBytes HexToBytes("B8 00 00 17 70 00 00 1F 40 00 00 F9 0B 00 00 05")
        .WriteBytes HexToBytes("00 00 00 0A 00 00 15 7C 00 00 14 B4 00 00 10 68")
        .WriteBytes HexToBytes("00 00 0B B8 00 00 17 70 00 00 1F 40 00 00 F9 0B")
        .WriteBytes HexToBytes("00 00 01 00 91 00 01 00 01 00 01 00 02 00 01 00")
        .WriteBytes HexToBytes("03 00 01 00 04 00 01 00 05 00 01 00 06 00 01 00")
        .WriteBytes HexToBytes("07 00 01 00 08 00 01 00 09 00 01 00 0A 00 01 00")
        .WriteBytes HexToBytes("0B 00 01 00 0C 00 01 00 0D 00 01 00 0E 00 01 00")
        .WriteBytes HexToBytes("0F 00 01 00 10 00 01 00 11 00 01 00 12 00 01 00")
        .WriteBytes HexToBytes("13 00 01 00 14 00 01 00 15 00 01 00 16 00 01 00")
        .WriteBytes HexToBytes("17 00 01 00 18 00 01 00 19 00 01 00 1A 00 01 00")
        .WriteBytes HexToBytes("1B 00 01 00 1C 00 01 00 1D 00 01 00 1E 00 01 00")
        .WriteBytes HexToBytes("1F 00 01 00 20 00 01 00 21 00 02 00 01 00 02 00")
        .WriteBytes HexToBytes("02 00 02 00 03 00 02 00 04 00 02 00 06 00 02 00")
        .WriteBytes HexToBytes("07 00 02 00 08 00 02 00 0A 00 02 00 0C 00 02 00")
        .WriteBytes HexToBytes("0D 00 02 00 0E 00 02 00 0F 00 02 00 10 00 02 00")
        .WriteBytes HexToBytes("11 00 02 00 12 00 02 00 13 00 02 00 14 00 02 00")
        .WriteBytes HexToBytes("15 00 03 00 01 00 03 00 02 00 03 00 03 00 03 00")
        .WriteBytes HexToBytes("06 00 03 00 07 00 03 00 08 00 03 00 09 00 03 00")
        .WriteBytes HexToBytes("0A 00 03 00 0B 00 03 00 0C 00 04 00 01 00 04 00")
        .WriteBytes HexToBytes("02 00 04 00 03 00 04 00 04 00 04 00 05 00 04 00")
        .WriteBytes HexToBytes("07 00 04 00 08 00 04 00 09 00 04 00 0A 00 04 00")
        .WriteBytes HexToBytes("0B 00 04 00 0C 00 04 00 0D 00 04 00 0E 00 04 00")
        .WriteBytes HexToBytes("0F 00 04 00 10 00 04 00 11 00 04 00 12 00 04 00")
        .WriteBytes HexToBytes("13 00 04 00 14 00 06 00 01 00 06 00 02 00 06 00")
        .WriteBytes HexToBytes("03 00 08 00 01 00 08 00 02 00 09 00 01 00 09 00")
        .WriteBytes HexToBytes("02 00 09 00 03 00 09 00 04 00 09 00 09 00 09 00")
        .WriteBytes HexToBytes("0A 00 09 00 0B 00 0A 00 01 00 0A 00 02 00 0A 00")
        .WriteBytes HexToBytes("03 00 0B 00 01 00 0B 00 02 00 0B 00 03 00 0B 00")
        .WriteBytes HexToBytes("04 00 0C 00 01 00 0C 00 02 00 0C 00 03 00 13 00")
        .WriteBytes HexToBytes("01 00 13 00 02 00 13 00 03 00 13 00 04 00 13 00")
        .WriteBytes HexToBytes("05 00 13 00 06 00 13 00 07 00 13 00 08 00 13 00")
        .WriteBytes HexToBytes("09 00 13 00 0A 00 13 00 0B 00 13 00 0C 00 13 00")
        .WriteBytes HexToBytes("0D 00 13 00 0E 00 13 00 0F 00 13 00 10 00 13 00")
        .WriteBytes HexToBytes("11 00 13 00 12 00 13 00 13 00 13 00 14 00 13 00")
        .WriteBytes HexToBytes("15 00 13 00 16 00 13 00 17 00 13 00 18 00 13 00")
        .WriteBytes HexToBytes("19 00 13 00 1A 00 13 00 1B 00 13 00 1C 00 13 00")
        .WriteBytes HexToBytes("1D 00 13 00 1E 00 13 00 1F 00 13 00 20 00 13 00")
        .WriteBytes HexToBytes("21 00 13 00 22 00 13 00 23 00 13 00 24 00 13 00")
        .WriteBytes HexToBytes("25 00 13 00 26 00 13 00 27 00 13 00 28 00 15 00")
        .WriteBytes HexToBytes("01 00 15 00 02 00 15 00 03 00 02 00 06 00 03 00")
        .WriteBytes HexToBytes("04 00 03 00 05 00 09 00 05 00 09 00 06 00 09 00")
        .WriteBytes HexToBytes("07 00 09 00 08 00 03 00 02 00 02 00 05 00 04 00")
        .WriteBytes HexToBytes("06 00 04 00 02 00 02 00 09 00 02 00 0B 00 05 00")
        .WriteBytes HexToBytes("00")
    
        ServiceRateParamsReply = .Buffer
    End With
End Function

Public Function ServiceSelfInfo(ByVal oAIMUser As clsAIMSession) As Byte()
    Dim oByteWriter As New clsByteBuffer
    Dim oTLVList As New clsTLVList
    
    With oByteWriter
        .WriteStringByte oAIMUser.FormattedScreenName
        .WriteU16 oAIMUser.WarningLevel
        
        ' NINA sends TLVs 0x22, 0x28, 0x2D, 0x2C, 0x29 however they are not at all
        ' documented on the wiki and some of the TLV's values are inconsistent
        ' across sessions, so thus I omitted them.
        With oTLVList
            .Add &H15, DWord(oAIMUser.ParentalControls)                     ' Parental controls
            .Add &H1E, DWord(oAIMUser.Subscriptions)                        ' Subscriptions
            .Add &HA, IPAddressToBytes(oAIMUser.IPAddress)                  ' IP address bytes
            .Add &H100A, StringToBytes(oAIMUser.IPAddress)                  ' IP address string
            .Add &H1, Word(oAIMUser.UserClass)                              ' User class
            .Add &H3, DWord(GetUnixTimestamp(oAIMUser.SignOnTime))          ' Sign on time as a UNIX timestamp
            .Add &HF, DWord(CDbl(DateDiff("s", oAIMUser.SignOnTime, Now)))  ' Online time in seconds
            .Add &H5, DWord(GetUnixTimestamp(oAIMUser.RegistrationTime))    ' Account creation time as a UNIX timestamp
        End With
        
        .WriteBytes oTLVList.GetSerializedBlock
        
        ServiceSelfInfo = .Buffer
    End With
End Function

Public Function ServiceUserInfo(ByVal oAIMUser As clsAIMSession) As Byte()
    Dim oByteWriter As New clsByteBuffer
    Dim oTLVList As New clsTLVList
    
    With oByteWriter
        .WriteStringByte oAIMUser.FormattedScreenName
        .WriteU16 oAIMUser.WarningLevel
        
        ' TODO(subpurple): NINA sends TLV the following TLVs that I have not added for one reason
        ' or another:
        '   - 0x1D (BART info)
        '   - 0x18 (alias)
        '   - 0x19 (short capabilities*)
        '   - 0x26 (profile set time)
        '
        ' * Doable, but would require me checking if the client supports short capabilities
        '   and I am far too lazy to do that.
        With oTLVList
            .Add &HD, oAIMUser.Capabilities                                 ' User capabilities
            .Add &H1, Word(oAIMUser.UserClass)                              ' User class
            .Add &H3, DWord(GetUnixTimestamp(oAIMUser.SignOnTime))          ' Sign on time as a UNIX timestamp
            .Add &HF, DWord(CDbl(DateDiff("s", oAIMUser.SignOnTime, Now)))  ' Online time in seconds
            .Add &H5, DWord(GetUnixTimestamp(oAIMUser.RegistrationTime))    ' Account creation time as a UNIX timestamp
        End With
        
        .WriteBytes oTLVList.GetSerializedBlock
        
        ServiceUserInfo = .Buffer
    End With
End Function

Public Function FeedbagReply(ByVal dblFeedbagTimestamp As Double, ByVal colFeedbagItems As Collection) As Byte()
    Dim oByteWriter As New clsByteBuffer
    Dim oFeedbagItem As clsFeedbagItem
    
    With oByteWriter
        .WriteByte 0                                            ' Number of classes in the feedbag (always 0)
        .WriteU16 colFeedbagItems.Count                         ' Number of items in the feedbag
        
        For Each oFeedbagItem In colFeedbagItems
            .WriteStringU16 oFeedbagItem.Name                   ' The item's name as UTF-8 string
            .WriteU16 oFeedbagItem.GroupID                      ' The item's group ID
            .WriteU16 oFeedbagItem.ItemID                       ' The item's ID
            .WriteU16 oFeedbagItem.ClassID                      ' The item's class ID
            .WriteU16 GetBytesLength(oFeedbagItem.Attributes)   ' The item's attributes' total length
            .WriteBytes oFeedbagItem.Attributes                 ' The item's attributes associated with the item (e.g. order)
        Next oFeedbagItem
        
        .WriteU32 dblFeedbagTimestamp
        
        FeedbagReply = .Buffer
    End With
End Function

Public Function FeedbagReplyNotModified(ByVal dblFeedbagTimestamp As Double, ByVal lngFeedbagItems As Long) As Byte()
    Dim oByteWriter As New clsByteBuffer
    
    With oByteWriter
        .WriteU32 dblFeedbagTimestamp
        .WriteU16 lngFeedbagItems
        
        FeedbagReplyNotModified = .Buffer
    End With
End Function

Public Function FeedbagRightsReply() As Byte()
    Dim oByteWriter As New clsByteBuffer
    Dim oMaxClassItemsWriter As New clsByteBuffer
    
    With oByteWriter
        With oMaxClassItemsWriter
            .WriteU16 1000              ' Maximum number of contacts
            .WriteU16 100               ' Maximum number of groups
            .WriteU16 1000              ' Maximum number of visible contacts
            .WriteU16 1000              ' Maximum number of invisible contacts
            .WriteU16 1                 ' Maximum visible/invisible bitmasks
            .WriteU16 1                 ' Maximum presence info fields
            .WriteU16 150               ' Limit for item type 0x06
            .WriteU16 12                ' Limit for item type 0x07
            .WriteU16 12                ' Limit for item type 0x08
            .WriteU16 3                 ' Limit for item type 0x09
            .WriteU16 50                ' Limit for item type 0x0A
            .WriteU16 50                ' Limit for item type 0x0B
            .WriteU16 0                 ' Limit for item type 0x0C
            .WriteU16 128               ' Limit for item type 0x0D
            .WriteU16 1000              ' Maximum number of ignore list entries
            .WriteU16 20                ' Limit for item type 0x0F
            .WriteU16 200               ' Limit for item 10
            .WriteU16 1                 ' Limit for item 11
            .WriteU16 100               ' Limit for item 12
            .WriteU16 1                 ' Limit for item 13
            .WriteU16 25                ' Limit for item 14
            
            ' These values are unknown, but are most likely more limits for specific items
            ' and are here to keep response parity with NINA:
            .WriteBytes HexToBytes( _
                "00 01 00 28 00 01 00 0A 00 C8 00 01 00 3C 00 C8 00 01 00 08" & _
                "00 14 00 01 27 10 03 E8 03 E8 00 32 00 01 00 05 01 F4 00 01" & _
                "00 08 27 10 00 01 00 01 00 01 27 10 00 00 00 00 00 01 07 D0" & _
                "00 00 00 3C 00 18 00 0A 00 01 00 00 00 00 00 00 00 00 00 01" & _
                "00 01 00 01 00 01 03 E8 00 01 00 01")
        End With
        
        .WriteBytes PutTLV(&H2, Word(254))                      ' Maximum class attributes
        .WriteBytes PutTLV(&H3, Word(1698))                     ' Maximum size of all the attributes on a single item
        .WriteBytes PutTLV(&H4, oMaxClassItemsWriter.Buffer)    ' Maximum items by class
        .WriteBytes PutTLV(&H5, Word(100))                      ' Maximum client items
        .WriteBytes PutTLV(&H6, Word(97))                       ' Maximum item name length that the database supports
        .WriteBytes PutTLV(&H7, Word(2000))                     ' Maximum recent buddies
        .WriteBytes PutTLV(&H8, Word(10))                       ' Interaction buddies
        .WriteBytes PutTLV(&H9, DWord(432000))                  ' Interaction half life - in 2^(-age/half_life) in seconds
        .WriteBytes PutTLV(&HA, DWord(14))                      ' Interaction max score
        .WriteBytes PutTLV(&HB, Word(0))                        ' Unknown
        .WriteBytes PutTLV(&HC, Word(600))                      ' Maximum buddies per group
        .WriteBytes PutTLV(&HD, Word(200))                      ' Maximum allowed bot buddies
        .WriteBytes PutTLV(&HE, Word(32))                       ' Maximum smart groups
        
        FeedbagRightsReply = .Buffer
    End With
End Function

Public Function LocateRightsReply() As Byte()
    Dim oTLVList As New clsTLVList
    
    With oTLVList
        .Add &H1, Word(4096)    ' Maximum signature length for this user
        .Add &H2, Word(128)     ' Maximum number of full UUID capabilities allowed
        .Add &H3, Word(30)      ' Maximum number of email addresses to look up at once
        .Add &H4, Word(4096)    ' Maximum CERT length for end to end encryption
        .Add &H5, Word(128)     ' Maximum number of short UUID capabilities allowed
        
        LocateRightsReply = .GetSerializedChain
    End With
End Function

Public Function BuddyRightsReply() As Byte()
    Dim oTLVList As New clsTLVList
    
    With oTLVList
        .Add &H1, Word(1000)    ' Number of usernames the user can have on their Buddy List
        .Add &H2, Word(3000)    ' Number of online users who can simultaneously watch this user
        .Add &H4, Word(160)     ' Number of temporary buddies
        
        BuddyRightsReply = .GetSerializedChain
    End With
End Function

Public Function IcbmParameterReply() As Byte()
    Dim oByteWriter As New clsByteBuffer
    
    ' These are the default server-side preferences.
    '
    ' If the client were to issue a ICBM__ADD_PARAMETERS packet - typically
    ' sent prior to OSERVICE__CLIENT_ONLINE - we should use the specified ones
    ' there instead.
    With oByteWriter
        .WriteU16 5         ' The maximum number of ICBM paramenter slots available
        .WriteU32 &H3       ' Controlling flags
        .WriteU16 512       ' The maximum size of an ICBM the client wants to receive from 80 - 8000
        .WriteU16 900       ' The maximum evil level of the sender when recieving a ICBM from 0 - 999
        .WriteU16 999       ' The maximum evil level of the destination when sending a ICBM from 0 - 999
        .WriteU32 1000      ' How often the client wants to receive ICBMs in milliseconds

        IcbmParameterReply = .Buffer
    End With
End Function

Public Function BosRightsReply() As Byte()
    Dim oTLVList As New clsTLVList
    
    With oTLVList
        .Add &H1, Word(1000)    ' Number of permit entries a user is allowed
        .Add &H2, Word(1000)    ' Number of deny entries a user is allowed
        .Add &H3, Word(1000)    ' Number of temporary permit entries a client is allowed
        
        BosRightsReply = .GetSerializedChain
    End With
End Function

Public Function ServiceResponse(ByVal lngFoodgroup As Long, ByVal strAddress As String, ByRef bytCookie() As Byte) As Byte()
    Dim oTLVList As New clsTLVList
    
    With oTLVList
        .Add &H5, StringToBytes(strAddress) ' Service address
        .Add &H6, bytCookie                 ' Authorization cookie
        .Add &HD, Word(lngFoodgroup)        ' Service type
        .Add &H8E, SingleByte(0)            ' SSL state
        
        ServiceResponse = .GetSerializedChain
    End With
End Function
