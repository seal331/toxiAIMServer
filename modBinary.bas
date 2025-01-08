Attribute VB_Name = "modBinary"
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

Private Declare Function CryptAcquireContext Lib "advapi32.dll" _
    Alias "CryptAcquireContextA" (ByRef hProv As Long, _
    ByVal pszContainer As String, ByVal pszProvider As String, _
    ByVal dwProvType As Long, ByVal dwFlags As Long) As Long

Private Declare Function CryptGenRandom Lib "advapi32.dll" _
    (ByVal hProv As Long, ByVal dwLen As Long, ByRef pbytBuffer As Byte) As Long

Private Declare Function CryptReleaseContext Lib "advapi32.dll" _
    (ByVal hProv As Long, ByVal dwFlags As Long) As Long

Private Const PROV_RSA_FULL = 1
Private Const CRYPT_VERIFYCONTEXT = &HF0000000

' Converts a byte array into a space-separated hexadecimal string.
Public Function ByteArrayToHexString(ByRef bytArray() As Byte) As String
    Dim i As Long
    
    If IsBytesEmpty(bytArray) Then
        Exit Function
    End If
    
    For i = LBound(bytArray) To UBound(bytArray)
        ByteArrayToHexString = ByteArrayToHexString & DecimalToHex(CLng(bytArray(i)))
        
        If i <> GetByteArrayLength(bytArray) - 1 Then
            ByteArrayToHexString = ByteArrayToHexString & " "
        End If
    Next i
End Function

' Converts a hexadecimal string (optionally space-separated) into a byte array.
Public Function HexStringToByteArray(ByVal strHex As String) As Byte()
    Dim bytResult() As Byte
    Dim i As Long
    
    strHex = Replace(strHex, " ", "")   ' Remove any spaces that might be in the input string
    
    If Len(strHex) Mod 2 <> 0 Then
        Err.Raise vbObjectError, "modBinary.HexStringToByteArray", "Invalid hex string length"
    End If
    
    ReDim bytResult(Len(strHex) \ 2 - 1)
    
    For i = 1 To Len(strHex) Step 2
        bytResult((i - 1) \ 2) = HexToDecimal(Mid(strHex, i, 2))
    Next i
    
    HexStringToByteArray = bytResult
End Function

' Copies a segment of one byte array to another.
Public Sub CopyBytes( _
    ByRef bytSource() As Byte, ByVal lngSourceOffset As Long, _
    ByRef bytDest() As Byte, ByVal lngDestOffset As Long, _
    ByVal lngLength As Long)
    
    If lngSourceOffset < 0 Or lngDestOffset < 0 Or lngLength < 0 Then
        Err.Raise vbObjectError, "modBinary.CopyBytes", "Invalid offset or length"
    End If
    
    If lngSourceOffset + lngLength > GetByteArrayLength(bytSource) Or lngDestOffset + lngLength > GetByteArrayLength(bytDest) Then
        Err.Raise vbObjectError, "modBinary.CopyBytes", "Offset and/or length exceed array bounds."
    End If
    
    CopyMemory bytDest(lngDestOffset), bytSource(lngSourceOffset), lngLength
End Sub

' Concatenates multiple byte arrays into one.
Public Function ConcatBytes(ParamArray bytSegments() As Variant) As Byte()
    Dim i As Long
    Dim lngTotal As Long
    Dim lngPos As Long
    Dim lngLength As Long
    Dim bytArray() As Byte
    Dim bytResult() As Byte
    
    ' Calculate the total length of each segment
    For i = LBound(bytSegments) To UBound(bytSegments)
        If Not IsArray(bytSegments(i)) Then
            Err.Raise vbObjectError, "modBinary.ConcatBytes", "All passed arguments must be byte arrays."
            Exit Function
        End If
        
        lngTotal = lngTotal + GetByteArrayLength(bytSegments(i))
    Next i
    
    ' Return an empty byte array if no or empty segments were provided
    If lngTotal <= 0 Then Exit Function
    
    ' Resize result byte array to fit all data segments
    ReDim bytResult(0 To lngTotal - 1)
    
    ' Append each data segment to the result
    For i = LBound(bytSegments) To UBound(bytSegments)
        bytArray = bytSegments(i)
        lngLength = GetByteArrayLength(bytArray)
        
        If lngLength > 0 Then
            CopyBytes bytArray, 0, bytResult, lngPos, lngLength
            lngPos = lngPos + lngLength
        End If
    Next i
    
    ' Return the concatenated byte array
    ConcatBytes = bytResult
End Function

' Compares two byte arrays for equality.
Public Function IsBytesEqual(ByRef bytArrayOne() As Byte, ByRef bytArrayTwo() As Byte) As Boolean
    Dim i As Long
    
    If GetByteArrayLength(bytArrayOne) <> GetByteArrayLength(bytArrayTwo) Then
        IsBytesEqual = False
        Exit Function
    End If
    
    For i = LBound(bytArrayOne) To UBound(bytArrayOne)
        If bytArrayOne(i) <> bytArrayTwo(i) Then
            IsBytesEqual = False
            Exit Function
        End If
    Next i

    IsBytesEqual = True
End Function

' Checks if a byte array is uninitialized or empty.
Public Function IsBytesEmpty(ByRef bytArray() As Byte) As Boolean
    Dim lngUBound As Long
    
    On Error Resume Next

    lngUBound = UBound(bytArray)
    IsBytesEmpty = (Err.Number <> 0)
End Function

' Returns an empty byte array.
Public Function GetEmptyByteArray() As Byte()
    Dim bytArray() As Byte
    
    GetEmptyByteArray = bytArray
End Function

' Calculates the length of a byte array.
Public Function GetByteArrayLength(vntArray) As Long
    On Error Resume Next
    
    GetByteArrayLength = UBound(vntArray) - LBound(vntArray) + 1
End Function

' Generates a byte array with random values of the specified length.
Public Function GenerateRandomBytes(ByVal lngLength As Long) As Byte()
    Dim lngHProv As Long
    Dim lngResult As Long
    Dim bytRand() As Byte
    
    ReDim bytRand(lngLength - 1)
    
    If CryptAcquireContext(lngHProv, vbNullString, vbNullString, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT) = 0 Then
        Err.Raise vbObjectError, "modBinary.GenerateRandomBytes", "Failed to acquire provider context for generating " & lngLength & " bytes"
    End If
    
    If CryptGenRandom(lngHProv, lngLength, bytRand(0)) = 0 Then
        Err.Raise vbObjectError, "modBinary.GenerateRandomBytes", "Failed to generate " & lngLength & " bytes"
    End If
    
    CryptReleaseContext lngHProv, 0
    
    GenerateRandomBytes = bytRand
End Function

' Creates a byte array containing a single byte value.
Public Function SingleByte(ByVal bytValue As Byte) As Byte()
    Dim bytArray(0) As Byte
    bytArray(0) = bytValue
    
    SingleByte = bytArray
End Function

' Converts a 16-bit integer (word) into a 2-byte array in big-endian format.
Public Function Word(ByVal lngValue As Long) As Byte()
    Dim bytArray(1) As Byte
    bytArray(0) = Int(lngValue / &H100) And &HFF
    bytArray(1) = lngValue And &HFF
    
    Word = bytArray
End Function

' Extracts a 16-bit integer (word) in big-endian format from a byte array at the specified offset.
Public Function GetWord(ByRef bytArray() As Byte, Optional lngOffset As Long = 0) As Long
    If lngOffset + 2 > GetByteArrayLength(bytArray) Then
        Err.Raise vbObjectError, "modBinary.GetWord", "Array is too small"
    End If
    
    GetWord = CLng(bytArray(lngOffset + 0)) * &H100 + _
              CLng(bytArray(lngOffset + 1))
End Function

' Converts a 32-bit integer (double word) into a 4-byte array in big-endian format.
Public Function DWord(ByVal dblValue As Double) As Byte()
    Dim dblMSB As Double, dblSecond As Double, dblThird As Double, dblLSB As Double
    
    dblMSB = Int(dblValue / &H1000000) And &HFF
    dblSecond = Int(dblValue / &H10000) And &HFF
    dblThird = Int(dblValue / &H100) And &HFF
    dblLSB = dblValue - (dblMSB * &H1000000 + dblSecond * &H10000 + dblThird * &H100) And &HFF
    
    Dim bytArray(3) As Byte
    bytArray(0) = dblMSB
    bytArray(1) = dblSecond
    bytArray(2) = dblThird
    bytArray(3) = dblLSB
    
    DWord = bytArray
End Function

' Extracts a 32-bit integer (double word) in big-endian format from a byte array at the specified offset.
Public Function GetDWord(ByRef bytArray() As Byte, Optional lngOffset As Long = 0) As Double
    If lngOffset + 4 > GetByteArrayLength(bytArray) Then
        Err.Raise vbObjectError, "modBinary.GetDWord", "Array is too small"
    End If
    
    GetDWord = CDbl(bytArray(lngOffset)) * &H1000000 + _
               CDbl(bytArray(lngOffset + 1)) * &H10000 + _
               CDbl(bytArray(lngOffset + 2)) * &H100 + _
               CDbl(bytArray(lngOffset + 3))
End Function

' Converts a string into a byte array with a 2-byte big-endian length prefix.
Public Function SWord(ByVal strValue As String) As Byte()
    Dim bytArray() As Byte
    ReDim bytArray(2 + Len(strValue) - 1)
    
    CopyBytes Word(Len(strValue)), 0, bytArray, 0, 2
    CopyBytes StringToBytes(strValue), 0, bytArray, 2, Len(strValue)
    
    SWord = bytArray
End Function

' Extracts a string with a 2-byte big-endian length prefix from a byte array at the specified offset.
Public Function GetSWord(ByRef bytArray() As Byte, Optional lngOffset As Long = 0) As String
    Dim bLength(1) As Byte      ' Temporary array to hold the 2-byte length prefix
    Dim lngLength As Long       ' Variable to store the computed string length
    Dim bStrData() As Byte      ' Array to store the extracted string data
    
    If lngOffset + 2 > GetByteArrayLength(bytArray) Then
        Exit Function
    End If
    
    ' Extract the 2-byte length prefix from the byte array into the temporary array
    CopyBytes bytArray, lngOffset, bLength, 0, 2
    
    ' Compute the 2-byte length prefix into a value
    lngLength = GetWord(bLength)
    
    ' Allocate space to hold the string data based on the extracted length
    ReDim bStrData(lngLength - 1)
    
    ' Copy the string data from the byte array (starting after the 2-byte length prefix)
    CopyBytes bytArray, lngOffset + 2, bStrData, 0, lngLength
    
    GetSWord = BytesToString(bStrData)
End Function

' Converts an IPv4 address string into a 4-byte array.
Public Function IPAddressToBytes(ByVal strIPAddress As String) As Byte()
    Dim strOctets() As String
    Dim bytIPAddress(3) As Byte
    Dim i As Long
    
    strOctets = Split(strIPAddress, ".")
    
    If UBound(strOctets) - LBound(strOctets) <> 3 Then
        Err.Raise vbObjectError, "modBinary.IPAddressToBytes", "Invalid IPv4 address format!"
    End If
    
    For i = LBound(strOctets) To UBound(strOctets)
        bytIPAddress(i) = CByte(strOctets(i))
    Next i
    
    IPAddressToBytes = bytIPAddress
End Function

' Converts a UTF-8 string to a byte array.
Public Function StringToBytes(ByVal strData As String) As Byte()
    StringToBytes = StrConv(strData, vbFromUnicode)
End Function

' Converts a byte array to a UTF-8 string.
Public Function BytesToString(ByRef bytData() As Byte) As String
    BytesToString = StrConv(bytData, vbUnicode)
End Function

' Converts a decimal value into a 2-character hexadecimal string.
Public Function DecimalToHex(ByVal lngVal As Long) As String
    DecimalToHex = IIf(lngVal >= 16, Hex(lngVal), "0" & Hex(lngVal))
End Function

' Converts a hexadecimal string into a decimal value.
Public Function HexToDecimal(ByVal strHexVal As String) As Long
    HexToDecimal = Val("&H" & strHexVal)
End Function

