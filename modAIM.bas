Attribute VB_Name = "modAIM"
Option Explicit

Public Type TLV
    Type As Long
    Length As Long
    Value() As Byte
End Type

' TODO(subpurple): possibly could turn this into `Public Const ...` statements into a module?
Public Enum UserFlags
    UserFlagUnconfirmed = &H1
    UserFlagAdministrator = &H2
    UserFlagAOL = &H4
    UserFlagOscarPay = &H8
    UserFlagOscarFree = &H10
    UserFlagUnavailable = &H20
    UserFlagICQ = &H40
    UserFlagWireless = &H80
    UserFlagInternal = &H100
    UserFlagFish = &H200
    UserFlagBot = &H400
    UserFlagBeast = &H800
End Enum

' Gets a random integer between the specified lower bound and upper bound.
Private Function RandomInteger(ByVal intLowerBound As Integer, ByVal intUpperBound As Integer) As Integer
    Randomize
    RandomInteger = Int((intUpperBound - intLowerBound + 1) * Rnd + intLowerBound)
End Function

'
Public Function RandomChallenge() As String
    Dim i As Integer
    For i = 1 To 10
        RandomChallenge = RandomChallenge & RandomInteger(0, 9)
    Next i
End Function

'
Public Function RandomCookie() As Byte()
    RandomCookie = GenerateRandomBytes(256)
End Function

'
Public Function GetTLV(ByVal lType As Long, ByRef bytData() As Byte, Optional ByVal lngOffset As Long = 0) As Byte()
    ' Set error handler as we may error out here attempting to read type / length values
    On Error GoTo ErrorHandler
    
    Dim bytTlvData() As Byte
    Dim i As Long, lngType As Long, lngLength As Long
    
    ' Offset our counter by what was passed
    i = i + lngOffset
    
    ' Loop unil we're at the end of the given data
    Do While i < GetByteArrayLength(bytData) - 1
        lngType = GetWord(bytData, i)
        lngLength = GetWord(bytData, i + 2)
        
        If lngType = lType Then
            ReDim bytTlvData(lngLength - 1)
            CopyBytes bytData, i + 4, bytTlvData, 0, lngLength
            
            GetTLV = bytTlvData
            Exit Function
        End If
        
        i = i + 4 + lngLength
    Loop
    
ErrorHandler:
    Erase GetTLV
End Function

'
Public Function TLVExists(ByVal lType As Long, ByRef bytData() As Byte, Optional ByVal lngOffset As Long = 0) As Boolean
    ' Set error handler as we may error out here attempting to read type / length values
    On Error GoTo ErrorHandler
    
    Dim i As Long, lngType As Long, lngLength As Long
    
    ' Offset our counter by what was passed
    i = i + lngOffset
    
    ' Loop unil we're at the end of the given data
    Do While i < GetByteArrayLength(bytData) - 1
        lngType = GetWord(bytData, i)
        lngLength = GetWord(bytData, i + 2)
        
        If lngType = lType Then
            TLVExists = True
            Exit Function
        End If
        
        i = i + 4 + lngLength
    Loop
    
ErrorHandler:
    TLVExists = False
End Function

' TODO(subpurple): make this support uninitialized arrays
Public Function PutTLV(ByVal lngType As Long, ByRef bytData() As Byte) As Byte()
    Dim bytArray() As Byte, lngLength As Long
    
    ' Compute the given byte array's length
    lngLength = GetByteArrayLength(bytData)
    
    ' Resize our byte array to fit the type, length of the byte array
    ' and the byte array itself
    ReDim bytArray(4 + lngLength - 1)
    
    CopyBytes Word(lngType), 0, bytArray, 0, 2      ' Add the TLV's type
    CopyBytes Word(lngLength), 0, bytArray, 2, 2    ' Add the TLV length
    CopyBytes bytData, 0, bytArray, 4, lngLength    ' Add the TLV data
    
    PutTLV = bytArray
End Function
