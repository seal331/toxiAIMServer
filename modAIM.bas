Attribute VB_Name = "modAIM"
Option Explicit

Public Type TLV
    Type As Long
    Length As Long
    Value() As Byte
End Type

' Gets a random integer between the specified lower bound and upper bound.
Public Function RandomInteger(ByVal intLowerBound As Integer, ByVal intUpperBound As Integer) As Integer
    Randomize
    RandomInteger = Int((intUpperBound - intLowerBound + 1) * Rnd + intLowerBound)
End Function

' Generates a random challenge consisting of 10 numbers from 0-9.
Public Function RandomChallenge() As String
    Dim i As Integer
    For i = 1 To 10
        RandomChallenge = RandomChallenge & RandomInteger(0, 9)
    Next i
End Function

' Generates a random cookie consisting of 256 truly random bytes.
Public Function RandomCookie() As Byte()
    RandomCookie = GenerateRandomBytes(256)
End Function

' Extracts a TLV's data, given the specified type, from a byte array.
Public Function GetTLV(ByVal lngType As Long, ByRef bytData() As Byte, Optional ByVal lngOffset As Long = 0) As Byte()
    ' Set error handler as we may error out here attempting to read type / length values
    On Error GoTo ErrorHandler
    
    Dim bytTLVData() As Byte
    Dim i As Long, lngTypeIter As Long, lngLength As Long
    
    ' Offset our counter by what was passed
    i = i + lngOffset
    
    ' Loop unil we're at the end of the given data
    Do While i < GetBytesLength(bytData) - 1
        lngTypeIter = GetWord(bytData, i)
        lngLength = GetWord(bytData, i + 2)
        
        If lngType = lngTypeIter Then
            ReDim bytTLVData(lngLength - 1)
            CopyBytes bytData, i + 4, bytTLVData, 0, lngLength
            
            GetTLV = bytTLVData
            Exit Function
        End If
        
        i = i + 4 + lngLength
    Loop
    
ErrorHandler:
    Erase GetTLV
End Function

' Determines whether a TLV exists in a given byte array.
Public Function TLVExists(ByVal lngType As Long, ByRef bytData() As Byte, Optional ByVal lngOffset As Long = 0) As Boolean
    ' Set error handler as we may error out here attempting to read type / length values
    On Error GoTo ErrorHandler
    
    Dim i As Long, lngTypeIter As Long, lngLength As Long
    
    ' Offset our counter by what was passed
    i = i + lngOffset
    
    ' Loop unil we're at the end of the given data
    Do While i < GetBytesLength(bytData) - 1
        lngTypeIter = GetWord(bytData, i)
        lngLength = GetWord(bytData, i + 2)
        
        If lngType = lngTypeIter Then
            TLVExists = True
            Exit Function
        End If
        
        i = i + 4 + lngLength
    Loop
    
ErrorHandler:
    TLVExists = False
End Function

' Constructs TLV bytes with the provided type and data.
Public Function PutTLV(ByVal lngType As Long, ByRef bytData() As Byte) As Byte()
    PutTLV = ConcatBytes( _
        Word(lngType), _
        Word(GetBytesLength(bytData)), _
        bytData)
End Function
