VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsByteBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' A byte array holding the buffer data.
Private m_bytBuffer() As Byte

' Tracks the current position in the buffer.
Private m_lngPosition As Long

' Tracks the current length of the buffer.
Private m_lngLength As Long

Private Sub Class_Initialize()
    Clear
End Sub

' Resets the buffer and its position.
Public Sub Clear()
    Erase m_bytBuffer
    m_lngPosition = 0
    m_lngLength = 0
End Sub

' Reads a single byte from the buffer.
Public Function ReadByte() As Byte
    CheckOverflow 1
    
    ' Return the byte value we read
    ReadByte = m_bytBuffer(m_lngPosition)
    
    ' Advance the position in the buffer by 1 (i.e. the size of a byte)
    m_lngPosition = m_lngPosition + 1
End Function

' Reads an array of bytes from the buffer.
Public Function ReadBytes(ByVal lngLength As Long) As Byte()
    CheckOverflow lngLength
    
    ' Create a new array to hold the bytes being read
    Dim bytResult() As Byte
    
    ' Return uninitialized array if length is 0
    If lngLength <= 0 Then
        ReadBytes = bytResult
        Exit Function
    End If
    
    ' Allocate bytes to fit the specified length
    ReDim bytResult(0 To lngLength - 1)
        
    ' Copy the bytes from the buffer into the result array
    Dim i As Long
    For i = 0 To lngLength - 1
        bytResult(i) = m_bytBuffer(m_lngPosition + i)
    Next i

    ' Advance the position in the buffer by the number of bytes read
    m_lngPosition = m_lngPosition + lngLength
    
    ' Return what we read
    ReadBytes = bytResult
End Function

' Reads a string from the buffer.
Public Function ReadString(ByVal lngLength As Long) As String
    If lngLength = 0 Then
        ReadString = ""
        Exit Function
    End If
    
    ReadString = StrConv(ReadBytes(lngLength), vbUnicode)
End Function

' Reads a 16-bit unsigned integer in big-endian format from the buffer.
Public Function ReadU16() As Long
    CheckOverflow 2
    
    ' Return the uint16 value we read
    ReadU16 = CLng(m_bytBuffer(m_lngPosition + 0)) * &H100 + _
              CLng(m_bytBuffer(m_lngPosition + 1))
    
    ' Advance the the position in the buffer by 2 (i.e. the size of u16)
    m_lngPosition = m_lngPosition + 2
End Function

' Reads a 32-bit unsigned integer in big-endian format from the buffer.
Public Function ReadU32() As Double
    CheckOverflow 4
    
    ' Return the uint32 value we read
    ReadU32 = CDbl(m_bytBuffer(m_lngPosition + 0)) * &H1000000 + _
              CDbl(m_bytBuffer(m_lngPosition + 1)) * &H10000 + _
              CDbl(m_bytBuffer(m_lngPosition + 2)) * &H100 + _
              CDbl(m_bytBuffer(m_lngPosition + 3))
              
    ' Advance the position in the buffer by 4 (i.e. the size of u32)
    m_lngPosition = m_lngPosition + 4
End Function

' Reads a string from the buffer via a byte describing the length
Public Function ReadStringByte() As String
    ReadStringByte = ReadString(ReadByte)
End Function

' Reads a string from the buffer via a 16-bit unsigned integer describing the length
Public Function ReadStringU16() As String
    ReadStringU16 = ReadString(ReadU16)
End Function

' Reads a string from the buffer via a 32-bit unsigned integer describing the length
Public Function ReadStringU32() As String
    ReadStringU32 = ReadString(ReadU32)
End Function

' Writes a single byte to the buffer.
Public Sub WriteByte(ByVal bytValue As Byte)
    EnsureCapacity 1
    
    m_bytBuffer(m_lngPosition) = bytValue
    m_lngPosition = m_lngPosition + 1
End Sub

' Writes an array of bytes to the buffer.
Public Sub WriteBytes(bytValues() As Byte)
    EnsureCapacity GetByteArrayLength(bytValues)
    
    Dim i As Long
    For i = 0 To GetByteArrayLength(bytValues) - 1
        m_bytBuffer(m_lngPosition) = bytValues(i)
        m_lngPosition = m_lngPosition + 1
    Next i
End Sub

' Wries a string to the buffer.
Public Sub WriteString(ByVal strValue As String)
    WriteBytes StrConv(strValue, vbFromUnicode)
End Sub

' Writes a big endian-typed uint16 value (2 bytes) to the buffer at the current position
Public Sub WriteU16(ByVal lngValue As Long)
    If lngValue < 0 Or lngValue > 65535 Then
        Err.Raise vbObjectError, "clsByteBuffer.WriteU16", "Value out of range for U16"
    End If
    
    EnsureCapacity 2
    
    m_bytBuffer(m_lngPosition) = Int(lngValue / &H100) And &HFF
    m_bytBuffer(m_lngPosition + 1) = lngValue And &HFF
    m_lngPosition = m_lngPosition + 2
End Sub

' Writes a big endian-typed uint32 value (4 bytes) to the buffer at the current position
Public Sub WriteU32(ByVal dblValue As Double)
    If dblValue < 0 Or dblValue > 4294967295# Then
        Err.Raise vbObjectError, "clsByteBuffer.WriteU32", "Value out of range for U32"
    End If
    
    EnsureCapacity 4
    
    Dim dblMSB As Double
    Dim dblSecond As Double
    Dim dblThird As Double
    Dim dblLSB As Double
    
    dblMSB = Int(dblValue / &H1000000) And &HFF
    dblSecond = Int(dblValue / &H10000) And &HFF
    dblThird = Int(dblValue / &H100) And &HFF
    dblLSB = dblValue - (dblMSB * &H1000000 _
                    + dblSecond * &H10000 _
                    + dblThird * &H100) And &HFF
    
    m_bytBuffer(m_lngPosition) = dblMSB
    m_bytBuffer(m_lngPosition + 1) = dblSecond
    m_bytBuffer(m_lngPosition + 2) = dblThird
    m_bytBuffer(m_lngPosition + 3) = dblLSB
    
    m_lngPosition = m_lngPosition + 4
End Sub

' Writes a string to the buffer prefixed with a byte describing the length
Public Sub WriteStringByte(ByVal strValue As String)
    WriteByte Len(strValue)
    WriteString strValue
End Sub

' Writes a string to the buffer prefixed with a 16-bit unsigned integer describing the length
Public Sub WriteStringU16(ByVal strValue As String)
    WriteU16 Len(strValue)
    WriteString strValue
End Sub

' Writes a string to the buffer prefixed with a 32-bit unsigned integer describing the length
Public Sub WriteStringU32(ByVal strValue As String)
    WriteU32 Len(strValue)
    WriteString strValue
End Sub

' Returns whether the buffer is at the end or not.
Public Function IsEnd() As Boolean
    If m_lngPosition >= m_lngLength Then
        IsEnd = True
    End If
End Function

' Ensures the read position does not exceed the buffer length.
Private Sub CheckOverflow(ByVal lngBytesToRead As Double)
    If m_lngPosition + lngBytesToRead > m_lngLength Then
        Err.Raise vbObjectError, "clsByteBuffer.CheckOverflow", "Buffer is too small"
    End If
End Sub

' Ensures the buffer has enough space before writing.
Private Sub EnsureCapacity(ByVal lngBytesToWrite As Double)
    ' Check if advancing the position by `lngBytesToWrite` would exceed the
    ' buffer's length.
    If m_lngPosition + lngBytesToWrite > m_lngLength Then
        ' Increase the buffer's length by `lngBytesToWrite`
        m_lngLength = m_lngPosition + lngBytesToWrite
        
        ' Resize the array to fit the new length
        ReDim Preserve m_bytBuffer(m_lngLength - 1)
    End If
End Sub

' Provides access to the underlying byte array.
Public Property Get Buffer() As Byte()
    Buffer = m_bytBuffer
End Property

' Sets a new byte array as the buffer.
Public Sub SetBuffer(bytBuffer() As Byte)
    m_bytBuffer = bytBuffer
    m_lngPosition = 0
    m_lngLength = GetByteArrayLength(bytBuffer)
End Sub

' Gets the current position within the buffer.
Public Property Get Position() As Long
    Position = m_lngPosition
End Property

' Sets the current position within the buffer.
Public Property Let Position(ByVal lngPos As Long)
    ' Ensure the given position doesn't exceed the bounds of the buffer
    If lngPos < 0 Or lngPos > m_lngLength Then
        Err.Raise vbObjectError, "clsByteBuffer.Position", "Invalid position"
    End If
    
    m_lngPosition = lngPos
End Property

' Gets the total length of the buffer.
Public Property Get Length() As Double
    Length = m_lngLength
End Property

' Sets the total length of the buffer.
Public Property Let Length(ByVal lngLength As Double)
    If lngLength < 0 Then
        Err.Raise vbObjectError, "clsByteBuffer.Length", "Invalid length"
    End If
    
    If m_lngPosition > lngLength Then
        m_lngPosition = lngLength
    End If
    
    m_lngLength = lngLength
    ReDim Preserve m_bytBuffer(m_lngLength)
End Property
