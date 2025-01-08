VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTLVList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_lngCount As Long
Private m_lngLength As Long
Private m_udtTLVs() As TLV

Public Sub Add(ByVal lngType As Long, ByRef bytValue() As Byte)
    ReDim Preserve m_udtTLVs(0 To m_lngCount)
    
    With m_udtTLVs(m_lngCount)
        .Type = lngType
        .Length = GetByteArrayLength(bytValue)
        .Value = bytValue
    End With
    
    m_lngCount = m_lngCount + 1
    m_lngLength = m_lngLength + 4 + GetByteArrayLength(bytValue)
End Sub

Public Sub Clear()
    m_lngCount = 0
    m_lngLength = 0
    Erase m_udtTLVs
End Sub

Public Sub LoadChain(ByRef bytData() As Byte)
    Dim oByteReader As New clsByteBuffer
    
    Clear
    
    With oByteReader
        .SetBuffer bytData
        
        Do Until .IsEnd
            Call Add(.ReadU16, .ReadBytes(.ReadU16))
        Loop
    End With
End Sub

Public Sub LoadBlock(ByRef bytData() As Byte)
    Dim oByteReader As New clsByteBuffer
    Dim i As Long
    
    Clear
    
    With oByteReader
        .SetBuffer bytData
    
        For i = 1 To .ReadU16
            Call Add(.ReadU16, .ReadBytes(.ReadU16))
        Next i
    End With
End Sub

Public Sub LoadLBlock(ByRef bytData() As Byte)
    Dim oByteReader As New clsByteBuffer
    Dim lngLength As Long
    
    Clear
    
    With oByteReader
        .SetBuffer bytData
        
        lngLength = .ReadU16
        
        Do While lngLength - 1 > .Position
            Call Add(.ReadU16, .ReadBytes(.ReadU16))
        Loop
    End With
End Sub

Public Function ContainsTLV(ByVal lngType As Long) As Boolean
    Dim i As Long
    
    For i = 0 To m_lngCount - 1
        If m_udtTLVs(i).Type = lngType Then
            ContainsTLV = True
            Exit Function
        End If
    Next i
    
    ContainsTLV = False
End Function

Public Function GetTLV(ByVal lngType As Long) As Byte()
    Dim i As Long
    
    For i = 0 To m_lngCount - 1
        If m_udtTLVs(i).Type = lngType Then
            GetTLV = m_udtTLVs(i).Value
            Exit Function
        End If
    Next i
End Function

Public Function GetTLVAsString(ByVal lngType As Long) As String
    Dim i As Long
    
    For i = 0 To m_lngCount - 1
        If m_udtTLVs(i).Type = lngType Then
            GetTLVAsString = BytesToString(m_udtTLVs(i).Value)
            Exit Function
        End If
    Next i
End Function

Public Function GetTLVAsWord(ByVal lngType As Long) As Long
    Dim i As Long
    
    For i = 0 To m_lngCount - 1
        If m_udtTLVs(i).Type = lngType Then
            GetTLVAsWord = GetWord(m_udtTLVs(i).Value)
            Exit Function
        End If
    Next i
End Function

Public Function GetTLVAsDWord(ByVal lngType As Long) As Double
    Dim i As Long
    
    For i = 0 To m_lngCount - 1
        If m_udtTLVs(i).Type = lngType Then
            GetTLVAsDWord = GetDWord(m_udtTLVs(i).Value)
            Exit Function
        End If
    Next i
End Function

Public Function GetSerializedChain() As Byte()
    Dim oByteWriter As New clsByteBuffer
    Dim i As Long
    
    For i = 0 To m_lngCount - 1
        With oByteWriter
            .WriteU16 m_udtTLVs(i).Type
            .WriteU16 m_udtTLVs(i).Length
            .WriteBytes m_udtTLVs(i).Value
        End With
    Next i
    
    GetSerializedChain = oByteWriter.Buffer
End Function

Public Function GetSerializedBlock() As Byte()
    Dim oByteWriter As New clsByteBuffer
    oByteWriter.WriteU16 m_lngCount
    oByteWriter.WriteBytes GetSerializedChain
    
    GetSerializedBlock = oByteWriter.Buffer
End Function

Public Function GetSerializedLBlock() As Byte()
    Dim oByteWriter As New clsByteBuffer
    oByteWriter.WriteU16 m_lngLength
    oByteWriter.WriteBytes GetSerializedChain
    
    GetSerializedLBlock = oByteWriter.Buffer
End Function

Public Property Get Count() As Long
    Count = m_lngCount
End Property

Public Property Get Length() As Long
    Length = m_lngLength
End Property

Private Sub Class_Initialize()
    Clear
End Sub
