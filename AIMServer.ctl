VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl AIMServer 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin MSWinsockLib.Winsock sckAIMServer 
      Index           =   0
      Left            =   600
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Image imgLogo 
      Height          =   480
      Left            =   0
      Picture         =   "AIMServer.ctx":0000
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "AIMServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event Connected(ByVal Index As Integer, ByVal RemoteHost As String)
Public Event SignOnFrame(ByVal Index As Integer, Data() As Byte)
Public Event DataFrame(ByVal Index As Integer, ByVal Foodgroup As Long, ByVal Subgroup As Long, ByVal Flags As Long, ByVal RequestID As Double, SnacData() As Byte)
Public Event SignOffFrame(ByVal Index As Integer)
Public Event Disconnected(ByVal Index As Integer)

Private Sequence() As Long
Private Buffers() As clsByteBuffer

Public Sub OpenServer(ByVal Port As Integer)
    sckAIMServer(0).Close
    sckAIMServer(0).LocalPort = Port
    sckAIMServer(0).Listen
End Sub

Public Sub CloseServer()
    Dim i As Integer
    For i = 1 To sckAIMServer.UBound
        sckAIMServer(i).Close
        
        RaiseEvent Disconnected(i)
        Unload sckAIMServer(i)
    Next i
    
    ReDim Sequence(0)
    ReDim Buffers(0)
    sckAIMServer(0).Close
End Sub

Public Sub CloseSocket(ByVal Index As Integer)
    Sequence(Index) = 0
    Set Buffers(Index) = Nothing
    
    RaiseEvent Disconnected(Index)
    sckAIMServer(Index).Close
End Sub

Public Function IsConnected(ByVal Index As Integer) As Boolean
    If sckAIMServer(Index).State = sckConnected Then
        IsConnected = True
    Else
        IsConnected = False
    End If
End Function

Public Sub SendFrame(ByVal Index As Integer, ByVal Frame As Byte, ByRef Payload() As Byte)
    If sckAIMServer(Index).State <> sckConnected Then Exit Sub
    
    If Sequence(Index) < 65535 Then
        Sequence(Index) = Sequence(Index) + 1
    Else
        Sequence(Index) = 0
    End If
    
    sckAIMServer(Index).SendData ConcatBytes( _
        SingleByte(&H2A), _
        SingleByte(Frame), _
        Word(Sequence(Index)), _
        Word(GetBytesLength(Payload)), _
        Payload _
    )
End Sub

Public Sub SendSNAC(ByVal Index As Integer, ByVal Foodgroup As Long, ByVal Subgroup As Long, ByVal Flags As Long, ByVal RequestID As Long, ByRef Data() As Byte)
    SendFrame Index, 2, ConcatBytes( _
        Word(Foodgroup), _
        Word(Subgroup), _
        Word(Flags), _
        DWord(RequestID), _
        Data _
    )
End Sub

Public Function GetIPAddress(ByVal Index As Integer) As String
    GetIPAddress = sckAIMServer(Index).RemoteHostIP
End Function

Private Function GetServerName() As String
    GetServerName = Replace(UserControl.Extender.Name, _
        "Server", "")
End Function

Private Function CreateSock() As Integer
    Dim i As Integer
    
    For i = 1 To sckAIMServer.UBound
        If sckAIMServer(i).State <> sckConnected Then
            CreateSock = i
            Exit Function
        End If
    Next i
    
    ReDim Preserve Sequence(0 To UBound(Sequence) + 1)
    ReDim Preserve Buffers(0 To UBound(Buffers) + 1)
    
    CreateSock = sckAIMServer.UBound + 1
    Load sckAIMServer(CreateSock)
End Function

Private Sub sckAIMServer_ConnectionRequest(Index As Integer, ByVal RequestID As Long)
    Dim i As Integer
    i = CreateSock
    
    sckAIMServer(i).Close
    sckAIMServer(i).Accept RequestID
    
    Sequence(i) = 0
    Set Buffers(i) = New clsByteBuffer
    
    RaiseEvent Connected(i, sckAIMServer(i).RemoteHostIP)
End Sub

Private Sub sckAIMServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim oBuffer As clsByteBuffer
    Dim bytPacketSegment() As Byte
    Dim lngOldPosition As Long
    
    Dim bytFrame As Byte
    Dim lngSequence As Long
    Dim lngPayloadLength As Long
    Dim bytPayload() As Byte
    
    sckAIMServer(Index).GetData bytPacketSegment, vbArray + vbByte
    
    Set oBuffer = Buffers(Index)
    
    lngOldPosition = oBuffer.Position
    
    oBuffer.WriteBytes bytPacketSegment
    oBuffer.Position = lngOldPosition

    Do
        ' If the buffer's length is zero, the remote end has closed the
        ' connection.
        If oBuffer.Length = 0 Then Exit Sub
        
        ' If the buffer length is below 6, that means we need to wait
        ' for more data as the full FLAP header has not been recieved.
        If oBuffer.Length < 6 Then Exit Sub
        
        ' Discard the buffer if it doesn't contain the FLAP marker as its
        ' first byte.
        If oBuffer.ReadByte <> &H2A Then
            oBuffer.Clear
            Exit Sub
        End If
        
        With oBuffer
            bytFrame = .ReadByte                ' Read the FLAP frame
            lngSequence = .ReadU16              ' Read the FLAP sequence
            lngPayloadLength = .ReadU16         ' Read the FLAP payload length
        End With
        
        ' Wait for more data if we haven't recieved the amount of data
        ' specified in FLAP's payload length field.
        If oBuffer.Length - 6 < lngPayloadLength Then Exit Sub
        
        ' Read the FLAP payload from the buffer.
        bytPayload = oBuffer.ReadBytes(lngPayloadLength)
        
        ' Route to the correct event depending on the frame.
        Select Case bytFrame
            Case 1
                RaiseEvent SignOnFrame(Index, bytPayload)
            
            Case 2
                ' Send a SNAC error signifying a busted payload if there
                ' isn't enough bytes for the SNAC header.
                If lngPayloadLength < 10 Then
                    LogError GetServerName, sckAIMServer(Index).RemoteHostIP & " gave an invalid SNAC header."
                    
                    SendSNAC Index, &H0, &H1, 0, 0, SnacError(&HE)
                Else
                    ' Route it via event
                    RaiseEvent DataFrame(Index, _
                        GetWord(bytPayload), _
                        GetWord(bytPayload, 2), _
                        GetWord(bytPayload, 4), _
                        GetDWord(bytPayload, 6), _
                        OffsetBytes(bytPayload, 10))
                End If
            
            Case 3
                ' Error frame: ignored
                
            Case 4
                RaiseEvent SignOffFrame(Index)
            
            Case 5
                ' Keep-alive frame: ignored
                
            Case Else
                LogWarning GetServerName, _
                    "Recieved an unknown frame from " & sckAIMServer(Index).RemoteHostIP & ": " & _
                    "0x" & DecimalToHex(bytFrame) & " with payload: " & BytesToHex(bytPayload)
        End Select
    Loop Until oBuffer.IsEnd
    
    oBuffer.Clear
End Sub

Private Sub sckAIMServer_Close(Index As Integer)
    CloseSocket Index
End Sub

Private Sub sckAIMServer_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    CloseSocket Index
End Sub

Private Sub UserControl_Initialize()
    ReDim Sequence(0)
    ReDim Buffers(0)
End Sub

Private Sub UserControl_Resize()
    UserControl.Width = imgLogo.Width
    UserControl.Height = imgLogo.Height
End Sub
