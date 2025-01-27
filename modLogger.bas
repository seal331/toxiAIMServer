Attribute VB_Name = "modLogger"
Option Explicit

' The following are additional non-stock colors used in specific log levels:
Public Const vbDarkGreen As Long = 25600
Public Const vbDarkRed As Long = 139

Private Sub LogInternal(ByVal strLevelName As String, ByVal lngLevelColor As Long, ByVal strService As String, ByVal strText As String)
    With frmMain.rtfServerLog
        .SelStart = Len(.Text)
        
        ' Add the current time and date
        .SelText = "[" & Format(Now, "mm/dd/yyyy h:mm:ss AM/PM") & "] ["
        
        ' Add the service text in bold
        .SelBold = True
        .SelText = strService
        .SelBold = False
        
        ' Add a seperator
        .SelText = "] ["
        
        ' Add the level in bold and its specific color
        .SelColor = lngLevelColor
        .SelBold = True
        .SelText = strLevelName
        .SelColor = vbBlack
        
        ' Add the specified text and a new line
        .SelBold = False
        .SelText = "] " & strText & vbCrLf
    End With
End Sub

Public Sub LogVerbose(ByVal strService As String, ByVal strText As String)
    LogInternal "Verbose", vbMagenta, strService, strText
End Sub

Public Sub LogDebug(ByVal strService As String, ByVal strText As String)
    LogInternal "Debug", vbGreen, strService, strText
End Sub

Public Sub LogInformation(ByVal strService As String, ByVal strText As String)
    LogInternal "Information", vbDarkGreen, strService, strText
End Sub

Public Sub LogWarning(ByVal strService As String, ByVal strText As String)
    LogInternal "Warning", vbYellow, strService, strText
End Sub

Public Sub LogError(ByVal strService As String, ByVal strText As String)
    LogInternal "Error", vbRed, strService, strText
End Sub

Public Sub LogFatal(ByVal strService As String, ByVal strText As String)
    LogInternal "Fatal", vbDarkRed, strService, strText
End Sub
