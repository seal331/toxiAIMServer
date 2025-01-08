VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "toxiAIMServer"
   ClientHeight    =   7485
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10380
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   10380
   StartUpPosition =   2  'CenterScreen
   Begin toxiAIMServer.AIMServer BOSServer 
      Left            =   600
      Top             =   6960
      _extentx        =   847
      _extenty        =   847
   End
   Begin toxiAIMServer.AIMServer BUCPServer 
      Left            =   0
      Top             =   6960
      _extentx        =   847
      _extenty        =   847
   End
   Begin TabDlg.SSTab tabMain 
      Height          =   7215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10125
      _ExtentX        =   17859
      _ExtentY        =   12726
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Dashboard"
      TabPicture(0)   =   "frmMain.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraOnlineUsers"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdServerToggle"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraServerLog"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lvwOnlineUsers"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtMessageToBroadcast"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdBroadcastMessage"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Account Management"
      TabPicture(1)   =   "frmMain.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblAcctMgrInDevelopment"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Settings"
      TabPicture(2)   =   "frmMain.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "tabSettings"
      Tab(2).Control(1)=   "cmdApplySettings"
      Tab(2).ControlCount=   2
      Begin VB.CommandButton cmdApplySettings 
         Caption         =   "Apply"
         Height          =   375
         Left            =   -66240
         TabIndex        =   9
         Top             =   6720
         Width           =   1215
      End
      Begin VB.CommandButton cmdBroadcastMessage 
         Caption         =   "Broadcast Message"
         Height          =   315
         Left            =   5160
         TabIndex        =   6
         Top             =   6720
         Width           =   1695
      End
      Begin VB.TextBox txtMessageToBroadcast 
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   6720
         Width           =   4935
      End
      Begin ComctlLib.ListView lvwOnlineUsers 
         Height          =   5895
         Left            =   7080
         TabIndex        =   3
         Top             =   1080
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   10398
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Screen Name"
            Object.Width           =   2205
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "IP Address"
            Object.Width           =   1587
         EndProperty
      End
      Begin VB.Frame fraServerLog 
         Caption         =   "Server Log"
         Height          =   5775
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   6735
         Begin RichTextLib.RichTextBox rtfServerLog 
            Height          =   5415
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   6495
            _ExtentX        =   11456
            _ExtentY        =   9551
            _Version        =   393217
            BorderStyle     =   0
            ReadOnly        =   -1  'True
            ScrollBars      =   2
            TextRTF         =   $"frmMain.frx":0054
         End
      End
      Begin VB.CommandButton cmdServerToggle 
         Caption         =   "Start Server"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
      Begin VB.Frame fraOnlineUsers 
         Caption         =   "Online Users"
         Height          =   6255
         Left            =   6960
         TabIndex        =   4
         Top             =   840
         Width           =   3015
      End
      Begin TabDlg.SSTab tabSettings 
         Height          =   6255
         Left            =   -74880
         TabIndex        =   8
         Top             =   360
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   11033
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "General"
         TabPicture(0)   =   "frmMain.frx":00CF
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fraDbInfo"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "fraConnectionInfo"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Error URLs"
         TabPicture(1)   =   "frmMain.frx":00EB
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fraErrorUrls"
         Tab(1).ControlCount=   1
         Begin VB.Frame fraErrorUrls 
            Caption         =   "Error URLs"
            Height          =   4095
            Left            =   -74880
            TabIndex        =   32
            Top             =   480
            Width           =   6135
            Begin VB.TextBox txtPasswdChangeUrl 
               Height          =   285
               Left            =   240
               TabIndex        =   42
               Text            =   "http://aim.aol.com/password/change_password.adp"
               Top             =   3480
               Width           =   5655
            End
            Begin VB.TextBox txtSuspendedAcctUrl 
               Height          =   285
               Left            =   240
               TabIndex        =   40
               Text            =   "http://www.aim.aol.com/errors/SUSPENDED.html"
               Top             =   2040
               Width           =   5655
            End
            Begin VB.TextBox txtDeletedAcctUrl 
               Height          =   285
               Left            =   240
               TabIndex        =   38
               Text            =   "http://www.aim.aol.com/errors/DELETED_ACCT.html"
               Top             =   2760
               Width           =   5655
            End
            Begin VB.TextBox txtIncorrectPasswdUrl 
               Height          =   285
               Left            =   240
               TabIndex        =   36
               Text            =   "http://www.aim.aol.com/errors/MISMATCH_PASSWD.html"
               Top             =   1320
               Width           =   5655
            End
            Begin VB.TextBox txtUnregisteredAcctUrl 
               Height          =   285
               Left            =   240
               TabIndex        =   33
               Text            =   "http://www.aim.aol.com/errors/UNREGISTERED_SCREENNAME.html"
               Top             =   600
               Width           =   5655
            End
            Begin VB.Label lblPasswdChangeUrl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Password Change URL:"
               Height          =   195
               Left            =   240
               TabIndex        =   41
               Top             =   3240
               Width           =   1680
            End
            Begin VB.Label lblSuspendedAcctUrl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Suspended Account URL:"
               Height          =   195
               Left            =   240
               TabIndex        =   39
               Top             =   1800
               Width           =   1815
            End
            Begin VB.Label lblDeletedAcctUrl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Deleted Account URL:"
               Height          =   195
               Left            =   240
               TabIndex        =   37
               Top             =   2520
               Width           =   1575
            End
            Begin VB.Label lblIncorrectPasswdUrl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Incorrect Password URL:"
               Height          =   195
               Left            =   240
               TabIndex        =   35
               Top             =   1080
               Width           =   1785
            End
            Begin VB.Label lblUnregisteredAcctUrl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Unregistered Account URL:"
               Height          =   195
               Left            =   240
               TabIndex        =   34
               Top             =   360
               Width           =   1950
            End
         End
         Begin VB.Frame fraConnectionInfo 
            Caption         =   "Connection Info"
            Height          =   2655
            Left            =   120
            TabIndex        =   10
            Top             =   480
            Width           =   4695
            Begin VB.TextBox txtDirectoryServerPort 
               Height          =   285
               Left            =   1920
               TabIndex        =   49
               Text            =   "5194"
               Top             =   2160
               Width           =   1095
            End
            Begin VB.TextBox txtAdminServerPort 
               Height          =   285
               Left            =   1920
               TabIndex        =   47
               Text            =   "5193"
               Top             =   1800
               Width           =   1095
            End
            Begin VB.TextBox txtAlertServerPort 
               Height          =   285
               Left            =   1920
               TabIndex        =   45
               Text            =   "5192"
               Top             =   1440
               Width           =   1095
            End
            Begin VB.TextBox txtBosServerPort 
               Height          =   285
               Left            =   1920
               TabIndex        =   16
               Text            =   "5191"
               Top             =   1080
               Width           =   1095
            End
            Begin VB.TextBox txtBucpServerPort 
               Height          =   285
               Left            =   1920
               TabIndex        =   14
               Text            =   "5190"
               Top             =   720
               Width           =   1095
            End
            Begin VB.TextBox txtServerHost 
               Height          =   285
               Left            =   1920
               TabIndex        =   12
               Text            =   "127.0.0.1"
               Top             =   360
               Width           =   2535
            End
            Begin VB.Label lblDirectoryServerPort 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Directory Server Port:"
               Height          =   195
               Left            =   240
               TabIndex        =   48
               Top             =   2190
               Width           =   1590
            End
            Begin VB.Label lblAdminServerPort 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Admin Server Port:"
               Height          =   195
               Left            =   480
               TabIndex        =   46
               Top             =   1830
               Width           =   1365
            End
            Begin VB.Label lblAlertServerPort 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Alert Server Port:"
               Height          =   195
               Left            =   600
               TabIndex        =   44
               Top             =   1470
               Width           =   1275
            End
            Begin VB.Label lblBosServerPort 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "BOS Server Port:"
               Height          =   195
               Left            =   600
               TabIndex        =   15
               Top             =   1110
               Width           =   1230
            End
            Begin VB.Label lblBucpServerPort 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "BUCP Server Port:"
               Height          =   195
               Left            =   480
               TabIndex        =   13
               Top             =   750
               Width           =   1320
            End
            Begin VB.Label lblServerHost 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Server Host:"
               Height          =   195
               Left            =   960
               TabIndex        =   11
               Top             =   390
               Width           =   915
            End
         End
         Begin VB.Frame fraDbInfo 
            Caption         =   "Database Info"
            Height          =   3375
            Left            =   4920
            TabIndex        =   17
            Top             =   480
            Width           =   4815
            Begin VB.Frame fraDbUser 
               Caption         =   "User"
               Height          =   1455
               Left            =   120
               TabIndex        =   22
               Top             =   1800
               Width           =   4575
               Begin VB.TextBox txtDbName 
                  Height          =   285
                  Left            =   1320
                  TabIndex        =   29
                  Top             =   960
                  Width           =   3015
               End
               Begin VB.TextBox txtDbPassword 
                  Height          =   285
                  IMEMode         =   3  'DISABLE
                  Left            =   1320
                  PasswordChar    =   "*"
                  TabIndex        =   28
                  Top             =   600
                  Width           =   3015
               End
               Begin VB.TextBox txtDbUserId 
                  Height          =   285
                  Left            =   1320
                  TabIndex        =   26
                  Top             =   240
                  Width           =   3015
               End
               Begin VB.Label lblDbName 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Database:"
                  Height          =   195
                  Left            =   480
                  TabIndex        =   30
                  Top             =   990
                  Width           =   750
               End
               Begin VB.Label lblDbPassword 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Password:"
                  Height          =   195
                  Left            =   480
                  TabIndex        =   27
                  Top             =   630
                  Width           =   750
               End
               Begin VB.Label lblDbUserId 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "User ID:"
                  Height          =   195
                  Left            =   600
                  TabIndex        =   25
                  Top             =   270
                  Width           =   600
               End
            End
            Begin VB.Frame fraDbConnection 
               Caption         =   "Connection"
               Height          =   1455
               Left            =   120
               TabIndex        =   18
               Top             =   240
               Width           =   4575
               Begin VB.ComboBox cmbDbDriver 
                  Height          =   315
                  Left            =   1320
                  TabIndex        =   31
                  Top             =   240
                  Width           =   3015
               End
               Begin VB.TextBox txtDbPort 
                  Height          =   285
                  Left            =   1320
                  TabIndex        =   24
                  Top             =   960
                  Width           =   1095
               End
               Begin VB.TextBox txtDbHost 
                  Height          =   285
                  Left            =   1320
                  TabIndex        =   23
                  Top             =   600
                  Width           =   3015
               End
               Begin VB.Label lblDbDriver 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Driver:"
                  Height          =   195
                  Left            =   720
                  TabIndex        =   21
                  Top             =   270
                  Width           =   495
               End
               Begin VB.Label lblDbHost 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   " Host:"
                  Height          =   195
                  Left            =   840
                  TabIndex        =   20
                  Top             =   630
                  Width           =   435
               End
               Begin VB.Label lblDbPort 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Port:"
                  Height          =   195
                  Left            =   915
                  TabIndex        =   19
                  Top             =   990
                  Width           =   360
               End
            End
         End
      End
      Begin VB.Label lblAcctMgrInDevelopment 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "This section is currently in development!"
         Height          =   195
         Left            =   -71250
         TabIndex        =   43
         Top             =   3645
         Width           =   2880
      End
   End
   Begin VB.Menu mnuUserActions 
      Caption         =   "User Actions"
      Visible         =   0   'False
      Begin VB.Menu mnuKickUser 
         Caption         =   "Kick"
      End
      Begin VB.Menu mnuSendUserMessage 
         Caption         =   "Send Message"
      End
      Begin VB.Menu mnuUserInfo 
         Caption         =   "Info"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim blnServerToggle As Boolean

''''''''
' BUCP '
''''''''
Private Sub BUCPServer_Connected(ByVal Index As Integer)
    LogInformation "BUCP", BUCPServer.GetIPAddress(Index) & " connected!"
    
    BUCPServer.SendFrame Index, 1, DWord(1)
End Sub

Private Sub BUCPServer_SignOnFrame(ByVal Index As Integer, Data() As Byte)
    If GetByteArrayLength(Data) > 4 Then
        LogDebug "BUCP", "Client is using FLAP-level authentication"
        
        ' TODO(subpurple): implement FLAP-level authentication
    End If
End Sub

Private Sub BUCPServer_DataFrame(ByVal Index As Integer, ByVal Foodgroup As Long, ByVal Subgroup As Long, ByVal Flags As Long, ByVal RequestID As Double, SnacData() As Byte)
    Dim oByteBuffer As clsByteBuffer
    Dim oAIMSession As clsAIMSession
    Dim strChallenge As String
    Dim bytScreenName() As Byte
    Dim strScreenName As String
    Dim enPasswordType As PasswordType
    Dim bytPasswordHash() As Byte
    
    If Foodgroup <> &H17 Then
        LogWarning "BUCP", "Client attempted to access foodgroup outside of BUCP"
        
        BOSServer.SendSNAC Index, Foodgroup, &H1, 0, 0, SnacError(&H4)
        Exit Sub
    End If
    
    Select Case Subgroup
        ' =====================================================================
        Case &H6    ' BUCP__CHALLENGE_REQUEST
        ' =====================================================================
            LogDebug "BUCP", "Recieved BUCP__CHALLENGE_REQUEST"
            
            bytScreenName = GetTLV(&H1, SnacData)
            
            If IsBytesEmpty(bytScreenName) Then
                LogError "BUCP", "Unable to find screen name TLV!"
                
                BUCPServer.SendSNAC Index, Foodgroup, &H1, 0, 0, SnacError(&HE)
                Exit Sub
            End If
            
            strScreenName = BytesToString(bytScreenName)
            strChallenge = RandomChallenge
            
            ' Remove any session that failed to authorize
            For Each oAIMSession In oAIMSessionManager
                If oAIMSession.ScreenName = TrimData(strScreenName) And oAIMSession.SignedOn = False Then
                    LogInformation "BUCP", "Removing ghost session for " & oAIMSession.ScreenName
                    oAIMSessionManager.Remove TrimData(oAIMSession.ScreenName)
                End If
            Next oAIMSession
            
            ' Add a new session with our screen name, IP address and challenge
            oAIMSessionManager.Add strScreenName, BUCPServer.GetIPAddress(Index), Index, strChallenge
            
            LogInformation "BUCP", "Generated challenge: " & strChallenge & " for Screen Name: " & strScreenName
            
            LogDebug "BUCP", "Sending BUCP__CHALLENGE_RESPONSE"
            BUCPServer.SendSNAC Index, Foodgroup, &H7, 0, 0, SWord(strChallenge)
        
        ' =====================================================================
        Case &H2    ' BUCP_LOGIN_REQUEST
        ' =====================================================================
            LogDebug "BUCP", "Recieved BUCP__LOGIN_REQUEST"
            
            bytScreenName = GetTLV(&H1, SnacData)
            bytPasswordHash = GetTLV(&H25, SnacData)
            
            ' Ensure that retrieval of screen name and password hash was successful
            If IsBytesEmpty(bytScreenName) Or IsBytesEmpty(bytPasswordHash) Then
                LogError "BUCP", "Unable to find required TLVs!"
                
                BUCPServer.SendSNAC Index, Foodgroup, &H1, 0, 0, SnacError(&HE)
                Exit Sub
            End If
            
            strScreenName = BytesToString(bytScreenName)
            
            ' Retrieve the previous session
            Set oAIMSession = oAIMSessionManager.Item(TrimData(strScreenName))
            If oAIMSession Is Nothing Then
                LogError "BUCP", "Previous session was not found!"
                Exit Sub
            End If
            
            ' Verify password via the stronger method if TLV 0x4A is present
            If TLVExists(&H4A, SnacData) Then
                enPasswordType = PasswordTypeStrongMD5
            Else
                enPasswordType = PasswordTypeWeakMD5
            End If
            
            ' TODO(subpurple): add state for sn format issues (e.g. zero or under 3 length)
            Select Case CheckLogin(strScreenName, bytPasswordHash, enPasswordType, oAIMSession.Challenge)
            
                Case LoginStateGood
                    LogInformation "BUCP", strScreenName & " was authenticated successfully!"
            
                    ' Generate the cookie used to authorize with BOS
                    oAIMSession.Cookie = RandomCookie
                    
                    ' Set-up the account with details in the database
                    Call SetupAccount(oAIMSession)
                    
                    BUCPServer.SendSNAC Index, &H1, &H3, 0, 0, BucpSuccessReply( _
                        oAIMSession.FormattedScreenName, _
                        oAIMSession.EmailAddress, _
                        oAIMSession.Cookie, _
                        oAIMSession.RegistrationStatus, _
                        AppSettings.Connection.ServerHost & ":" & AppSettings.Connection.BosPort, _
                        AppSettings.ErrorURLs.PasswordChange)
            
                Case LoginStateUnregistered
                    LogError "BUCP", strScreenName & " gave an unregistered screen name."
            
                    BUCPServer.SendSNAC Index, &H1, &H3, 0, 0, BucpErrorReply( _
                        strScreenName, _
                        1, _
                        AppSettings.ErrorURLs.UnregisteredAccount)
                
                Case LoginStateIncorrectPassword
                    LogError "BUCP", strScreenName & " gave a incorrect password."
            
                    BUCPServer.SendSNAC Index, &H1, &H3, 0, 0, BucpErrorReply( _
                        strScreenName, _
                        5, _
                        AppSettings.ErrorURLs.IncorrectPassword)
                
                Case LoginStateSuspended
                    LogError "BUCP", strScreenName & " attempted to sign on but is suspended."
                    
                    BUCPServer.SendSNAC Index, &H1, &H3, 0, 0, BucpErrorReply( _
                        strScreenName, _
                        17, _
                        AppSettings.ErrorURLs.SuspendedAccount)
                
                Case LoginStateDeleted
                    LogError "BUCP", strScreenName & " attempted to sign on but is deleted."
                    
                    BUCPServer.SendSNAC Index, &H1, &H3, 0, 0, BucpErrorReply( _
                        strScreenName, _
                        8, _
                        AppSettings.ErrorURLs.DeletedAccount)
                        
            End Select
            
    End Select
End Sub

Private Sub BUCPServer_SignOffFrame(ByVal Index As Integer)
    BUCPServer.SendFrame Index, 4, GetEmptyByteArray
    BUCPServer.CloseSocket Index
End Sub

Private Sub BUCPServer_Disconnected(ByVal Index As Integer)
    Dim oAIMSession As clsAIMSession
    
    For Each oAIMSession In oAIMSessionManager
        If oAIMSession.Authorized = False And oAIMSession.AuthSocket = Index Then
            LogDebug "BUCP", "Removing unauthorized session for " & oAIMSession.ScreenName
            oAIMSessionManager.Remove TrimData(oAIMSession.ScreenName)
        End If
    Next oAIMSession
End Sub

'''''''
' BOS '
'''''''
Private Sub BOSServer_Connected(ByVal Index As Integer)
    LogInformation "BOS", BUCPServer.GetIPAddress(Index) & " connected!"
    
    BOSServer.SendFrame Index, 1, DWord(1)
End Sub

Private Sub BOSServer_SignOnFrame(ByVal Index As Integer, Data() As Byte)
    Dim bytCookie() As Byte
    Dim oAIMSession As clsAIMSession
    
    If GetByteArrayLength(Data) > 4 Then
        bytCookie = GetTLV(&H6, Data, 4)
        
        For Each oAIMSession In oAIMSessionManager
            If IsBytesEqual(oAIMSession.Cookie, bytCookie) Then
                LogInformation "BOS", "Found session for " & oAIMSession.FormattedScreenName & "!"
                
                ' Assign the index for the session to our session
                oAIMSession.Index = Index
                
                ' Send OSERVICE__HOST_ONLINE with our list of supported foodgroups
                LogDebug "BOS", "Sending OSERVICE__HOST_ONLINE"
                
                BOSServer.SendSNAC Index, &H1, &H3, 0, 0, ServiceHostOnline
                Exit Sub
            End If
        Next oAIMSession
        
        ' Log that we were unable to find the cookie that the client provided
        LogError "BOS", "Unable to find session! Client provided cookie: " & ByteArrayToHexString(bytCookie)
    End If
    
    ' Disconnect the client
    BOSServer.CloseSocket Index
End Sub

Private Sub BOSServer_DataFrame(ByVal Index As Integer, ByVal Foodgroup As Long, ByVal Subgroup As Long, ByVal Flags As Long, ByVal RequestID As Double, SnacData() As Byte)
    Dim oByteReader As New clsByteBuffer
    Dim oByteWriter As New clsByteBuffer
    Dim oTLVList As New clsTLVList
    Dim oAIMSession As clsAIMSession
    Dim oAIMSessionTemp As clsAIMSession
    Dim lngFoodgroup As Long
    Dim strAddress As String
    Dim bytCookie() As Byte
    Dim dblFeedbagTimestamp As Double
    Dim lngFeedbagItems As Long
    Dim strName As String
    Dim lngGroupID As Long
    Dim lngItemID As Long
    Dim lngClassID As Long
    Dim oListItem As ListItem

    ' Retrieve the client's session
    For Each oAIMSessionTemp In oAIMSessionManager
        If oAIMSessionTemp.Index = Index Then
            Set oAIMSession = oAIMSessionTemp
        End If
    Next oAIMSessionTemp
    
    ' Verify we found the session and send a SNAC error signifying to the client
    ' that we haven't logged in if not.
    If oAIMSession Is Nothing Then
        LogError "BOS", "Unable to find session!"
        
        BOSServer.SendSNAC Index, Foodgroup, &H1, 0, 0, SnacError(&H4)
        Exit Sub
    End If
    
    Select Case Foodgroup
        
        ' =====================================================================
        Case &H1                ' OSERVICE__
        ' =====================================================================
            Select Case Subgroup
            
                ' =============================================================
                Case &H2        ' OSERVICE__CLIENT_ONLINE
                ' =============================================================
                    LogDebug "BOS", "Recieved OSERVICE__CLIENT_ONLINE"
                    
                    oByteReader.SetBuffer SnacData
                    
                    With oByteReader
                        Do Until .IsEnd
                            LogVerbose "BOS", _
                                "Foodgroup 0x" & DecimalToHex(.ReadU16) & " / " & _
                                "version 0x" & DecimalToHex(.ReadU16) & " / " & _
                                "tool ID: 0x" & DecimalToHex(.ReadU16) & " / " & _
                                "tool version: 0x" & DecimalToHex(.ReadU16)
                        Loop
                    End With
                
                ' =============================================================
                Case &H4        ' OSERVICE__SERVICE_REQUEST
                ' =============================================================
                    LogDebug "BOS", "Recieved OSERVICE__SERVICE_REQUEST"
                    
                    ' Get the service the client is requesting
                    lngFoodgroup = GetWord(SnacData)
                    
                    ' Log the service and set the address to the appropriate one if found
                    Select Case lngFoodgroup
                        Case &H18
                            LogInformation "BOS", "Client requested ALERT service"
                            strAddress = AppSettings.Connection.ServerHost & ":" & AppSettings.Connection.AlertPort
                            
                        Case &H7
                            LogInformation "BOS", "Client requested ADMIN service"
                            strAddress = AppSettings.Connection.ServerHost & ":" & AppSettings.Connection.AdminPort
                            
                        Case &HF
                           LogInformation "BOS", "Client requested ODIR service"
                           strAddress = AppSettings.Connection.ServerHost & ":" & AppSettings.Connection.DirectoryPort
                            
                        Case Else
                            LogError "BOS", "Client requested unknown service: 0x" & DecimalToHex(lngFoodgroup)
                            
                            BOSServer.SendSNAC Index, &H1, &H1, 0, 0, SnacError(&H6)
                            Exit Sub
                    End Select
                    
                    ' Since I don't currently have OSCAR services implemented, we will always
                    ' send a SNAC error back meaning the service isn't available right now.
                    LogWarning "BOS", "OSCAR services not implemented - sending back SNAC error"
                    BOSServer.SendSNAC Index, &H1, &H1, 0, 0, SnacError(&H5)
                    
                ' =============================================================
                Case &H6        ' OSERVICE__RATE_PARAMS_QUERY
                ' =============================================================
                    LogDebug "BOS", "Recieved OSERVICE__RATE_PARAMS_QUERY"
                    
                    LogDebug "BOS", "Sending OSERVICE__RATE_PARAMS_REPLY"
                    BOSServer.SendSNAC Index, &H1, &H7, 0, 0, ServiceRateParamsReply
                
                ' =============================================================
                Case &H8        ' OSERVICE__RATE_PARAMS_SUB_ADD
                ' =============================================================
                    LogDebug "BOS", "Recieved OSERVICE__RATE_PARAMS_SUB_ADD"
                    
                    ' TODO(subpurple): implement OSCAR rate limits
                    
                ' =============================================================
                Case &HE        ' OSERVICE__USER_INFO_QUERY
                ' =============================================================
                    LogDebug "BOS", "Recieved OSERVICE__USER_INFO_QUERY"
                    
                    LogDebug "BOS", "Sending OSERVICE__USER_INFO_UPDATE"
                    BOSServer.SendSNAC Index, &H1, &HF, 0, 0, ServiceUserInfoUpdate(oAIMSession)
                
                ' =============================================================
                Case &H16       ' OSERVICE__NOOP
                ' =============================================================
                    ' Keep-alive SNAC: ignored
                
                ' =============================================================
                Case &H17       ' OSERVICE__CLIENT_VERSIONS
                ' =============================================================
                    LogDebug "BOS", "Recieved OSERVICE__CLIENT_VERSIONS"
                    
                    LogDebug "BOS", "Sending OSERVICE__HOST_VERSIONS"
                    BOSServer.SendSNAC Index, &H1, &H18, 0, 0, ServiceHostVersions
                
                Case Else
                    LogError "BOS", "Unknown SNAC - foodgroup: OSERVICE, subgroup: 0x" & DecimalToHex(Subgroup)
            
            End Select
        
        ' =====================================================================
        Case &H2                ' LOCATE__
        ' =====================================================================
            Select Case Subgroup
            
                ' =============================================================
                Case &H2        ' LOCATE__RIGHTS_QUERY
                ' =============================================================
                    LogDebug "BOS", "Recieved LOCATE__RIGHTS_QUERY"
                    
                    LogDebug "BOS", "Sending LOCATE__RIGHTS_REPLY"
                    BOSServer.SendSNAC Index, &H2, &H3, 0, 0, LocateRightsReply
                    
                ' =============================================================
                Case &H4        ' LOCATE__SET_INFO
                ' =============================================================
                    LogDebug "BOS", "Recieved LOCATE__SET_INFO"
                    
                    oTLVList.LoadChain SnacData
                    
                    ' TODO(subpurple): in the future, this should be saved in either the database or session
                    If oTLVList.ContainsTLV(&H2) Then   ' User profile
                        LogInformation "BOS", oAIMSession.FormattedScreenName & " set user profile: " & oTLVList.GetTLVAsString(&H2)
                    End If
                    
                    If oTLVList.ContainsTLV(&H4) Then   ' Away message
                        LogInformation "BOS", oAIMSession.FormattedScreenName & " set away message: " & oTLVList.GetTLVAsString(&H4)
                    End If
                    
                    If oTLVList.ContainsTLV(&H5) Then   ' Client capabilities
                        LogInformation "BOS", oAIMSession.FormattedScreenName & " set capabilities"
                    End If
                
                ' =============================================================
                Case &H8        ' LOCATE__WATCHER_NOTIFICATION
                ' =============================================================
                    LogDebug "BOS", "Recieved LOCATE__WATCHER_NOTIFICATION: " & ByteArrayToHexString(SnacData)
                    
                Case Else
                    LogError "BOS", "Unknown SNAC - foodgroup: LOCATE, subgroup: 0x" & DecimalToHex(Subgroup)
                
            End Select
        
        ' =====================================================================
        Case &H3                ' BUDDY__
        ' =====================================================================
            Select Case Subgroup
            
                ' =============================================================
                Case &H2        ' BUDDY__RIGHTS_QUERY
                ' =============================================================
                    LogDebug "BOS", "Recieved BUDDY__RIGHTS_QUERY"
                    
                    LogDebug "BOS", "Sending BUDDY__RIGHTS_REPLY"
                    BOSServer.SendSNAC Index, &H3, &H3, 0, 0, BuddyRightsReply
            
                Case Else
                    LogError "BOS", "Unknown SNAC - foodgroup: BUDDY, subgroup: 0x" & DecimalToHex(Subgroup)
                    
            End Select
        
        ' =====================================================================
        Case &H4                ' ICBM__
        ' =====================================================================
            Select Case Subgroup
                
                ' =============================================================
                Case &H2        ' ICBM__ADD_PARAMETERS
                ' =============================================================
                    LogDebug "BOS", "Recieved ICBM__ADD_PARAMETERS"
                    
                ' =============================================================
                Case &H4        ' ICBM__PARAMETER_QUERY
                ' =============================================================
                    LogDebug "BOS", "Recieved ICBM__PARAMETER_QUERY"
                    
                    LogDebug "BOS", "Sending ICBM__PARAMETER_REPLY"
                    BOSServer.SendSNAC Index, &H4, &H5, 0, 0, IcbmParameterReply
            
                Case Else
                    LogError "BOS", "Unknown SNAC - foodgroup: ICBM, subgroup: 0x" & DecimalToHex(Subgroup)
            
            End Select
        
        ' =====================================================================
        Case &H9                ' BOS__
        ' =====================================================================
            Select Case Subgroup
            
                ' =============================================================
                Case &H2        ' BOS__RIGHTS_QUERY
                ' =============================================================
                    LogDebug "BOS", "Recieved BOS__RIGHTS_QUERY"
                    
                    LogDebug "BOS", "Sending BOS__RIGHTS_REPLY"
                    BOSServer.SendSNAC Index, &H9, &H3, 0, 0, BosRightsReply
                
                Case Else
                    LogError "BOS", "Unknown SNAC - foodgroup: BOS, subgroup: 0x" & DecimalToHex(Subgroup)
            
            End Select
        
        ' =====================================================================
        Case &H13               ' FEEDBAG__
        ' =====================================================================
            Select Case Subgroup
                
                ' =============================================================
                Case &H2        ' FEEDBAG__RIGHTS_QUERY
                ' =============================================================
                    LogDebug "BOS", "Recieved FEEDBAG__RIGHTS_QUERY"
                    
                    LogDebug "BOS", "Sending FEEDBAG__RIGHTS_REPLY"
                    BOSServer.SendSNAC Index, &H13, &H3, 0, 0, FeedbagRightsReply
                
                ' =============================================================
                Case &H4        ' FEEDBAG__QUERY
                ' =============================================================
                    LogDebug "BOS", "Recieved FEEDBAG__QUERY"
                    
                    LogDebug "BOS", "Sending FEEDBAG__REPLY"
                    BOSServer.SendSNAC Index, &H13, &H6, 0, RequestID, FeedbagReply(GetFeedbagData(oAIMSession))
                    
                ' =============================================================
                Case &H5        ' FEEDBAG__QUERY_IF_MODIFIED
                ' =============================================================
                    LogDebug "BOS", "Recieved FEEDBAG__QUERY_IF_MODIFIED"
                    
                    If GetByteArrayLength(SnacData) <> 6 Then
                        LogError "BOS", "Client gave wrong data length to FEEDBAG__QUERY_IF_MODIFIED!"
                        
                        BOSServer.SendSNAC Index, &H13, &H1, 0, RequestID, SnacError(&HE)
                        Exit Sub
                    End If
                    
                    oByteReader.SetBuffer SnacData
                    
                    dblFeedbagTimestamp = oByteReader.ReadU32
                    lngFeedbagItems = oByteReader.ReadU16
                    
                    LogInformation "BOS", "Cached feedbag timestamp: " & Format(ConvertUnixTimestamp(dblFeedbagTimestamp), "mm/dd/yyyy h:mm:ss AM/PM")
                    LogInformation "BOS", "Cached feedbag items: " & lngFeedbagItems
                    
                    If FeedbagCheckIfNew(oAIMSession, dblFeedbagTimestamp, lngFeedbagItems) Then
                        LogDebug "BOS", "Sending FEEDBAG__REPLY"
                        BOSServer.SendSNAC Index, &H13, &H6, 0, RequestID, FeedbagReply(GetFeedbagData(oAIMSession))
                    Else
                        LogDebug "BOS", "Sending FEEDBAG__REPLY_NOT_MODIFIED"
                        BOSServer.SendSNAC Index, &H13, &HF, 0, RequestID, FeedbagReplyNotModified(dblFeedbagTimestamp, lngFeedbagItems)
                    End If
                
                ' =============================================================
                Case &H7            ' FEEBAG__USE
                ' =============================================================
                    LogDebug "BOS", "Recieved FEEDBAG__USE"
                    
                    LogInformation "BOS", oAIMSession.FormattedScreenName & " has recieved their buddy list."
                
                ' =============================================================
                Case &H8            ' FEEDBAG__INSERT_ITEM
                ' =============================================================
                    LogDebug "BOS", "Recieved FEEDBAG__INSERT_ITEM"
                    
                    oByteReader.SetBuffer SnacData
                    
                    With oByteReader
                        Do Until .IsEnd
                            strName = .ReadStringU16
                            lngGroupID = .ReadU16
                            lngItemID = .ReadU16
                            lngClassID = .ReadU16
                            
                            oTLVList.LoadChain .ReadBytes(.ReadU16)
                            
                            oByteWriter.WriteU16 FeedbagAddItem(oAIMSession, strName, lngGroupID, lngItemID, lngClassID, oTLVList)
                        Loop
                    End With
                    
                    LogDebug "BOS", "Sending FEEDBAG__STATUS"
                    BOSServer.SendSNAC Index, &H13, &HE, 0, RequestID, FeedbagStatus(oByteWriter.Buffer)
                    
                ' =============================================================
                Case &H9        ' FEEDBAG__UPDATE_ITEM
                ' =============================================================
                    LogDebug "BOS", "Recieved FEEDBAG__UPDATE_ITEM"
                    
                    oByteReader.SetBuffer SnacData
                    
                    With oByteReader
                        Do Until .IsEnd
                            strName = .ReadStringU16
                            lngGroupID = .ReadU16
                            lngItemID = .ReadU16
                            lngClassID = .ReadU16
                            
                            oTLVList.LoadChain .ReadBytes(.ReadU16)
                            
                            oByteWriter.WriteU16 FeedbagUpdateItem(oAIMSession, strName, lngGroupID, lngItemID, lngClassID, oTLVList)
                        Loop
                    End With
                    
                    LogDebug "BOS", "Sending FEEDBAG__STATUS"
                    BOSServer.SendSNAC Index, &H13, &HE, 0, RequestID, FeedbagStatus(oByteWriter.Buffer)
                    
                ' =============================================================
                Case &HA        ' FEEDBAG__DELETE_ITEM
                ' =============================================================
                    LogDebug "BOS", "Recieved FEEDBAG__DELETE_ITEM"
                    
                    oByteReader.SetBuffer SnacData
                    
                    With oByteReader
                        Do Until .IsEnd
                            strName = .ReadStringU16
                            lngGroupID = .ReadU16
                            lngItemID = .ReadU16
                            lngClassID = .ReadU16
                            
                            oTLVList.LoadChain .ReadBytes(.ReadU16)
                            
                            oByteWriter.WriteU16 FeedbagDeleteItem(oAIMSession, strName, lngGroupID, lngItemID, lngClassID, oTLVList)
                        Loop
                    End With
                    
                    LogDebug "BOS", "Sending FEEDBAG__STATUS"
                    BOSServer.SendSNAC Index, &H13, &HE, 0, RequestID, FeedbagStatus(oByteWriter.Buffer)
                    
                ' =============================================================
                Case &H11       ' FEEDBAG__START_CLUSTER
                ' =============================================================
                    LogInformation "BOS", "Client has started data burst of inserting/updating/deleting feedbag items"
                    
                ' =============================================================
                Case &H12       ' FEEDBAG__END_CLUSTER
                ' =============================================================
                    LogInformation "BOS", "Client has ended data burst of inserting/updating/deleting feedbag items"
                
                Case Else
                    LogError "BOS", "Unknown SNAC - foodgroup: FEEDBAG, subgroup: 0x" & DecimalToHex(Subgroup)
            
            End Select
            
        Case Else
            LogError "BOS", "Unknown SNAC - foodgroup: 0x" & DecimalToHex(Foodgroup) & ", subgroup: 0x" & DecimalToHex(Subgroup)
    End Select
End Sub

Private Sub BOSServer_SignOffFrame(ByVal Index As Integer)
    BOSServer.CloseSocket Index
End Sub

Private Sub BOSServer_Disconnected(ByVal Index As Integer)
    Dim oAIMSession As clsAIMSession
    
    BOSServer.SendFrame Index, 4, GetEmptyByteArray
    
    For Each oAIMSession In oAIMSessionManager
        If oAIMSession.Index = Index Then
            oAIMSession.SignedOn = False
            
            'Call UpdateUserStatus(oAIMSession)
            
            oAIMSessionManager.Remove TrimData(oAIMSession.ScreenName)
            LogInformation "BOS", "Removing session for " & oAIMSession.FormattedScreenName
        End If
    Next oAIMSession
End Sub

Private Sub cmdServerToggle_Click()
    If blnServerToggle = False Then
        StartServer
    Else
        StopServer
    End If
End Sub

Private Sub Form_Load()
    blnServerToggle = False
    
    frmMain.cmdServerToggle.Enabled = InitializeDatabase()
    
    SyncLocalSettings
End Sub

Private Sub Form_Unload(Cancel As Integer)
    TerminateDatabase
End Sub

Private Sub StartServer()
    blnServerToggle = True
    cmdApplySettings.Enabled = False
    cmdServerToggle.Caption = "Stop Server"
    
    LogInformation "BUCP", "Server started on port " & AppSettings.Connection.BucpPort
    BUCPServer.OpenServer AppSettings.Connection.BucpPort
    
    LogInformation "BOS", "Server started on port " & AppSettings.Connection.BosPort
    BOSServer.OpenServer AppSettings.Connection.BosPort
    
    LogError "Alert", "Failed to start: not implemented"
    LogError "Admin", "Failed to start: not implemented"
    LogError "Directory", "Failed to start: not implemented"
End Sub

Private Sub StopServer()
    blnServerToggle = False
    cmdApplySettings.Enabled = True
    cmdServerToggle.Caption = "Start Server"
    
    LogInformation "BUCP", "Server stopped"
    BUCPServer.CloseServer
    
    LogInformation "BOS", "Server stopped"
    BOSServer.CloseServer
End Sub

' TODO(subpurple): this settings code is kind of bad, I should rewrite it
Private Sub SyncLocalSettings()
    With AppSettings
        ' Read connection-related settings:
        txtServerHost.Text = .Connection.ServerHost
        txtBucpServerPort.Text = CStr(.Connection.BucpPort)
        txtBosServerPort.Text = CStr(.Connection.BosPort)
        txtAlertServerPort.Text = CStr(.Connection.AlertPort)
        txtAdminServerPort.Text = CStr(.Connection.AdminPort)
        txtDirectoryServerPort.Text = CStr(.Connection.DirectoryPort)
        
        ' Read database-related settings:
        cmbDbDriver.Text = .Database.Driver
        txtDbHost.Text = .Database.Host
        txtDbPort.Text = .Database.Port
        txtDbUserId.Text = .Database.UserID
        txtDbPassword.Text = .Database.Password
        txtDbName.Text = .Database.Name
        
        ' Read error URL-related settings:
        txtUnregisteredAcctUrl.Text = .ErrorURLs.UnregisteredAccount
        txtIncorrectPasswdUrl.Text = .ErrorURLs.IncorrectPassword
        txtSuspendedAcctUrl.Text = .ErrorURLs.SuspendedAccount
        txtDeletedAcctUrl.Text = .ErrorURLs.DeletedAccount
        txtPasswdChangeUrl.Text = .ErrorURLs.PasswordChange
    End With
End Sub

Private Sub SyncUISettings()
    With AppSettings
        ' Write connection-related settings:
        .Connection.ServerHost = txtServerHost.Text
        .Connection.BucpPort = CLng(txtBucpServerPort.Text)
        .Connection.BosPort = CLng(txtBosServerPort.Text)
        .Connection.AlertPort = CLng(txtAlertServerPort.Text)
        .Connection.AdminPort = CLng(txtAdminServerPort.Text)
        .Connection.DirectoryPort = CLng(txtDirectoryServerPort.Text)
        
        ' Write database-related settings:
        .Database.Driver = cmbDbDriver.Text
        .Database.Host = txtDbHost.Text
        .Database.Port = CLng(txtDbPort.Text)
        .Database.UserID = txtDbUserId.Text
        .Database.Password = txtDbPassword.Text
        .Database.Name = txtDbName.Text
        
        ' Write error URL-related settings:
        .ErrorURLs.UnregisteredAccount = txtUnregisteredAcctUrl.Text
        .ErrorURLs.IncorrectPassword = txtIncorrectPasswdUrl.Text
        .ErrorURLs.SuspendedAccount = txtSuspendedAcctUrl.Text
        .ErrorURLs.DeletedAccount = txtDeletedAcctUrl.Text
        .ErrorURLs.PasswordChange = txtPasswdChangeUrl.Text
    End With
    
    WriteSettings
    
    ' Re-try initializing the database if failed before
    If cmdServerToggle.Enabled = False Then
        cmdServerToggle.Enabled = InitializeDatabase()
    End If
End Sub

Private Sub ValidateSetting(ByVal sFieldName As String, ByVal sFieldText As String, Optional ByVal blnNumerical As Boolean = False)
    If Trim(sFieldText) = "" Then
        Err.Raise vbObjectError, "frmMain.ValidateSetting", "The " & sFieldName & " field must not be blank!"
    ElseIf (blnNumerical And Not IsNumeric(sFieldText)) Then
        Err.Raise vbObjectError, "frmMain.ValidateSetting", "The " & sFieldName & " field must be numerical!"
    End If
End Sub

Private Sub OnlyNumbers(ByRef KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9
        Case vbKeyBack, vbKeyClear, vbKeyDelete
        Case vbKeyLeft, vbKeyRight, vbKeyUp, vbKeyDown, vbKeyTab
        Case Else
            KeyAscii = 0
            Beep
    End Select
End Sub

Private Sub cmdApplySettings_Click()
    On Error GoTo ErrorHandler
    
    ' Validate connection-related settings:
    ValidateSetting "Server Host", txtServerHost.Text
    ValidateSetting "BUCP Server Port", txtBucpServerPort.Text, True
    ValidateSetting "BOS Server Port", txtBosServerPort.Text, True
    ValidateSetting "Alert Server Port", txtAlertServerPort.Text, True
    ValidateSetting "Admin Server Port", txtAdminServerPort.Text, True
    ValidateSetting "Directory Server Port", txtDirectoryServerPort.Text, True
    
    ' Validate database-related settings:
    ValidateSetting "Database Driver", cmbDbDriver.Text
    ValidateSetting "Database Host", txtDbHost.Text
    ValidateSetting "Database Port", txtDbPort.Text, True
    ValidateSetting "Database User ID", txtDbUserId.Text
    ValidateSetting "Database Password", txtDbPassword.Text
    ValidateSetting "Database Name", txtDbName.Text
    
    ' Validate error URL-related settings:
    ValidateSetting "Unregistered Account URL", txtUnregisteredAcctUrl.Text
    ValidateSetting "Incorrect Password URL", txtIncorrectPasswdUrl.Text
    ValidateSetting "Suspended Account URL", txtSuspendedAcctUrl.Text
    ValidateSetting "Deleted Account URL", txtDeletedAcctUrl.Text
    ValidateSetting "Password Change URL", txtPasswdChangeUrl.Text
    
    SyncUISettings
    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub lvwOnlineUsers_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuUserActions
    End If
End Sub

Private Sub txtBosServerPort_KeyPress(KeyAscii As Integer)
    OnlyNumbers KeyAscii
End Sub

Private Sub txtBucpServerPort_KeyPress(KeyAscii As Integer)
    OnlyNumbers KeyAscii
End Sub

Private Sub txtDbPort_KeyPress(KeyAscii As Integer)
    OnlyNumbers KeyAscii
End Sub
