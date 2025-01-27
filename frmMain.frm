VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "toxiAIMServer"
   ClientHeight    =   7485
   ClientLeft      =   45
   ClientTop       =   540
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
   Begin toxiAIMServer.AIMServer DirectoryServer 
      Left            =   2520
      Top             =   6960
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin toxiAIMServer.AIMServer AdminServer 
      Left            =   1920
      Top             =   6960
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin toxiAIMServer.AIMServer AlertServer 
      Left            =   1320
      Top             =   6960
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin toxiAIMServer.AIMServer BOSServer 
      Left            =   720
      Top             =   6960
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin toxiAIMServer.AIMServer BUCPServer 
      Left            =   120
      Top             =   6960
      _ExtentX        =   847
      _ExtentY        =   847
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
      Tab(0).Control(3)=   "txtMessageToBroadcast"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdBroadcastMessage"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lvwOnlineUsers"
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
      Tab(2).Control(0)=   "cmdApplySettings"
      Tab(2).Control(1)=   "tabSettings"
      Tab(2).ControlCount=   2
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
               Begin VB.ComboBox cboDbDriver 
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

' ========================================
' BUCP Server
' ========================================
Private Sub BUCPServer_Connected(ByVal Index As Integer, ByVal RemoteHost As String)
    LogInformation "BUCP", RemoteHost & " connected!"
    
    BUCPServer.SendFrame Index, 1, DWord(1)
End Sub

Private Sub BUCPServer_SignOnFrame(ByVal Index As Integer, Data() As Byte)
    Dim oAIMUser As clsAIMSession
    Dim bytScreenName() As Byte
    Dim strScreenName As String
    Dim bytRoastedPassword() As Byte
    
    If GetBytesLength(Data) > 4 Then
        LogDebug "BUCP", "Client is using FLAP-level authentication"
        
        bytScreenName = GetTLV(&H1, Data, 4)
        bytRoastedPassword = GetTLV(&H2, Data, 4)
        
        If IsBytesEmpty(bytScreenName) Or IsBytesEmpty(bytRoastedPassword) Then
            ' I'm not sure what the correct behavior is for required TLVs being missing
            ' via FLAP-level authentication.  I'd have to check NINA packet logs.
            Exit Sub
        End If
        
        strScreenName = BytesToString(bytScreenName)
            
        ' Remove any session that failed to authorize
        For Each oAIMUser In oAIMSessionManager
            If oAIMUser.ScreenName = TrimData(strScreenName) And oAIMUser.SignedOn = False Then
                oAIMSessionManager.Remove TrimData(oAIMUser.ScreenName)
            End If
        Next oAIMUser
        
        Select Case CheckLogin(strScreenName, bytRoastedPassword, PasswordTypeXor)
            
            Case LoginStateGood
                LogInformation "BUCP", strScreenName & " was authenticated successfully!"
                
                ' Add the session to the manager
                Set oAIMUser = oAIMSessionManager.Add(strScreenName, BUCPServer.GetIPAddress(Index), Index)
                
                ' Generate the cookie used to authorize with BOS
                oAIMUser.Cookie = RandomCookie
                    
                ' Set-up the account with details in the database
                Call SetupAccount(oAIMUser)
                
                BUCPServer.SendFrame Index, 4, LoginSuccessReply( _
                    oAIMUser.FormattedScreenName, _
                    oAIMUser.EmailAddress, _
                    oAIMUser.Cookie, _
                    oAIMUser.RegistrationStatus, _
                    g_strServerHost & ":" & g_lngBosPort, _
                    g_strPasswordChangeURL)
            
            Case LoginStateUnregistered
                LogError "BUCP", strScreenName & " gave an unregistered screen name."
                
                BUCPServer.SendFrame Index, 4, LoginErrorReply( _
                    strScreenName, _
                    1, _
                    g_strUnregisteredAccountURL)
                
            Case LoginStateIncorrectPassword
                LogError "BUCP", strScreenName & " gave a incorrect password."
                
                BUCPServer.SendFrame Index, 4, LoginErrorReply( _
                    strScreenName, _
                    5, _
                    g_strIncorrectPasswordURL)
                
            Case LoginStateSuspended
                LogError "BUCP", strScreenName & " attempted to sign on but is suspended."
                
                BUCPServer.SendFrame Index, 4, LoginErrorReply( _
                    strScreenName, _
                    17, _
                    g_strSuspendedAccountURL)
                
            Case LoginStateDeleted
                LogError "BUCP", strScreenName & " attempted to sign on but is deleted."
                
                BUCPServer.SendFrame Index, 4, LoginErrorReply( _
                    strScreenName, _
                    8, _
                    g_strDeletedAccountURL)
            
        End Select
    End If
End Sub

Private Sub BUCPServer_DataFrame(ByVal Index As Integer, ByVal Foodgroup As Long, ByVal Subgroup As Long, ByVal Flags As Long, ByVal RequestID As Double, SnacData() As Byte)
    Dim oByteBuffer As clsByteBuffer
    Dim oAIMUser As clsAIMSession
    Dim strChallenge As String
    Dim bytScreenName() As Byte
    Dim strScreenName As String
    Dim enuPasswordType As PasswordType
    Dim bytPasswordHash() As Byte
    
    If Foodgroup <> &H17 Then
        LogError "BUCP", "Unavailable SNAC - foodgroup: 0x" & DecimalToHex(Foodgroup) & ", subgroup: 0x" & DecimalToHex(Subgroup)
        
        'BUCPServer.SendSNAC Index, Foodgroup, &H1, 0, 0, SnacError(&H4)
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
            For Each oAIMUser In oAIMSessionManager
                If oAIMUser.ScreenName = TrimData(strScreenName) And oAIMUser.SignedOn = False Then
                    LogInformation "BUCP", "Removing ghost session for " & oAIMUser.ScreenName
                    oAIMSessionManager.Remove TrimData(oAIMUser.ScreenName)
                End If
            Next oAIMUser
            
            ' Add a new session with our screen name, IP address and challenge
            Set oAIMUser = oAIMSessionManager.Add(strScreenName, BUCPServer.GetIPAddress(Index), Index)
            oAIMUser.Challenge = strChallenge
            
            LogInformation "BUCP", strScreenName & " generated challenge: " & strChallenge
            
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
            Set oAIMUser = oAIMSessionManager.Item(TrimData(strScreenName))
            If oAIMUser Is Nothing Then
                LogError "BUCP", "Previous session was not found!"
                Exit Sub
            End If
            
            ' Verify password via the stronger method if TLV 0x4A is present
            If TLVExists(&H4A, SnacData) Then
                enuPasswordType = PasswordTypeStrongMD5
            Else
                enuPasswordType = PasswordTypeWeakMD5
            End If
            
            ' TODO(subpurple): add state for sn format issues (e.g. zero or under 3 length)
            Select Case CheckLogin(strScreenName, bytPasswordHash, enuPasswordType, oAIMUser.Challenge)
            
                Case LoginStateGood
                    LogInformation "BUCP", strScreenName & " was authenticated successfully!"
            
                    ' Generate the cookie used to authorize with BOS
                    oAIMUser.Cookie = RandomCookie
                    
                    ' Set-up the account with details in the database
                    Call SetupAccount(oAIMUser)
                    
                    BUCPServer.SendSNAC Index, &H1, &H3, 0, 0, LoginSuccessReply( _
                        oAIMUser.FormattedScreenName, _
                        oAIMUser.EmailAddress, _
                        oAIMUser.Cookie, _
                        oAIMUser.RegistrationStatus, _
                        g_strServerHost & ":" & g_lngBosPort, _
                        g_strPasswordChangeURL)
            
                Case LoginStateUnregistered
                    LogError "BUCP", strScreenName & " gave an unregistered screen name."
            
                    BUCPServer.SendSNAC Index, &H1, &H3, 0, 0, LoginErrorReply( _
                        strScreenName, _
                        1, _
                        g_strUnregisteredAccountURL)
                
                Case LoginStateIncorrectPassword
                    LogError "BUCP", strScreenName & " gave a incorrect password."
            
                    BUCPServer.SendSNAC Index, &H1, &H3, 0, 0, LoginErrorReply( _
                        strScreenName, _
                        5, _
                        g_strIncorrectPasswordURL)
                
                Case LoginStateSuspended
                    LogError "BUCP", strScreenName & " attempted to sign on but is suspended."
                    
                    BUCPServer.SendSNAC Index, &H1, &H3, 0, 0, LoginErrorReply( _
                        strScreenName, _
                        17, _
                        g_strSuspendedAccountURL)
                
                Case LoginStateDeleted
                    LogError "BUCP", strScreenName & " attempted to sign on but is deleted."
                    
                    BUCPServer.SendSNAC Index, &H1, &H3, 0, 0, LoginErrorReply( _
                        strScreenName, _
                        8, _
                        g_strDeletedAccountURL)
                        
            End Select
            
    End Select
End Sub

Private Sub BUCPServer_SignOffFrame(ByVal Index As Integer)
    BUCPServer.SendFrame Index, 4, GetEmptyBytes
    BUCPServer.CloseSocket Index
End Sub

Private Sub BUCPServer_Disconnected(ByVal Index As Integer)
    Dim oAIMUser As clsAIMSession
    
    For Each oAIMUser In oAIMSessionManager
        If oAIMUser.Authorized = False And oAIMUser.AuthSocket = Index Then
            LogDebug "BUCP", "Removing unauthorized session for " & oAIMUser.ScreenName
            oAIMSessionManager.Remove TrimData(oAIMUser.ScreenName)
        End If
    Next oAIMUser
End Sub

' ========================================
' BOS Server
' ========================================
Private Sub BOSServer_Connected(ByVal Index As Integer, ByVal RemoteHost As String)
    LogInformation "BOS", RemoteHost & " connected!"
    
    BOSServer.SendFrame Index, 1, DWord(1)
End Sub

Private Sub BOSServer_SignOnFrame(ByVal Index As Integer, Data() As Byte)
    Dim bytCookie() As Byte
    Dim oAIMUser As clsAIMSession
    
    If GetBytesLength(Data) > 4 Then
        bytCookie = GetTLV(&H6, Data, 4)
        
        For Each oAIMUser In oAIMSessionManager
            If IsBytesEqual(oAIMUser.Cookie, bytCookie) Then
                LogInformation "BOS", "Found session for " & oAIMUser.FormattedScreenName & "!"
                
                oAIMUser.Index = Index
                
                LogDebug "BOS", "Sending OSERVICE__HOST_ONLINE"
                
                BOSServer.SendSNAC Index, &H1, &H3, 0, 0, ServiceHostOnline
                Exit Sub
            End If
        Next oAIMUser
        
        LogError "BOS", "Unable to find session! Client provided cookie: " & BytesToHex(bytCookie)
    End If
    
    BOSServer.CloseSocket Index
End Sub

Private Sub BOSServer_DataFrame(ByVal Index As Integer, ByVal Foodgroup As Long, ByVal Subgroup As Long, ByVal Flags As Long, ByVal RequestID As Double, SnacData() As Byte)
    Dim oByteReader As New clsByteBuffer
    Dim oByteWriter As New clsByteBuffer
    Dim oTLVList As New clsTLVList
    Dim oAIMUser As clsAIMSession
    Dim oAIMUserTemp As clsAIMSession

    ' Retrieve the client's session
    For Each oAIMUserTemp In oAIMSessionManager
        If oAIMUserTemp.Index = Index Then
            Set oAIMUser = oAIMUserTemp
        End If
    Next oAIMUserTemp
    
    ' Verify we found the session and send a SNAC error signifying to the client
    ' that we haven't logged in if not.
    If oAIMUser Is Nothing Then
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
                    
                    ' Log that this user has signed on successfully
                    LogInformation "BOS", oAIMUser.FormattedScreenName & " has signed on successfully."
                    
                    ' Set the signed on flag for this session
                    oAIMUser.SignedOn = True
                
                ' =============================================================
                Case &H4        ' OSERVICE__SERVICE_REQUEST
                ' =============================================================
                    Dim lngFoodgroup As Long
                    Dim strAddress As String
                    Dim bytCookie() As Byte
                    
                    LogDebug "BOS", "Recieved OSERVICE__SERVICE_REQUEST"
                    
                    ' Get the service the client is requesting
                    lngFoodgroup = GetWord(SnacData)
                    
                    ' Log the service and set the address to the appropriate one if found
                    Select Case lngFoodgroup
                        Case &H18
                            LogInformation "BOS", "Client requested ALERT service"
                            strAddress = g_strServerHost & ":" & g_lngAlertPort
                            
                        Case &H7
                            LogInformation "BOS", "Client requested ADMIN service"
                            strAddress = g_strServerHost & ":" & g_lngAdminPort
                            
                        Case &HF
                           LogInformation "BOS", "Client requested ODIR service"
                           strAddress = g_strServerHost & ":" & g_lngDirectoryPort
                            
                        Case Else
                            LogError "BOS", "Client requested unknown service: 0x" & DecimalToHex(lngFoodgroup)
                            
                            BOSServer.SendSNAC Index, &H1, &H1, 0, RequestID, SnacError(&H5)
                            Exit Sub
                    End Select
                    
                    ' Generate a cookie
                    bytCookie = RandomCookie
                    
                    ' Add the service to the session with the foodgroup and cookie
                    oAIMUser.AddService lngFoodgroup, bytCookie
                    
                    LogDebug "BOS", "Sending OSERVICE__SERVICE_RESPONSE"
                    BOSServer.SendSNAC Index, &H1, &H5, 0, RequestID, ServiceResponse(lngFoodgroup, strAddress, bytCookie)
                    
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
                    BOSServer.SendSNAC Index, &H1, &HF, 0, 0, ServiceSelfInfo(oAIMUser)
                
                ' =============================================================
                Case &H11       ' OSERVICE__IDLE_NOTIFICATION
                ' =============================================================
                    Dim dblIdleTime As Double
                    
                    LogDebug "BOS", "Recieved OSERVICE__IDLE_NOTIFICATION"
                    
                    dblIdleTime = GetDWord(SnacData)
                    
                    If dblIdleTime > 0 Then
                        LogInformation "BOS", oAIMUser.FormattedScreenName & " has been idle for " & dblIdleTime & " seconds"
                        
                        oAIMUser.Idle = True
                        oAIMUser.IdleTime = Now
                    Else
                        LogInformation "BOS", oAIMUser.FormattedScreenName & " is no longer idle"
                        
                        oAIMUser.Idle = False
                    End If
                    
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
                    Dim bytCapabilities() As Byte
                    
                    LogDebug "BOS", "Recieved LOCATE__SET_INFO"
                    
                    oTLVList.LoadChain SnacData
                    
                    If oTLVList.ContainsTLV(&H1) Then   ' User profile encoding
                        ' TODO(subpurple)
                        LogInformation "BOS", oAIMUser.FormattedScreenName & " set user profile encoding: " & oTLVList.GetTLVAsString(&H1)
                    End If
                    
                    If oTLVList.ContainsTLV(&H2) Then   ' User profile
                        ' TODO(subpurple)
                        LogInformation "BOS", oAIMUser.FormattedScreenName & " set user profile: " & oTLVList.GetTLVAsString(&H2)
                    End If
                    
                    If oTLVList.ContainsTLV(&H3) Then   ' Away message encoding
                        oAIMUser.AwayMessageEncoding = oTLVList.GetTLVAsString(&H3)
                        
                        LogInformation "BOS", oAIMUser.FormattedScreenName & " set away message encoding: " & oTLVList.GetTLVAsString(&H3)
                    End If
                    
                    If oTLVList.ContainsTLV(&H4) Then   ' Away message
                        oAIMUser.AwayMessage = oTLVList.GetTLVAsString(&H4)
                        
                        LogInformation "BOS", oAIMUser.FormattedScreenName & " set away message: " & oTLVList.GetTLVAsString(&H4)
                    End If
                    
                    If oTLVList.ContainsTLV(&H5) Then   ' Client capabilities
                        bytCapabilities = oTLVList.GetTLV(&H5)
                        
                        If GetBytesLength(bytCapabilities) Mod 16 Then
                            LogError "BOS", "Capability list must be an array of 16-byte values!"
                            Exit Sub
                        End If
                        
                        oAIMUser.SetCapabilities bytCapabilities
                        
                        LogInformation "BOS", oAIMUser.FormattedScreenName & " set capabilities."
                    End If
                
                ' =============================================================
                Case &HB        ' LOCATE__GET_DIR_INFO
                ' =============================================================
                    Dim strScreenName As String
                    
                    strScreenName = GetSByte(SnacData)
                    
                    LogDebug "BOS", "Recieved LOCATE__GET_DIR_INFO"
                    LogInformation "BOS", "Getting directory info for " & strScreenName
                    
                    ' TODO(subpurple)
                    ' should use modServer.GetDirectoryInfo(...) for this
                    
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
                    
                    oByteReader.SetBuffer SnacData
                    
                    With oByteReader
                        LogDebug "BOS", "Client set ICBM parameters for channel 0x" & DecimalToHex(.ReadU16)
                        LogDebug "BOS", "ICBM flags: 0x" & DecimalToHex(.ReadU32)
                        LogDebug "BOS", "Maximum incoming ICBM length: " & .ReadU16
                        LogDebug "BOS", "Maximum sender warning level: " & .ReadU16
                        LogDebug "BOS", "Maximum reciever warning level: " & .ReadU16
                        LogDebug "BOS", "Minimum ICBM interval (milliseconds): " & .ReadU32
                    End With
                    
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
                    BOSServer.SendSNAC Index, &H13, &H6, 0, RequestID, FeedbagReply(FeedbagGetTime(oAIMUser), FeedbagGetData(oAIMUser))
                    
                ' =============================================================
                Case &H5        ' FEEDBAG__QUERY_IF_MODIFIED
                ' =============================================================
                    Dim dblFeedbagTimestamp As Double
                    Dim lngFeedbagItems As Long
                    
                    LogDebug "BOS", "Recieved FEEDBAG__QUERY_IF_MODIFIED"
                    
                    If GetBytesLength(SnacData) <> 6 Then
                        LogError "BOS", "Client gave wrong data length to FEEDBAG__QUERY_IF_MODIFIED!"
                        
                        BOSServer.SendSNAC Index, &H13, &H1, 0, RequestID, SnacError(&HE)
                        Exit Sub
                    End If
                    
                    oByteReader.SetBuffer SnacData
                    
                    dblFeedbagTimestamp = oByteReader.ReadU32
                    lngFeedbagItems = oByteReader.ReadU16
                    
                    LogDebug "BOS", "Cached feedbag timestamp: " & Format(ConvertUnixTimestamp(dblFeedbagTimestamp), "mm/dd/yyyy h:mm:ss AM/PM")
                    LogDebug "BOS", "Cached feedbag items: " & lngFeedbagItems
                    
                    If FeedbagIsModified(oAIMUser, dblFeedbagTimestamp, lngFeedbagItems) = True Then
                        LogDebug "BOS", "Sending FEEDBAG__REPLY"
                        BOSServer.SendSNAC Index, &H13, &H6, 0, RequestID, FeedbagReply(FeedbagGetTime(oAIMUser), FeedbagGetData(oAIMUser))
                    Else
                        LogDebug "BOS", "Sending FEEDBAG__REPLY_NOT_MODIFIED"
                        BOSServer.SendSNAC Index, &H13, &HF, 0, RequestID, FeedbagReplyNotModified(dblFeedbagTimestamp, lngFeedbagItems)
                    End If
                
                ' =============================================================
                Case &H7            ' FEEBAG__USE
                ' =============================================================
                    LogDebug "BOS", "Recieved FEEDBAG__USE"
                    
                    LogInformation "BOS", oAIMUser.FormattedScreenName & " has recieved their buddy list."
                
                ' =============================================================
                Case &H8            ' FEEDBAG__INSERT_ITEM
                ' =============================================================
                    LogDebug "BOS", "Recieved FEEDBAG__INSERT_ITEM"
                    
                    oByteReader.SetBuffer SnacData
                    
                    With oByteReader
                        Do Until .IsEnd
                            oByteWriter.WriteU16 FeedbagAddItem(oAIMUser, _
                                .ReadStringU16, _
                                .ReadU16, _
                                .ReadU16, _
                                .ReadU16, _
                                .ReadBytes(.ReadU16))
                        Loop
                    End With
                    
                    LogDebug "BOS", "Sending FEEDBAG__STATUS"
                    BOSServer.SendSNAC Index, &H13, &HE, 0, RequestID, oByteWriter.Buffer
                    
                ' =============================================================
                Case &H9        ' FEEDBAG__UPDATE_ITEM
                ' =============================================================
                    LogDebug "BOS", "Recieved FEEDBAG__UPDATE_ITEM"
                    
                    oByteReader.SetBuffer SnacData
                    
                    With oByteReader
                        Do Until .IsEnd
                            oByteWriter.WriteU16 FeedbagUpdateItem(oAIMUser, _
                                .ReadStringU16, _
                                .ReadU16, _
                                .ReadU16, _
                                .ReadU16, _
                                .ReadBytes(.ReadU16))
                        Loop
                    End With
                    
                    LogDebug "BOS", "Sending FEEDBAG__STATUS"
                    BOSServer.SendSNAC Index, &H13, &HE, 0, RequestID, oByteWriter.Buffer
                    
                ' =============================================================
                Case &HA        ' FEEDBAG__DELETE_ITEM
                ' =============================================================
                    LogDebug "BOS", "Recieved FEEDBAG__DELETE_ITEM"
                    
                    oByteReader.SetBuffer SnacData
                    
                    With oByteReader
                        Do Until .IsEnd
                            oByteWriter.WriteU16 FeedbagDeleteItem(oAIMUser, _
                                .ReadStringU16, _
                                .ReadU16, _
                                .ReadU16, _
                                .ReadU16, _
                                .ReadBytes(.ReadU16))
                        Loop
                    End With
                    
                    LogDebug "BOS", "Sending FEEDBAG__STATUS"
                    BOSServer.SendSNAC Index, &H13, &HE, 0, RequestID, oByteWriter.Buffer
                    
                ' =============================================================
                Case &H11       ' FEEDBAG__START_CLUSTER
                ' =============================================================
                    ' Ignored
                    
                ' =============================================================
                Case &H12       ' FEEDBAG__END_CLUSTER
                ' =============================================================
                    ' Ignored
                    
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
    Dim oAIMUser As clsAIMSession
    
    BOSServer.SendFrame Index, 4, GetEmptyBytes
    
    For Each oAIMUser In oAIMSessionManager
        If oAIMUser.Index = Index Then
            oAIMUser.SignedOn = False
            
            'Call UpdateUserStatus(oAIMUser)
            
            oAIMSessionManager.Remove TrimData(oAIMUser.ScreenName)
            LogInformation "BOS", "Removing session for " & oAIMUser.FormattedScreenName
        End If
    Next oAIMUser
End Sub

' ========================================
' Stubs
' ========================================

Private Sub AlertServer_Connected(ByVal Index As Integer, ByVal RemoteHost As String)
    ' @todo
    LogInformation "Alert", RemoteHost & " connected!"
    
    AlertServer.SendFrame Index, 1, DWord(1)
End Sub

Private Sub AlertServer_SignOnFrame(ByVal Index As Integer, Data() As Byte)
    ' @todo
End Sub

Private Sub AlertServer_DataFrame(ByVal Index As Integer, ByVal Foodgroup As Long, ByVal Subgroup As Long, ByVal Flags As Long, ByVal RequestID As Double, SnacData() As Byte)
    ' @todo
End Sub

Private Sub AlertServer_SignOffFrame(ByVal Index As Integer)
     ' @todo
End Sub

Private Sub AlertServer_Disconnected(ByVal Index As Integer)
     ' @todo
End Sub

' ========================================

Private Sub AdminServer_Connected(ByVal Index As Integer, ByVal RemoteHost As String)
    ' @todo
    LogInformation "Admin", RemoteHost & " connected!"
    
    AdminServer.SendFrame Index, 1, DWord(1)
End Sub

Private Sub AdminServer_SignOnFrame(ByVal Index As Integer, Data() As Byte)
    ' @todo
End Sub

Private Sub AdminServer_DataFrame(ByVal Index As Integer, ByVal Foodgroup As Long, ByVal Subgroup As Long, ByVal Flags As Long, ByVal RequestID As Double, SnacData() As Byte)
    ' @todo
End Sub

Private Sub AdminServer_SignOffFrame(ByVal Index As Integer)
    ' @todo
End Sub

Private Sub AdminServer_Disconnected(ByVal Index As Integer)
    ' @todo
End Sub

' ========================================

Private Sub DirectoryServer_Connected(ByVal Index As Integer, ByVal RemoteHost As String)
    ' @todo
    LogInformation "Directory", RemoteHost & " connected!"
    
    DirectoryServer.SendFrame Index, 1, DWord(1)
End Sub

Private Sub DirectoryServer_SignOnFrame(ByVal Index As Integer, Data() As Byte)
    ' @todo
End Sub

Private Sub DirectoryServer_DataFrame(ByVal Index As Integer, ByVal Foodgroup As Long, ByVal Subgroup As Long, ByVal Flags As Long, ByVal RequestID As Double, SnacData() As Byte)
    ' @todo
End Sub

Private Sub DirectoryServer_SignOffFrame(ByVal Index As Integer)
    ' @todo
End Sub

Private Sub DirectoryServer_Disconnected(ByVal Index As Integer)
    ' @todo
End Sub

' ========================================
' UI
' ========================================
Private Sub StartServer()
    blnServerToggle = True
    cmdApplySettings.Enabled = False
    cmdServerToggle.Caption = "Stop Server"
    
    LogInformation "BUCP", "Server started on port " & g_lngBucpPort
    BUCPServer.OpenServer g_lngBucpPort
    
    LogInformation "BOS", "Server started on port " & g_lngBosPort
    BOSServer.OpenServer g_lngBosPort
    
    LogInformation "Alert", "Server started on port " & g_lngAlertPort
    AlertServer.OpenServer g_lngAlertPort
    
    LogInformation "Admin", "Server started on port " & g_lngAdminPort
    AdminServer.OpenServer g_lngAdminPort
    
    LogInformation "Directory", "Server started on port " & g_lngDirectoryPort
    DirectoryServer.OpenServer g_lngDirectoryPort
End Sub

Private Sub StopServer()
    blnServerToggle = False
    cmdApplySettings.Enabled = True
    cmdServerToggle.Caption = "Start Server"
    
    LogInformation "BUCP", "Server stopped"
    BUCPServer.CloseServer
    
    LogInformation "BOS", "Server stopped"
    BOSServer.CloseServer
    
    LogInformation "Alert", "Server stopped"
    AlertServer.CloseServer
    
    LogInformation "Admin", "Server stopped"
    AdminServer.CloseServer
    
    LogInformation "Directory", "Server stopped"
    DirectoryServer.CloseServer
End Sub

Private Sub Form_Load()
    blnServerToggle = False
    
    frmMain.cmdServerToggle.Enabled = InitializeDatabase()
End Sub

Private Sub Form_Unload(Cancel As Integer)
    TerminateDatabase
End Sub

Private Sub cmdServerToggle_Click()
    If blnServerToggle = False Then
        StartServer
    Else
        StopServer
    End If
End Sub

Private Sub cmdApplySettings_Click()
    WriteSettings
    
    MsgBox "These changes will take effect next time you relaunch the server.", vbInformation
End Sub

Private Sub lvwOnlineUsers_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuUserActions
    End If
End Sub

Private Sub EnsureKeyIsNumerical(ByRef KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9
        Case vbKeyBack, vbKeyClear, vbKeyDelete
        Case vbKeyLeft, vbKeyRight, vbKeyUp, vbKeyDown, vbKeyTab
        Case Else
            KeyAscii = 0
            Beep
    End Select
End Sub

Private Sub txtBosServerPort_KeyPress(KeyAscii As Integer)
    EnsureKeyIsNumerical KeyAscii
End Sub

Private Sub txtBucpServerPort_KeyPress(KeyAscii As Integer)
    EnsureKeyIsNumerical KeyAscii
End Sub

Private Sub txtDbPort_KeyPress(KeyAscii As Integer)
    EnsureKeyIsNumerical KeyAscii
End Sub
