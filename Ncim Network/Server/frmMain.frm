VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ncim Network XP- Server"
   ClientHeight    =   7005
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6060
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   467
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   404
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   6060
      TabIndex        =   13
      Top             =   0
      Width           =   6060
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome to the Ncim Network Server"
         Height          =   195
         Left            =   210
         TabIndex        =   15
         Top             =   330
         Width           =   2685
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ncim Network XP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   210
         TabIndex        =   14
         Top             =   135
         Width           =   1500
      End
   End
   Begin VB.ListBox lstUnKickable 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Left            =   630
      TabIndex        =   2
      Top             =   7410
      Width           =   2010
   End
   Begin MSWinsockLib.Winsock wskServer 
      Index           =   0
      Left            =   3180
      Top             =   7770
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSysTray 
      Caption         =   "System Tray"
      Height          =   345
      Left            =   3630
      TabIndex        =   1
      Top             =   6555
      Width           =   1125
   End
   Begin VB.CommandButton cmdListen 
      Caption         =   "&Start Server"
      Height          =   345
      Left            =   4845
      TabIndex        =   0
      Top             =   6555
      Width           =   1125
   End
   Begin VB.PictureBox picTab 
      BorderStyle     =   0  'None
      Height          =   5400
      Index           =   0
      Left            =   105
      Picture         =   "frmMain.frx":628A
      ScaleHeight     =   360
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   388
      TabIndex        =   3
      Top             =   1020
      Visible         =   0   'False
      Width           =   5820
      Begin VB.ListBox lstUsers 
         Height          =   3090
         IntegralHeight  =   0   'False
         ItemData        =   "frmMain.frx":6C7AC
         Left            =   480
         List            =   "frmMain.frx":6C7AE
         TabIndex        =   7
         Top             =   450
         Width           =   4995
      End
      Begin VB.Label lblfra01 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Server Information"
         ForeColor       =   &H00D54600&
         Height          =   195
         Left            =   375
         TabIndex        =   12
         Top             =   3840
         Width           =   1290
      End
      Begin VB.Label lblfra02 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User(s) Online"
         ForeColor       =   &H00D54600&
         Height          =   195
         Left            =   375
         TabIndex        =   11
         Top             =   210
         Width           =   990
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Server Status:"
         Height          =   195
         Left            =   480
         TabIndex        =   10
         Top             =   4800
         Width           =   1005
      End
      Begin VB.Label lblIpPort 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Server IP + Port:"
         Height          =   195
         Left            =   480
         TabIndex        =   9
         Top             =   4440
         Width           =   1170
      End
      Begin VB.Label lblCount 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Count:"
         Height          =   195
         Left            =   480
         TabIndex        =   8
         Top             =   4080
         Width           =   840
      End
   End
   Begin VB.PictureBox picTab 
      BorderStyle     =   0  'None
      Height          =   5400
      Index           =   2
      Left            =   105
      Picture         =   "frmMain.frx":6C7B0
      ScaleHeight     =   360
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   388
      TabIndex        =   5
      Top             =   1020
      Visible         =   0   'False
      Width           =   5820
      Begin VB.TextBox txtPort 
         Alignment       =   1  'Right Justify
         Height          =   295
         Left            =   480
         MaxLength       =   4
         TabIndex        =   22
         Text            =   "0"
         Top             =   705
         Width           =   480
      End
      Begin VB.Label lblRestart 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To restart Ncim Network server, Click Here."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000015&
         Height          =   195
         Left            =   480
         MouseIcon       =   "frmMain.frx":70EFD
         MousePointer    =   99  'Custom
         TabIndex        =   24
         ToolTipText     =   "Click here to visit the Ncim Network homepage."
         Top             =   1710
         Width           =   3075
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Your change will take effect next time you start Ncim Network Server."
         Height          =   195
         Left            =   480
         TabIndex        =   23
         Top             =   1515
         Width           =   4905
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblfraServerSet 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Server Settings"
         ForeColor       =   &H00D54600&
         Height          =   195
         Left            =   375
         TabIndex        =   21
         Top             =   210
         Width           =   1080
      End
      Begin VB.Label lblPortText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "This is the port that the server will listen on for connections."
         Height          =   195
         Left            =   480
         TabIndex        =   20
         Top             =   450
         Width           =   4830
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox picTab 
      BorderStyle     =   0  'None
      Height          =   5400
      Index           =   3
      Left            =   105
      Picture         =   "frmMain.frx":7104F
      ScaleHeight     =   360
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   388
      TabIndex        =   6
      Top             =   1020
      Visible         =   0   'False
      Width           =   5820
      Begin VB.CheckBox chkSplash 
         Height          =   195
         Left            =   810
         TabIndex        =   32
         Top             =   2625
         Width           =   195
      End
      Begin VB.CheckBox chkSS 
         Height          =   195
         Left            =   810
         TabIndex        =   31
         Top             =   2295
         Width           =   195
      End
      Begin VB.CheckBox chkSysTray 
         Height          =   195
         Left            =   810
         TabIndex        =   27
         Top             =   1005
         Width           =   195
      End
      Begin VB.Label lblSplash 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Show splash screen on start up"
         Height          =   195
         Left            =   1065
         TabIndex        =   34
         Top             =   2625
         Width           =   2220
      End
      Begin VB.Label lblSS 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Start server on application start up"
         Height          =   195
         Left            =   1065
         TabIndex        =   33
         Top             =   2295
         Width           =   2415
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Application Options"
         ForeColor       =   &H00D54600&
         Height          =   195
         Left            =   375
         TabIndex        =   30
         Top             =   1500
         Width           =   1365
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "These are the options that will change features of Ncim Network Server."
         Height          =   390
         Left            =   480
         TabIndex        =   29
         Top             =   1740
         Width           =   4860
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblSysTray 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Show icon in system tray"
         Height          =   195
         Left            =   1065
         TabIndex        =   28
         Top             =   1005
         Width           =   1740
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Windows Options"
         ForeColor       =   &H00D54600&
         Height          =   195
         Left            =   375
         TabIndex        =   26
         Top             =   210
         Width           =   1245
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "These are options that you can change within the Windows environment."
         Height          =   390
         Left            =   480
         TabIndex        =   25
         Top             =   450
         Width           =   4830
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox picTab 
      BorderStyle     =   0  'None
      Height          =   5400
      Index           =   4
      Left            =   105
      Picture         =   "frmMain.frx":7579C
      ScaleHeight     =   360
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   388
      TabIndex        =   16
      Top             =   1020
      Visible         =   0   'False
      Width           =   5820
      Begin VB.Label lblSite 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Please go to the website for more information"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000015&
         Height          =   195
         Left            =   480
         MouseIcon       =   "frmMain.frx":79EE9
         MousePointer    =   99  'Custom
         TabIndex        =   19
         ToolTipText     =   "Click here to visit the Ncim Network homepage."
         Top             =   4905
         Width           =   3165
      End
      Begin VB.Label lblAbout 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "About..."
         Height          =   195
         Left            =   480
         TabIndex        =   18
         Top             =   450
         Width           =   4830
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "About Ncim Network XP"
         ForeColor       =   &H00D54600&
         Height          =   195
         Left            =   375
         TabIndex        =   17
         Top             =   210
         Width           =   1725
      End
   End
   Begin VB.PictureBox picTab 
      BorderStyle     =   0  'None
      Height          =   5400
      Index           =   1
      Left            =   105
      Picture         =   "frmMain.frx":7A03B
      ScaleHeight     =   360
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   388
      TabIndex        =   35
      Top             =   1020
      Visible         =   0   'False
      Width           =   5820
      Begin VB.TextBox txtLog 
         Height          =   4530
         Left            =   480
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   37
         Top             =   450
         Width           =   4995
      End
      Begin VB.Label lblEventLog 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Event Log of Session"
         ForeColor       =   &H00D54600&
         Height          =   195
         Left            =   375
         TabIndex        =   36
         Top             =   210
         Width           =   1515
      End
   End
   Begin ComctlLib.TabStrip tbsOpt 
      Height          =   5760
      Left            =   90
      TabIndex        =   4
      Top             =   705
      Width           =   5880
      _ExtentX        =   10372
      _ExtentY        =   10160
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   5
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Users"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Event Log"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Connection"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Preferences"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "About"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image imgNotXP 
      Height          =   75
      Left            =   1980
      Top             =   6570
      Width           =   75
   End
   Begin VB.Menu mnu 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuGetIP 
         Caption         =   "GetIP"
      End
      Begin VB.Menu mnuKick 
         Caption         =   "Kick User"
      End
      Begin VB.Menu mnuUsrIndex 
         Caption         =   "Users Indexs"
      End
   End
   Begin VB.Menu mnu2 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuShow 
         Caption         =   "Show Server"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function InitCommonControls Lib "Comctl32.dll" () As Long       'Windows XP Controls Declaration
Dim Num As Long
Dim StartTime As String
Dim strUsers() As String
Dim MsgEnd

Const SvrPort = "1290"

Private Declare Function GetVersionEx& Lib "kernel32" Alias _
    "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) 'As Long
    Private Const VER_PLATFORM_WIN32_NT = 2
    Private Const VER_PLATFORM_WIN32_WINDOWS = 1
    Private Const VER_PLATFORM_WIN32s = 0


Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128 'Maintenance string for PSS usage
    End Type

Private Sub FindMyOS()
    Dim junk
    Dim osvi As OSVERSIONINFO
    osvi.dwOSVersionInfoSize = 148
    junk = GetVersionEx(osvi)


    If junk <> 0 Then
        Select Case osvi.dwPlatformId
            Case VER_PLATFORM_WIN32s '0
                MsgEnd = "Microsoft Win32s"
                
            Case VER_PLATFORM_WIN32_WINDOWS '1
            
                If ((osvi.dwMajorVersion > 4) Or ((osvi.dwMajorVersion = 4) And (osvi.dwMinorVersion > 0))) Then
                    If osvi.dwMinorVersion = 1 Then
                        MsgEnd = "Microsoft Windows 98 Second Edition"
                    Else
                        MsgEnd = "Microsoft Windows 98"
                    End If
                Else
                    MsgEnd = "Microsoft Windows 95"
                End If
                
            Case VER_PLATFORM_WIN32_NT '2
                
                If osvi.dwMajorVersion <= 4 Then
                    MsgEnd = "Microsoft Windows NT"
                ElseIf osvi.dwMajorVersion = 5 Then
                    If osvi.dwMinorVersion >= 1 Then
                        MsgEnd = "Microsoft Windows XP"
                    Else
                        MsgEnd = "Microsoft Windows 2000"
                    End If
                End If
                
End Select
       
End If
End Sub

Private Sub cmdListen_Click()
On Error Resume Next
If cmdListen.Caption = "&Start Server" Then
    wskServer(0).Close
    wskServer(0).Listen
    lblStatus.Caption = "Server Status: " & "Waiting for Connections..."
    cmdListen.Caption = "&Close Server"
ElseIf cmdListen.Caption = "&Close Server" Then
    For i = 0 To wskServer().UBound
        wskServer(i).Close
        strUsers(i) = ""
    Next i
    lblStatus.Caption = "Server Status: " & "Closed..."
    cmdListen.Caption = "&Start Server"
    lstUsers.Clear
    Num = "0"
    lblCount.Caption = "User Count: " & Num
End If
End Sub

Private Sub cmdSysTray_Click()
Me.Hide
End Sub

Private Sub Form_Initialize()
    
'Skin the controls
    Dim X As Long
        X = InitCommonControls
        
End Sub

Private Sub Form_Load()
If App.PrevInstance = True Then
    Call MsgBox("Another instance of Ncim Network Server Is already running." & vbCrLf & "This instance will be closed.", vbExclamation, "Ncim Network")
    End
End If

FindMyOS

If MsgEnd = "Microsoft Windows XP" Then
Else
picTab(0).Picture = imgNotXP.Picture
picTab(1).Picture = imgNotXP.Picture
picTab(2).Picture = imgNotXP.Picture
picTab(3).Picture = imgNotXP.Picture
picTab(4).Picture = imgNotXP.Picture
End If

'Get settings
txtPort.Text = GetSetting(App.Title, "Connection", "Port", "1290")

wskServer(0).LocalPort = txtPort.Text
lblIpPort.Caption = "Server IP + Port: " & wskServer(0).LocalIP & ":" & wskServer(0).LocalPort
Num = "0"
lblCount.Caption = "User Count: " & Num
lblStatus.Caption = "Server Status: " & "Closed..."
StartTime = Time

picTab(tbsOpt.SelectedItem.Index - 1).Visible = True
lblAbout.Caption = "Ncim Network is a network communications application, that has been designed to allow people to send/receive messages and files." & vbCrLf & vbCrLf
lblAbout.Caption = lblAbout.Caption & "This can be any form of network such as;" & vbCrLf & "     -  LAN (Local Area Network)" & vbCrLf & "     -  WAN (Wide Area Network)" & vbCrLf & "     -  Internet Connection" & vbCrLf & vbCrLf
lblAbout.Caption = lblAbout.Caption & "This version of Ncim Network (" & App.Major & "." & App.Minor & "." & App.Revision & ") has been designed for Windows XP." & vbCrLf & vbCrLf & vbCrLf & "Designed and Coded by Mike Simmonds"

If GetSetting(App.Title, "Preferences", "Splash", 1) = 1 Then
    frmSplash.Show vbModal, Me
    chkSplash.Value = 1
End If

If GetSetting(App.Title, "Preferences", "System Tray", 1) = 1 Then
    Call AddIcon(Me, "Ncim Network XP Server" & vbCrLf & "Click to restore Server Window")
    chkSysTray.Value = 1
End If

If GetSetting(App.Title, "Preferences", "Start Server", 0) = 1 Then
    cmdListen_Click
    chkSS.Value = 1
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
'Save settings
SaveSetting App.Title, "Connection", "Port", txtPort.Text
SaveSetting App.Title, "Preferences", "Splash", chkSplash.Value
SaveSetting App.Title, "Preferences", "System Tray", chkSysTray.Value
SaveSetting App.Title, "Preferences", "Start Server", chkSS.Value

Call DeleteIcon
DoEvents
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Dim Result As Long, Msg As Long
If Me.ScaleMode = vbPixels Then
    Msg = X
Else
    Msg = X / Screen.TwipsPerPixelX
End If
    
Select Case Msg
Case 517
Me.PopupMenu mnu2
        
Case 514
Result = SetForegroundWindow(Me.hwnd)
Me.Show
            
End Select
End Sub

Private Sub lblRestart_Click()
frmRestart.Show
Unload Me
End Sub

Private Sub lblSite_Click()
ShellExecute Me.hwnd, "open", "http://www.ncim.co.uk/networkxp/server", "", "", 1
End Sub

Private Sub lblSplash_Click()

    If chkSplash.Value = 1 Then
        chkSplash.Value = 0
    Else
        chkSplash.Value = 1
    End If

End Sub

Private Sub lblSS_Click()

    If chkSS.Value = 1 Then
        chkSS.Value = 0
    Else
        chkSS.Value = 1
    End If

End Sub

Private Sub lblSUF_Click()

    If chkSUF.Value = 1 Then
        chkSUF.Value = 0
    Else
        chkSUF.Value = 1
    End If
    
End Sub

Private Sub lblSysTray_Click()

    If chkSysTray.Value = 1 Then
        chkSysTray.Value = 0
    Else
        chkSysTray.Value = 1
    End If

End Sub

Private Sub lstUsers_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lstUsers.SelCount = "0" Then
Exit Sub
Else
    If Button = 1 Then
        Exit Sub
    Else
        Call Me.PopupMenu(mnu, , lstUsers.Left + 160)
    End If
End If
End Sub

Private Sub mnuGetIP_Click()
Call MsgBox(lstUsers.Text & "'s IP is: " & GetIP(lstUsers.Text), vbInformation, "User's IP")
End Sub

Private Sub mnuKick_Click()
Call KickUser(Trim(lstUsers.Text), "Server")
End Sub

Private Sub mnuShow_Click()
Me.Show
End Sub

Private Sub mnuUsrIndex_Click()
Dim strList As String
For i = LBound(strUsers()) To UBound(strUsers())
    If Not strUsers(i) = "" Then
        strList = strList & strUsers(i) & "(" & i & ")" & vbCrLf
        DoEvents
    End If
Next i
MsgBox strList, vbInformation, "Users Index's"
End Sub

Private Sub tbsOpt_Click()

picTab(0).Visible = False
picTab(1).Visible = False
picTab(2).Visible = False
picTab(3).Visible = False
picTab(4).Visible = False

picTab(tbsOpt.SelectedItem.Index - 1).Visible = True

End Sub

Private Sub txtPort_KeyPress(KeyAscii As Integer)
    If ((KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 9) Then
        KeyAscii = 0
        Beep
    End If
End Sub

Private Sub wskServer_ConnectionRequest(Index As Integer, ByVal requestID As Long)
On Error Resume Next
Dim lngIndex As Long, blnFlag As Boolean
   
For lngIndex& = 1 To wskServer().UBound
    If wskServer(lngIndex&).State = sckClosed Then
        blnFlag = True
        Exit For
    End If
Next lngIndex&
      
If blnFlag = False Then
    lngIndex& = wskServer().UBound + 1
    Load wskServer(lngIndex&)
    ReDim Preserve strUsers(lngIndex&) As String
    strUsers(lngIndex&) = ""
End If

Call wskServer(lngIndex&).Accept(requestID&)

End Sub

Private Sub wskServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim strData As String

wskServer(Index).GetData strData$

Call DataParsing(strData, Index)
DoEvents
End Sub

Function DataParsing(strData As String, Index As Integer)
On Error Resume Next
Dim Command As String, Info As String, i As Integer

Command$ = Left$(strData$, InStr(strData$, " ") - 1)
Info$ = Right$(strData$, Len(strData$) - InStr(strData$, " "))

Select Case Command

Case "Message":
    For i = 1 To wskServer().UBound
        Call wskServer(i).SendData("Message " & Info)
        DoEvents
    Next i
    
    txtLog.Text = txtLog.Text & Info & vbCrLf
    
Case "PMMessage":
    Dim User As String, UserText As String, UserE As Boolean
    User$ = Left$(Info$, InStr(Info$, ":") - 1)
    UserText$ = Right$(Info$, Len(Info$) - InStrRev(Info$, ":"))
    DoEvents
    
    For i = LBound(strUsers()) To UBound(strUsers())
        If strUsers(i) = User Then
            UserE = True
            Call wskServer(i).SendData("PMMessage " & UserText)
            DoEvents
            Call wskServer(Index).SendData("PMMessage " & UserText)
            DoEvents
        End If
    Next i

    If Not UserE = True Then
        Call wskServer(Index).SendData("PMError " & User)
        DoEvents
    End If
    
    txtLog.Text = txtLog.Text & vbCrLf & "Private IM: " & User & ">" & UserE
    
Case "Join":
    If CheckUsername(Info) = True Then
        Call wskServer(Index).SendData("ErrUsername " & Info)
        DoEvents
        Exit Function
    End If
    
    Num = Num + 1
    lblCount.Caption = "User Count: " & Num
    
    Call lstUsers.AddItem(Info)
    strUsers(Index) = Info
    DoEvents
    
    For i = 0 To lstUsers.ListCount
        If lstUsers.List(i) <> "" Then
            Call wskServer(Index).SendData("UserList " & lstUsers.List(i))
            DoEvents
        End If
        DoEvents
    Next i
    
    For i = 1 To wskServer().UBound
        If Not i = Index Then
            Call wskServer(i).SendData("Joined " & Info)
            DoEvents
        End If
        DoEvents
    Next i
    
    txtLog.Text = txtLog.Text & vbCrLf & "Join: " & Info
    
Case "Leave":
    For i = 1 To wskServer().UBound
        If Not i = Index Then
            Call wskServer(i).SendData("Left " & Info)
            DoEvents
        End If
        DoEvents
    Next i
    
    For i = 0 To lstUsers.ListCount
        If lstUsers.List(i) = Info Then
            Call lstUsers.RemoveItem(i)
            DoEvents
        End If
        DoEvents
    Next i
    DoEvents
    
    If CheckUnKick(Info) = True Then
        For i = 0 To lstUnKickable.ListCount
            If lstUnKickable.List(i) = Info Then
                Call lstUnKickable.RemoveItem(i)
                DoEvents
            End If
        Next i
    End If
    
    strUsers(Index) = ""
    
    Num = Num - 1
    If Num < "0" Then Num = "0"
    lblCount.Caption = "User Count: " & Num

    wskServer(Index).Close
    
    txtLog.Text = txtLog.Text & vbCrLf & "Leave: " & Info
    
Case "Prompt":
    Call Prompt(Info$, Index)

End Select
End Function

Function Prompt(Info As String, Index As Integer)
    Dim Command As String, Data As String
    Command$ = Left$(Info$, InStr(Info$, ":") - 1)
    Data$ = Right$(Info$, Len(Info$) - InStr(Info$, ":"))
    DoEvents
    
    If Command = "KickUser" Then
        Call KickUser(Data, strUsers(Index))
        DoEvents
    ElseIf Command = "ServerInfo" Then
        Dim strText As String
        strText = "Server IP/Port: " & lblIpPort.Caption & vbCrLf
        strText = strText & "User Count: " & lblCount.Caption & vbCrLf
        strText = strText & "Server Vesion: " & lblVersion.Caption & vbCrLf
        strText = strText & "Server Start Time: " & StartTime
        DoEvents
        Call wskServer(Index).SendData("PMMessage " & "Server|" & strText)
        DoEvents
    ElseIf Command = "GetIP" Then
        Call wskServer(Index).SendData("PMMessage " & "Server|" & GetIP(Data))
        DoEvents
    ElseIf Command = "MakemeUnKickable" Then
        Call lstUnKickable.AddItem(Data)
        DoEvents
        Call wskServer(Index).SendData("PMMessage " & "Server|" & "You are unkickable")
        DoEvents
    ElseIf Command = "Help" Then
        Dim sText As String
        sText = "KickUser:Username" & vbCrLf
        sText = sText & "ServerInfo:username" & vbCrLf
        sText = sText & "GetIP:username" & vbCrLf
        sText = sText & "Help:username"
        DoEvents
        Call wskServer(Index).SendData("PMMessage " & "Server|" & sText)
        DoEvents
    End If
End Function

Function GetIP(Username As String) As String
For i = 0 To UBound(strUsers())
    If strUsers(i) = Username Then
    GetIP = wskServer(i).RemoteHostIP
    End If
Next i
If GetIP = "" Then
GetIP = "None"
End If
End Function

Function KickUser(Username As String, By As String)
On Error Resume Next
Dim i As Integer
For i = 0 To UBound(strUsers())
    
    If CheckUnKick(Username) = True Then
        If Not By = "Server" Then
            wskServer(UIndex(By)).SendData ("PMMessage " & "Server|" & "User " & Username & "is unkickable")
            DoEvents
            Exit Function
        End If
    End If
    
    If strUsers(i) = Username Then
        Call wskServer(i).SendData("UKicked " & By)
        DoEvents
        Call wskServer(i).Close
        DoEvents
            
    For z = 1 To wskServer().UBound
        If Not i = Index Then
            Call wskServer(z).SendData("Kicked " & Username)
            DoEvents
        End If
        DoEvents
    Next z
    DoEvents
    
    For z = 0 To lstUsers.ListCount
        If lstUsers.List(z) = Username Then
            Call lstUsers.RemoveItem(z)
            DoEvents
        End If
        DoEvents
    Next z
    DoEvents
    
    If CheckUnKick(Username) = True Then
        For a = 0 To lstUnKickable.ListCount
            If lstUnKickable.List(a) = Username Then
                Call lstUnKickable.RemoveItem(i)
                DoEvents
            End If
        Next a
    End If
    
    strUsers(i) = ""
    
    Num = Num - 1
    If Num < "0" Then Num = "0"
    lblCount.Caption = "User Count: " & Num
        
    End If
Next i

txtLog.Text = txtLog.Text & "Kicked: " & Username

End Function

Function UIndex(Username As String) As Integer
Dim i As Integer
For i = o To UBound(strUsers())
    If strUsers(i) = Username Then
        UIndex = i
        Exit Function
    End If
Next i
End Function

Function CheckUnKick(Username As String) As Boolean
For i = 0 To lstUnKickable.ListCount
If lstUnKickable.List(i) = Username Then
    CheckUnKick = True
    Exit Function
End If
Next i
CheckUnKick = False
End Function

Function CheckUsername(Username As String) As Boolean
For i = 0 To lstUsers.ListCount
    If lstUsers.List(i) = Username Then
        CheckUsername = True
        Exit Function
    End If
Next i
CheckUsername = False
End Function
