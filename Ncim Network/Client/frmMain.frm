VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ncim Network XP - Username"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   540
   ClientWidth     =   6060
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   467
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   404
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUser 
      Height          =   285
      Left            =   1710
      Locked          =   -1  'True
      TabIndex        =   31
      Top             =   2.45745e5
      Width           =   1335
   End
   Begin VB.Timer tmrConnection 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   945
      Top             =   2.45745e5
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   2.45745e5
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   345
      Left            =   4845
      TabIndex        =   4
      Top             =   6555
      Width           =   1125
   End
   Begin VB.CommandButton cmdSignOut 
      Caption         =   "Sign Out"
      Height          =   345
      Left            =   3630
      TabIndex        =   3
      Top             =   6555
      Width           =   1125
   End
   Begin VB.PictureBox picHeader 
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   6060
      TabIndex        =   0
      Top             =   0
      Width           =   6060
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
         TabIndex        =   2
         Top             =   135
         Width           =   1500
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome to the Ncim Network Client"
         Height          =   195
         Left            =   210
         TabIndex        =   1
         Top             =   330
         Width           =   2610
      End
   End
   Begin MSWinsockLib.Winsock wskClient 
      Left            =   465
      Top             =   2.45745e5
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox rt 
      Height          =   255
      Left            =   3030
      TabIndex        =   29
      Top             =   2.45745e5
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      _Version        =   393217
      ReadOnly        =   -1  'True
      TextRTF         =   $"frmMain.frx":628A
   End
   Begin RichTextLib.RichTextBox rt2 
      Height          =   255
      Left            =   3030
      TabIndex        =   30
      Top             =   2.45745e5
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      _Version        =   393217
      ReadOnly        =   -1  'True
      TextRTF         =   $"frmMain.frx":630C
   End
   Begin VB.PictureBox picTab 
      BorderStyle     =   0  'None
      Height          =   5400
      Index           =   0
      Left            =   105
      Picture         =   "frmMain.frx":638E
      ScaleHeight     =   360
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   388
      TabIndex        =   8
      Top             =   1020
      Visible         =   0   'False
      Width           =   5820
      Begin VB.TextBox txtText 
         Height          =   300
         Left            =   480
         TabIndex        =   28
         Top             =   4200
         Width           =   4995
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "&Send"
         Default         =   -1  'True
         Height          =   345
         Left            =   480
         TabIndex        =   27
         Top             =   4605
         Width           =   1125
      End
      Begin RichTextLib.RichTextBox txtDisplay 
         Height          =   3150
         Left            =   510
         TabIndex        =   25
         Top             =   930
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   5556
         _Version        =   393217
         BorderStyle     =   0
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         Appearance      =   0
         TextRTF         =   $"frmMain.frx":AADB
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
      Begin VB.TextBox txtHolder 
         Enabled         =   0   'False
         Height          =   3210
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   900
         Width           =   4995
      End
      Begin VB.Label lblPortText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To send an IM (Instant Message) type your message in the text field at the bottom of the window and click Send."
         Height          =   390
         Left            =   480
         TabIndex        =   12
         Top             =   450
         Width           =   4830
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblfraServerSet 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Instant Messaging"
         ForeColor       =   &H00D54600&
         Height          =   195
         Left            =   375
         TabIndex        =   11
         Top             =   210
         Width           =   1290
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Your change will take effect next time you start NcimNetwork Server."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   480
         TabIndex        =   10
         Top             =   1515
         Width           =   4905
         WordWrap        =   -1  'True
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
         ForeColor       =   &H00666666&
         Height          =   195
         Left            =   480
         MouseIcon       =   "frmMain.frx":AB78
         MousePointer    =   99  'Custom
         TabIndex        =   9
         ToolTipText     =   "Click here to visit the Ncim Network homepage."
         Top             =   1710
         Width           =   3075
      End
   End
   Begin VB.PictureBox picTab 
      BorderStyle     =   0  'None
      Height          =   5400
      Index           =   1
      Left            =   105
      Picture         =   "frmMain.frx":ACCA
      ScaleHeight     =   360
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   388
      TabIndex        =   5
      Top             =   1020
      Visible         =   0   'False
      Width           =   5820
      Begin VB.ListBox lstUsers 
         Height          =   4530
         IntegralHeight  =   0   'False
         ItemData        =   "frmMain.frx":F417
         Left            =   480
         List            =   "frmMain.frx":F419
         TabIndex        =   6
         ToolTipText     =   "Double-click a user to send a private message..."
         Top             =   450
         Width           =   4995
      End
      Begin VB.Label lblfra02 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User(s) Online"
         ForeColor       =   &H00D54600&
         Height          =   195
         Left            =   375
         TabIndex        =   7
         Top             =   210
         Width           =   990
      End
   End
   Begin VB.PictureBox picTab 
      BorderStyle     =   0  'None
      Height          =   5400
      Index           =   3
      Left            =   105
      Picture         =   "frmMain.frx":F41B
      ScaleHeight     =   360
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   388
      TabIndex        =   20
      Top             =   1020
      Visible         =   0   'False
      Width           =   5820
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "About Ncim Network XP"
         ForeColor       =   &H00D54600&
         Height          =   195
         Left            =   375
         TabIndex        =   23
         Top             =   210
         Width           =   1725
      End
      Begin VB.Label lblAbout 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "About..."
         Height          =   195
         Left            =   480
         TabIndex        =   22
         Top             =   450
         Width           =   4830
         WordWrap        =   -1  'True
      End
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
         MouseIcon       =   "frmMain.frx":13B68
         MousePointer    =   99  'Custom
         TabIndex        =   21
         ToolTipText     =   "Click here to visit the Ncim Network homepage."
         Top             =   4905
         Width           =   3165
      End
   End
   Begin VB.PictureBox picTab 
      BorderStyle     =   0  'None
      Height          =   5400
      Index           =   2
      Left            =   105
      Picture         =   "frmMain.frx":13CBA
      ScaleHeight     =   360
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   388
      TabIndex        =   13
      Top             =   1020
      Visible         =   0   'False
      Width           =   5820
      Begin VB.CheckBox chkSS 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   810
         TabIndex        =   15
         Top             =   1005
         Width           =   195
      End
      Begin VB.CheckBox chkSplash 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   810
         TabIndex        =   14
         Top             =   1335
         Width           =   195
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "These are the options that will change features of Ncim Network Client."
         Height          =   390
         Left            =   480
         TabIndex        =   19
         Top             =   450
         Width           =   4860
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Application Options"
         ForeColor       =   &H00D54600&
         Height          =   195
         Left            =   375
         TabIndex        =   18
         Top             =   210
         Width           =   1365
      End
      Begin VB.Label lblSS 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Play sound when a message is sent and received"
         Height          =   195
         Left            =   1065
         TabIndex        =   17
         Top             =   1005
         Width           =   3495
      End
      Begin VB.Label lblSplash 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Show splash screen on start up"
         Height          =   195
         Left            =   1065
         TabIndex        =   16
         Top             =   1335
         Width           =   2220
      End
   End
   Begin ComctlLib.TabStrip tbsOpt 
      Height          =   5760
      Left            =   90
      TabIndex        =   24
      Top             =   705
      Width           =   5880
      _ExtentX        =   10372
      _ExtentY        =   10160
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   4
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Instant Messaging"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Users"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Preferences"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "About"
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
      Left            =   1755
      Top             =   6540
      Width           =   75
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      Caption         =   "Status:"
      Height          =   195
      Left            =   90
      TabIndex        =   32
      Top             =   6630
      Width           =   495
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Dim MsgEnd As String
Dim Kicked As Boolean
Dim nTime As Long
Private SvrIPAdd As String

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

Function Connect(svrIP As String)
wskClient.Close
wskClient.Connect svrIP, "1290"
SvrIPAdd = svrIP
lblStatus.Caption = "Connecting..."
End Function

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdSend_Click()
If cmdSend.Enabled = True Then
If wskClient.State = sckConnected Then
    If txtText.Text <> "" Then
        Dim allText As String
        allText = vbCrLf & vbCrLf & txtUser.Text & ":" & vbCrLf & txtText.Text
        DoEvents
        Call wskClient.SendData("Message " & allText)
        If PlaySnd = True Then
        Call PlayWav("Send.wav")
        End If
    Else
        Call MsgBox("You must enter text to send it", vbInformation, "Ncim Network XP")
    End If
Else
    Call MsgBox("You must be connected to some to send text", vbInformation, "Ncim Network XP")
End If
txtText.Text = ""
End If
End Sub

Private Sub cmdSignOut_Click()
Call wskClient.SendData("Leave " & txtUser.Text)
DoEvents
wskClient.Close
DoEvents
frmSignon.txtUser.Text = txtUser.Text
frmSignon.Show
DoEvents
Unload Me
End Sub

Private Sub Form_Load()
'If App.PrevInstance = True Then
'    Call MsgBox("Another instance of Ncim Network XP Client is already running." & vbCrLf & "This instance will be closed.", vbExclamation, "Ncim Network XP")
'    End
'End If

FindMyOS

If MsgEnd = "Microsoft Windows XP" Then
Else
picTab(0).Picture = imgNotXP.Picture
picTab(1).Picture = imgNotXP.Picture
picTab(2).Picture = imgNotXP.Picture
picTab(3).Picture = imgNotXP.Picture
End If


nTime = "0"
tmrConnection.Enabled = True
If Prompt = True Then
    mnuPrompt.Visible = True
End If
cmdSend.Enabled = False

picTab(tbsOpt.SelectedItem.Index - 1).Visible = True
lblAbout.Caption = "Ncim Network is a network communications application, that has been designed to allow people to send/receive messages and files." & vbCrLf & vbCrLf
lblAbout.Caption = lblAbout.Caption & "This can be any form of network such as;" & vbCrLf & "     -  LAN (Local Area Network)" & vbCrLf & "     -  WAN (Wide Area Network)" & vbCrLf & "     -  Internet Connection" & vbCrLf & vbCrLf
lblAbout.Caption = lblAbout.Caption & "This version of Ncim Network (" & App.Major & "." & App.Minor & "." & App.Revision & ") has been designed for Windows XP." & vbCrLf & vbCrLf & vbCrLf & "Designed and Coded by Mike Simmonds"

txtDisplay.Text = "* Chat Session - " & Format(Date, "DDDD, DD MMMM YYYY") & "   " & Format(Time, "HH:MM:SS am/pm") & " *"

chkSplash.Value = GetSetting(App.Title, "Preferences", "Splash", 1)
chkSS.Value = GetSetting(App.Title, "Preferences", "Sound", 1)

End Sub

Private Sub tbsOpt_Click()

picTab(0).Visible = False
picTab(1).Visible = False
picTab(2).Visible = False
picTab(3).Visible = False

picTab(tbsOpt.SelectedItem.Index - 1).Visible = True

End Sub


Private Sub Form_Unload(Cancel As Integer)
If wskClient.State = sckConnected Then
Call wskClient.SendData("Leave " & txtUser.Text)
DoEvents
End If
End Sub

Private Sub lstUsers_DblClick()
If lstUsers.Text = txtUser.Text Then
    Call MsgBox("You can not send a Private Message to yourself.", vbInformation, "Ncim Network XP")
Else
frmPM.lblUserPM.Caption = lstUsers.Text
frmPM.Show vbModal, Me
End If
End Sub

Private Sub mnuAAC_Click()
frmAbout.Show
End Sub

Private Sub mnuExit_Click()
If wskClient.State = sckConnected Then
Call wskClient.SendData("Leave " & txtUser.Text)
DoEvents
End If
End
End Sub

Private Sub mnuOpt_Click()
frmOptions.Show
End Sub

Private Sub mnuPrompt_Click()
frmPrompt.Show
End Sub

Private Sub tmrConnection_Timer()
If wskClient.State = sckConnected Then
    Exit Sub
Else
If Kicked = False Then
      Call MsgBox("Ncr cannot connect to the server." & Chr(10) & "Please make sure that you are connected to the Internet and the server IP address is correct.", vbInformation, "Ncim Network XP")
        Unload Me
        frmSignon.Show
    End If
End If

End Sub

Private Sub txtText_Change()
If txtText.Text > "" Then
    cmdSend.Enabled = True
Else
    cmdSend.Enabled = False
End If

End Sub

Private Sub txtText_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdSend_Click
    KeyAscii = 0
End If
End Sub

Public Sub wskClient_Close()
lblStatus.Caption = "Closed"
If Kicked = False Then
Call MsgBox("Connection was unexpectedly dropped.", vbInformation, "Ncim Network XP")
frmSignon.Show
Unload Me
End If
End Sub

Private Sub wskClient_Connect()
lblStatus.Caption = "Connected..."
Call wskClient.SendData("Join " & txtUser.Text)
tmrConnection.Enabled = False
End Sub

Private Sub wskClient_DataArrival(ByVal bytesTotal As Long)
Dim strData As String

wskClient.GetData strData$

Call DataParsing(strData)
End Sub

Function DataParsing(strData As String)
On Error Resume Next
Dim Command As String, Info As String, User As String, Text As String

Command$ = Left$(strData$, InStr(strData$, " ") - 1)
Info$ = Right$(strData$, Len(strData$) - InStr(strData$, " "))

Select Case Command

Case "Message":
    User$ = Left$(Info$, InStr(Info$, ":") - 1)
    Text$ = Right$(Info$, Len(Info$) - InStr(Info$, ":"))
    DoEvents
    
    Call AddChat(User, Text)
    DoEvents
    If Not User = txtUser.Text Then
        If PlaySnd = True Then
            Call PlayWav("Recv.wav")
        End If
    End If
    
Case "PMMessage":
    User$ = Left$(Info$, InStr(Info$, "|") - 1)
    Text$ = Right$(Info$, Len(Info$) - InStr(Info$, "|"))
    DoEvents
    
    Call AddChatPM(User, Text)
    DoEvents
    
Case "ErrUsername":
    Dim Answer As VbMsgBoxResult
    Answer = MsgBox("The username " & Info$ & " is in use." & vbCrLf & "Would you like to enter a new one.", vbExclamation + vbYesNo, "Ncim Network XP")
        If Answer = vbYes Then
            Dim strUser As String
            strUser = InputBox("Please enter a new username.", "New Username")
            If strUser <> "" Then
            Call wskClient.SendData("Join " & strUser)
            DoEvents
            txtUser.Text = strUser
            lblUser.Caption = txtUser.Text
            End If
        ElseIf Answer = vbNo Then
            wskClient.Close
            frmSignon.Show
            DoEvents
            Unload Me
        End If

Case "PMError":
    Call AddText(vbCrLf & vbCrLf & "  // User: " & Info & " doesn't exist, or has left")
    DoEvents

Case "UserList":
    Call lstUsers.AddItem(Info)
    DoEvents
    
Case "Joined":
    Call lstUsers.AddItem(Info)
    DoEvents
    Call AddText(vbCrLf & vbCrLf & " // User: " & Info & " has joined the chat")
    
Case "Left":
    For i = 0 To lstUsers.ListCount
        If lstUsers.List(i) = Info Then
            Call lstUsers.RemoveItem(i)
            DoEvents
        End If
        DoEvents
    Next i
    
    Call AddText(vbCrLf & vbCrLf & "  // User: " & Info & " left the chat room")
    
Case "Kicked":
    For i = 0 To lstUsers.ListCount
        If lstUsers.List(i) = Info Then
            Call lstUsers.RemoveItem(i)
            DoEvents
        End If
        DoEvents
    Next i
    
    Call AddText(vbCrLf & vbCrLf & "  // User: " & Info & " was removed from the chat session")
    
Case "UKicked":
        Kicked = True
        Call MsgBox("You have been removed from this session by the server.", vbCritical, "Ncim Network XP")
        Unload Me
    
   
End Select
End Function

Function AddText(Text As String)

rt2.SelStart = 0
rt2.SelLength = 0
rt2.TextRTF = Text
rt2.SelStart = 2
rt2.SelLength = Len(Text)
rt2.SelColor = &H8000&
rt2.SelStart = 0
rt2.SelLength = 0
DoEvents
txtDisplay.SelStart = Len(txtDisplay)
DoEvents
txtDisplay.SelRTF = rt2.TextRTF

DoEvents
txtDisplay.SelStart = Len(txtDisplay)

rt2.TextRTF = ""
End Function

Function AddChat(User As String, Text As String)

User = User & " says"

rt.SelBold = False
rt.SelColor = vbBlack

rt.SelStart = 0
rt.SelLength = 0
rt.SelRTF = User & ": " & Text
rt.SelStart = 0
rt.SelLength = Len(User) + 1
If User = txtUser.Text & " says" Then
rt.SelColor = vbRed
Else
rt.SelColor = vbBlue
End If
rt.SelBold = True
rt.SelStart = 0
rt.SelLength = 0
DoEvents
txtDisplay.SelStart = Len(txtDisplay)
DoEvents
txtDisplay.SelRTF = rt.TextRTF

DoEvents
txtDisplay.SelStart = Len(txtDisplay)

rt.TextRTF = ""
End Function

Function AddChatPM(User As String, Text As String)
rt.SelBold = False
rt.SelColor = vbBlack

rt.SelStart = 0
rt.SelLength = 0
rt.SelRTF = User & ": " & Text
rt.SelStart = 0
rt.SelLength = Len(User) + 1
rt.SelColor = &H8000&
rt.SelBold = True
rt.SelStart = 0
rt.SelLength = 0
DoEvents
txtDisplay.SelStart = Len(txtDisplay)
DoEvents
txtDisplay.SelRTF = rt.TextRTF

DoEvents
txtDisplay.SelStart = Len(txtDisplay)

rt.TextRTF = ""
End Function

Private Sub lblSplash_Click()

    If chkSplash.Value = 1 Then
        chkSplash.Value = 0
        
    Else
        chkSplash.Value = 1
    End If
    
        SaveSetting App.Title, "Preferences", "Splash", chkSplash.Value

End Sub

Private Sub lblSS_Click()

    If chkSS.Value = 1 Then
        chkSS.Value = 0
        PlaySnd = False
    Else
        chkSS.Value = 1
        PlaySnd = True
    End If
    
            SaveSetting App.Title, "Preferences", "Sound", chkSS.Value

End Sub
