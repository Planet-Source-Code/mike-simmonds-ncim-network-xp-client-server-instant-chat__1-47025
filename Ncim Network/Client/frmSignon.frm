VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmSignon 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ncim Network XP"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4200
   Icon            =   "frmSignon.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   338
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   280
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picFooter 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   0
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   280
      TabIndex        =   8
      Top             =   4365
      Width           =   4200
      Begin VB.CommandButton cmdConnect 
         Caption         =   "Connect"
         Default         =   -1  'True
         Height          =   345
         Left            =   1725
         TabIndex        =   10
         Top             =   195
         Width           =   1125
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   345
         Left            =   2925
         TabIndex        =   9
         Top             =   195
         Width           =   1125
      End
      Begin VB.Line lne3D1 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   0
         X2              =   279
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line lne3D2 
         BorderColor     =   &H80000014&
         Index           =   0
         X1              =   0
         X2              =   280
         Y1              =   1
         Y2              =   1
      End
      Begin VB.Line lne3D3 
         BorderColor     =   &H80000014&
         Index           =   0
         X1              =   279
         X2              =   280
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.PictureBox picPB 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   3705
      Left            =   0
      ScaleHeight     =   247
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   280
      TabIndex        =   11
      Top             =   915
      Visible         =   0   'False
      Width           =   4200
      Begin VB.Timer tmr1 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   870
         Top             =   2010
      End
      Begin ComctlLib.ProgressBar pb1 
         Height          =   195
         Left            =   600
         TabIndex        =   12
         Top             =   1215
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   344
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label lblPleasewait 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Please wait..."
         Height          =   195
         Left            =   600
         TabIndex        =   13
         Top             =   990
         Width           =   945
      End
   End
   Begin VB.PictureBox picHeader 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   915
      Left            =   0
      ScaleHeight     =   61
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   280
      TabIndex        =   5
      Top             =   0
      Width           =   4200
      Begin VB.Line lne3D1 
         BorderColor     =   &H80000010&
         Index           =   1
         X1              =   0
         X2              =   279
         Y1              =   59
         Y2              =   59
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Please login to the Ncim Network"
         Height          =   195
         Left            =   210
         TabIndex        =   7
         Top             =   420
         Width           =   2355
      End
      Begin VB.Label Label3 
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
         TabIndex        =   6
         Top             =   225
         Width           =   1500
      End
   End
   Begin VB.TextBox txtSvrIp 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Text            =   "127.0.0.1"
      Top             =   2340
      Width           =   2790
   End
   Begin VB.TextBox txtUser 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   1890
      Width           =   2790
   End
   Begin MSWinsockLib.Winsock wskSignon 
      Left            =   6000
      Top             =   2355
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblServer 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To view your online contacts, have online conversations and share files, please sign on."
      ForeColor       =   &H00000000&
      Height          =   390
      Left            =   210
      TabIndex        =   4
      Top             =   1095
      Width           =   3780
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Server IP:"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   210
      TabIndex        =   3
      Top             =   2385
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   210
      TabIndex        =   2
      Top             =   1950
      Width           =   765
   End
End
Attribute VB_Name = "frmSignon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function InitCommonControls Lib "Comctl32.dll" () As Long       'Windows XP Controls Declaration

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdConnect_Click()
If txtUser.Text = "" Or txtSvrIp.Text = "" Then
    Call MsgBox("Please complete all the fields to log in to Ncim Network.", vbInformation, "Ncim Network XP")
Else
    tmr1.Enabled = True
End If

picPB.Visible = True
cmdConnect.Enabled = False
cmdCancel.Enabled = False

End Sub

Private Sub Form_Initialize()
'Skin the controls
    Dim X As Long
        X = InitCommonControls
End Sub

Private Sub tmr1_Timer()

If pb1.Value = pb1.Max Then
    frmMain.Show
    frmMain.txtUser.Text = txtUser.Text
    Call frmMain.Connect(txtSvrIp.Text)
    frmMain.Caption = "Ncim Network XP - " & txtUser.Text
    DoEvents
    Unload Me
Else
    pb1.Value = pb1.Value + 1
End If

End Sub
