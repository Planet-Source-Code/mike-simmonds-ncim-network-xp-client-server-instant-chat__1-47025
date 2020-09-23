VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   Caption         =   "Ncim Network XP - Server"
   ClientHeight    =   4500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3300
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   300
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrClose 
      Interval        =   5000
      Left            =   2160
      Top             =   3075
   End
   Begin VB.Image imgIcon 
      Height          =   750
      Left            =   435
      Picture         =   "frmSplash.frx":628A
      Top             =   1410
      Width           =   465
   End
   Begin VB.Label lblUser 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(username)"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1020
      TabIndex        =   2
      Top             =   4245
      Width           =   780
   End
   Begin VB.Label lblApp 
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
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1020
      TabIndex        =   1
      Top             =   1605
      Width           =   1500
   End
   Begin VB.Label lblServer 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Server: 6.2.2003"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1020
      TabIndex        =   0
      Top             =   1800
      Width           =   1185
   End
   Begin VB.Shape shpBar 
      BackColor       =   &H00E7EFF1&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   4485
      Left            =   360
      Top             =   15
      Width           =   615
   End
   Begin VB.Shape shpBorder 
      BackStyle       =   1  'Opaque
      Height          =   4500
      Left            =   0
      Top             =   0
      Width           =   3300
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

lblUser.Caption = GetSetting(App.Title, "Owner", "Name", "")

End Sub

Private Sub tmrClose_Timer()

FadeOut Me, True

End Sub
