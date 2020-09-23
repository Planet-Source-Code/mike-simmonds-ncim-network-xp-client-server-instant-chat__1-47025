VERSION 5.00
Begin VB.Form frmPM 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ncim Network XP - Private Message"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5115
   Icon            =   "frmPM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   206
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   341
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSendPM 
      Caption         =   "Send PM"
      Default         =   -1  'True
      Height          =   345
      Left            =   3885
      TabIndex        =   1
      Top             =   2640
      Width           =   1125
   End
   Begin VB.TextBox txtPM 
      Height          =   1695
      Left            =   90
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   840
      Width           =   4935
   End
   Begin VB.Label lblUserPM 
      Height          =   255
      Left            =   810
      TabIndex        =   4
      Top             =   555
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Send to:"
      Height          =   255
      Left            =   90
      TabIndex        =   3
      Top             =   555
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Whatever you type in this window only the user you selected will see."
      Height          =   255
      Left            =   90
      TabIndex        =   2
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "frmPM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub cmdSendPM_Click()
If txtPM.Text <> "" Then
Dim allText As String
allText = vbCrLf & vbCrLf & frmMain.txtUser.Text & "|" & txtPM.Text
Call frmMain.wskClient.SendData("PMMessage " & lblUserPM.Caption & ":" & allText)
txtPM.Text = ""
DoEvents
End If
Unload Me
End Sub

Private Sub txtPM_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdSendPM_Click
    KeyAscii = 0
End If
End Sub
