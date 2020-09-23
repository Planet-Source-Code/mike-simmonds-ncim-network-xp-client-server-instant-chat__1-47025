Attribute VB_Name = "modMain"
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public PlaySnd As Boolean

Function PlayWav(File As String)
Dim Play
File = App.Path & "\" & File
Play = sndPlaySound(File, 1)
End Function

Sub Main()

If GetSetting(App.Title, "Preferences", "Sound", 1) = 1 Then
    PlaySnd = True
Else
    PlaySnd = False
End If

If GetSetting(App.Title, "Preferences", "Splash", 1) = 1 Then
    frmSplash.Show
Else
    frmSignon.Show
End If

End Sub

