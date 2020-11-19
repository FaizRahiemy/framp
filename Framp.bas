Attribute VB_Name = "Module1"
    Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lparam _
    As Any) As Long
    Declare Sub ReleaseCapture Lib "user32" ()
     
    Public Sub GeserJendela(TheForm As Form)
    ReleaseCapture
    Call SendMessage(TheForm.hWnd, &HA1, 2, 0&)
    End Sub
Public Sub Mute(isMute As Boolean)
    If isMute = True Then
        Call mciSendString("set mp3play audio all off", 0, 0, 0)
    ElseIf isMute = False Then
        Call mciSendString("set mp3play audio all on", 0, 0, 0)
    End If
End Sub
