'C:\Users\andregoro\AppData\Local\Programs\Opera\launcher.exe  
Private Declare Function DisplaySize Lib "user32" Alias _
    "GetSystemMetrics" (ByVal nIndex As Long) As Long
Private Const SM_CXSCREEN = 0
Private Const SM_CYSCREEN = 1
Sub CheckDisplayResolution()
    Dim VideoInfo As String
    Dim Msg1 As String, Msg2 As String, Msg3 As String

    VideoInfo = VideoRes
    Msg1 = "            Sua atual Resolução é " & VideoInfo & Chr(10)
    Msg2 = "A resoluções ótima para esta aplicação é 1024 x 768 " & Chr(10)
    Msg3 = "   Pedimos a gentileza de ajustar a resolução"
    Select Case VideoInfo
    Case Is = "640 x 480 "
        MsgBox Msg1 & Msg2 & Msg3
    Case Is = "800 x 600 "
        MsgBox Msg1 & Msg2
    Case Is = "1024 x 768"
        MsgBox Msg1
    Case Else
        'MsgBox Msg2 & Msg3
        MsgBox Msg1 & Msg2 & Msg3
    End Select
End Sub

Private Function VideoRes() As String
    Dim vidWidth
    Dim vidHeight
   
    vidWidth = DisplaySize(SM_CXSCREEN)
    vidHeight = DisplaySize(SM_CYSCREEN)
    Select Case (vidWidth * vidHeight)
    Case 307200
        VideoRes = "640 x 480"
    Case 480000
        VideoRes = "800 x 600"
    Case 786432
        VideoRes = "1024 x 768"
    Case Else
        VideoRes = vidWidth & " x " & vidHeight
    End Select
End Function