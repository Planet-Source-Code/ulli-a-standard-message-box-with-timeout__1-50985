Attribute VB_Name = "Module1"
Option Explicit

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long

Private Const WM_CLOSE      As Long = 16
Private CurrMsgBoxTitle     As String

Public Function TimedMsgBox(Prompt As String, Optional ByVal TimeOut As Long = 0, Optional Icon As VbMsgBoxStyle = vbOKOnly, Optional Title As String = vbNullString) As VbMsgBoxResult

  'display a timed message box

  Dim TimerId   As Long

    If Title = vbNullString Then
        CurrMsgBoxTitle = App.Title
      Else 'NOT TITLE...
        CurrMsgBoxTitle = Title
    End If
    If TimeOut = 0 Then
        TimeOut = (Len(Prompt) + Len(CurrMsgBoxTitle) + 20) * 40 'adjust timeout depending on prompt length
    End If
    TimerId = SetTimer(0, 0, TimeOut, AddressOf TimeoutMsgBox)
    TimedMsgBox = MsgBox(Prompt, Icon, CurrMsgBoxTitle)
    If CurrMsgBoxTitle = vbNullString Then 'closed by timer
        TimedMsgBox = 0
    End If
    KillTimer 0, TimerId

End Function

Private Sub TimeoutMsgBox(hWnd As Long, uMsg As Long, idEvent As Long, dwTime As Long)

  'close timed message box

    SendMessage FindWindow(vbNullString, CurrMsgBoxTitle), WM_CLOSE, 0&, 0&
    CurrMsgBoxTitle = vbNullString

End Sub

':) Ulli's VB Code Formatter V2.16.13 (2004-Jan-14 02:38) 9 + 34 = 43 Lines
