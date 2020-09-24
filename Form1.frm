VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Form1"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   3090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.HScrollBar scrDelay 
      Height          =   255
      LargeChange     =   1000
      Left            =   435
      Max             =   10000
      SmallChange     =   100
      TabIndex        =   1
      Top             =   1665
      Width           =   2220
   End
   Begin VB.CommandButton btShowMsgBox 
      Caption         =   "Show timed message box"
      Height          =   495
      Left            =   938
      TabIndex        =   0
      Top             =   210
      Width           =   1215
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Height          =   195
      Index           =   2
      Left            =   960
      TabIndex        =   4
      Top             =   810
      Width           =   45
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Automatic"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   1
      Left            =   435
      TabIndex        =   3
      Top             =   1980
      Width           =   705
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Adjust delay"
      Height          =   180
      Index           =   0
      Left            =   435
      TabIndex        =   2
      Top             =   1425
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btShowMsgBox_Click()

  Dim Delay    As String

    If scrDelay = 0 Then
        Delay = "an automatic delay,"
      Else 'NOT SCRDELAY...
        Delay = Round(scrDelay / 1000, 1) & " seconds,"
    End If

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    lbl(2) = "Return code is " & _
              TimedMsgBox("This message box will disappear after " & Delay & vbCrLf & _
                          "however, you may click any button before it disappears.", _
                           scrDelay, _
                           vbOKCancel Or vbInformation Or vbDefaultButton2, _
                          "Timed Message Box")
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

End Sub

Private Sub scrDelay_Change()

    scrDelay_Scroll

End Sub

Private Sub scrDelay_Scroll()

    lbl(1).Visible = (scrDelay = 0)

End Sub

':) Ulli's VB Code Formatter V2.16.13 (2004-Jan-14 02:38) 1 + 35 = 36 Lines
