VERSION 5.00
Begin VB.Form frmRecorder 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mouse Recorder"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4290
   Icon            =   "frmRecorder.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   4290
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrPlay 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   420
      Top             =   0
   End
   Begin VB.Timer tmrRecord 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play"
      Height          =   555
      Left            =   1140
      TabIndex        =   1
      Top             =   1260
      Width           =   1875
   End
   Begin VB.CommandButton cmdRecord 
      Caption         =   "Record"
      Height          =   555
      Left            =   1140
      TabIndex        =   0
      Top             =   660
      Width           =   1875
   End
End
Attribute VB_Name = "frmRecorder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************'
'                                                       '
'   By:         Waleed A. Aly                           '
'   ASL:        [20 M Egypt]                            '
'   eMail:      wa_aly@tdcspace.dk                      '
'   Thanks to:  www.allapi.net                          '
'                                                       '
'     Please eMail me any Comments and|or Suggestions.  '
'   I hope you like my work and think is usefull !  :)  '
'   I'd love to know how many people are using my Code  '
'   so you can always eMail me if you are goin' to use  '
'   it :)                                               '
'                                      Thanks.          '
'                                                       '
'*******************************************************'

Private Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long

Private Type PointAPI
    x As Long
    y As Long
End Type

Private Const SPS As Long = 100    ' Recorded Samples Per Second
Private Pos() As PointAPI, i As Long, m As Long, Samples As Long

Private Sub Form_Load()

    tmrRecord.Interval = 1000 / SPS
    tmrPlay.Interval = 1000 / SPS

End Sub

Private Sub cmdRecord_Click()

    Samples = SPS * Val(InputBox("Number of seconds to record :", "Recording Time"))
    If Samples <= 0 Then Exit Sub
    i = 0
    ReDim Pos(Samples)
    tmrRecord.Enabled = True
    tmrPlay.Enabled = False

End Sub

Private Sub cmdPlay_Click()

    m = i - 1: i = 0
    If m < 0 Then Exit Sub
    tmrRecord.Enabled = False
    tmrPlay.Enabled = True

End Sub

Private Sub tmrRecord_Timer()

    GetCursorPos Pos(i)
    If i < Samples Then
        i = i + 1
    Else
        tmrRecord.Enabled = False
        tmrPlay.Enabled = False
        MsgBox "Record finished.", vbInformation, "finished!"
    End If

End Sub

Private Sub tmrPlay_Timer()

    SetCursorPos Pos(i).x, Pos(i).y
    If i < m Then
        i = i + 1
    Else
        tmrRecord.Enabled = False
        tmrPlay.Enabled = False
        MsgBox "Play finished.", vbInformation, "finished!"
    End If

End Sub
