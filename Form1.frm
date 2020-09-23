VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4185
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   3825
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4185
   ScaleWidth      =   3825
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Timer7 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2625
      Top             =   1950
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   525
      Top             =   2550
   End
   Begin VB.Timer Timer5 
      Interval        =   10000
      Left            =   1425
      Top             =   975
   End
   Begin MCI.MMControl mm 
      Height          =   540
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   3150
      _ExtentX        =   5556
      _ExtentY        =   953
      _Version        =   393216
      PlayEnabled     =   -1  'True
      NextVisible     =   0   'False
      DeviceType      =   "sequencer"
      FileName        =   "C:\programming\troj\christmaskey\Jingle.mid"
   End
   Begin VB.Timer Timer4 
      Interval        =   500
      Left            =   2250
      Top             =   1950
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2775
      Top             =   2550
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2400
      Top             =   2550
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2025
      Top             =   2550
   End
   Begin VB.Line Line3 
      X1              =   2625
      X2              =   2625
      Y1              =   2400
      Y2              =   2475
   End
   Begin VB.Line Line2 
      X1              =   1725
      X2              =   2550
      Y1              =   1500
      Y2              =   1875
   End
   Begin VB.Line Line1 
      X1              =   1500
      X2              =   750
      Y1              =   1500
      Y2              =   2400
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cmdm$
Dim cmdp$
Private Sub Form_Load()
mm.filename = App.Path & "\" & "mu.dat"
mm.DeviceType = "Sequencer"
mm.Command = "open"
mm.Command = "play"
'cmdm$ = "open " & App.Path & "\" & "mu.dat" & " type sequencer alias music"
'cmdp$ = "play music"
'x& = mciSendString(cmdm$, 0&, 0, 0)
'x& = mciSendString(cmdp$, 0&, 0, 0)
End Sub
Private Sub mm_Done(NotifyCode As Integer)
mm.Command = "close"
mm.DeviceType = "Sequencer"
mm.Command = "open"
mm.Command = "play"
End Sub

Private Sub Timer1_Timer()
Static n
Dim keys(255) As Byte
keys(&H91) = n
tmp = SetKeyboardState(keys(0))
n = n + 1
If n = 2 Then n = 0
End Sub

Private Sub Timer2_Timer()
Static n
Dim keys(255) As Byte
keys(&H90) = n
tmp = SetKeyboardState(keys(0))
n = n + 1
If n = 2 Then n = 0
End Sub

Private Sub Timer3_Timer()
Static n
Dim keys(255) As Byte
keys(&H14) = n
tmp = SetKeyboardState(keys(0))
n = n + 1
If n = 2 Then n = 0
End Sub

Private Sub Timer4_Timer()
Static t
t = t + 1
If t = 1 Then
Timer1.Interval = 1
Timer2.Interval = 1
Timer3.Interval = 1
End If
If t = 3 Then
Timer1.Enabled = True
Timer2.Enabled = False
Timer3.Enabled = False
t = 0
End If
If t = 1 Then
Timer2.Enabled = True
Timer1.Enabled = False
Timer3.Enabled = False
End If
If t = 2 Then
Timer3.Enabled = True
Timer2.Enabled = False
Timer1.Enabled = False
End If

End Sub

Private Sub Timer5_Timer()
Static n
If n = 0 Then
Timer6.Enabled = True
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Timer7.Enabled = False
End If
If n = 1 Then
Timer7.Interval = 500
Timer7.Enabled = True
Timer6.Enabled = False
Timer4.Enabled = False
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
End If
If n = 2 Then
Timer4.Enabled = True
Timer6.Enabled = False
Timer7.Enabled = False
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
n = -1
End If
n = n + 1
End Sub

Private Sub Timer6_Timer()
Static n
Dim keys(255) As Byte
keys(&H14) = n
keys(&H90) = n
keys(&H91) = n
tmp = SetKeyboardState(keys(0))
n = n + 1
If n = 2 Then n = 0
End Sub

Private Sub Timer7_Timer()
Static t
Dim keys(255) As Byte
If Timer7.Interval > 70 Then Timer7.Interval = Timer7.Interval - 20
t = t + 1
If t = 1 Then
Timer1.Interval = 200
Timer2.Interval = 200
Timer3.Interval = 200
End If
If t = 3 Then
keys(&H14) = 0
keys(&H90) = 0
keys(&H91) = 1
tmp = SetKeyboardState(keys(0))
t = 0
End If
If t = 1 Then
keys(&H14) = 0
keys(&H90) = 1
keys(&H91) = 0
tmp = SetKeyboardState(keys(0))
End If
If t = 2 Then
keys(&H14) = 1
keys(&H90) = 0
keys(&H91) = 0
tmp = SetKeyboardState(keys(0))
End If



End Sub
