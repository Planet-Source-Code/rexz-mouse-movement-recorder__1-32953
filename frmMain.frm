VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Mouse Recorder"
   ClientHeight    =   630
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2670
   LinkTopic       =   "Form1"
   ScaleHeight     =   630
   ScaleWidth      =   2670
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrPlay 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   480
      Top             =   360
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Play"
      Height          =   255
      Left            =   1320
      TabIndex        =   2
      Top             =   360
      Width           =   495
   End
   Begin VB.Timer tmrRecord 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   360
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   255
      Left            =   1920
      TabIndex        =   0
      Top             =   360
      Width           =   735
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim temp As POINTAPI
Dim temp2 As Long, playTemp As Long
Dim MouseXPos(10000) As Long
Dim MouseYPos(10000) As Long
Private Sub Command1_Click()
If Command1.Caption = "Start" Then
tmrRecord.Enabled = True
Command1.Caption = "Stop"
ElseIf Command1.Caption = "Stop" Then
tmrRecord.Enabled = False
Command1.Caption = "Start"
End If
End Sub

Private Sub Command2_Click()
If Command2.Caption = "Play" Then
tmrPlay.Enabled = True
Command2.Caption = "Stop"
ElseIf Command2.Caption = "Stop" Then
tmrPlay.Enabled = False
Command2.Caption = "Play"
End If
End Sub

Private Sub tmrPlay_Timer()
If MouseXPos(playTemp) < 1 Then
playTemp = 0
Command2.Caption = "Play"
Text1.Text = "X: 0  Y: 0 State: Nothing"
tmrPlay.Enabled = False
Else
Text1.Text = "X: " & MouseXPos(playTemp) & " Y: " & MouseYPos(playTemp) & " Play State: " & playTemp
SetCursorPos MouseXPos(playTemp), MouseYPos(playTemp)
playTemp = playTemp + 1
End If
End Sub

Private Sub tmrRecord_Timer()
GetCursorPos temp
MouseXPos(temp2) = temp.x
MouseYPos(temp2) = temp.y
temp2 = temp2 + 1
Text1.Text = "X: " & temp.x & " Y: " & temp.y & " Rec State: " & temp2
End Sub

