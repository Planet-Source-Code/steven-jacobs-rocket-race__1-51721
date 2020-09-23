VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6120
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6345
   LinkTopic       =   "Form1"
   ScaleHeight     =   6120
   ScaleWidth      =   6345
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Left            =   5880
      Top             =   5400
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   3480
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   3
      Top             =   4680
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2400
      Picture         =   "Form1.frx":0442
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   2
      Top             =   4680
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   495
      Left            =   3240
      TabIndex        =   1
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start Race"
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label Label2 
      Height          =   975
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label Label1 
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   5400
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Simple "Rocket Race" Program
'For educational purposes only
'Feel free to adjust/amend/etc. at will
'Programmer: Steven Jacobs
'Date:  02/13/2004

Dim iCounter As Integer
Dim distance As Integer

Private Sub Command1_Click()

Picture1.Top = 4680 'starting position of picture1
Picture1.Top = 4680 'starting position of picture2
Timer1.Enabled = True
Timer1.Interval = 1
Command1.Enabled = False

End Sub

Private Sub Command2_Click()
Form_QueryUnload 0, 1
End Sub

Private Sub Form_Load()
Form1.Caption = "Rocket Race Program"
iCounter = 1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If MsgBox("Are you sure you want to quit??", vbYesNoCancel + vbQuestion, "Quitting Already?") = vbYes Then
If Timer1.Enabled = True Then
Timer1.Enabled = False
End If
MsgBox "Bye"
Form_Unload (0)
End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Timer1_Timer()

'use timer instead of a recursive or conditional looping structure

iCounter = iCounter + 1

Randomize

If Picture1.Top = 0 And Picture2.Top = 0 Then
Label2.Caption = "Race Status: Its A Tie!!!"
Timer1.Enabled = False
Command1.Enabled = True
ElseIf Picture1.Top = 0 Then
Label2.Caption = "Race Status: Rocket #1 Wins"
Timer1.Enabled = False
Command1.Enabled = True
ElseIf Picture2.Top = 0 Then
Label2.Caption = "Race Status: Rocket #2 Wins"
Timer1.Enabled = False
Command1.Enabled = True
Else

Picture1.Top = (Picture1.Top - Rnd(iCounter))
Picture2.Top = (Picture2.Top - Rnd(iCounter))

If Picture1.Top < Picture2.Top Then
DoEvents

distance = CInt(Picture2.Top - Picture1.Top)
Label2.Caption = "Race Status: " & vbCrLf & "Rocket #1 Leads by: " & CStr(distance)

ElseIf Picture2.Top < Picture1.Top Then
DoEvents

distance = CInt(Picture1.Top - Picture2.Top)
Label2.Caption = "Race Status : " & vbCrLf & "Rocket #2 Leads by: " & CStr(distance)
Else
Label2.Caption = "Race Status: Running Neck and Neck"
End If


Label1.Caption = "Rocket #1: " & Picture1.Top & vbCrLf & _
"Rocket #2: " & Picture2.Top
End If
End Sub
