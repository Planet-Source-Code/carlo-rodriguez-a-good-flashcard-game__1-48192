VERSION 5.00
Begin VB.Form frmDiv 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FlashCard (Division)"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7740
   Icon            =   "frmflashcarddiv.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   7740
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   2310
      Top             =   3840
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   465
      Left            =   6870
      TabIndex        =   7
      Top             =   3810
      Width           =   795
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next Problem"
      Enabled         =   0   'False
      Height          =   465
      Left            =   6000
      TabIndex        =   6
      Top             =   3810
      Width           =   795
   End
   Begin VB.TextBox txtanswer 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   5640
      MaxLength       =   2
      TabIndex        =   4
      Text            =   "5"
      Top             =   480
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   2310
      Picture         =   "frmflashcarddiv.frx":0442
      Top             =   840
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Score:"
      Height          =   255
      Left            =   1410
      TabIndex        =   8
      Top             =   3540
      Width           =   1095
   End
   Begin VB.Label lblmessage 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3180
      TabIndex        =   5
      Top             =   2460
      Width           =   4095
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   41.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   4440
      TabIndex        =   3
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lblscore 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   41.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1050
      TabIndex        =   2
      Top             =   2340
      Width           =   1965
   End
   Begin VB.Label lblNum2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1215
      Left            =   3240
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label lblNum1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1215
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0080FF80&
      BackStyle       =   1  'Opaque
      Height          =   1755
      Left            =   360
      Top             =   240
      Width           =   7125
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   1755
      Left            =   630
      Top             =   450
      Width           =   7005
   End
End
Attribute VB_Name = "frmDiv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Sum As Integer
Dim NumProb As Integer, NumRight As Integer, NumWrong As Integer
Dim number1 As Integer
Dim number2 As Integer
Private Sub cmdExit_Click()
frmMain.Show
Unload Me
End Sub
Private Sub cmdNext_Click()
Timer1.Enabled = True
txtanswer.Text = ""
lblmessage.Caption = ""
NumProb = NumProb + 1
number1 = (Rnd * 30)
number2 = (Rnd * 30)
Sum = number1 * number2
lblNum1.Caption = Format(Sum, "#0")
lblNum2.Caption = Format(number2, "#0")
cmdNext.Enabled = False
txtanswer.SetFocus
End Sub
Private Sub Form_Activate()
Call cmdNext_Click
End Sub
Private Sub Form_Load()
Randomize Timer
NumProb = 0
NumRight = 0
End Sub

Private Sub Timer1_Timer()
Call Answer
End Sub
Private Sub txtAnswer_KeyPress(KeyAscii As Integer)
Dim Ans As Integer
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack Then
  Exit Sub
ElseIf KeyAscii = vbKeyReturn Then
  Ans = Val(txtanswer.Text)
  If Ans = number1 Then
    NumRight = NumRight + 1
    lblmessage.Caption = "That's correct!"
  Else
    lblmessage.Caption = "Answer is " + Format(number1, "#0")
    NumWrong = NumWrong + 1
  End If
  lblscore.Caption = NumRight - NumWrong
  cmdNext.Enabled = True
  cmdNext.SetFocus
  Timer1.Enabled = False
Else
  KeyAscii = 0
End If
End Sub
Private Sub Answer()
Dim Ans As Integer
If txtanswer.Text = "" Then
  Ans = 0
  End If
  Ans = Val(txtanswer.Text)
  If Ans = number1 Then
    NumRight = NumRight + 1
    lblmessage.Caption = "That's correct!"
  Else
    lblmessage.Caption = "Answer is " + Format(number1, "#0")
    NumWrong = NumWrong + 1
  End If
  lblscore.Caption = NumRight - NumWrong
  cmdNext.Enabled = True
  cmdNext.SetFocus
  Timer1.Enabled = False
 End Sub

