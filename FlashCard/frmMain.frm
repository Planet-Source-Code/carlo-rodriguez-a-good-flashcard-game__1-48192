VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FlashCard"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4350
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   4350
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDiv 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Division"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   690
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2490
      Width           =   2745
   End
   Begin VB.CommandButton cmdMulti 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Multiplication"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   690
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1920
      Width           =   2745
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Subtraction"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   690
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1350
      Width           =   2745
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Addition"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   690
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   780
      Width           =   2745
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Quit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1680
      TabIndex        =   1
      Top             =   4080
      Width           =   885
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "FlashCard Main Menu"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   270
      TabIndex        =   8
      Top             =   90
      Width           =   3525
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "FlashCard Main Menu"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   405
      Left            =   300
      TabIndex        =   7
      Top             =   120
      Width           =   3525
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Created by: Carlo S. Rodriguez"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   690
      TabIndex        =   2
      Top             =   3780
      Width           =   2355
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Height          =   4035
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   3825
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   3675
      Left            =   240
      Top             =   60
      Width           =   3645
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   3645
      Left            =   630
      Top             =   390
      Width           =   3495
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
frmAdd.Show
Unload Me
End Sub

Private Sub cmdDiv_Click()
frmDiv.Show
Unload Me
End Sub

Private Sub cmdMulti_Click()
frmMulti.Show
Unload Me
End Sub

Private Sub cmdSub_Click()
frmSub.Show
Unload Me
End Sub

Private Sub Command1_Click()
Unload Me
End Sub
