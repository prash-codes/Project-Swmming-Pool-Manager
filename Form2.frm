VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   Caption         =   "VISITORS FORM "
   ClientHeight    =   3015
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   12615
   ScaleWidth      =   21360
   Begin VB.CommandButton Command6 
      BackColor       =   &H0080FF80&
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   18000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   10440
      Width           =   3000
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H008080FF&
      Caption         =   "TIMETABLE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   18000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8520
      Width           =   3000
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H000080FF&
      Caption         =   "FEES STRUCTURE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   18000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6600
      Width           =   3000
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FF8080&
      Caption         =   "MEMBERSHIP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   18000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4680
      Width           =   3000
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF80FF&
      Caption         =   "RULES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   18000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2760
      Width           =   3000
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "INTRODUCTION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   18000
      MaskColor       =   &H00C0C000&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   840
      Width           =   3000
   End
   Begin VB.Image Image1 
      Height          =   12570
      Left            =   0
      Picture         =   "Form2.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   21285
   End
   Begin VB.Menu Entry 
      Caption         =   ""
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Form3.Show

End Sub

Private Sub Command2_Click()
Form4.Show

End Sub

Private Sub Command3_Click()
Form5.Show

End Sub

Private Sub Command4_Click()
Form6.Show

End Sub

Private Sub Command5_Click()
Form7.Show

End Sub

Private Sub Command6_Click()
Form8.Show
Unload Me

End Sub

