VERSION 5.00
Begin VB.Form Form5 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000D&
   Caption         =   "MEMBERSHIP"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form5"
   MDIChild        =   -1  'True
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "# MEMBERSHIP #"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   1590
      Left            =   4830
      TabIndex        =   0
      Top             =   -120
      Width           =   11355
   End
   Begin VB.Image Image1 
      Height          =   10800
      Left            =   2640
      Picture         =   "Form5.frx":0000
      Top             =   1440
      Width           =   14400
   End
   Begin VB.Image Image2 
      Height          =   18840
      Left            =   0
      Picture         =   "Form5.frx":B69D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   30120
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
