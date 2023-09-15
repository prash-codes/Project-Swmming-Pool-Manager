VERSION 5.00
Begin VB.Form Form9 
   Caption         =   "Form9"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form9"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   Begin VB.Frame Frame1 
      Caption         =   "CUSTOMER DETAILS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   12495
      Left            =   1200
      TabIndex        =   0
      Top             =   0
      Width           =   15135
      Begin VB.TextBox Text9 
         Height          =   735
         Left            =   3720
         TabIndex        =   24
         Top             =   9840
         Width           =   2535
      End
      Begin VB.OptionButton Option2 
         Caption         =   "FEMALE"
         Height          =   555
         Left            =   11400
         TabIndex        =   21
         Top             =   7320
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "MALE"
         Height          =   615
         Left            =   10200
         TabIndex        =   20
         Top             =   7320
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "CANCLE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1260
         Left            =   5880
         TabIndex        =   18
         Top             =   11160
         Width           =   3000
      End
      Begin VB.CommandButton Command1 
         Caption         =   "SAVE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1260
         Left            =   2040
         TabIndex        =   17
         Top             =   11160
         Width           =   3000
      End
      Begin VB.TextBox Text8 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   3
         EndProperty
         Height          =   600
         Left            =   9240
         TabIndex        =   16
         Top             =   8880
         Width           =   2415
      End
      Begin VB.TextBox Text7 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   3
         EndProperty
         Height          =   600
         Left            =   3720
         TabIndex        =   15
         Top             =   8880
         Width           =   2500
      End
      Begin VB.TextBox Text6 
         Height          =   720
         Left            =   3720
         TabIndex        =   14
         Top             =   7440
         Width           =   1215
      End
      Begin VB.TextBox Text5 
         Height          =   600
         Left            =   3720
         TabIndex        =   13
         Top             =   6240
         Width           =   3495
      End
      Begin VB.TextBox Text4 
         Height          =   1000
         Left            =   3840
         TabIndex        =   12
         Top             =   4560
         Width           =   4935
      End
      Begin VB.TextBox Text3 
         Height          =   600
         Left            =   3840
         TabIndex        =   11
         Top             =   3360
         Width           =   4815
      End
      Begin VB.TextBox Text2 
         Height          =   600
         Left            =   3840
         TabIndex        =   10
         Top             =   2160
         Width           =   6735
      End
      Begin VB.TextBox Text1 
         Height          =   600
         Left            =   3840
         TabIndex        =   9
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label11 
         BackColor       =   &H80000004&
         Caption         =   "FEES"
         Height          =   495
         Left            =   840
         TabIndex        =   23
         Top             =   10080
         Width           =   1575
      End
      Begin VB.Label Label9 
         Caption         =   "GENDER :"
         Height          =   465
         Left            =   9000
         TabIndex        =   19
         Top             =   7560
         Width           =   1005
      End
      Begin VB.Label Label8 
         Caption         =   "EXP DATE"
         Height          =   600
         Left            =   6600
         TabIndex        =   8
         Top             =   8880
         Width           =   2500
      End
      Begin VB.Label Label7 
         Caption         =   "JOINING DATE"
         Height          =   600
         Left            =   840
         TabIndex        =   7
         Top             =   8880
         Width           =   2505
      End
      Begin VB.Label Label6 
         Caption         =   "AGE"
         Height          =   495
         Left            =   840
         TabIndex        =   6
         Top             =   7560
         Width           =   2415
      End
      Begin VB.Label Label5 
         Caption         =   "EMAIL ID"
         Height          =   500
         Left            =   840
         TabIndex        =   5
         Top             =   6240
         Width           =   2505
      End
      Begin VB.Label Label4 
         Caption         =   "ADDRESS"
         Height          =   500
         Left            =   960
         TabIndex        =   4
         Top             =   4800
         Width           =   2505
      End
      Begin VB.Label Label3 
         Caption         =   "PHONE NUMBER"
         Height          =   500
         Left            =   960
         TabIndex        =   3
         Top             =   3480
         Width           =   2505
      End
      Begin VB.Label Label2 
         Caption         =   "FULL NAME"
         Height          =   500
         Left            =   960
         TabIndex        =   2
         Top             =   2160
         Width           =   2505
      End
      Begin VB.Label Label1 
         Caption         =   "BATCH NUMBER"
         Height          =   500
         Left            =   960
         TabIndex        =   1
         Top             =   960
         Width           =   2505
      End
   End
   Begin VB.Label Label10 
      Caption         =   "Label10"
      Height          =   495
      Left            =   10080
      TabIndex        =   22
      Top             =   6240
      Width           =   1215
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rs As Recordset

Private Sub Command1_Click()
Set db = OpenDatabase("D:\Visual Basic6.0\PROJECT\Swiming pool\SWIMMING POOL.mdb")
Set rs = db.OpenRecordset("SELECT * from CUSTOMER")
rs.AddNew
rs.Fields(0).Value = CInt(Text1.Text)

rs.Fields(1).Value = (Text2.Text)

rs.Fields(2).Value = CDbl(Text3.Text)

rs.Fields(3).Value = (Text4.Text)

rs.Fields(4).Value = (Text5.Text)

rs.Fields(5).Value = (Text6.Text)


If Option1.Value = True Then
rs.Fields(6).Value = "Male"
End If


If Option2.Value = True Then
rs.Fields(6).Value = "Female"
End If

rs.Fields(7).Value = (Text7.Text)

rs.Fields(8).Value = (Text8.Text)

rs.Fields(9).Value = CDbl(Text9.Text)


MsgBox ("record saved")
rs.Update


End Sub

Private Sub Command2_Click()
Unload Me
Form8.Show

End Sub

Private Sub Form_Load()
Text1.Text = " "
Text2.Text = " "
Text3.Text = " "
Text4.Text = " "
Text5.Text = " "
Text6.Text = " "
Text7.Text = " "
Text8.Text = " "
Text9.Text = " "

End Sub

