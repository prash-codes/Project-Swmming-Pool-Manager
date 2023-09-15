VERSION 5.00
Begin VB.Form Form8 
   BackColor       =   &H80000005&
   Caption         =   "MANAGER WORK FORM"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form8"
   MDIChild        =   -1  'True
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\Visual Basic6.0\PROJECT\Swiming pool\SWIMMING POOL.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   6600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "CUSTOMER"
      Top             =   5400
      Width           =   5535
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H8000000B&
      Caption         =   "PRINT"
      Height          =   900
      Left            =   11280
      TabIndex        =   10
      Top             =   6480
      Width           =   2500
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SHOW ALL"
      Height          =   900
      Left            =   8520
      TabIndex        =   9
      Top             =   6480
      Width           =   2500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ADD NEW"
      Height          =   900
      Left            =   5640
      TabIndex        =   8
      Top             =   6480
      Width           =   2500
   End
   Begin VB.TextBox Text4 
      DataField       =   "FEES"
      DataSource      =   "Data1"
      Height          =   600
      Left            =   9480
      TabIndex        =   7
      Top             =   4440
      Width           =   2500
   End
   Begin VB.TextBox Text3 
      DataField       =   "EXPDATE"
      DataSource      =   "Data1"
      Height          =   600
      Left            =   9480
      TabIndex        =   6
      Top             =   3360
      Width           =   2500
   End
   Begin VB.TextBox Text2 
      DataField       =   "FULLNAME"
      DataSource      =   "Data1"
      Height          =   600
      Left            =   9480
      TabIndex        =   5
      Top             =   2280
      Width           =   2500
   End
   Begin VB.TextBox Text1 
      DataField       =   "BATCHNUMBER"
      DataSource      =   "Data1"
      Height          =   600
      Left            =   9480
      TabIndex        =   4
      Top             =   1080
      Width           =   2500
   End
   Begin VB.Label Label4 
      Caption         =   "FEES"
      Height          =   600
      Left            =   6720
      TabIndex        =   3
      Top             =   4440
      Width           =   2505
   End
   Begin VB.Label Label3 
      Caption         =   "EXPIRY DATE"
      Height          =   600
      Left            =   6720
      TabIndex        =   2
      Top             =   3360
      Width           =   2505
   End
   Begin VB.Label Label2 
      Caption         =   "NAME"
      Height          =   600
      Left            =   6720
      TabIndex        =   1
      Top             =   2280
      Width           =   2505
   End
   Begin VB.Label Label1 
      Caption         =   "BATCH NUMBER"
      Height          =   600
      Index           =   0
      Left            =   6720
      TabIndex        =   0
      Top             =   1080
      Width           =   2505
   End
   Begin VB.Menu MEMBERSHIP 
      Caption         =   "MEMBERSHIP"
      Begin VB.Menu NEW 
         Caption         =   "NEW"
      End
      Begin VB.Menu OPEN 
         Caption         =   "OPEN"
      End
   End
   Begin VB.Menu BILLING 
      Caption         =   "BILLING"
      Begin VB.Menu SAVEBILL 
         Caption         =   "SAVEBILL"
      End
      Begin VB.Menu PRINTBILL 
         Caption         =   "PRINTBILL"
      End
   End
   Begin VB.Menu EXIT 
      Caption         =   "EXIT"
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form9.Show

End Sub

Private Sub Command2_Click()
Form10.Show

End Sub

Private Sub Command3_Click()
PrintForm

End Sub

Private Sub EXIT_Click()
From8.Show

Unload Me

End Sub

Private Sub Image1_Click()

End Sub

P
End Sub

Private Sub Form_Activate()





Data1.Visible = False
Command3.Visible = False
C = 0
Text3 = Date

Text4 = 0
Data1.Recordset.MoveFirst
For A = 0 To Data1.Recordset.RecordCount - 1
    Data1.Recordset.MoveNext
Next A
End Sub

Private Sub NEW_Click()
Form9.Show

End Sub

Private Sub OPEN_Click()
Form10.Show

End Sub

Private Sub PRINTBILL_Click()
PrintForm

End Sub

Private Sub SAVEBILL_Click()
PrintForm

End Sub

