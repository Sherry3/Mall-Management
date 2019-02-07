VERSION 5.00
Begin VB.Form payroll 
   Caption         =   "PAYROLL"
   ClientHeight    =   3705
   ClientLeft      =   7200
   ClientTop       =   4305
   ClientWidth     =   6900
   Icon            =   "payroll.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3705
   ScaleWidth      =   6900
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1080
      ItemData        =   "payroll.frx":058A
      Left            =   480
      List            =   "payroll.frx":058C
      TabIndex        =   5
      Top             =   2040
      Width           =   5895
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   5400
      Picture         =   "payroll.frx":058E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   760
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Info"
      Height          =   855
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   5775
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2400
         TabIndex        =   2
         Top             =   300
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Enter the Employee Name"
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Line Line4 
      X1              =   6360
      X2              =   6360
      Y1              =   1560
      Y2              =   1920
   End
   Begin VB.Line Line3 
      X1              =   480
      X2              =   480
      Y1              =   1560
      Y2              =   1920
   End
   Begin VB.Line Line2 
      X1              =   480
      X2              =   6360
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line1 
      X1              =   480
      X2              =   6360
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Emp_Name             Month              Date              Basic Salary            Allowances"
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   1680
      Width           =   5775
   End
End
Attribute VB_Name = "payroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim res As New ADODB.Recordset

Private Sub Command11_Click()

 MsgBox ("Enter the valid Product Code")
End Sub

Private Sub Command1_Click()
Do Until res.EOF
If res.Fields("fullname").Value = (Text1.Text) Then
 List1.AddItem res.Fields("fullname") & "     " & res.Fields("month") & "    " & Date & "   " & res.Fields("salary")
 Exit Sub
 Else
 res.MoveNext
 End If
 Loop
 MsgBox ("User Not Found!!!!")
End Sub

Private Sub Form_Load()
Call Connection
Set res = New ADODB.Recordset
res.Open "select*from pay_roll", cn, adOpenDynamic, adLockOptimistic
End Sub

