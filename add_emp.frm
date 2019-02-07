VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form add_emp 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4425
   ClientLeft      =   6615
   ClientTop       =   3690
   ClientWidth     =   7410
   ClipControls    =   0   'False
   Icon            =   "add_emp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "add_emp.frx":038A
   ScaleHeight     =   4425
   ScaleWidth      =   7410
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6960
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Height          =   495
      Left            =   5760
      Picture         =   "add_emp.frx":18F30
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton cmdadd 
      Height          =   495
      Left            =   4320
      Picture         =   "add_emp.frx":198F7
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H0080FFFF&
      Height          =   975
      Left            =   4320
      TabIndex        =   23
      Top             =   2400
      Width           =   2775
      Begin VB.TextBox emppwd 
         ForeColor       =   &H00FF0000&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1080
         PasswordChar    =   "*"
         TabIndex        =   27
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox empname 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1080
         TabIndex        =   26
         Top             =   200
         Width           =   1455
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   630
         Width           =   855
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "User name"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H0080FFFF&
      Caption         =   "General"
      Height          =   1695
      Left            =   360
      TabIndex        =   13
      Top             =   2400
      Width           =   3735
      Begin VB.TextBox duration 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1440
         TabIndex        =   22
         Top             =   1180
         Width           =   1215
      End
      Begin VB.TextBox experience 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1440
         TabIndex        =   20
         Top             =   700
         Width           =   1815
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Option5"
         Height          =   195
         Left            =   2400
         TabIndex        =   16
         Top             =   340
         Width           =   255
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Option4"
         Height          =   255
         Left            =   1560
         TabIndex        =   15
         Top             =   340
         Width           =   255
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Bond Duration"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Mention if any"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   740
         Width           =   1575
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "No"
         Height          =   255
         Left            =   2760
         TabIndex        =   18
         Top             =   345
         Width           =   495
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Yes"
         Height          =   255
         Left            =   1920
         TabIndex        =   17
         Top             =   345
         Width           =   495
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Work experience"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   15
      Left            =   360
      TabIndex        =   12
      Top             =   2520
      Width           =   255
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Info"
      Height          =   1815
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   4815
      Begin VB.OptionButton Option3 
         Caption         =   "Option3"
         Height          =   195
         Left            =   2640
         TabIndex        =   10
         Top             =   1340
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Option2"
         Height          =   195
         Left            =   1680
         TabIndex        =   8
         Top             =   1340
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   195
         Left            =   1560
         TabIndex        =   7
         Top             =   13450
         Width           =   135
      End
      Begin VB.TextBox age 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   720
         TabIndex        =   6
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox address 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2280
         TabIndex        =   4
         Top             =   810
         Width           =   2295
      End
      Begin VB.TextBox iname 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2280
         TabIndex        =   2
         Top             =   310
         Width           =   2295
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Female"
         Height          =   255
         Left            =   3000
         TabIndex        =   11
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Male"
         Height          =   255
         Left            =   2040
         TabIndex        =   9
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Age"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Employee address"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter the employee name"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   2055
      End
   End
End
Attribute VB_Name = "add_emp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim reu As New ADODB.Recordset
Dim reu1 As New ADODB.Recordset
Dim str As String


Private Sub cmdadd_Click()
Set reu = New ADODB.Recordset
Set reu1 = New ADODB.Recordset
reu.Open "select*from emp_record", cn, adOpenDynamic, adLockOptimistic
reu1.Open "select*from pay_roll", cn, adOpenDynamic, adLockOptimistic
reu.AddNew
reu1.AddNew
reu.Fields("fullname") = iname.Text
reu.Fields("address") = address.Text
reu.Fields("experience") = experience.Text
reu.Fields("age") = age.Text
reu.Fields("duration") = duration.Text
If Option1.Value = True Then
reu.Fields("sex") = Male
ElseIf Option2.Value = True Then
reu.Fields("sex") = Female
End If
reu.Fields("emp_name") = empname.Text
reu.Fields("emp_pwd") = emppwd.Text
reu.Fields("photo") = str
reu1.Fields("fullname") = iname.Text
reu1.Fields("salary") = 5000
reu1.Fields("month") = 1
MsgBox ("User Added Successfully")
reu.Update
reu.MoveNext
reu1.Update
reu1.MoveNext
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call Connection
End Sub

Private Sub Label13_Click()

CommonDialog1.ShowOpen
CommonDialog1.Filter = "JPEG(*.jpeg)|*.jpeg|BITMAP(*.bmp)|*.bmp|PNG(*.png)|*.png"
Image1.Picture = LoadPicture(CommonDialog1.FileName)
str = CommonDialog1.FileName
Label13.Visible = False
Image1.Stretch = False
End Sub
