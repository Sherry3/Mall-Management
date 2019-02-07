VERSION 5.00
Begin VB.Form del_emp 
   BackColor       =   &H00FFFF00&
   ClientHeight    =   3120
   ClientLeft      =   7455
   ClientTop       =   3720
   ClientWidth     =   5955
   Icon            =   "del_emp.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "del_emp.frx":038A
   ScaleHeight     =   3120
   ScaleWidth      =   5955
   Begin VB.CommandButton Command3 
      Height          =   495
      Left            =   1200
      Picture         =   "del_emp.frx":41CFB
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Height          =   495
      Left            =   3240
      Picture         =   "del_emp.frx":4277D
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF80&
      Caption         =   "Info"
      Height          =   1695
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   5415
      Begin VB.OptionButton Option2 
         Caption         =   "Option2"
         Height          =   195
         Left            =   1920
         TabIndex        =   7
         Top             =   1330
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   195
         Left            =   1080
         TabIndex        =   5
         Top             =   1330
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox Text1 
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   2160
         TabIndex        =   2
         Top             =   320
         Width           =   2055
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "No"
         Height          =   255
         Left            =   2280
         TabIndex        =   8
         Top             =   1320
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Yes"
         Height          =   255
         Left            =   1440
         TabIndex        =   6
         Top             =   1320
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Dues if any"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter the employee name"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1935
      End
   End
End
Attribute VB_Name = "del_emp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim record As New ADODB.Recordset
Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Set record = New ADODB.Recordset
record.Open "select*from emp_record", cn, adOpenDynamic, adLockOptimistic
Do Until record.EOF
If record.Fields("fullname").Value = Text1.Text Then
 Label2.Visible = True
  Label2.Caption = Label2.Caption + "     " + record.Fields("address")
  Label3.Visible = True
  Label4.Visible = True
  Label5.Visible = True
  Option1.Visible = True
  Option2.Visible = True
  Command3.Picture = LoadPicture("D:\psp\mall management\CNF.jpg")
  If Option2.Value = True Then
  record.Delete
  MsgBox ("User Deleted Successfully")
  End If
  If Option1.Value = True Then
  MsgBox ("Clear the previous dues first!!!")
  End If
  Exit Sub
  Else
  record.MoveNext
  End If
  Loop
  MsgBox ("Sorry Provided User not found!!!!")
  Text1.SetFocus
  Text1.Text = ""
End Sub

Private Sub Form_Load()
Option1.Value = False
Option2.Value = False
Call Connection
End Sub


