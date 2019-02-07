VERSION 5.00
Begin VB.Form search_emp 
   Caption         =   "Employee Informattion"
   ClientHeight    =   3735
   ClientLeft      =   7080
   ClientTop       =   3720
   ClientWidth     =   5700
   Icon            =   "search_emp.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "search_emp.frx":058A
   ScaleHeight     =   3735
   ScaleWidth      =   5700
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Info"
      Height          =   1815
      Left            =   360
      TabIndex        =   3
      Top             =   1560
      Visible         =   0   'False
      Width           =   5055
      Begin VB.Image Image1 
         Height          =   1455
         Left            =   3480
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Years"
         Height          =   255
         Left            =   1680
         TabIndex        =   14
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   1440
         TabIndex        =   13
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   2640
         TabIndex        =   12
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   720
         TabIndex        =   11
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Bond Duration::"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1440
         Width           =   3495
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Sex::"
         Height          =   255
         Left            =   2040
         TabIndex        =   8
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label empsex 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   1560
         TabIndex        =   7
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Age::"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Address::"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   3495
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Name::"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   3495
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "General"
      Height          =   855
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   5055
      Begin VB.CommandButton Command1 
         Height          =   495
         Left            =   4320
         Picture         =   "search_emp.frx":29F2B
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   210
         Width           =   495
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2280
         TabIndex        =   2
         Top             =   310
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter the Employee Name"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   2175
      End
   End
End
Attribute VB_Name = "search_emp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim record As New ADODB.Recordset
Private Sub Command1_Click()
Set record = New ADODB.Recordset
record.Open "select*from emp_record", cn, adOpenDynamic, adLockOptimistic
Do Until record.EOF
If record.Fields("fullname").Value = Text1.Text Then
 Frame2.Visible = True
  Label2.Caption = Label2.Caption + "     " + record.Fields("fullname")
  Label3.Caption = Label3.Caption + "     " + record.Fields("address")
  Label7.Caption = record.Fields("age")
  Label8.Caption = record.Fields("sex")
  Label9.Caption = record.Fields("duration")
  'Image1.Picture = LoadPicture(record.Fields("photo"))
  Exit Sub
  Else
  record.MoveNext
  End If
  Loop
  MsgBox ("Sorry Provided User not found!!!!")
  Text1.SetFocus
  Text1.Text = ""
  Frame2.Visible = False
End Sub

Private Sub Form_Load()
Call Connection
End Sub

