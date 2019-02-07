VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   1545
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3750
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmLogin.frx":058A
   ScaleHeight     =   912.837
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtUserName 
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   1290
      TabIndex        =   1
      Top             =   135
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   495
      TabIndex        =   4
      Top             =   1020
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2100
      TabIndex        =   5
      Top             =   1020
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      ForeColor       =   &H00FF0000&
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   525
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&User Name:"
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Password:"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   540
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As New ADODB.Recordset
Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    LoginSucceeded = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
 Set rs = New ADODB.Recordset
 rs.Open "select*from emp_record", cn, adOpenDynamic, adLockOptimistic
 Dim uname As String
 Dim upwd As String
 uname = txtUserName
 upwd = txtPassword
 Do Until rs.EOF
 If rs.Fields("emp_name").Value = uname And rs.Fields("emp_pwd") = upwd Then
 main.StatusBar1.Panels(2).Text = uname
 rs.Fields("status") = 1
 'Dim speaks, speech
 'speaks = "Welcome to Future Mall"
 'Set speech = CreateObject("sapi.spvoice")
 'speech.Speak speaks
 MDIForm1.mnufile = True
 MDIForm1.mnuview = True
 main.Show
 main.WindowState = 2
 LoginSucceeded = True
 txtPassword.Text = ""
 txtUserName = ""
 Me.Hide
 Exit Sub
 Else
 rs.MoveNext
 End If
 Loop
 MsgBox "Invalid Password, try again!", , "Login"
        txtPassword.SetFocus
        txtPassword.Text = ""
 
 
End Sub

Private Sub Form_Load()
Call Connection
End Sub
