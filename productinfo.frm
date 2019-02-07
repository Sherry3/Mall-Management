VERSION 5.00
Begin VB.Form productinfo 
   Caption         =   "PRODUCT INFO"
   ClientHeight    =   6960
   ClientLeft      =   4470
   ClientTop       =   2040
   ClientWidth     =   10035
   Icon            =   "productinfo.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "productinfo.frx":058A
   ScaleHeight     =   6960
   ScaleWidth      =   10035
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Command2"
      Height          =   195
      Left            =   480
      TabIndex        =   6
      Top             =   9720
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CLICK FOR ANY OFFERS"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4800
      TabIndex        =   2
      Top             =   5520
      Width           =   4935
   End
   Begin VB.Frame Frame2 
      Caption         =   "AVAILABLE PRODUCTS"
      Height          =   4815
      Left            =   4800
      TabIndex        =   1
      Top             =   480
      Width           =   4935
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   4380
         ItemData        =   "productinfo.frx":19130
         Left            =   120
         List            =   "productinfo.frx":19132
         TabIndex        =   4
         Top             =   240
         Width           =   4695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "AVAILABLE BRANDS"
      Height          =   6255
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   4335
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   5460
         ItemData        =   "productinfo.frx":19134
         Left            =   120
         List            =   "productinfo.frx":19136
         TabIndex        =   3
         Top             =   360
         Width           =   4095
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   9720
      Width           =   975
   End
End
Attribute VB_Name = "productinfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim re As New ADODB.Recordset

Private Sub Command1_Click()
offers.Show
productinfo.Visible = False
Call Connection
Set re = New ADODB.Recordset
re.Open "select * from offer", cn, adOpenDynamic, adLockOptimistic
Do Until re.EOF
If re.Fields("section").Value = Label3.Caption Then
offers.List1.AddItem (re.Fields("offer"))
 re.MoveNext
 Else
 re.MoveNext
 End If
 Loop
End Sub


Private Sub Command2_Click()
productlist.Visible = True
Unload Me
End Sub

Private Sub Form_Load()
Label3.Visible = False
End Sub


Private Sub List1_Click()
List2.Clear
Call Connection
Set re = New ADODB.Recordset
re.Open "select * from productlist", cn, adOpenDynamic, adLockOptimistic
Do Until re.EOF
If re.Fields("prod_type").Value = Label3.Caption And re.Fields("company") = List1.Text Then
List2.AddItem (re.Fields("prod_name"))
 re.MoveNext
 Else
 re.MoveNext
 End If
 Loop
End Sub
