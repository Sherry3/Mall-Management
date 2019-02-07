VERSION 5.00
Begin VB.Form screen_update 
   Caption         =   "Update Screen Info"
   ClientHeight    =   3555
   ClientLeft      =   5895
   ClientTop       =   3765
   ClientWidth     =   8145
   Icon            =   "screen_update.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "screen_update.frx":058A
   ScaleHeight     =   3555
   ScaleWidth      =   8145
   Begin VB.Frame screen_update 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Info"
      Height          =   2175
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   7095
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "screen_update.frx":19130
         Left            =   1680
         List            =   "screen_update.frx":19149
         TabIndex        =   5
         Text            =   "Food"
         Top             =   1680
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Height          =   375
         Left            =   3720
         Picture         =   "screen_update.frx":19195
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1600
         Width           =   735
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   975
         Left            =   2880
         TabIndex        =   2
         Top             =   360
         Width           =   3975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter the Section"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter the Information to be display"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   2535
      End
   End
End
Attribute VB_Name = "screen_update"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim reu As New ADODB.Recordset

Private Sub Command1_Click()
Call Connection
Set reu = New ADODB.Recordset
reu.Open "select * from offer", cn, adOpenDynamic, adLockOptimistic
reu.AddNew
reu.Fields("section") = Combo1.Text
reu.Fields("offer") = Text1.Text
MsgBox ("Offer Added Successfully")
reu.Update
reu.MoveNext
End Sub
