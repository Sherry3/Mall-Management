VERSION 5.00
Begin VB.Form prev_form 
   Caption         =   "Search For Bill"
   ClientHeight    =   3990
   ClientLeft      =   6480
   ClientTop       =   3375
   ClientWidth     =   5250
   Icon            =   "prev_form.frx":0000
   LinkTopic       =   "Form3"
   Picture         =   "prev_form.frx":058A
   ScaleHeight     =   3990
   ScaleWidth      =   5250
   Begin VB.Frame Frame2 
      Caption         =   "Details"
      Height          =   1815
      Left            =   720
      TabIndex        =   4
      Top             =   1680
      Width           =   3855
      Begin VB.Label Label5 
         Caption         =   "Total price"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Label Label4 
         Caption         =   "Items purchased"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   2895
      End
      Begin VB.Label Label3 
         Caption         =   "Time::"
         Height          =   375
         Left            =   2040
         TabIndex        =   6
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Date::"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   3600
      Picture         =   "prev_form.frx":4B777
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   740
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Info"
      Height          =   855
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   3855
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   2
         Top             =   290
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Enter The Bill no"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
   End
End
Attribute VB_Name = "prev_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rcs As ADODB.Recordset

Private Sub Command1_Click()
'Do Until rcs.EOF
'If rcs.
End Sub

Private Sub Form_Load()
Call Connection
Set rcs = New ADODB.Recordset
rcs.Open "select*from bill", cn, adOpenDynamic, adLockOptimistic
End Sub
