VERSION 5.00
Begin VB.Form offers 
   Caption         =   "OFFERS"
   ClientHeight    =   3795
   ClientLeft      =   6285
   ClientTop       =   4305
   ClientWidth     =   6990
   Icon            =   "offers.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3795
   ScaleWidth      =   6990
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Command1"
      Height          =   195
      Left            =   840
      TabIndex        =   2
      Top             =   5040
      Width           =   135
   End
   Begin VB.Frame Frame1 
      Caption         =   "Current offers"
      Height          =   3135
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   6495
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   2760
         ItemData        =   "offers.frx":058A
         Left            =   120
         List            =   "offers.frx":058C
         TabIndex        =   1
         Top             =   240
         Width           =   6255
      End
   End
End
Attribute VB_Name = "offers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
productinfo.Show
End Sub
