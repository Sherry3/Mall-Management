VERSION 5.00
Begin VB.Form database_update 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "database_update"
   ClientHeight    =   4170
   ClientLeft      =   7800
   ClientTop       =   3690
   ClientWidth     =   8085
   Icon            =   "database.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "database.frx":1BF62
   ScaleHeight     =   4170
   ScaleWidth      =   8085
   Begin VB.CommandButton cancel 
      Height          =   495
      Left            =   2880
      Picture         =   "database.frx":27ED7
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton save 
      Height          =   495
      Left            =   1560
      Picture         =   "database.frx":2889E
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton add 
      Height          =   495
      Left            =   240
      Picture         =   "database.frx":292C0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFF80&
      Caption         =   "Info"
      Height          =   1575
      Left            =   4320
      TabIndex        =   3
      Top             =   2280
      Width           =   3495
      Begin VB.TextBox addpname 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1800
         TabIndex        =   28
         Top             =   1100
         Width           =   1575
      End
      Begin VB.TextBox addprice 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1800
         TabIndex        =   27
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox addqty 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1800
         TabIndex        =   26
         Top             =   310
         Width           =   1575
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Provider Name"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   1130
         Width           =   1335
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   740
         Width           =   1335
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "ID"
      Height          =   855
      Left            =   240
      TabIndex        =   2
      Top             =   2280
      Width           =   3855
      Begin VB.TextBox addid 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1560
         TabIndex        =   22
         Top             =   340
         Width           =   1815
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Product_ID No"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   370
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFF80&
      Caption         =   "Date"
      Height          =   1695
      Left            =   4320
      TabIndex        =   1
      Top             =   360
      Width           =   3495
      Begin VB.TextBox addidate 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1800
         TabIndex        =   20
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox addedate 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1800
         TabIndex        =   19
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox addmdate 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1800
         TabIndex        =   18
         Top             =   350
         Width           =   1575
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Import Date"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Expiry Date"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   880
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Manufacturing Date"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Name"
      Height          =   1695
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   3855
      Begin VB.ComboBox selecttype 
         ForeColor       =   &H00FF0000&
         Height          =   315
         ItemData        =   "database.frx":29B74
         Left            =   1560
         List            =   "database.frx":29B8D
         TabIndex        =   13
         Top             =   1290
         Width           =   1695
      End
      Begin VB.TextBox addcmpny 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1560
         TabIndex        =   12
         Top             =   820
         Width           =   1695
      End
      Begin VB.TextBox addname 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1560
         TabIndex        =   11
         Top             =   340
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Product Type"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Product Company"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Product Name"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Label Label6 
      Caption         =   "Product Name"
      Height          =   255
      Left            =   4560
      TabIndex        =   15
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Product Name"
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   1200
      Width           =   1215
   End
End
Attribute VB_Name = "database_update"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim rec As New ADODB.Recordset


Private Sub add_Click()
addname.Text = ""
addcmpny.Text = ""
addedate.Text = ""
addmdate.Text = ""
addqty.Text = ""
addprice.Text = ""
addid.Text = ""
addidate.Text = ""
addpname.Text = ""
addname.SetFocus
rec.AddNew
End Sub

Private Sub cancel_Click()
Unload Me
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
Call Connection
main.WindowState = 2
Set rec = New ADODB.Recordset
rec.Open "select*from productlist", cn, adOpenDynamic, adLockOptimistic
rec.AddNew
End Sub

Private Sub save_Click()
rec.Fields("prod_name") = addname.Text
rec.Fields("company") = addcmpny.Text
rec.Fields("exp_date") = addedate.Text
rec.Fields("mfg_date") = addmdate.Text
rec.Fields("prod_qty") = addqty.Text
rec.Fields("prod_price") = addprice.Text
rec.Fields("prod_id") = addid.Text
rec.Fields("provider") = addpname.Text
rec.Fields("imp_date") = addidate.Text
rec.Fields("prod_type") = selecttype.List(selecttype.ListIndex)
rec.Update
rec.MoveNext
End Sub

