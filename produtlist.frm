VERSION 5.00
Begin VB.Form productlist 
   Caption         =   "PRODUCTS"
   ClientHeight    =   8565
   ClientLeft      =   5370
   ClientTop       =   1590
   ClientWidth     =   8145
   Icon            =   "produtlist.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "produtlist.frx":058A
   ScaleHeight     =   8565
   ScaleWidth      =   8145
   Begin VB.CommandButton Command15 
      Cancel          =   -1  'True
      Caption         =   "Command15"
      Height          =   135
      Left            =   480
      TabIndex        =   15
      Top             =   9840
      Width           =   75
   End
   Begin VB.Frame Frame1 
      Caption         =   "PRODUCT SECTION"
      Height          =   7935
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7575
      Begin VB.CommandButton Command14 
         Height          =   615
         Left            =   4200
         Picture         =   "produtlist.frx":41EFB
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   6960
         Width           =   2895
      End
      Begin VB.CommandButton Command13 
         Height          =   615
         Left            =   4200
         Picture         =   "produtlist.frx":433FE
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   5880
         Width           =   2895
      End
      Begin VB.CommandButton Command12 
         Height          =   615
         Left            =   4200
         Picture         =   "produtlist.frx":44345
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   4800
         Width           =   2895
      End
      Begin VB.CommandButton Command11 
         Height          =   615
         Left            =   4200
         Picture         =   "produtlist.frx":45BB5
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   3720
         Width           =   2895
      End
      Begin VB.CommandButton Command10 
         Height          =   615
         Left            =   4200
         Picture         =   "produtlist.frx":47500
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   2640
         Width           =   2895
      End
      Begin VB.CommandButton Command9 
         Height          =   615
         Left            =   4200
         Picture         =   "produtlist.frx":48F91
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1560
         Width           =   2895
      End
      Begin VB.CommandButton Command8 
         Height          =   615
         Left            =   4200
         Picture         =   "produtlist.frx":4ABCE
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   480
         Width           =   2895
      End
      Begin VB.CommandButton Command7 
         Height          =   615
         Left            =   360
         Picture         =   "produtlist.frx":4C2C0
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   6960
         Width           =   2895
      End
      Begin VB.CommandButton Command6 
         Height          =   615
         Left            =   360
         Picture         =   "produtlist.frx":4DD5C
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   5880
         Width           =   2895
      End
      Begin VB.CommandButton Command5 
         Height          =   615
         Left            =   360
         Picture         =   "produtlist.frx":4F77E
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   4800
         Width           =   2895
      End
      Begin VB.CommandButton Command4 
         Height          =   615
         Left            =   360
         Picture         =   "produtlist.frx":50E3C
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3720
         Width           =   2895
      End
      Begin VB.CommandButton Command3 
         Height          =   615
         Left            =   360
         Picture         =   "produtlist.frx":51FC2
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2640
         Width           =   2895
      End
      Begin VB.CommandButton Command2 
         Height          =   615
         Left            =   360
         Picture         =   "produtlist.frx":5367A
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1560
         Width           =   2895
      End
      Begin VB.CommandButton Command1 
         Height          =   615
         Left            =   360
         Picture         =   "produtlist.frx":54AD9
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   480
         Width           =   2895
      End
   End
End
Attribute VB_Name = "productlist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim re As New ADODB.Recordset

Private Sub Command1_Click()
productinfo.Show
productlist.Visible = False

Call Connection
Set re = New ADODB.Recordset
re.Open "select * from productlist", cn, adOpenDynamic, adLockOptimistic
Do Until re.EOF
If re.Fields("prod_type").Value = "Appliances" Then
i = 1
productinfo.Label3.Caption = re.Fields("prod_type").Value
productinfo.List1.AddItem (re.Fields("company"))
 re.MoveNext
 Else
 re.MoveNext
 End If
 Loop

End Sub

Private Sub Command10_Click()
productinfo.Show
productlist.Visible = False

Call Connection
Set re = New ADODB.Recordset
re.Open "select * from productlist", cn, adOpenDynamic, adLockOptimistic
Do Until re.EOF
If re.Fields("prod_type").Value = "Computers" Then
i = 1
productinfo.Label3.Caption = re.Fields("prod_type").Value
productinfo.List1.AddItem (re.Fields("company"))
 re.MoveNext
 Else
 re.MoveNext
 End If
 Loop
End Sub

Private Sub Command11_Click()
productinfo.Show
productlist.Visible = False

Call Connection
Set re = New ADODB.Recordset
re.Open "select * from productlist", cn, adOpenDynamic, adLockOptimistic
Do Until re.EOF
If re.Fields("prod_type").Value = "Footwear" Then
i = 1
productinfo.Label3.Caption = re.Fields("prod_type").Value
productinfo.List1.AddItem (re.Fields("company"))
 re.MoveNext
 Else
 re.MoveNext
 End If
 Loop
End Sub

Private Sub Command12_Click()
productinfo.Show
productlist.Visible = False

Call Connection
Set re = New ADODB.Recordset
re.Open "select * from productlist", cn, adOpenDynamic, adLockOptimistic
Do Until re.EOF
If re.Fields("prod_type").Value = "Home & Kitchen" Then
i = 1
productinfo.Label3.Caption = re.Fields("prod_type").Value
productinfo.List1.AddItem (re.Fields("company"))
 re.MoveNext
 Else
 re.MoveNext
 End If
 Loop
End Sub

Private Sub Command13_Click()
productinfo.Show
productlist.Visible = False

Call Connection
Set re = New ADODB.Recordset
re.Open "select * from productlist", cn, adOpenDynamic, adLockOptimistic
Do Until re.EOF
If re.Fields("prod_type").Value = "Food" Then
i = 1
productinfo.Label3.Caption = re.Fields("prod_type").Value
productinfo.List1.AddItem (re.Fields("company"))
 re.MoveNext
 Else
 re.MoveNext
 End If
 Loop
End Sub

Private Sub Command14_Click()
productinfo.Show
productlist.Visible = False
i = 0
Call Connection
Set re = New ADODB.Recordset
re.Open "select * from productlist", cn, adOpenDynamic, adLockOptimistic
Do Until re.EOF
If re.Fields("prod_type").Value = "Cinemas" Then
i = 1
productinfo.Label3.Caption = re.Fields("prod_type").Value
productinfo.List1.AddItem (re.Fields("company"))
 re.MoveNext
 Else
 re.MoveNext
 End If
 Loop
End Sub

Private Sub Command15_Click()
Dialog.Show
Unload Me
End Sub

Private Sub Command2_Click()
productinfo.Show
productlist.Visible = False

Call Connection
Set re = New ADODB.Recordset
re.Open "select * from productlist", cn, adOpenDynamic, adLockOptimistic
Do Until re.EOF
If re.Fields("prod_type").Value = "Fashion" Then
i = 1
productinfo.Label3.Caption = re.Fields("prod_type").Value
productinfo.List1.AddItem (re.Fields("company"))
 re.MoveNext
 Else
 re.MoveNext
 End If
 Loop
End Sub

Private Sub Command3_Click()
productinfo.Show
productlist.Visible = False

Call Connection
Set re = New ADODB.Recordset
re.Open "select * from productlist", cn, adOpenDynamic, adLockOptimistic
Do Until re.EOF
If re.Fields("prod_type").Value = "Grossery" Then
i = 1
productinfo.Label3.Caption = re.Fields("prod_type").Value
productinfo.List1.AddItem (re.Fields("company"))
 re.MoveNext
 Else
 re.MoveNext
 End If
 Loop
End Sub

Private Sub Command4_Click()
productinfo.Show
productlist.Visible = False

Call Connection
Set re = New ADODB.Recordset
re.Open "select * from productlist", cn, adOpenDynamic, adLockOptimistic
Do Until re.EOF
If re.Fields("prod_type").Value = "Books" Then
i = 1
productinfo.Label3.Caption = re.Fields("prod_type").Value
productinfo.List1.AddItem (re.Fields("company"))
 re.MoveNext
 Else
 re.MoveNext
 End If
 Loop
End Sub

Private Sub Command5_Click()
productinfo.Show
productlist.Visible = False

Call Connection
Set re = New ADODB.Recordset
re.Open "select * from productlist", cn, adOpenDynamic, adLockOptimistic
Do Until re.EOF
If re.Fields("prod_type").Value = "Jewellery" Then
i = 1
productinfo.Label3.Caption = re.Fields("prod_type").Value
productinfo.List1.AddItem (re.Fields("company"))
 re.MoveNext
 Else
 re.MoveNext
 End If
 Loop
End Sub

Private Sub Command6_Click()
productinfo.Show
productlist.Visible = False

Call Connection
Set re = New ADODB.Recordset
re.Open "select * from productlist", cn, adOpenDynamic, adLockOptimistic
Do Until re.EOF
If re.Fields("prod_type").Value = "Stationary" Then
i = 1
productinfo.Label3.Caption = re.Fields("prod_type").Value
productinfo.List1.AddItem (re.Fields("company"))
 re.MoveNext
 Else
 re.MoveNext
 End If
 Loop
End Sub

Private Sub Command7_Click()
productinfo.Show
productlist.Visible = False

Call Connection
Set re = New ADODB.Recordset
re.Open "select * from productlist", cn, adOpenDynamic, adLockOptimistic
Do Until re.EOF
If re.Fields("prod_type").Value = "Toys & Games" Then
i = 1
productinfo.Label3.Caption = re.Fields("prod_type").Value
productinfo.List1.AddItem (re.Fields("company"))
 re.MoveNext
 Else
 re.MoveNext
 End If
 Loop
End Sub

Private Sub Command8_Click()
productinfo.Show
productlist.Visible = False

Call Connection
Set re = New ADODB.Recordset
re.Open "select * from productlist", cn, adOpenDynamic, adLockOptimistic
Do Until re.EOF
If re.Fields("prod_type").Value = "Clothing" Then
i = 1
productinfo.Label3.Caption = re.Fields("prod_type").Value
productinfo.List1.AddItem (re.Fields("company"))
 re.MoveNext
 Else
 re.MoveNext
 End If
 Loop
 
End Sub

Private Sub Command9_Click()
Call Connection
Set re = New ADODB.Recordset
re.Open "select * from productlist", cn, adOpenDynamic, adLockOptimistic
Do Until re.EOF
If re.Fields("prod_type").Value = "Electronics" Then
i = 1
productinfo.Label3.Caption = re.Fields("prod_type").Value
productinfo.List1.AddItem (re.Fields("company"))
 re.MoveNext
 Else
 re.MoveNext
 End If
 Loop
 
 productinfo.Show
productlist.Visible = False

End Sub

