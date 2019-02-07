VERSION 5.00
Begin VB.Form new_trans 
   BackColor       =   &H8000000C&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8130
   ClientLeft      =   5040
   ClientTop       =   2055
   ClientWidth     =   12060
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0B3A
   ScaleHeight     =   8130
   ScaleWidth      =   12060
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdprint 
      Height          =   735
      Left            =   9000
      Picture         =   "Form1.frx":424AB
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   7080
      Width           =   1575
   End
   Begin VB.CommandButton Command15 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8640
      Picture         =   "Form1.frx":431EE
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   6240
      Width           =   2295
   End
   Begin VB.CommandButton Command13 
      Height          =   615
      Left            =   9840
      Picture         =   "Form1.frx":44537
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton Command12 
      Height          =   615
      Left            =   8040
      Picture         =   "Form1.frx":44F8D
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton Command11 
      Height          =   375
      Left            =   5160
      Picture         =   "Form1.frx":45975
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   960
      Width           =   855
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3630
      ItemData        =   "Form1.frx":462B9
      Left            =   490
      List            =   "Form1.frx":462BB
      TabIndex        =   24
      Top             =   4225
      Width           =   6450
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   5520
      TabIndex        =   23
      Top             =   3000
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   2760
      TabIndex        =   21
      Top             =   3000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   5520
      TabIndex        =   19
      Top             =   2520
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2280
      TabIndex        =   17
      Top             =   2520
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3360
      TabIndex        =   12
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton Command10 
      Height          =   735
      Left            =   9360
      Picture         =   "Form1.frx":462BD
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4200
      Width           =   735
   End
   Begin VB.CommandButton Command9 
      Height          =   735
      Left            =   10560
      Picture         =   "Form1.frx":46BDC
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      Height          =   735
      Left            =   9360
      Picture         =   "Form1.frx":4750D
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton Command7 
      Height          =   735
      Left            =   8160
      Picture         =   "Form1.frx":47E7D
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton Command6 
      Height          =   735
      Left            =   10560
      Picture         =   "Form1.frx":48603
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      Height          =   735
      Left            =   9360
      Picture         =   "Form1.frx":48F3E
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Height          =   735
      Left            =   8160
      Picture         =   "Form1.frx":49769
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Height          =   735
      Left            =   10560
      Picture         =   "Form1.frx":49F41
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Height          =   735
      Left            =   9360
      Picture         =   "Form1.frx":4A816
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Height          =   735
      Left            =   8160
      Picture         =   "Form1.frx":4B052
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   960
      Width           =   735
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00C0E0FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      X1              =   7920
      X2              =   11520
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00C0E0FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      X1              =   7920
      X2              =   11520
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C0E0FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      X1              =   11520
      X2              =   11520
      Y1              =   600
      Y2              =   5160
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0E0FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      X1              =   7920
      X2              =   7920
      Y1              =   600
      Y2              =   5160
   End
   Begin VB.Line Line15 
      BorderWidth     =   2
      X1              =   6960
      X2              =   6960
      Y1              =   4200
      Y2              =   7920
   End
   Begin VB.Line Line14 
      BorderWidth     =   2
      X1              =   480
      X2              =   480
      Y1              =   4200
      Y2              =   7920
   End
   Begin VB.Line Line13 
      BorderWidth     =   2
      X1              =   480
      X2              =   6960
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Expiry Date:"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   4200
      TabIndex        =   22
      Top             =   3000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Manufacturing Date:"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   720
      TabIndex        =   20
      Top             =   3000
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Quantity:"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   3720
      TabIndex        =   18
      Top             =   2520
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Price:"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   720
      TabIndex        =   16
      Top             =   2520
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   2280
      TabIndex        =   15
      Top             =   2100
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Type:"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   720
      TabIndex        =   14
      Top             =   2040
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   2280
      TabIndex        =   13
      Top             =   1600
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Name:"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   720
      TabIndex        =   11
      Top             =   1560
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter the product code"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   720
      TabIndex        =   10
      Top             =   960
      Width           =   2415
   End
   Begin VB.Line Line12 
      X1              =   6960
      X2              =   6960
      Y1              =   600
      Y2              =   3960
   End
   Begin VB.Line Line11 
      X1              =   480
      X2              =   6960
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line10 
      X1              =   480
      X2              =   6960
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line9 
      X1              =   480
      X2              =   480
      Y1              =   600
      Y2              =   3960
   End
   Begin VB.Line Line8 
      BorderWidth     =   2
      X1              =   480
      X2              =   6960
      Y1              =   7920
      Y2              =   7920
   End
   Begin VB.Line Line7 
      X1              =   6960
      X2              =   6960
      Y1              =   4200
      Y2              =   7920
   End
   Begin VB.Line Line6 
      X1              =   480
      X2              =   6960
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line Line5 
      X1              =   480
      X2              =   480
      Y1              =   4200
      Y2              =   7920
   End
End
Attribute VB_Name = "new_trans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim r As New ADODB.Recordset
Dim ado2 As New ADODB.Connection
Dim r1 As New ADODB.Recordset
Dim ado3 As New ADODB.Connection
Dim r2 As New ADODB.Recordset
Dim i, dis, res, g As Integer

Private Sub cmdprint_Click()
Dim u As Long
For u = 0 To List1.ListCount - 1
Printer.Print List1.List(u)
Next
Printer.EndDoc
r1.MoveLast
 g = r1.Fields("bill_no")
 r1.AddNew
 r1.Fields("bill_no") = g + 1
 r1.Fields("purchase_time") = Time
 r1.Fields("purchase_date") = Date
 r1.Fields("quantity") = i
 r1.Fields("price") = res
 r1.Update
 r1.MoveNext
 End Sub

Private Sub Command1_Click()
If Text3.Visible = False Then
Text1.Text = Text1.Text + "1"
End If
If Text1.Text <> " " Then
Text3.Text = Text3.Text + "1"
End If
End Sub

Private Sub Command10_Click()
If Text3.Visible = False Then
Text1.Text = Text1.Text + "0"
End If
If Text1.Text <> " " Then
Text3.Text = Text3.Text + "0"

End If

End Sub

Private Sub Command11_Click()
Set r = New ADODB.Recordset
r.Open "select*from productlist", cn, adOpenDynamic, adLockOptimistic
Do Until r.EOF
If r.Fields("prod_id").Value = Val(Text1.Text) Then
 Label2.Visible = True
 Label3.Visible = True
 Label4.Visible = True
 Label5.Visible = True
 Label6.Visible = True
 Label7.Visible = True
 Label8.Visible = True
 Label9.Visible = True
 Text2.Visible = True
 Text3.Visible = True
 Text4.Visible = True
 Text5.Visible = True
 Label3.Caption = r.Fields("prod_name").Value
 Label5.Caption = r.Fields("prod_type").Value
 Text2.Text = r.Fields("prod_price").Value
 Text4.Text = r.Fields("mfg_date").Value
 Text5.Text = r.Fields("exp_date").Value
 Text3.SetFocus
 Text3.Text = ""
 Exit Sub
 Else
 r.MoveNext
 End If
 Loop
 MsgBox ("Enter the valid Product Code")
 Text1.Text = ""
 Text1.SetFocus
 Label2.Visible = False
 Label3.Visible = False
 Label4.Visible = False
 Label5.Visible = False
 Label6.Visible = False
 Label7.Visible = False
 Label8.Visible = False
 Label9.Visible = False
 Text2.Visible = False
 Text3.Visible = False
 Text4.Visible = False
 Text5.Visible = False
End Sub

Private Sub Command12_Click()

i = i + 1
List1.AddItem "  " & (i) & "                   " & Label3.Caption & "                                     " & Val(Text3.Text) & "                     " & Val(Text2.Text) & "                   " & Label5.Caption
res = res + Val(Text2.Text) * Val(Text3.Text)
r2.MoveFirst
 Do Until r2.EOF
 If r2.Fields("prod_id").Value = Val(Text1.Text) Then
 r2.Fields("prod_qty").Value = r2.Fields("prod_qty").Value - Val(Text3.Text)
 r2.Update
  Text1.Text = ""
 Text1.SetFocus
 Label2.Visible = False
 Label3.Visible = False
 Label4.Visible = False
 Label5.Visible = False
 Label6.Visible = False
 Label7.Visible = False
 Label8.Visible = False
 Label9.Visible = False
 Text2.Visible = False
 Text3.Visible = False
 Text4.Visible = False
 Text5.Visible = False
 Exit Sub
 Else
 r2.MoveNext
 End If
 Loop
 Text1.Text = ""
 Text1.SetFocus
 Label2.Visible = False
 Label3.Visible = False
 Label4.Visible = False
 Label5.Visible = False
 Label6.Visible = False
 Label7.Visible = False
 Label8.Visible = False
 Label9.Visible = False
 Text2.Visible = False
 Text3.Visible = False
 Text4.Visible = False
 Text5.Visible = False
End Sub

Private Sub Command13_Click()
Unload Me
End Sub

Private Sub Command15_Click()
List1.AddItem "-------------------------------------------------------------------------------------------------------------------------------------------------"
List1.AddItem "      Total item purchased::  " & (i) & "       Discount if any:: " & dis & "%" & "      Total Cost:: " & res & " Rs"
List1.AddItem "-------------------------------------------------------------------------------------------------------------------------------------------------"
List1.AddItem "                                                   Thank you for purchasing                                                                       "
List1.AddItem "                                                      Have a Good Day                                                                       "
End Sub

Private Sub Command2_Click()
If Text3.Visible = False Then
Text1.Text = Text1.Text + "2"
End If
If Text1.Text <> " " Then
Text3.Text = Text3.Text + "2"

End If

End Sub

Private Sub Command3_Click()
If Text3.Visible = False Then
Text1.Text = Text1.Text + "3"
End If
If Text1.Text <> " " Then
Text3.Text = Text3.Text + "3"

End If

End Sub

Private Sub Command4_Click()
If Text3.Visible = False Then
Text1.Text = Text1.Text + "4"
End If
If Text1.Text <> " " Then
Text3.Text = Text3.Text + "4"

End If

End Sub

Private Sub Command5_Click()
If Text3.Visible = False Then
Text1.Text = Text1.Text + "5"
End If
If Text1.Text <> " " Then
Text3.Text = Text3.Text + "5"

End If

End Sub

Private Sub Command6_Click()
If Text3.Visible = False Then
Text1.Text = Text1.Text + "6"
End If
If Text1.Text <> " " Then
Text3.Text = Text3.Text + "6"

End If

End Sub

Private Sub Command7_Click()
If Text3.Visible = False Then
Text1.Text = Text1.Text + "7"
End If
If Text1.Text = " " Then
Text3.Text = Text3.Text + "7"

End If

End Sub

Private Sub Command8_Click()
If Text3.Visible = False Then
Text1.Text = Text1.Text + "8"
End If
If Text1.Text <> " " Then
Text3.Text = Text3.Text + "8"

End If

End Sub

Private Sub Command9_Click()
If Text3.Visible = False Then
Text1.Text = Text1.Text + "9"
End If
If Text1.Text <> " " Then
Text3.Text = Text3.Text + "9"

End If

End Sub

Private Sub Form_Load()
Call Connection
Dim str, str1 As String
Set ado2 = Nothing
Set ado3 = Nothing
ado2.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=database.mdb;Persist Security Info=False"
str = "select*from bill"
r1.Open str, ado2, adOpenDynamic, adLockOptimistic
ado3.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=database.mdb;Persist Security Info=False"
str1 = "select*from productlist"
r2.Open str1, ado3, adOpenDynamic, adLockOptimistic
i = 0
dis = 0
res = 0
List1.AddItem "                                                    FUTURE MALL                                                             "
List1.AddItem "                         23/C SOUTH TUKOGANJ SAROVAR PORTICA                                              "
List1.AddItem "                                                    0731-2493112                                              "
List1.AddItem "-------------------------------------------------------------------------------------------------------------------------------------------------"
List1.AddItem "   S.No             ProductName                     Qty                 Price               Product Type"
List1.AddItem "-------------------------------------------------------------------------------------------------------------------------------------------------"
End Sub


