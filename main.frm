VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form main 
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   Picture         =   "main.frx":1BF62
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   3240
      Left            =   0
      TabIndex        =   40
      Top             =   0
      Width           =   4560
      _ExtentX        =   8043
      _ExtentY        =   5715
      ButtonWidth     =   1640
      ButtonHeight    =   1852
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   11
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "New"
            Key             =   ""
            Description     =   "Start newly"
            Object.ToolTipText     =   "Start newly"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Billing"
            Key             =   ""
            Object.ToolTipText     =   "Make a transaction"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Database"
            Key             =   ""
            Object.ToolTipText     =   "Update database"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "View"
            Key             =   ""
            Object.ToolTipText     =   "View User/Employee Screen"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Employee"
            Key             =   ""
            Object.ToolTipText     =   "Employee Information"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Announce"
            Key             =   ""
            Object.ToolTipText     =   "Click for any announcement"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Calculator"
            Key             =   ""
            Object.ToolTipText     =   "Use Calculator"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Print"
            Key             =   ""
            Object.ToolTipText     =   "Print Bill"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Restart"
            Key             =   ""
            Object.ToolTipText     =   "Restart the system"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Help"
            Key             =   ""
            Object.ToolTipText     =   "Show help"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Exit"
            Key             =   ""
            Object.ToolTipText     =   "Exit from the program"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
      EndProperty
      Begin VB.PictureBox Picture1 
         Height          =   1095
         Left            =   10320
         Picture         =   "main.frx":C07EE
         ScaleHeight     =   1035
         ScaleWidth      =   9915
         TabIndex        =   42
         Top             =   0
         Width           =   9975
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   7815
      Left            =   360
      TabIndex        =   33
      Top             =   1440
      Width           =   4095
      Begin VB.Line Line5 
         BorderWidth     =   5
         X1              =   0
         X2              =   4080
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Label Label36 
         BackStyle       =   0  'Transparent
         Caption         =   "View Product Section"
         BeginProperty Font 
            Name            =   "Poor Richard"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   39
         Top             =   3360
         Width           =   2655
      End
      Begin VB.Shape Shape12 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   135
         Left            =   360
         Shape           =   3  'Circle
         Top             =   3480
         Width           =   135
      End
      Begin VB.Label Label35 
         BackColor       =   &H000080FF&
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   4440
         Width           =   3855
      End
      Begin VB.Label Label34 
         BackColor       =   &H000080FF&
         Caption         =   "Product Section"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   2880
         Width           =   3855
      End
      Begin VB.Label Label33 
         BackColor       =   &H000080FF&
         Caption         =   "Billing"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   2040
         Width           =   3855
      End
      Begin VB.Label Label32 
         BackColor       =   &H000080FF&
         Caption         =   "Employee Information"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   1200
         Width           =   3855
      End
      Begin VB.Label Label31 
         BackColor       =   &H000080FF&
         Caption         =   "VIEW"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   360
         Width           =   3855
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   7815
      Left            =   360
      TabIndex        =   25
      Top             =   1440
      Width           =   4095
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "View Previous Bill"
         BeginProperty Font 
            Name            =   "Poor Richard"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   32
         Top             =   3000
         Width           =   2415
      End
      Begin VB.Shape Shape10 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   135
         Left            =   360
         Shape           =   3  'Circle
         Top             =   3120
         Width           =   135
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "New Transaction"
         BeginProperty Font 
            Name            =   "Poor Richard"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   31
         Top             =   2520
         Width           =   2055
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   135
         Left            =   360
         Shape           =   3  'Circle
         Top             =   2640
         Width           =   135
      End
      Begin VB.Label Label27 
         BackColor       =   &H000080FF&
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   5760
         Width           =   3855
      End
      Begin VB.Label Label26 
         BackColor       =   &H000080FF&
         Caption         =   "Product Section"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   4920
         Width           =   3855
      End
      Begin VB.Label Label25 
         BackColor       =   &H000080FF&
         Caption         =   "Billing"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   2040
         Width           =   3855
      End
      Begin VB.Label Label24 
         BackColor       =   &H000080FF&
         Caption         =   "Employee Information"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1200
         Width           =   3855
      End
      Begin VB.Label Label23 
         BackColor       =   &H000080FF&
         Caption         =   "VIEW"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   360
         Width           =   3855
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   7815
      Left            =   360
      TabIndex        =   15
      Top             =   1440
      Width           =   4095
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Payroll"
         BeginProperty Font 
            Name            =   "Poor Richard"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   24
         Top             =   3120
         Width           =   1695
      End
      Begin VB.Shape Shape9 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   135
         Left            =   360
         Shape           =   3  'Circle
         Top             =   3240
         Width           =   135
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Delete Employee"
         BeginProperty Font 
            Name            =   "Poor Richard"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   23
         Top             =   2640
         Width           =   2055
      End
      Begin VB.Shape Shape8 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   135
         Left            =   360
         Shape           =   3  'Circle
         Top             =   2760
         Width           =   135
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Search Employee"
         BeginProperty Font 
            Name            =   "Poor Richard"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   22
         Top             =   2160
         Width           =   2295
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Add Employee"
         BeginProperty Font 
            Name            =   "Poor Richard"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   21
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Shape Shape7 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   135
         Left            =   360
         Shape           =   3  'Circle
         Top             =   2280
         Width           =   135
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   135
         Left            =   360
         Shape           =   3  'Circle
         Top             =   1800
         Width           =   135
      End
      Begin VB.Label Label18 
         BackColor       =   &H000080FF&
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   5880
         Width           =   3855
      End
      Begin VB.Label Label17 
         BackColor       =   &H000080FF&
         Caption         =   "Product Section"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   5040
         Width           =   3855
      End
      Begin VB.Label Label16 
         BackColor       =   &H000080FF&
         Caption         =   "Billing"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   4200
         Width           =   3855
      End
      Begin VB.Label Label14 
         BackColor       =   &H000080FF&
         Caption         =   "Employee Information"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1200
         Width           =   3855
      End
      Begin VB.Label Label13 
         BackColor       =   &H000080FF&
         Caption         =   "VIEW"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   3855
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   7815
      Left            =   360
      TabIndex        =   9
      Top             =   1440
      Width           =   4095
      Begin VB.Label Label15 
         BackColor       =   &H000080FF&
         Caption         =   "VIEW"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   3855
      End
      Begin VB.Label Label12 
         BackColor       =   &H000080FF&
         Caption         =   "Employee Information"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1200
         Width           =   3855
      End
      Begin VB.Label Label11 
         BackColor       =   &H000080FF&
         Caption         =   "Billing"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   2040
         Width           =   3855
      End
      Begin VB.Label Label10 
         BackColor       =   &H000080FF&
         Caption         =   "Product Section"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2880
         Width           =   3855
      End
      Begin VB.Label Label9 
         BackColor       =   &H000080FF&
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   3720
         Width           =   3855
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   7815
      Left            =   360
      TabIndex        =   0
      Top             =   1440
      Width           =   4095
      Begin VB.Label Label8 
         BackColor       =   &H000080FF&
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   4440
         Width           =   3855
      End
      Begin VB.Label Label7 
         BackColor       =   &H000080FF&
         Caption         =   "Product Section"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   3720
         Width           =   3855
      End
      Begin VB.Label Label6 
         BackColor       =   &H000080FF&
         Caption         =   "Billing"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   3000
         Width           =   3855
      End
      Begin VB.Label Label5 
         BackColor       =   &H000080FF&
         Caption         =   "Employee Information"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   2280
         Width           =   3855
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Screen Update"
         BeginProperty Font 
            Name            =   "Poor Richard"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   4
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   135
         Left            =   360
         Shape           =   3  'Circle
         Top             =   1560
         Width           =   135
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "DataBase Update"
         BeginProperty Font 
            Name            =   "Poor Richard"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   2
         Top             =   960
         Width           =   2295
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   135
         Left            =   360
         Shape           =   3  'Circle
         Top             =   1080
         Width           =   135
      End
      Begin VB.Label Label1 
         BackColor       =   &H000080FF&
         Caption         =   "VIEW"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   3855
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   41
      Top             =   2520
      Width           =   4560
      _ExtentX        =   8043
      _ExtentY        =   873
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   6
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Object.Width           =   2884
            MinWidth        =   2028
            Picture         =   "main.frx":C907E
            Text            =   "Logged in as:"
            TextSave        =   "Logged in as:"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Current User"
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Object.Width           =   5292
            MinWidth        =   5292
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   3528
            MinWidth        =   3528
            Picture         =   "main.frx":C9790
            TextSave        =   "4/8/2013"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   3175
            MinWidth        =   3175
            Picture         =   "main.frx":C9EA2
            TextSave        =   "3:55 PM"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Bevel           =   0
            Object.Width           =   8468
            MinWidth        =   8468
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "*Developed By: Ashish,Devendra,Pallavi,Sourabh"
            TextSave        =   "*Developed By: Ashish,Devendra,Pallavi,Sourabh"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Poor Richard"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      OLEDropMode     =   1
   End
   Begin VB.Line Line4 
      BorderWidth     =   5
      X1              =   0
      X2              =   4080
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line3 
      BorderWidth     =   5
      X1              =   360
      X2              =   4440
      Y1              =   9240
      Y2              =   9240
   End
   Begin VB.Line Line2 
      BorderWidth     =   5
      X1              =   360
      X2              =   360
      Y1              =   1440
      Y2              =   9240
   End
   Begin VB.Line Line1 
      BorderWidth     =   5
      X1              =   4440
      X2              =   4440
      Y1              =   1440
      Y2              =   9240
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   7440
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   11
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "main.frx":CA5B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "main.frx":CC106
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "main.frx":CDC58
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "main.frx":CF7AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "main.frx":D12FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "main.frx":D2E4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "main.frx":D49A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "main.frx":D64F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "main.frx":D8044
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "main.frx":D9B96
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "main.frx":DB6E8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   480
      Shape           =   3  'Circle
      Top             =   2880
      Width           =   135
   End
   Begin VB.Label Label3 
      Caption         =   "DataBase Update"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   2880
      Width           =   2295
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   480
      Shape           =   3  'Circle
      Top             =   3000
      Width           =   135
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rr As New ADODB.Recordset

Private Sub Command1_Click()
Dialog.Visible = True
main.Visible = False
End Sub

Private Sub Form_Load()
Call Connection
Set rr = New ADODB.Recordset
rr.Open "select*from emp_record", cn, adOpenDynamic, adLockOptimistic
Do Until rr.EOF
If rr.Fields("status") = 1 Then
StatusBar1.Panels(2).Text = rr.Fields("emp_name")
Exit Sub
Else
rr.MoveNext
End If
Loop
Frame2.Visible = True
Frame1.Visible = False
Frame3.Visible = False
Frame4.Visible = False
Frame5.Visible = False
End Sub

Private Sub Label10_Click()
Frame5.Visible = True
Frame2.Visible = False
End Sub

Private Sub Label11_Click()
Frame4.Visible = True
Frame2.Visible = False
End Sub

Private Sub Label12_Click()
Frame3.Visible = True
Frame2.Visible = False
End Sub

Private Sub Label13_Click()
Frame1.Visible = True
Frame3.Visible = False
End Sub

Private Sub Label15_Click()
Frame1.Visible = True
Frame2.Visible = False
End Sub

Private Sub Label16_Click()
Frame4.Visible = True
Frame2.Visible = False
End Sub

Private Sub Label17_Click()
Frame5.Visible = True
Frame3.Visible = False
End Sub

Private Sub Label18_Click()
help.Show
End Sub

Private Sub Label19_Click()
add_emp.Show
End Sub

Private Sub Label2_Click()
database_update.Show
End Sub

Private Sub Label20_Click()
search_emp.Show
End Sub

Private Sub Label21_Click()
del_emp.Show
End Sub

Private Sub Label22_Click()
payroll.Show
End Sub

Private Sub Label23_Click()
Frame1.Visible = True
Frame4.Visible = False
End Sub

Private Sub Label24_Click()
Frame3.Visible = True
Frame4.Visible = False
End Sub

Private Sub Label26_Click()
Frame5.Visible = True
Frame4.Visible = False
End Sub

Private Sub Label27_Click()
help.Show
End Sub

Private Sub Label28_Click()
new_trans.Show
End Sub

Private Sub Label29_Click()
prev_form.Show
End Sub

Private Sub Label30_Click()
search_bill.Show
End Sub

Private Sub Label31_Click()
Frame1.Visible = True
Frame5.Visible = False
End Sub

Private Sub Label32_Click()
Frame3.Visible = True
Frame5.Visible = False
End Sub

Private Sub Label33_Click()
Frame4.Visible = True
Frame5.Visible = False
End Sub

Private Sub Label35_Click()
help.Show
End Sub

Private Sub Label36_Click()
productlist.Show
End Sub

Private Sub Label4_Click()
screen_update.Show
End Sub

Private Sub Label5_Click()
Frame3.Visible = True
Frame1.Visible = False
End Sub

Private Sub Label6_Click()
Frame4.Visible = True
Frame1.Visible = False
End Sub

Private Sub Label7_Click()
Frame5.Visible = True
Frame1.Visible = False
End Sub

Private Sub Label8_Click()
help.Show
End Sub

Private Sub Label9_Click()
help.Show
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
Select Case Button.Caption
Case "New"
      main.Show
Case "Exit"
      Unload Me
      Unload MDIForm1
Case "Billing"
      new_trans.Show
Case "Database"
      database_update.Show
Case "View"
      screen_update.Show
Case "Print"
      Dim u As Long
      For u = 0 To new_trans.List1.ListCount - 1
      Printer.Print new_trans.List1.List(u)
      Next
      Printer.EndDoc
Case "Restart"
      main.Visible = False
      frmLogin.Show
      frmLogin.Left = (Screen.Width / 2) - (frmLogin.Width / 2)
      frmLogin.Top = (Screen.Height / 2) - (frmLogin.Height - 50 / 2)
Case "Employee"
      add_emp.Show
Case "Calculator"
      Shell "calc"
Case "Announce"
      Dim Message, Speak
      Message = InputBox("Enter text", "Speak")
      Set Speak = CreateObject("sapi.spvoice")
      Speak.Speak Message
End Select
End Sub
