VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "FUTURE MALL"
   ClientHeight    =   8805
   ClientLeft      =   1905
   ClientTop       =   1140
   ClientWidth     =   15840
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":058A
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnuopen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuprint 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuprintpreview 
         Caption         =   "&Print Preview"
      End
      Begin VB.Menu mnudd 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "&Exit"
         Shortcut        =   %{BKSP}
      End
   End
   Begin VB.Menu mnuview 
      Caption         =   "&View"
      Begin VB.Menu mnusupdate 
         Caption         =   "Screen_Update"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnudataupdate 
         Caption         =   "DataBase_Update"
         Shortcut        =   {F2}
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
      Begin VB.Menu mhelp 
         Caption         =   "Help"
      End
      Begin VB.Menu mnuabout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()

Dialog.Show
Dialog.Left = (Screen.Width / 2) - (Dialog.Width / 2)
Dialog.Top = (Screen.Height / 2) - (Dialog.Height - 50 / 2)

End Sub

Private Sub mnuabout_Click()
help.Show
End Sub

Private Sub mnudataupdate_Click()
database_update.Show
End Sub

Private Sub mnuexit_Click()
Unload Me
End Sub

Private Sub mnuopen_Click()
main.Show

End Sub

Private Sub mnuprint_Click()
Dim u As Long
    
    For u = 0 To new_trans.List1.ListCount - 1
        Printer.Print new_trans.List1.List(u)
    Next
    Printer.EndDoc
End Sub
