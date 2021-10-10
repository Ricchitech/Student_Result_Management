VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "HOME"
   ClientHeight    =   10635
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   17040
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   Picture         =   "Form3.frx":2C2B9
   ScaleHeight     =   10635
   ScaleWidth      =   17040
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Image Image1 
      Height          =   9975
      Left            =   720
      Picture         =   "Form3.frx":34A92
      Stretch         =   -1  'True
      Top             =   720
      Width           =   11655
   End
   Begin VB.Menu mnustud 
      Caption         =   "Student data"
      Begin VB.Menu mnunew 
         Caption         =   "New Data"
      End
      Begin VB.Menu mnumod 
         Caption         =   "Modify Data"
      End
   End
   Begin VB.Menu menures 
      Caption         =   "Result"
      Begin VB.Menu menuaddnew 
         Caption         =   "New Result"
      End
      Begin VB.Menu menuupdateres 
         Caption         =   "UpdateRes"
      End
   End
   Begin VB.Menu menurep 
      Caption         =   "Result Report"
   End
   Begin VB.Menu mnuclose 
      Caption         =   ">>"
      Begin VB.Menu mnuaddnewuser 
         Caption         =   "Add New User"
      End
      Begin VB.Menu menulg 
         Caption         =   "Logout"
      End
      Begin VB.Menu mexit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub menuaddnew_Click()
Form6.Show
Unload Me
End Sub

Private Sub menulg_Click()
Form2.Show
Unload Me
End Sub

Private Sub menurep_Click()
Form8.Show
Unload Me
End Sub

Private Sub menuupdateres_Click()
Form7.Show
Unload Me
End Sub

Private Sub mexit_Click()
End
End Sub

Private Sub mnumod_Click()
Form5.Show
Unload Me
End Sub

Private Sub mnunew_Click()
Form4.Show
Unload Me
End Sub
