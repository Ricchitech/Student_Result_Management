VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form9 
   Caption         =   "Admin Login"
   ClientHeight    =   3600
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   6195
   Icon            =   "Form9.frx":0000
   LinkTopic       =   "Form9"
   Picture         =   "Form9.frx":2C2B9
   ScaleHeight     =   3600
   ScaleWidth      =   6195
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtUser 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   840
      Width           =   2535
   End
   Begin VB.TextBox txtPass 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2640
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1560
      Width           =   2535
   End
   Begin VB.CommandButton logincmd 
      BackColor       =   &H00FF00FF&
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2280
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   360
      Top             =   2880
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Ricchi\Desktop\SRMS\result.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Ricchi\Desktop\SRMS\result.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from Login"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Menu mnuend 
      Caption         =   "Home"
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub logincmd_Click()
If txtUser = "" Or txtPass = "" Then
MsgBox "Please Enter Details"
txtUser.SetFocus
Exit Sub
End If
USER = txtUser.Text
PASS = txtPass.Text
Adodc1.RecordSource = "select * from Login where Username = '" + USER + "' and Password = '" + PASS + "'"
Adodc1.Refresh
If Adodc1.Recordset.RecordCount = 0 Then
MsgBox " Login Failed"
txtUser.Text = ""
txtPass.Text = ""
Else
MsgBox "Login Successfull"
Form3.Show
Form2.Hide
Unload Me
Exit Sub
End If
End Sub

Private Sub mnuend_Click()
Form3.Show
Unload Me
End Sub
