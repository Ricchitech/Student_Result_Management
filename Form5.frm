VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form5 
   Caption         =   "Update Student Details"
   ClientHeight    =   11040
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   12915
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   Picture         =   "Form5.frx":2C2B9
   ScaleHeight     =   11040
   ScaleWidth      =   12915
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Check"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox Text1 
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   6
      Top             =   480
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   1200
      Width           =   2775
   End
   Begin VB.ComboBox Combo1 
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      ItemData        =   "Form5.frx":77974
      Left            =   3600
      List            =   "Form5.frx":77976
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   1920
      Width           =   2775
   End
   Begin VB.CommandButton cmdsave 
      BackColor       =   &H0000FF00&
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4200
      Width           =   1335
   End
   Begin VB.ComboBox Combo2 
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      ItemData        =   "Form5.frx":77978
      Left            =   3600
      List            =   "Form5.frx":7797A
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   2640
      Width           =   2775
   End
   Begin VB.CommandButton uploadbtn 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Upload Picture"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2760
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   840
      Top             =   9960
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
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
      RecordSource    =   "Select * from Student"
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
   Begin MSComDlg.CommonDialog Cd1 
      Left            =   11040
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   3360
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   -2147483638
      Format          =   103809025
      CurrentDate     =   34700
      MaxDate         =   55153
      MinDate         =   34700
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Name of Student"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   360
      TabIndex        =   11
      Top             =   1200
      Width           =   3015
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Course"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   360
      TabIndex        =   10
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Register  Number"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   360
      TabIndex        =   9
      Top             =   480
      Width           =   3135
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Scheme"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   360
      TabIndex        =   8
      Top             =   2640
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Birth"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   3360
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   2055
      Left            =   8040
      Stretch         =   -1  'True
      Top             =   480
      Width           =   1695
   End
   Begin VB.Menu mnuhm 
      Caption         =   "Home"
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pic As String

Private Sub cmdsave_Click()
Adodc1.Recordset.Fields("Regno").Value = Text1.Text
Adodc1.Recordset.Fields("Studname").Value = Text2.Text
Adodc1.Recordset.Fields("Course").Value = Combo1.Text
Adodc1.Recordset.Fields("Scheme").Value = Combo2.Text
Adodc1.Recordset.Fields("Photo").Value = pic
Adodc1.Recordset.Update
MsgBox "Data Updated"
End Sub

Private Sub Command1_Click()
Adodc1.RecordSource = "select * from Student where Regno = '" + Text1.Text + "'"
Adodc1.Refresh
If Adodc1.Recordset.RecordCount = 0 Then
MsgBox "Invalid Register Number / Data Not Found"
ElseIf Adodc1.Recordset.RecordCount = 1 Then
Text2 = Adodc1.Recordset("Studname")
Combo1 = Adodc1.Recordset("Course")
Combo2 = Adodc1.Recordset("Scheme")
pic = Adodc1.Recordset("Photo")
Image1.Picture = LoadPicture(pic)
 cmdsave.Visible = True
uploadbtn.Visible = True
End If
End Sub

Private Sub Form_Load()
 cmdsave.Visible = False
uploadbtn.Visible = False
Combo1.AddItem "BA-HEP"
Combo1.AddItem "BSC-PMCs"
Combo1.AddItem "BCOM"
Combo1.AddItem "BBA"
Combo1.AddItem "BCA"
Combo1.AddItem "BSC-CBZ"
Combo1.AddItem "BSC-PCM"
Combo1.AddItem "BA-HES"
Combo1.AddItem "BA-HEJ"
Combo1.AddItem "BA-HEK"
Combo2.AddItem "CBCS"
End Sub

Private Sub mnuhm_Click()
Form3.Show
Unload Me
End Sub

Private Sub uploadbtn_Click()
Cd1.ShowOpen
Cd1.Filter = "Jpeg|*.jpg"
pic = Cd1.FileName
Image1.Picture = LoadPicture(pic)
End Sub
