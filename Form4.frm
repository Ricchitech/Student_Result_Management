VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form4 
   Caption         =   "Add Student Details "
   ClientHeight    =   10935
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   17040
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   Picture         =   "Form4.frx":2C2B9
   ScaleHeight     =   10935
   ScaleWidth      =   17040
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog Cd1 
      Left            =   9120
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2640
      Width           =   1455
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
      ItemData        =   "Form4.frx":38F45
      Left            =   3480
      List            =   "Form4.frx":38F47
      Sorted          =   -1  'True
      TabIndex        =   11
      Top             =   2640
      Width           =   2775
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   3480
      TabIndex        =   10
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
      Format          =   113311745
      CurrentDate     =   34700
      MaxDate         =   55153
      MinDate         =   34700
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FF00FF&
      Caption         =   "HOME"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   18840
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   240
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   120
      Top             =   6120
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
      RecordSource    =   "select * from Student"
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
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4200
      Width           =   1335
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
      ItemData        =   "Form4.frx":38F49
      Left            =   3480
      List            =   "Form4.frx":38F4B
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   1920
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
      Left            =   3480
      TabIndex        =   5
      Top             =   1200
      Width           =   2775
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
      Left            =   3480
      TabIndex        =   4
      Top             =   480
      Width           =   2775
   End
   Begin VB.Image Image2 
      Height          =   4935
      Left            =   5640
      Picture         =   "Form4.frx":38F4D
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   9015
   End
   Begin VB.Image Image1 
      Height          =   2055
      Left            =   6840
      Stretch         =   -1  'True
      Top             =   360
      Width           =   1695
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
      Left            =   240
      TabIndex        =   9
      Top             =   3360
      Width           =   3015
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
      Left            =   240
      TabIndex        =   3
      Top             =   2640
      Width           =   3135
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
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   3135
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
      Left            =   240
      TabIndex        =   1
      Top             =   1920
      Width           =   1815
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
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   3015
   End
   Begin VB.Menu mnuhome 
      Caption         =   "Home"
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pic As String

Private Sub cmdsave_Click()
If Text1.Text = "" Or Text2.Text = "" Or Combo1.Text = "" Or Combo2.Text = "" Then
 MsgBox "Enter Correct Details"
 Else
Adodc1.Recordset.AddNew
Adodc1.Recordset.Fields("Regno").Value = Text1.Text
Adodc1.Recordset.Fields("Studname").Value = Text2.Text
Adodc1.Recordset.Fields("Course").Value = Combo1.Text
Adodc1.Recordset.Fields("Scheme").Value = Combo2.Text
Adodc1.Recordset.Fields("DOB").Value = DTPicker1.Value
Adodc1.Recordset.Fields("Photo").Value = pic
Adodc1.Recordset.Update
 MsgBox "Updated"
 Image1.Visible = False
 cmdsave.Visible = False
uploadbtn.Visible = False
Text1.Text = ""
Text2.Text = ""
Combo1.Text = ""
Combo2.Text = ""
 End If
End Sub

Private Sub Combo2_Click()
uploadbtn.Visible = True
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

Private Sub mnuhome_Click()
Form3.Show
Unload Me
End Sub

Private Sub uploadbtn_Click()
Image1.Visible = True
Cd1.ShowOpen
Cd1.Filter = "Jpeg|*.jpg"
pic = Cd1.FileName
Image1.Picture = LoadPicture(pic)
cmdsave.Visible = True
End Sub
