VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form7 
   Caption         =   "Update Result"
   ClientHeight    =   10935
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   17040
   Icon            =   "Form7.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   Picture         =   "Form7.frx":2C2B9
   ScaleHeight     =   10935
   ScaleWidth      =   17040
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
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
      Left            =   3240
      TabIndex        =   13
      Top             =   480
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
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
      Left            =   3240
      TabIndex        =   12
      Top             =   1200
      Width           =   2775
   End
   Begin VB.TextBox Text3 
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
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
      Left            =   3240
      TabIndex        =   11
      Top             =   1920
      Width           =   2775
   End
   Begin VB.TextBox Text4 
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
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
      Left            =   3240
      TabIndex        =   10
      Top             =   2640
      Width           =   2775
   End
   Begin VB.TextBox Text5 
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
      Left            =   3240
      TabIndex        =   9
      Top             =   4800
      Width           =   2775
   End
   Begin VB.TextBox Text6 
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
      Left            =   3240
      TabIndex        =   8
      Top             =   5520
      Width           =   2775
   End
   Begin VB.TextBox Text7 
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
      Left            =   3240
      TabIndex        =   7
      Top             =   6240
      Width           =   2775
   End
   Begin VB.TextBox Text8 
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
      Left            =   3240
      TabIndex        =   6
      Top             =   6960
      Width           =   2775
   End
   Begin VB.TextBox Text9 
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
      Left            =   3240
      TabIndex        =   5
      Top             =   7680
      Width           =   2775
   End
   Begin VB.TextBox Text10 
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
      Left            =   3240
      TabIndex        =   4
      Top             =   8400
      Width           =   2775
   End
   Begin VB.CommandButton cmdcheck 
      BackColor       =   &H0000FFFF&
      Caption         =   "Get Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton cmdsave 
      BackColor       =   &H0000FF00&
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9120
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   3360
      Width           =   2775
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   4080
      Width           =   2775
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   360
      Top             =   10440
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
      RecordSource    =   "select * from Result"
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
      Left            =   10560
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Register Number"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   360
      TabIndex        =   25
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Student Name"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   360
      TabIndex        =   24
      Top             =   1200
      Width           =   2535
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
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   360
      TabIndex        =   23
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Month/Year"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   360
      TabIndex        =   22
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Subject 1"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   360
      TabIndex        =   21
      Top             =   6240
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Languege 2"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   360
      TabIndex        =   20
      Top             =   5520
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Semester"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   360
      TabIndex        =   19
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Languege 1"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   360
      TabIndex        =   18
      Top             =   4800
      Width           =   2175
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Subject 2"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   360
      TabIndex        =   17
      Top             =   6960
      Width           =   1815
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Additional Subject"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   360
      TabIndex        =   16
      Top             =   8400
      Width           =   2895
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Subject 3"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   360
      TabIndex        =   15
      Top             =   7680
      Width           =   1815
   End
   Begin VB.Label Label13 
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
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   360
      TabIndex        =   14
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   2055
      Left            =   8400
      Stretch         =   -1  'True
      Top             =   480
      Width           =   1695
   End
   Begin VB.Menu mnuhm 
      Caption         =   "Home"
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pic As String
Dim grd As String
Dim res As String
Dim perc As Integer
Dim Total As Integer

Private Sub cmdsrch_Click()
Command4.Visible = False
Command5.Visible = False
Command6.Visible = False
Command7.Visible = False
End Sub

Private Sub Command1_Click()
Command4.Visible = True
Command5.Visible = True
Command6.Visible = True
Command7.Visible = True
End Sub

Private Sub Command2_Click()
Form3.Show
Unload Me
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub cmdcheck_Click()
Adodc1.RecordSource = "select * from Result where Regno = '" + Text1.Text + "'"
Adodc1.Refresh
If Adodc1.Recordset.RecordCount = 0 Then
MsgBox "Invalid Register Number / Data Not Found"
ElseIf Adodc1.Recordset.RecordCount = 1 Then
Image1.Visible = True
Text2 = Adodc1.Recordset("Studname")
Text3 = Adodc1.Recordset("Course")
Text4 = Adodc1.Recordset("Scheme")
Combo1 = Adodc1.Recordset("Semester")
Combo2 = Adodc1.Recordset("MonthnYear")
Text5 = Adodc1.Recordset("Lang1")
Text6 = Adodc1.Recordset("Lang2")
Text7 = Adodc1.Recordset("Sub1")
Text8 = Adodc1.Recordset("Sub2")
Text9 = Adodc1.Recordset("Sub3")
Text10 = Adodc1.Recordset("Sub4")
pic = Adodc1.Recordset("Photo")
Image1.Picture = LoadPicture(pic)
cmdsave.Visible = True
End If
End Sub

Private Sub cmdsave_Click()
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Text7.Text = "" Or Text8.Text = "" Or Text9.Text = "" Or Text10.Text = "" Or Combo1.Text = "" Or Combo2.Text = "" Then
MsgBox "Enter Correct Details"
Else
Adodc1.Refresh
With Adodc1.Recordset
Adodc1.RecordSource = "select * from Result where Regno = '" + Text1.Text + "'"
Adodc1.Refresh
Adodc1.Recordset.Fields("Regno").Value = Text1.Text
Adodc1.Recordset.Fields("Studname").Value = Text2.Text
Adodc1.Recordset.Fields("Course").Value = Text3.Text
Adodc1.Recordset.Fields("Scheme").Value = Text4.Text
Adodc1.Recordset.Fields("Semester").Value = Combo1.Text
Adodc1.Recordset.Fields("MonthnYear").Value = Combo2.Text
Adodc1.Recordset.Fields("Lang1").Value = Text5.Text
Adodc1.Recordset.Fields("Lang2").Value = Text6.Text
Adodc1.Recordset.Fields("Sub1").Value = Text7.Text
Adodc1.Recordset.Fields("Sub2").Value = Text8.Text
Adodc1.Recordset.Fields("Sub3").Value = Text9.Text
Adodc1.Recordset.Fields("Sub4").Value = Text10.Text
Adodc1.Recordset.Fields("Photo").Value = pic
Total = Val(Text5) + Val(Text6) + Val(Text7) + Val(Text8) + Val(Text9) + Val(Text10)
perc = (Total / 600) * 100
If perc >= 80 Then
grd = "First Class Distinction"
ElseIf perc >= 60 And perc < 80 Then
grd = "First Class"
ElseIf perc >= 50 And perc < 60 Then
grd = "Second Class"
ElseIf perc >= 35 And perc < 50 Then
grd = "Pass Class"
Else
grd = "Fail"
End If
If Val(Text5) < 30 Or Val(Text6) < 30 Or Val(Text7) < 30 Or Val(Text8) < 30 Or Val(Text9) < 30 Or Val(Text10) < 30 Then
res = "Fail"
Else
res = "Pass"
End If
Adodc1.Recordset.Fields("Percentage").Value = perc
Adodc1.Recordset.Fields("Grade").Value = grd
Adodc1.Recordset.Fields("Res").Value = res
Adodc1.Recordset.Update
MsgBox "Data Updated"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Combo1.Text = ""
Combo2.Text = ""
Image1.Visible = False

Exit Sub
End With
End If
End Sub

Private Sub Form_Load()
cmdsave.Visible = False
Combo1.AddItem "I Semester"
Combo1.AddItem "II Semester"
Combo1.AddItem "III Semester"
Combo1.AddItem "IV Semester"
Combo1.AddItem "V Semester"
Combo1.AddItem "VI Semester"
Combo1.AddItem "VII Semester"
Combo1.AddItem "VIII Semester"
Combo2.AddItem "May/2020"
Combo2.AddItem "June/2020"
Combo2.AddItem "Nov/2020"
Combo2.AddItem "Dec/2020"
Combo2.AddItem "May/2021"
Combo2.AddItem "June/2021"
Combo2.AddItem "Nov/2021"
Combo2.AddItem "Dec/2021"
Combo2.AddItem "May/2022"
Combo2.AddItem "June/2022"
Combo2.AddItem "Nov/2022"
Combo2.AddItem "Dec/2022"
Combo2.AddItem "May/2023"
Combo2.AddItem "June/2023"
Combo2.AddItem "Nov/2023"
Combo2.AddItem "Dec/2023"
Combo2.AddItem "May/2024"
Combo2.AddItem "June/2024"
Combo2.AddItem "Nov/2024"
Combo2.AddItem "Dec/2024"
Combo2.AddItem "May/2025"
Combo2.AddItem "June/2025"
Combo2.AddItem "Nov/2025"
Combo2.AddItem "Dec/2025"
End Sub

Private Sub mnuhm_Click()
Form3.Show
Unload Me
End Sub
