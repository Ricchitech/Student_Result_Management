VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form6 
   Caption         =   "Add New Result"
   ClientHeight    =   10935
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   17040
   Icon            =   "Form6.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   Picture         =   "Form6.frx":2C2B9
   ScaleHeight     =   10935
   ScaleWidth      =   17040
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog Cd1 
      Left            =   10440
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Left            =   3120
      TabIndex        =   26
      Top             =   3960
      Width           =   2775
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
      Left            =   3120
      TabIndex        =   25
      Top             =   3240
      Width           =   2775
   End
   Begin VB.CommandButton cmdsave 
      BackColor       =   &H0000FF00&
      Caption         =   "Save"
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
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   9000
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
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
      Left            =   19200
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   120
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   240
      Top             =   10680
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
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   360
      Width           =   1455
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
      Left            =   3120
      TabIndex        =   20
      Top             =   8280
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
      Left            =   3120
      TabIndex        =   19
      Top             =   7560
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
      Left            =   3120
      TabIndex        =   18
      Top             =   6840
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
      Left            =   3120
      TabIndex        =   17
      Top             =   6120
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
      Left            =   3120
      TabIndex        =   16
      Top             =   5400
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
      Left            =   3120
      TabIndex        =   15
      Top             =   4680
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
      Left            =   3120
      TabIndex        =   14
      Top             =   2520
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
      Left            =   3120
      TabIndex        =   13
      Top             =   1800
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
      Left            =   3120
      TabIndex        =   12
      Top             =   1080
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
      Left            =   3120
      TabIndex        =   11
      Top             =   360
      Width           =   2775
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   1680
      Top             =   10560
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
   Begin VB.Image Image1 
      Height          =   2055
      Left            =   8280
      Stretch         =   -1  'True
      Top             =   360
      Width           =   1695
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
      Height          =   495
      Left            =   240
      TabIndex        =   24
      Top             =   2520
      Width           =   1815
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
      Height          =   495
      Left            =   240
      TabIndex        =   10
      Top             =   7560
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
      Height          =   495
      Left            =   240
      TabIndex        =   9
      Top             =   8280
      Width           =   2895
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
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   6840
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
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   4680
      Width           =   2175
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
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   3240
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
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   5400
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
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   6120
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
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   3960
      Width           =   1815
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
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1800
      Width           =   1815
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
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   2535
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
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2535
   End
   Begin VB.Menu mnuhm 
      Caption         =   "Home"
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pic As String
Dim grd As String
Dim res As String
Dim perc As String
Dim Total As Integer

Private Sub cmdcheck_Click()
If Text1.Text = "" Then
MsgBox "Enter Register Number"
Else
Adodc2.RecordSource = "select * from Result where Regno = '" + Text1.Text + "'"
Adodc2.Refresh
If Adodc2.Recordset.RecordCount = 1 Then
MsgBox "Student Score Already Existed", vbInformation
Else
Adodc1.RecordSource = "select * from Student where Regno = '" + Text1.Text + "'"
Adodc1.Refresh
If Adodc1.Recordset.RecordCount = 0 Then
MsgBox "Invalid Register Number / Data Not Found"
ElseIf Adodc1.Recordset.RecordCount = 1 Then
Image1.Visible = True
Text2 = Adodc1.Recordset("Studname")
Text3 = Adodc1.Recordset("Course")
Text4 = Adodc1.Recordset("Scheme")
pic = Adodc1.Recordset("Photo")
Image1.Picture = LoadPicture(pic)
cmdsave.Visible = True
End If
End If
End If
End Sub

Private Sub cmdsave_Click()
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Text7.Text = "" Or Text8.Text = "" Or Text9.Text = "" Or Text10.Text = "" Or Combo1.Text = "" Or Combo2.Text = "" Then
MsgBox "Enter Correct Details"
Else
Adodc2.RecordSource = "select * from Result where Regno = '" + Text1.Text + "'"
Adodc2.Refresh
If Adodc2.Recordset.RecordCount = 0 Then
With Adodc2.Recordset
Adodc2.Recordset.AddNew
Adodc2.Recordset.Fields("Regno").Value = Text1.Text
Adodc2.Recordset.Fields("Studname").Value = Text2.Text
Adodc2.Recordset.Fields("Course").Value = Text3.Text
Adodc2.Recordset.Fields("Scheme").Value = Text4.Text
Adodc2.Recordset.Fields("Semester").Value = Combo1.Text
Adodc2.Recordset.Fields("MonthnYear").Value = Combo2.Text
Adodc2.Recordset.Fields("Lang1").Value = Text5.Text
Adodc2.Recordset.Fields("Lang2").Value = Text6.Text
Adodc2.Recordset.Fields("Sub1").Value = Text7.Text
Adodc2.Recordset.Fields("Sub2").Value = Text8.Text
Adodc2.Recordset.Fields("Sub3").Value = Text9.Text
Adodc2.Recordset.Fields("Sub4").Value = Text10.Text
Adodc2.Recordset.Fields("Photo").Value = pic
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
res = "FAIL"
Else
res = "PASS"
End If
Adodc2.Recordset.Fields("Percentage").Value = perc
Adodc2.Recordset.Fields("Grade").Value = grd
Adodc2.Recordset.Fields("Res").Value = res
Adodc2.Recordset.Update
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
End With
ElseIf Adodc2.Recordset.RecordCount = 1 Then
MsgBox "Student Score Already Existed", vbInformation
End If
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
