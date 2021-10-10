VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form2 
   Caption         =   "STUDENT RESULT MANAGEMENT SYSTEM"
   ClientHeight    =   12315
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   17040
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   Picture         =   "Form2.frx":2C2B9
   ScaleHeight     =   12315
   ScaleMode       =   0  'User
   ScaleWidth      =   14000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      BackColor       =   &H80000001&
      Caption         =   "Score details"
      BeginProperty Font 
         Name            =   "News701 BT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   6615
      Left            =   240
      TabIndex        =   6
      Top             =   4080
      Width           =   13695
      Begin VB.Label Label42 
         BackStyle       =   0  'Transparent
         Caption         =   "PASS"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   495
         Left            =   10080
         TabIndex        =   47
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label Label41 
         BackStyle       =   0  'Transparent
         Caption         =   "Marks"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   10080
         TabIndex        =   46
         Top             =   2400
         Width           =   3015
      End
      Begin VB.Label Label40 
         BackStyle       =   0  'Transparent
         Caption         =   "Marks"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   10080
         TabIndex        =   45
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label39 
         BackStyle       =   0  'Transparent
         Caption         =   "P"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   5280
         TabIndex        =   44
         Top             =   5040
         Width           =   975
      End
      Begin VB.Label Label38 
         BackStyle       =   0  'Transparent
         Caption         =   "P"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   5280
         TabIndex        =   43
         Top             =   4320
         Width           =   975
      End
      Begin VB.Label Label37 
         BackStyle       =   0  'Transparent
         Caption         =   "P"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   5280
         TabIndex        =   42
         Top             =   3600
         Width           =   975
      End
      Begin VB.Label Label36 
         BackStyle       =   0  'Transparent
         Caption         =   "P"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   5280
         TabIndex        =   41
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label35 
         BackStyle       =   0  'Transparent
         Caption         =   "P"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   5280
         TabIndex        =   40
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label34 
         BackStyle       =   0  'Transparent
         Caption         =   "P"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   5280
         TabIndex        =   39
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   "Result"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   495
         Left            =   5040
         TabIndex        =   38
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "Marks"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   495
         Left            =   3240
         TabIndex        =   37
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   "Marks"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   3240
         TabIndex        =   36
         Top             =   5160
         Width           =   975
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Marks"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   3240
         TabIndex        =   35
         Top             =   4440
         Width           =   975
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "Marks"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   3240
         TabIndex        =   34
         Top             =   3720
         Width           =   975
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "Marks"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   3240
         TabIndex        =   33
         Top             =   3000
         Width           =   975
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Marks"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   3240
         TabIndex        =   32
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Marks"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   3240
         TabIndex        =   31
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "Result           :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7560
         TabIndex        =   30
         Top             =   3120
         Width           =   2055
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Percentage    :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7560
         TabIndex        =   29
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Grade            :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7560
         TabIndex        =   28
         Top             =   2400
         Width           =   2055
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Grade Points."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   495
         Left            =   7560
         TabIndex        =   27
         Top             =   960
         Width           =   2775
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H8000000D&
         BorderWidth     =   3
         Height          =   5655
         Left            =   7200
         Top             =   720
         Width           =   6135
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H8000000D&
         BorderWidth     =   3
         Height          =   5655
         Left            =   360
         Top             =   720
         Width           =   6495
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   26
         Top             =   5880
         Width           =   2055
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Additional Sub."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   25
         Top             =   5160
         Width           =   2055
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Subject 3"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   24
         Top             =   4440
         Width           =   2055
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Subject 2"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   23
         Top             =   3720
         Width           =   2055
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Subject 1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   22
         Top             =   3000
         Width           =   2055
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Language 2"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   21
         Top             =   2280
         Width           =   2055
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Language 1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   20
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Subject Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   495
         Left            =   600
         TabIndex        =   19
         Top             =   840
         Width           =   1935
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Personal data"
      BeginProperty Font 
         Name            =   "News706 BT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2535
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   13695
      Begin MSComDlg.CommonDialog Cd1 
         Left            =   11520
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Image Image1 
         Height          =   2055
         Left            =   10560
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "M/Y"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   7560
         TabIndex        =   18
         Top             =   1680
         Width           =   2415
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Scheme"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   7560
         TabIndex        =   17
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Sem"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   7560
         TabIndex        =   16
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Course"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   2400
         TabIndex        =   15
         Top             =   1680
         Width           =   2415
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Scheme : "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4920
         TabIndex        =   14
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Month/Year of Exam :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4920
         TabIndex        =   13
         Top             =   1680
         Width           =   2655
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Semester   : "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4920
         TabIndex        =   12
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Course      :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   2400
         TabIndex        =   10
         Top             =   1080
         Width           =   2535
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "regno:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   2400
         TabIndex        =   9
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Student Name:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Register Number:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Student Result"
      BeginProperty Font 
         Name            =   "News706 BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1095
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   7335
      Begin VB.CommandButton findbtn 
         BackColor       =   &H0000FFFF&
         Caption         =   "CHECK"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox Text1 
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
         Left            =   3240
         TabIndex        =   3
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Register Number"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.CommandButton exitcmd 
      BackColor       =   &H000000FF&
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   19080
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9720
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   12720
      Top             =   480
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
   Begin VB.Image Image2 
      Height          =   9975
      Left            =   4800
      Picture         =   "Form2.frx":8326D
      Stretch         =   -1  'True
      Top             =   840
      Width           =   12135
   End
   Begin VB.Menu mnures 
      Caption         =   "Result"
   End
   Begin VB.Menu mnuadmin 
      Caption         =   "Admin"
   End
   Begin VB.Menu mnuclose 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub findbtn_Click()
Dim pic As String
Adodc1.RecordSource = "select * from Result where Regno = '" + Text1.Text + "'"
Adodc1.Refresh
If Adodc1.Recordset.RecordCount = 0 Then
MsgBox "Invalid Register Number / Data Not Found"
ElseIf Adodc1.Recordset.RecordCount = 1 Then
Frame2.Visible = True
Frame3.Visible = True
Image1.Visible = True
Label4 = Adodc1.Recordset("Regno")
Label5 = Adodc1.Recordset("Studname")
Label10 = Adodc1.Recordset("Course")
Label11 = Adodc1.Recordset("Semester")
Label12 = Adodc1.Recordset("Scheme")
Label13 = Adodc1.Recordset("MonthnYear")
Label26 = Adodc1.Recordset("Lang1")
Label27 = Adodc1.Recordset("Lang2")
Label28 = Adodc1.Recordset("Sub1")
Label29 = Adodc1.Recordset("Sub2")
Label30 = Adodc1.Recordset("Sub3")
Label31 = Adodc1.Recordset("Sub4")
perce = Adodc1.Recordset("Percentage")
Label40.Caption = perce + " % "
Label41 = Adodc1.Recordset("Grade")
Label42 = Adodc1.Recordset("Res")
pic = Adodc1.Recordset("Photo")
Image1.Picture = LoadPicture(pic)
If Val(Label26) < 30 Then
Label34.Caption = "F"
Label34.ForeColor = &HFF&
Else
Label34.Caption = "P"
Label34.ForeColor = &H8000&
End If
If Val(Label27) < 30 Then
Label35.Caption = "F"
Label35.ForeColor = &HFF&
Else
Label35.Caption = "P"
Label35.ForeColor = &H8000&
End If
If Val(Label28) < 30 Then
Label36.Caption = "F"
Label36.ForeColor = &HFF&
Else
Label36.Caption = "P"
Label36.ForeColor = &H8000&
End If
If Val(Label29) < 30 Then
Label37.Caption = "F"
Label37.ForeColor = &HFF&
Else
Label37.Caption = "P"
Label37.ForeColor = &H8000&
End If
If Val(Label30) < 30 Then
Label38.Caption = "F"
Label38.ForeColor = &HFF&
Else
Label38.Caption = "P"
Label38.ForeColor = &H8000&
End If
If Val(Label31) < 30 Then
Label39.Caption = "F"
Label39.ForeColor = &HFF&
Else
Label39.Caption = "P"
Label39.ForeColor = &H8000&
End If
If Label42.Caption = "FAIL" Then
Label41.Caption = "FAIL"
Label41.ForeColor = &HFF&
Label42.ForeColor = &HFF&
Else
Label42.ForeColor = &H8000&
End If
If Label10.Caption = "BA-HEP" Then
Label17.Caption = "History"
Label18.Caption = "Economics"
Label19.Caption = "Politics"
ElseIf Label10.Caption = "BA-HES" Then
Label17.Caption = "History"
Label18.Caption = "Economics"
Label19.Caption = "Socialogy"
ElseIf Label10.Caption = "BA-HEK" Then
Label17.Caption = "History"
Label18.Caption = "Economics"
Label19.Caption = "Opt. Kannada"
ElseIf Label10.Caption = "BA-HEJ" Then
Label17.Caption = "History"
Label18.Caption = "Economics"
Label19.Caption = "Journolism"
ElseIf Label10.Caption = "BSC-PMCs" Then
Label17.Caption = "Physics"
Label18.Caption = "Mathematics"
Label19.Caption = "Comp. Science"
ElseIf Label10.Caption = "BCOM" Then
Label17.Caption = "Subject 1"
Label18.Caption = "Subject 2"
Label19.Caption = "Subject 3"
ElseIf Label10.Caption = "BBA" Then
Label17.Caption = "Subject 1"
Label18.Caption = "Subject 2"
Label19.Caption = "Subject 3"
ElseIf Label10.Caption = "BCA" Then
Label17.Caption = "Comp. Science 1"
Label18.Caption = "Comp. Science 2"
Label19.Caption = "Comp. Science 3"
ElseIf Label10.Caption = "BSC-PCM" Then
Label17.Caption = "Physics"
Label18.Caption = "Chemistry"
Label19.Caption = "MAthematics"
ElseIf Label10.Caption = "BSC-CBZ" Then
Label17.Caption = "Chemistry"
Label18.Caption = "Botany"
Label19.Caption = "Zoology"
End If
End If
End Sub

Private Sub Form_Load()
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
End Sub



Private Sub mnuadmin_Click()
Form9.Show
End Sub

Private Sub mnuclose_Click()
End
End Sub

Private Sub mnures_Click()
Frame1.Visible = True

End Sub
