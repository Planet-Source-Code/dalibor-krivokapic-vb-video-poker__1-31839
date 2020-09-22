VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00008000&
   Caption         =   "VB Video Poker v1.0"
   ClientHeight    =   6705
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   7425
   LinkTopic       =   "Form1"
   ScaleHeight     =   6705
   ScaleWidth      =   7425
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List4 
      Height          =   1425
      Left            =   4320
      Sorted          =   -1  'True
      TabIndex        =   88
      Top             =   5160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ListBox List3 
      Height          =   1230
      Left            =   5520
      TabIndex        =   87
      Top             =   5400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ListBox List2 
      Height          =   1425
      Left            =   2160
      Sorted          =   -1  'True
      TabIndex        =   86
      Top             =   4800
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   6360
      Sorted          =   -1  'True
      TabIndex        =   85
      Top             =   5400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   3600
      Top             =   5760
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   3000
      Top             =   5760
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   855
      Left            =   4320
      TabIndex        =   83
      Top             =   4200
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1508
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmMain.frx":0000
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   0
         Format          =   """€""#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   285
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   80
      TabStop         =   0   'False
      Text            =   "5"
      Top             =   5400
      Width           =   855
   End
   Begin ComCtl2.UpDown UpDown1 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Top             =   5400
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   503
      _Version        =   327681
      Enabled         =   -1  'True
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00008000&
      Height          =   1815
      Left            =   120
      TabIndex        =   71
      Top             =   0
      Width           =   7095
      Begin VB.Label Label14 
         BackColor       =   &H00008000&
         Caption         =   "Straight flush:"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   5280
         TabIndex        =   89
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label13 
         BackColor       =   &H00008000&
         Caption         =   "Flush:"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   5280
         TabIndex        =   79
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label12 
         BackColor       =   &H00008000&
         Caption         =   "Full house:"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   5280
         TabIndex        =   78
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label11 
         BackColor       =   &H00008000&
         Caption         =   "Four of a kind:"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   5280
         TabIndex        =   77
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label10 
         BackColor       =   &H00008000&
         Caption         =   "Flush royal:"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   2760
         TabIndex        =   76
         Top             =   900
         Width           =   1575
      End
      Begin VB.Label Label9 
         BackColor       =   &H00008000&
         Caption         =   "Jacks or better:"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   75
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label8 
         BackColor       =   &H00008000&
         Caption         =   "Two pairs:"
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   120
         TabIndex        =   74
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label7 
         BackColor       =   &H00008000&
         Caption         =   "Three of a kind:"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   73
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label6 
         BackColor       =   &H00008000&
         Caption         =   "Straight:"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   72
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Deal"
      Height          =   375
      Left            =   3060
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   4440
      Width           =   1215
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   5520
      ScaleHeight     =   1545
      ScaleWidth      =   1185
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   2520
      Width           =   1215
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   4320
      ScaleHeight     =   1545
      ScaleWidth      =   1185
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   2520
      Width           =   1215
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   3120
      ScaleHeight     =   1545
      ScaleWidth      =   1185
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   2520
      Width           =   1215
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   1920
      ScaleHeight     =   1545
      ScaleWidth      =   1185
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   2520
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   720
      ScaleHeight     =   1545
      ScaleWidth      =   1185
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   2520
      Width           =   1215
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   1
      Left            =   0
      Picture         =   "frmMain.frx":00AE
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   57
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   2
      Left            =   600
      Picture         =   "frmMain.frx":0578
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   56
      TabStop         =   0   'False
      Tag             =   "2"
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   3
      Left            =   1200
      Picture         =   "frmMain.frx":0A42
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   55
      TabStop         =   0   'False
      Tag             =   "3"
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   4
      Left            =   1800
      Picture         =   "frmMain.frx":0F0C
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   54
      TabStop         =   0   'False
      Tag             =   "4"
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   5
      Left            =   2400
      Picture         =   "frmMain.frx":13D6
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   53
      TabStop         =   0   'False
      Tag             =   "5"
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   6
      Left            =   3000
      Picture         =   "frmMain.frx":18A0
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   52
      TabStop         =   0   'False
      Tag             =   "6"
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   7
      Left            =   3600
      Picture         =   "frmMain.frx":1D6A
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   51
      TabStop         =   0   'False
      Tag             =   "7"
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   8
      Left            =   4200
      Picture         =   "frmMain.frx":2234
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   50
      TabStop         =   0   'False
      Tag             =   "8"
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   9
      Left            =   4800
      Picture         =   "frmMain.frx":26FE
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   49
      TabStop         =   0   'False
      Tag             =   "9"
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   10
      Left            =   5400
      Picture         =   "frmMain.frx":2BC8
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   48
      TabStop         =   0   'False
      Tag             =   "10"
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   11
      Left            =   0
      Picture         =   "frmMain.frx":3092
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   47
      TabStop         =   0   'False
      Tag             =   "12"
      Top             =   2000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   12
      Left            =   600
      Picture         =   "frmMain.frx":3E94
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   46
      TabStop         =   0   'False
      Tag             =   "13"
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   13
      Left            =   1200
      Picture         =   "frmMain.frx":4C96
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   45
      TabStop         =   0   'False
      Tag             =   "14"
      Top             =   2000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   14
      Left            =   1800
      Picture         =   "frmMain.frx":5A98
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   44
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   15
      Left            =   2400
      Picture         =   "frmMain.frx":5F62
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   43
      TabStop         =   0   'False
      Tag             =   "2"
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   16
      Left            =   3000
      Picture         =   "frmMain.frx":642C
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   42
      TabStop         =   0   'False
      Tag             =   "3"
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   17
      Left            =   3600
      Picture         =   "frmMain.frx":68F6
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   41
      TabStop         =   0   'False
      Tag             =   "4"
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   18
      Left            =   4200
      Picture         =   "frmMain.frx":6DC0
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   40
      TabStop         =   0   'False
      Tag             =   "5"
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   19
      Left            =   4800
      Picture         =   "frmMain.frx":728A
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   39
      TabStop         =   0   'False
      Tag             =   "6"
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   20
      Left            =   5400
      Picture         =   "frmMain.frx":7754
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   38
      TabStop         =   0   'False
      Tag             =   "7"
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   21
      Left            =   840
      Picture         =   "frmMain.frx":7C1E
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   37
      TabStop         =   0   'False
      Tag             =   "8"
      Top             =   1080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   22
      Left            =   720
      Picture         =   "frmMain.frx":80E8
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   36
      TabStop         =   0   'False
      Tag             =   "9"
      Top             =   960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   23
      Left            =   840
      Picture         =   "frmMain.frx":85B2
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   35
      TabStop         =   0   'False
      Tag             =   "10"
      Top             =   1320
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   24
      Left            =   6360
      Picture         =   "frmMain.frx":8A7C
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   34
      TabStop         =   0   'False
      Tag             =   "12"
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   25
      Left            =   6360
      Picture         =   "frmMain.frx":987E
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   33
      TabStop         =   0   'False
      Tag             =   "13"
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   26
      Left            =   6360
      Picture         =   "frmMain.frx":A680
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   32
      TabStop         =   0   'False
      Tag             =   "14"
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   27
      Left            =   3600
      Picture         =   "frmMain.frx":B482
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   31
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   28
      Left            =   4200
      Picture         =   "frmMain.frx":B94C
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   30
      TabStop         =   0   'False
      Tag             =   "2"
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   29
      Left            =   4800
      Picture         =   "frmMain.frx":BE16
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   29
      TabStop         =   0   'False
      Tag             =   "3"
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   30
      Left            =   5400
      Picture         =   "frmMain.frx":C2E0
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   28
      TabStop         =   0   'False
      Tag             =   "4"
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   31
      Left            =   120
      Picture         =   "frmMain.frx":C7AA
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   27
      TabStop         =   0   'False
      Tag             =   "5"
      Top             =   960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   32
      Left            =   360
      Picture         =   "frmMain.frx":CC74
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   26
      TabStop         =   0   'False
      Tag             =   "6"
      Top             =   1080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   33
      Left            =   2280
      Picture         =   "frmMain.frx":D13E
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   25
      TabStop         =   0   'False
      Tag             =   "7"
      Top             =   1320
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   34
      Left            =   2760
      Picture         =   "frmMain.frx":D608
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   24
      TabStop         =   0   'False
      Tag             =   "8"
      Top             =   720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   35
      Left            =   2520
      Picture         =   "frmMain.frx":DAD2
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   23
      TabStop         =   0   'False
      Tag             =   "9"
      Top             =   840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   36
      Left            =   3120
      Picture         =   "frmMain.frx":DF9C
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   22
      TabStop         =   0   'False
      Tag             =   "10"
      Top             =   720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   37
      Left            =   3600
      Picture         =   "frmMain.frx":E466
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   21
      TabStop         =   0   'False
      Tag             =   "12"
      Top             =   1000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   38
      Left            =   4320
      Picture         =   "frmMain.frx":F268
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   20
      TabStop         =   0   'False
      Tag             =   "13"
      Top             =   1000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   39
      Left            =   4800
      Picture         =   "frmMain.frx":1006A
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   19
      TabStop         =   0   'False
      Tag             =   "14"
      Top             =   960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   40
      Left            =   5400
      Picture         =   "frmMain.frx":10E6C
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   18
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   41
      Left            =   2040
      Picture         =   "frmMain.frx":11C6E
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   17
      TabStop         =   0   'False
      Tag             =   "2"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   42
      Left            =   2280
      Picture         =   "frmMain.frx":12138
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   16
      TabStop         =   0   'False
      Tag             =   "3"
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   43
      Left            =   1320
      Picture         =   "frmMain.frx":12602
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   15
      TabStop         =   0   'False
      Tag             =   "4"
      Top             =   1320
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   44
      Left            =   1800
      Picture         =   "frmMain.frx":12ACC
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   14
      TabStop         =   0   'False
      Tag             =   "5"
      Top             =   1080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   45
      Left            =   3960
      Picture         =   "frmMain.frx":12F96
      ScaleHeight     =   555
      ScaleWidth      =   915
      TabIndex        =   13
      TabStop         =   0   'False
      Tag             =   "6"
      Top             =   600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   46
      Left            =   2400
      Picture         =   "frmMain.frx":13460
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   12
      TabStop         =   0   'False
      Tag             =   "7"
      Top             =   1440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   47
      Left            =   3720
      Picture         =   "frmMain.frx":1392A
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   11
      TabStop         =   0   'False
      Tag             =   "8"
      Top             =   1320
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   48
      Left            =   4200
      Picture         =   "frmMain.frx":13DF4
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   10
      TabStop         =   0   'False
      Tag             =   "9"
      Top             =   960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   49
      Left            =   4920
      Picture         =   "frmMain.frx":142BE
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   9
      TabStop         =   0   'False
      Tag             =   "10"
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   50
      Left            =   5400
      Picture         =   "frmMain.frx":14788
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   8
      TabStop         =   0   'False
      Tag             =   "12"
      Top             =   10
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   51
      Left            =   840
      Picture         =   "frmMain.frx":1558A
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   7
      TabStop         =   0   'False
      Tag             =   "13"
      Top             =   480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   52
      Left            =   960
      Picture         =   "frmMain.frx":1638C
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   6
      TabStop         =   0   'False
      Tag             =   "14"
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   53
      Left            =   3240
      Picture         =   "frmMain.frx":1718E
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   5
      TabStop         =   0   'False
      Tag             =   "Joker"
      Top             =   1920
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   54
      Left            =   3000
      Picture         =   "frmMain.frx":17F90
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1320
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   55
      Left            =   6000
      Picture         =   "frmMain.frx":18D92
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   56
      Left            =   2280
      Picture         =   "frmMain.frx":19B94
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   57
      Left            =   4080
      Picture         =   "frmMain.frx":1A996
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   960
      Visible         =   0   'False
      Width           =   615
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer3 
      Height          =   375
      Left            =   360
      TabIndex        =   84
      Top             =   4560
      Visible         =   0   'False
      Width           =   615
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   0   'False
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -600
      WindowlessVideo =   0   'False
   End
   Begin VB.Label Label16 
      BackColor       =   &H00008000&
      Caption         =   "0"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1080
      TabIndex        =   82
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Label Label15 
      BackColor       =   &H00008000&
      Caption         =   "Cash:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   480
      TabIndex        =   81
      Top             =   5760
      Width           =   495
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer2 
      Height          =   855
      Left            =   6120
      TabIndex        =   70
      TabStop         =   0   'False
      Top             =   1320
      Visible         =   0   'False
      Width           =   735
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   0   'False
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -600
      WindowlessVideo =   0   'False
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   495
      Left            =   1200
      TabIndex        =   69
      TabStop         =   0   'False
      Top             =   3960
      Visible         =   0   'False
      Width           =   975
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   0   'False
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -600
      WindowlessVideo =   0   'False
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "HOLD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   5505
      TabIndex        =   68
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "HOLD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   4305
      TabIndex        =   67
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "HOLD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3105
      TabIndex        =   66
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "HOLD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1905
      TabIndex        =   65
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "HOLD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   705
      TabIndex        =   64
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Menu mnuGame 
      Caption         =   "&Game"
      Begin VB.Menu mnuNew 
         Caption         =   "New game"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   {F12}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim number1 As Integer
Dim number2 As Integer
Dim number3 As Integer
Dim number4 As Integer
Dim number5 As Integer
Dim num1 As Integer
Dim num2 As Integer
Dim num3 As Integer
Dim num4 As Integer
Dim num5 As Integer
Dim Ulog As Integer
Dim Cash As Integer
Dim Jacks As Integer
Dim Quins As Integer
Dim Kings As Integer
Dim Aces As Integer
Dim Earn As Integer
Dim Dva As Integer
Dim Tri As Integer
Dim Cetiri As Integer
Dim Pet As Integer
Dim Sest As Integer
Dim Sedam As Integer
Dim Osam As Integer
Dim Devet As Integer
Dim Deset As Integer
Dim i As Integer
Dim X1, X2, X3, X4, X5 As Integer

Private Sub Command1_Click()
If Timer1.Enabled = False Then
  If Command1.Caption = "Deal" Then
    If Cash < Ulog Then
      MsgBox "You don't have enough money", vbCritical, "You Looser!!!"
      Exit Sub
    End If
  UpDown1.Enabled = False
  Cash = Cash - Ulog
  Label16.Caption = Cash
  Jacks = 0
  Quins = 0
  Kings = 0
  Aces = 0
  Dva = 0
  Tri = 0
  Cetiri = 0
  Pet = 0
  Sest = 0
  Sedam = 0
  Osam = 0
  Devet = 0
  Deset = 0
  i = 0
  Earn = 0
  Label1.Visible = False
  Label2.Visible = False
  Label3.Visible = False
  Label4.Visible = False
  Label5.Visible = False
  MediaPlayer1.FileName = App.Path & "\" & "hold.wav"
  Picture1.Enabled = True
  Picture2.Enabled = True
  Picture3.Enabled = True
  Picture4.Enabled = True
  Picture5.Enabled = True
    Do
      number1 = Int(Rnd * 52 + 1)
      number2 = Int(Rnd * 52 + 1)
      number3 = Int(Rnd * 52 + 1)
      number4 = Int(Rnd * 52 + 1)
      number5 = Int(Rnd * 52 + 1)
    Loop Until number1 <> number2 And number1 <> number3 And number1 <> number4 And number1 <> number5 And number2 <> number3 And number2 <> number4 And number2 <> number5 And number3 <> number4 And number3 <> number5 And number4 <> number5
    List1.Clear
    List3.Clear
    List4.Clear
    Picture1.Picture = card(number1).Picture
    Picture1.Tag = card(number1).Tag
    X1 = Int(Picture1.Tag)
    List1.AddItem (X1)
    List3.AddItem (number1)
    List1.Refresh
    List3.Refresh
    Picture2.Picture = card(number2).Picture
    Picture2.Tag = card(number2).Tag
    X2 = Int(Picture2.Tag)
    List1.AddItem (X2)
    List3.AddItem (number2)
    List1.Refresh
    List3.Refresh
    Picture3.Picture = card(number3).Picture
    Picture3.Tag = card(number3).Tag
    X3 = Int(Picture3.Tag)
    List1.AddItem (X3)
    List3.AddItem (number3)
    List1.Refresh
    List3.Refresh
    Picture4.Picture = card(number4).Picture
    Picture4.Tag = card(number4).Tag
    X4 = Int(Picture4.Tag)
    List1.AddItem (X4)
    List3.AddItem (number4)
    List1.Refresh
    List3.Refresh
    Picture5.Picture = card(number5).Picture
    Picture5.Tag = card(number5).Tag
    X5 = Int(Picture5.Tag)
    List1.AddItem (X5)
    List3.AddItem (number5)
    List1.Refresh
    List3.Refresh
    List2.Clear
    List2.AddItem (List1.List(0))
    List2.AddItem (List1.List(1))
    List2.AddItem (List1.List(2))
    List2.AddItem (List1.List(3))
    List2.AddItem (List1.List(4))
    List4.Clear
    List4.AddItem (List3.List(0))
    List4.AddItem (List3.List(1))
    List4.AddItem (List3.List(2))
    List4.AddItem (List3.List(3))
    List4.AddItem (List3.List(4))
    Command1.Caption = "Deal again"
  Else
   Do
    num1 = Int(Rnd * 52 + 1)
    num2 = Int(Rnd * 52 + 1)
    num3 = Int(Rnd * 52 + 1)
    num4 = Int(Rnd * 52 + 1)
    num5 = Int(Rnd * 52 + 1)
   Loop Until num1 <> num2 And num1 <> num3 And num1 <> num4 And num1 <> num5 And num2 <> num3 And num2 <> num4 And num2 <> num5 And num3 <> num4 And num3 <> num5 And num4 <> num5 And num1 <> number1 And num1 <> number2 And num1 <> number3 And num1 <> number4 And num1 <> number5 And num2 <> number1 And num2 <> number2 And num2 <> number3 And num2 <> number4 And num2 <> number5 And num3 <> number1 And num3 <> number2 And num3 <> number3 And num3 <> number4 And num3 <> number5 And num4 <> number1 And num4 <> number2 And num4 <> number3 And num4 <> number4 And num4 <> number5 And num5 <> number1 And num5 <> number2 And num5 <> number3 And num5 <> number4 And num5 <> number5
   Picture1.Enabled = False
   Picture2.Enabled = False
   Picture3.Enabled = False
   Picture4.Enabled = False
   Picture5.Enabled = False
   UpDown1.Enabled = True
   Command1.Caption = "Deal"
   If Label1.Visible = False Then
     Picture1.Picture = card(num1).Picture
     Picture1.Tag = card(num1).Tag
     X1 = Int(Picture1.Tag)
     List1.List(0) = X1
     List3.List(0) = num1
     List1.Refresh
     List3.Refresh
   Else
     Picture1.Picture = card(number1).Picture
     Picture1.Tag = card(number1).Tag
     X1 = Int(Picture1.Tag)
     List1.List(0) = X1
     List3.List(0) = number1
     List1.Refresh
     List3.Refresh
   End If
   If Label2.Visible = False Then
     Picture2.Picture = card(num2).Picture
     Picture2.Tag = card(num2).Tag
     X2 = Int(Picture2.Tag)
     List1.List(1) = X2
     List3.List(1) = num2
     List1.Refresh
     List3.Refresh
   Else
     Picture2.Picture = card(number2).Picture
     Picture2.Tag = card(number2).Tag
     X2 = Int(Picture2.Tag)
     List1.List(1) = X2
     List3.List(1) = number2
     List1.Refresh
   End If
   If Label3.Visible = False Then
     Picture3.Picture = card(num3).Picture
     Picture3.Tag = card(num3).Tag
     X3 = Int(Picture3.Tag)
     List1.List(2) = X3
     List3.List(2) = num3
     List1.Refresh
     List3.Refresh
   Else
     Picture3.Picture = card(number3).Picture
     Picture3.Tag = card(number3).Tag
     X3 = Int(Picture3.Tag)
     List1.List(2) = X3
     List3.List(2) = number3
     List1.Refresh
     List3.Refresh
   End If
   If Label4.Visible = False Then
     Picture4.Picture = card(num4).Picture
     Picture4.Tag = card(num4).Tag
     X4 = Int(Picture4.Tag)
     List1.List(3) = X4
     List3.List(3) = num4
     List1.Refresh
     List3.Refresh
   Else
     Picture4.Picture = card(number4).Picture
     Picture4.Tag = card(number4).Tag
     X4 = Int(Picture4.Tag)
     List1.List(3) = X4
     List3.List(3) = number4
     List1.Refresh
     List3.Refresh
   End If
   If Label5.Visible = False Then
     Picture5.Picture = card(num5).Picture
     Picture5.Tag = card(num5).Tag
     X5 = Int(Picture5.Tag)
     List1.List(4) = X5
     List3.List(4) = num5
     List1.Refresh
     List3.Refresh
   Else
     Picture5.Picture = card(number5).Picture
     Picture5.Tag = card(number5).Tag
     X5 = Int(Picture5.Tag)
     List1.List(4) = X5
     List3.List(4) = number5
     List1.Refresh
     List3.Refresh
   End If
   List2.Clear
   List2.AddItem (List1.List(0))
   List2.AddItem (List1.List(1))
   List2.AddItem (List1.List(2))
   List2.AddItem (List1.List(3))
   List2.AddItem (List1.List(4))
   List4.Clear
   List4.AddItem (List3.List(0))
   List4.AddItem (List3.List(1))
   List4.AddItem (List3.List(2))
   List4.AddItem (List3.List(3))
   List4.AddItem (List3.List(4))
   If Picture1.Tag = "12" Then Jacks = Jacks + 1
   If Picture2.Tag = "12" Then Jacks = Jacks + 1
   If Picture3.Tag = "12" Then Jacks = Jacks + 1
   If Picture4.Tag = "12" Then Jacks = Jacks + 1
   If Picture5.Tag = "12" Then Jacks = Jacks + 1
   If Picture1.Tag = "13" Then Quins = Quins + 1
   If Picture2.Tag = "13" Then Quins = Quins + 1
   If Picture3.Tag = "13" Then Quins = Quins + 1
   If Picture4.Tag = "13" Then Quins = Quins + 1
   If Picture5.Tag = "13" Then Quins = Quins + 1
   If Picture1.Tag = "14" Then Kings = Kings + 1
   If Picture2.Tag = "14" Then Kings = Kings + 1
   If Picture3.Tag = "14" Then Kings = Kings + 1
   If Picture4.Tag = "14" Then Kings = Kings + 1
   If Picture5.Tag = "14" Then Kings = Kings + 1
   If Picture1.Tag = "1" Then Aces = Aces + 1
   If Picture2.Tag = "1" Then Aces = Aces + 1
   If Picture3.Tag = "1" Then Aces = Aces + 1
   If Picture4.Tag = "1" Then Aces = Aces + 1
   If Picture5.Tag = "1" Then Aces = Aces + 1
   If Picture1.Tag = "2" Then Dva = Dva + 1
   If Picture2.Tag = "2" Then Dva = Dva + 1
   If Picture3.Tag = "2" Then Dva = Dva + 1
   If Picture4.Tag = "2" Then Dva = Dva + 1
   If Picture5.Tag = "2" Then Dva = Dva + 1
   If Picture1.Tag = "3" Then Tri = Tri + 1
   If Picture2.Tag = "3" Then Tri = Tri + 1
   If Picture3.Tag = "3" Then Tri = Tri + 1
   If Picture4.Tag = "3" Then Tri = Tri + 1
   If Picture5.Tag = "3" Then Tri = Tri + 1
   If Picture1.Tag = "4" Then Cetiri = Cetiri + 1
   If Picture2.Tag = "4" Then Cetiri = Cetiri + 1
   If Picture3.Tag = "4" Then Cetiri = Cetiri + 1
   If Picture4.Tag = "4" Then Cetiri = Cetiri + 1
   If Picture5.Tag = "4" Then Cetiri = Cetiri + 1
   If Picture1.Tag = "5" Then Pet = Pet + 1
   If Picture2.Tag = "5" Then Pet = Pet + 1
   If Picture3.Tag = "5" Then Pet = Pet + 1
   If Picture4.Tag = "5" Then Pet = Pet + 1
   If Picture5.Tag = "5" Then Pet = Pet + 1
   If Picture1.Tag = "6" Then Sest = Sest + 1
   If Picture2.Tag = "6" Then Sest = Sest + 1
   If Picture3.Tag = "6" Then Sest = Sest + 1
   If Picture4.Tag = "6" Then Sest = Sest + 1
   If Picture5.Tag = "6" Then Sest = Sest + 1
   If Picture1.Tag = "7" Then Sedam = Sedam + 1
   If Picture2.Tag = "7" Then Sedam = Sedam + 1
   If Picture3.Tag = "7" Then Sedam = Sedam + 1
   If Picture4.Tag = "7" Then Sedam = Sedam + 1
   If Picture5.Tag = "7" Then Sedam = Sedam + 1
   If Picture1.Tag = "8" Then Osam = Osam + 1
   If Picture2.Tag = "8" Then Osam = Osam + 1
   If Picture3.Tag = "8" Then Osam = Osam + 1
   If Picture4.Tag = "8" Then Osam = Osam + 1
   If Picture5.Tag = "8" Then Osam = Osam + 1
   If Picture1.Tag = "9" Then Devet = Devet + 1
   If Picture2.Tag = "9" Then Devet = Devet + 1
   If Picture3.Tag = "9" Then Devet = Devet + 1
   If Picture4.Tag = "9" Then Devet = Devet + 1
   If Picture5.Tag = "9" Then Devet = Devet + 1
   If Picture1.Tag = "10" Then Deset = Deset + 1
   If Picture2.Tag = "10" Then Deset = Deset + 1
   If Picture3.Tag = "10" Then Deset = Deset + 1
   If Picture4.Tag = "10" Then Deset = Deset + 1
   If Picture5.Tag = "10" Then Deset = Deset + 1
   If Aces = 4 Or Dva = 4 Or Tri = 4 Or Cetiri = 4 Or Pet = 4 Or Sest = 4 Or Sedam = 4 Or Osam = 4 Or Devet = 4 Or Deset = 4 Or Jacks = 4 Or Quins = 4 Or Kings = 4 Then
     Earn = Ulog * 25
     Timer1.Enabled = True
   Else
   If (List4.List(0) = "1" And List4.List(1) = "10" And List4.List(2) = "11" And List4.List(3) = "12" And List4.List(4) = "13") Or (List4.List(0) = "14" And List4.List(1) = "23" And List4.List(2) = "24" And List4.List(3) = "25" And List4.List(4) = "26") Or (List4.List(0) = "27" And List4.List(1) = "36" And List4.List(2) = "37" And List4.List(3) = "38" And List4.List(4) = "39") Or (List4.List(0) = "40" And List4.List(1) = "49" And List4.List(2) = "50" And List4.List(3) = "51" And List4.List(4) = "52") Then
       Earn = Ulog * 100
       Timer1.Enabled = True
   Else
     If (List4.List(0) = "1" And List4.List(1) = "2" And List4.List(2) = "3" And List4.List(3) = "4" And List4.List(4) = "5") Or (List4.List(0) = "14" And List4.List(1) = "15" And List4.List(2) = "16" And List4.List(3) = "17" And List4.List(4) = "18") Or (List4.List(0) = "27" And List4.List(1) = "28" And List4.List(2) = "29" And List4.List(3) = "30" And List4.List(4) = "31") Or (List4.List(0) = "40" And List4.List(1) = "41" And List4.List(2) = "42" And List4.List(3) = "43" And List4.List(4) = "44") Then
       Earn = Ulog * 50
       Timer1.Enabled = True
   Else
    If (List4.List(0) = "2" And List4.List(1) = "3" And List4.List(2) = "4" And List4.List(3) = "5" And List4.List(4) = "6") Or (List4.List(0) = "15" And List4.List(1) = "16" And List4.List(2) = "17" And List4.List(3) = "18" And List4.List(4) = "19") Or (List4.List(0) = "28" And List4.List(1) = "29" And List4.List(2) = "30" And List4.List(3) = "31" And List4.List(4) = "32") Or (List4.List(0) = "41" And List4.List(1) = "42" And List4.List(2) = "43" And List4.List(3) = "44" And List4.List(4) = "45") Then
       Earn = Ulog * 50
       Timer1.Enabled = True
   Else
     If (List4.List(0) = "3" And List4.List(1) = "4" And List4.List(2) = "5" And List4.List(3) = "6" And List4.List(4) = "7") Or (List4.List(0) = "16" And List4.List(1) = "17" And List4.List(2) = "18" And List4.List(3) = "19" And List4.List(4) = "20") Or (List4.List(0) = "29" And List4.List(1) = "30" And List4.List(2) = "31" And List4.List(3) = "32" And List4.List(4) = "33") Or (List4.List(0) = "42" And List4.List(1) = "43" And List4.List(2) = "44" And List4.List(3) = "45" And List4.List(4) = "46") Then
       Earn = Ulog * 50
       Timer1.Enabled = True
   Else
     If (List4.List(0) = "4" And List4.List(1) = "5" And List4.List(2) = "6" And List4.List(3) = "7" And List4.List(4) = "8") Or (List4.List(0) = "17" And List4.List(1) = "18" And List4.List(2) = "19" And List4.List(3) = "20" And List4.List(4) = "21") Or (List4.List(0) = "30" And List4.List(1) = "31" And List4.List(2) = "32" And List4.List(3) = "33" And List4.List(4) = "34") Or (List4.List(0) = "43" And List4.List(1) = "44" And List4.List(2) = "45" And List4.List(3) = "46" And List4.List(4) = "47") Then
       Earn = Ulog * 50
       Timer1.Enabled = True
   Else
     If (List4.List(0) = "5" And List4.List(1) = "6" And List4.List(2) = "7" And List4.List(3) = "8" And List4.List(4) = "9") Or (List4.List(0) = "18" And List4.List(1) = "19" And List4.List(2) = "20" And List4.List(3) = "21" And List4.List(4) = "22") Or (List4.List(0) = "31" And List4.List(1) = "32" And List4.List(2) = "33" And List4.List(3) = "34" And List4.List(4) = "35") Or (List4.List(0) = "44" And List4.List(1) = "45" And List4.List(2) = "46" And List4.List(3) = "47" And List4.List(4) = "48") Then
       Earn = Ulog * 50
       Timer1.Enabled = True
   Else
     If (List4.List(0) = "10" And List4.List(1) = "6" And List4.List(2) = "7" And List4.List(3) = "8" And List4.List(4) = "9") Or (List4.List(0) = "19" And List4.List(1) = "20" And List4.List(2) = "21" And List4.List(3) = "22" And List4.List(4) = "23") Or (List4.List(0) = "32" And List4.List(1) = "33" And List4.List(2) = "34" And List4.List(3) = "35" And List4.List(4) = "36") Or (List4.List(0) = "45" And List4.List(1) = "46" And List4.List(2) = "47" And List4.List(3) = "48" And List4.List(4) = "49") Then
       Earn = Ulog * 50
       Timer1.Enabled = True
    Else
     If (List4.List(0) = "10" And List4.List(1) = "12" And List4.List(2) = "7" And List4.List(3) = "8" And List4.List(4) = "9") Or (List4.List(0) = "20" And List4.List(1) = "21" And List4.List(2) = "22" And List4.List(3) = "23" And List4.List(4) = "24") Or (List4.List(0) = "33" And List4.List(1) = "34" And List4.List(2) = "35" And List4.List(3) = "36" And List4.List(4) = "37") Or (List4.List(0) = "46" And List4.List(1) = "47" And List4.List(2) = "48" And List4.List(3) = "49" And List4.List(4) = "50") Then
       Earn = Ulog * 50
       Timer1.Enabled = True
   Else
     If (List4.List(0) = "10" And List4.List(1) = "12" And List4.List(2) = "13" And List4.List(3) = "8" And List4.List(4) = "9") Or (List4.List(0) = "21" And List4.List(1) = "22" And List4.List(2) = "23" And List4.List(3) = "24" And List4.List(4) = "25") Or (List4.List(0) = "34" And List4.List(1) = "35" And List4.List(2) = "36" And List4.List(3) = "37" And List4.List(4) = "38") Or (List4.List(0) = "47" And List4.List(1) = "48" And List4.List(2) = "49" And List4.List(3) = "50" And List4.List(4) = "51") Then
       Earn = Ulog * 50
       Timer1.Enabled = True
   Else
     If (List4.List(0) = "10" And List4.List(1) = "12" And List4.List(2) = "13" And List4.List(3) = "14" And List4.List(4) = "9") Or (List4.List(0) = "22" And List4.List(1) = "23" And List4.List(2) = "24" And List4.List(3) = "25" And List4.List(4) = "26") Or (List4.List(0) = "35" And List4.List(1) = "36" And List4.List(2) = "37" And List4.List(3) = "38" And List4.List(4) = "39") Or (List4.List(0) = "48" And List4.List(1) = "49" And List4.List(2) = "50" And List4.List(3) = "51" And List4.List(4) = "52") Then
       Earn = Ulog * 50
       Timer1.Enabled = True
     Else
   If List2.List(0) = "1" And List2.List(1) = "2" And List2.List(2) = "3" And List2.List(3) = "4" And List2.List(4) = "5" Then
     Earn = Ulog * 4
     Timer1.Enabled = True
   Else
   If List2.List(0) = "2" And List2.List(1) = "3" And List2.List(2) = "4" And List2.List(3) = "5" And List2.List(4) = "6" Then
     Earn = Ulog * 4
     Timer1.Enabled = True
   Else
   If List2.List(0) = "3" And List2.List(1) = "4" And List2.List(2) = "5" And List2.List(3) = "6" And List2.List(4) = "7" Then
     Earn = Ulog * 4
     Timer1.Enabled = True
   Else
   If List2.List(0) = "4" And List2.List(1) = "5" And List2.List(2) = "6" And List2.List(3) = "7" And List2.List(4) = "8" Then
     Earn = Ulog * 4
     Timer1.Enabled = True
   Else
   If List2.List(0) = "5" And List2.List(1) = "6" And List2.List(2) = "7" And List2.List(3) = "8" And List2.List(4) = "9" Then
     Earn = Ulog * 4
     Timer1.Enabled = True
   Else
   If List2.List(0) = "10" And List2.List(1) = "6" And List2.List(2) = "7" And List2.List(3) = "8" And List2.List(4) = "9" Then
     Earn = Ulog * 4
     Timer1.Enabled = True
   Else
   If List2.List(0) = "10" And List2.List(1) = "12" And List2.List(2) = "7" And List2.List(3) = "8" And List2.List(4) = "9" Then
     Earn = Ulog * 4
     Timer1.Enabled = True
   Else
   If List2.List(0) = "10" And List2.List(1) = "12" And List2.List(2) = "13" And List2.List(3) = "8" And List2.List(4) = "9" Then
     Earn = Ulog * 4
     Timer1.Enabled = True
   Else
   If List2.List(0) = "10" And List2.List(1) = "12" And List2.List(2) = "13" And List2.List(3) = "14" And List2.List(4) = "9" Then
     Earn = Ulog * 4
     Timer1.Enabled = True
   Else
   If List2.List(0) = "1" And List2.List(1) = "10" And List2.List(2) = "12" And List2.List(3) = "13" And List2.List(4) = "14" Then
     Earn = Ulog * 4
     Timer1.Enabled = True
   Else
   If (Aces = 2 And Dva = 2) Or (Aces = 2 And Tri = 2) Or (Aces = 2 And Cetiri = 2) Or (Aces = 2 And Pet = 2) Or (Aces = 2 And Sest = 2) Or (Aces = 2 And Sedam = 2) Or (Aces = 2 And Osam = 2) Or (Aces = 2 And Devet = 2) Or (Aces = 2 And Deset = 2) Or (Aces = 2 And Jacks = 2) Or (Aces = 2 And Quins = 2) Or (Aces = 2 And Kings = 2) Or _
     (Dva = 2 And Tri = 2) Or (Dva = 2 And Cetiri = 2) Or (Dva = 2 And Pet = 2) Or (Dva = 2 And Sest = 2) Or (Dva = 2 And Sedam = 2) Or (Dva = 2 And Osam = 2) Or (Dva = 2 And Devet = 2) Or (Dva = 2 And Deset = 2) Or (Dva = 2 And Jacks = 2) Or (Dva = 2 And Quins = 2) Or (Dva = 2 And Kings = 2) Or (Tri = 2 And Cetiri = 2) Or (Tri = 2 And Pet = 2) Or (Tri = 2 And Sest = 2) Or (Tri = 2 And Sedam = 2) Or (Tri = 2 And Osam = 2) Or (Tri = 2 And Devet = 2) Or (Tri = 2 And Deset = 2) Or (Tri = 2 And Jacks = 2) Or (Tri = 2 And Quins = 2) Or (Tri = 2 And Kings = 2) Or _
     (Cetiri = 2 And Pet = 2) Or (Cetiri = 2 And Sest = 2) Or (Cetiri = 2 And Sedam = 2) Or (Cetiri = 2 And Osam = 2) Or (Cetiri = 2 And Devet = 2) Or (Cetiri = 2 And Deset = 2) Or (Cetiri = 2 And Jacks = 2) Or (Cetiri = 2 And Quins = 2) Or (Cetiri = 2 And Kings = 2) Or (Pet = 2 And Sest = 2) Or (Pet = 2 And Sedam = 2) Or (Pet = 2 And Osam = 2) Or (Pet = 2 And Devet = 2) Or (Pet = 2 And Deset = 2) Or (Pet = 2 And Jacks = 2) Or (Pet = 2 And Quins = 2) Or (Pet = 2 And Kings = 2) Or (Sest = 2 And Sedam = 2) Or (Sest = 2 And Osam = 2) Or (Sest = 2 And Devet = 2) Or (Sest = 2 And Deset = 2) Or (Sest = 2 And Jacks = 2) Or (Sest = 2 And Quins = 2) Or (Sest = 2 And Kings = 2) Or _
     (Sedam = 2 And Osam = 2) Or (Sedam = 2 And Devet = 2) Or (Sedam = 2 And Deset = 2) Or (Sedam = 2 And Jacks = 2) Or (Sedam = 2 And Quins = 2) Or (Sedam = 2 And Kings = 2) Or (Osam = 2 And Devet = 2) Or (Osam = 2 And Deset = 2) Or (Osam = 2 And Jacks = 2) Or (Osam = 2 And Quins = 2) Or (Osam = 2 And Kings = 2) Or (Devet = 2 And Deset = 2) Or (Devet = 2 And Jacks = 2) Or (Devet = 2 And Quins = 2) Or (Devet = 2 And Kings = 2) Or (Deset = 2 And Jacks = 2) Or (Deset = 2 And Quins = 2) Or (Deset = 2 And Kings = 2) Or (Jacks = 2 And Quins = 2) Or (Jacks = 2 And Kings = 2) Or (Quins = 2 And Kings = 2) Then
     Earn = Ulog * 2
     Timer1.Enabled = True
   Else
  If (Aces = 3 And Dva = 2) Or (Aces = 3 And Tri = 2) Or (Aces = 3 And Cetiri = 2) Or (Aces = 3 And Pet = 2) Or (Aces = 3 And Sest = 2) Or (Aces = 3 And Sedam = 2) Or (Aces = 3 And Osam = 2) Or (Aces = 3 And Devet = 2) Or (Aces = 3 And Deset = 2) Or (Aces = 3 And Jacks = 2) Or (Aces = 3 And Quins = 2) Or (Aces = 3 And Kings = 2) Or (Dva = 3 And Aces = 2) Or (Dva = 3 And Tri = 2) Or (Dva = 3 And Cetiri = 2) Or (Dva = 3 And Pet = 2) Or (Dva = 3 And Sest = 2) Or (Dva = 3 And Sedam = 2) Or (Dva = 3 And Osam = 2) Or (Dva = 3 And Devet = 2) Or (Dva = 3 And Deset = 2) Or (Dva = 3 And Jacks = 2) Or (Dva = 3 And Quins = 2) Or (Dva = 3 And Kings = 2) Or _
  (Tri = 3 And Aces = 2) Or (Tri = 3 And Dva = 2) Or (Tri = 3 And Cetiri = 2) Or (Tri = 3 And Pet = 2) Or (Tri = 3 And Sest = 2) Or (Tri = 3 And Sedam = 2) Or (Tri = 3 And Osam = 2) Or (Tri = 3 And Devet = 2) Or (Tri = 3 And Deset = 2) Or (Tri = 3 And Jacks = 2) Or (Tri = 3 And Quins = 2) Or (Tri = 3 And Kings = 2) Or (Cetiri = 3 And Aces = 2) Or (Cetiri = 3 And Dva = 2) Or (Cetiri = 3 And Tri = 2) Or (Cetiri = 3 And Pet = 2) Or (Cetiri = 3 And Sest = 2) Or (Cetiri = 3 And Sedam = 2) Or (Cetiri = 3 And Osam = 2) Or (Cetiri = 3 And Devet = 2) Or (Cetiri = 3 And Deset = 2) Or (Cetiri = 3 And Jacks = 2) Or (Cetiri = 3 And Quins = 2) Or (Cetiri = 3 And Kings = 2) Or _
  (Pet = 3 And Aces = 2) Or (Pet = 3 And Dva = 2) Or (Pet = 3 And Tri = 2) Or (Pet = 3 And Cetiri = 2) Or (Pet = 3 And Sest = 2) Or (Pet = 3 And Sedam = 2) Or (Pet = 3 And Osam = 2) Or (Pet = 3 And Devet = 2) Or (Pet = 3 And Deset = 2) Or (Pet = 3 And Jacks = 2) Or (Pet = 3 And Quins = 2) Or (Pet = 3 And Kings = 2) Or (Sest = 3 And Aces = 2) Or (Sest = 3 And Dva = 2) Or (Sest = 3 And Tri = 2) Or (Sest = 3 And Cetiri = 2) Or (Sest = 3 And Pet = 2) Or (Sest = 3 And Sedam = 2) Or (Sest = 3 And Osam = 2) Or (Sest = 3 And Devet = 2) Or (Sest = 3 And Deset = 2) Or (Sest = 3 And Jacks = 2) Or (Sest = 3 And Quins = 2) Or (Sest = 3 And Kings = 2) Or _
  (Sedam = 3 And Aces = 2) Or (Sedam = 3 And Dva = 2) Or (Sedam = 3 And Tri = 2) Or (Sedam = 3 And Cetiri = 2) Or (Sedam = 3 And Pet = 2) Or (Sedam = 3 And Sest = 2) Or (Sedam = 3 And Osam = 2) Or (Sedam = 3 And Devet = 2) Or (Sedam = 3 And Deset = 2) Or (Sedam = 3 And Jacks = 2) Or (Sedam = 3 And Quins = 2) Or (Sedam = 3 And Kings = 2) Or (Osam = 3 And Aces = 2) Or (Osam = 3 And Dva = 2) Or (Osam = 3 And Tri = 2) Or (Osam = 3 And Cetiri = 2) Or (Osam = 3 And Pet = 2) Or (Osam = 3 And Sest = 2) Or (Osam = 3 And Sedam = 2) Or (Osam = 3 And Devet = 2) Or (Osam = 3 And Deset = 2) Or (Osam = 3 And Jacks = 2) Or (Osam = 3 And Quins = 2) Or (Osam = 3 And Kings = 2) Or _
  (Devet = 3 And Aces = 2) Or (Devet = 3 And Dva = 2) Or (Devet = 3 And Tri = 2) Or (Devet = 3 And Cetiri = 2) Or (Devet = 3 And Pet = 2) Or (Devet = 3 And Sest = 2) Or (Devet = 3 And Sedam = 2) Or (Devet = 3 And Osam = 2) Or (Devet = 3 And Deset = 2) Or (Devet = 3 And Jacks = 2) Or (Devet = 3 And Quins = 2) Or (Devet = 3 And Kings = 2) Or (Deset = 3 And Aces = 2) Or (Deset = 3 And Dva = 2) Or (Deset = 3 And Tri = 2) Or (Deset = 3 And Cetiri = 2) Or (Deset = 3 And Pet = 2) Or (Deset = 3 And Sest = 2) Or (Deset = 3 And Sedam = 2) Or (Deset = 3 And Osam = 2) Or (Deset = 3 And Devet = 2) Or (Deset = 3 And Jacks = 2) Or (Deset = 3 And Quins = 2) Or (Deset = 3 And Kings = 2) Or _
  (Jacks = 3 And Aces = 2) Or (Jacks = 3 And Dva = 2) Or (Jacks = 3 And Tri = 2) Or (Jacks = 3 And Cetiri = 2) Or (Jacks = 3 And Pet = 2) Or (Jacks = 3 And Sest = 2) Or (Jacks = 3 And Sedam = 2) Or (Jacks = 3 And Osam = 2) Or (Jacks = 3 And Devet = 2) Or (Jacks = 3 And Deset = 2) Or (Jacks = 3 And Quins = 2) Or (Jacks = 3 And Kings = 2) Or (Quins = 3 And Aces = 2) Or (Quins = 3 And Dva = 2) Or (Quins = 3 And Tri = 2) Or (Quins = 3 And Cetiri = 2) Or (Quins = 3 And Pet = 2) Or (Quins = 3 And Sest = 2) Or (Quins = 3 And Sedam = 2) Or (Quins = 3 And Osam = 2) Or (Quins = 3 And Devet = 2) Or (Quins = 3 And Deset = 2) Or (Quins = 3 And Jacks = 2) Or (Quins = 3 And Kings = 2) Or _
  (Kings = 3 And Aces = 2) Or (Kings = 3 And Dva = 2) Or (Kings = 3 And Tri = 2) Or (Kings = 3 And Cetiri = 2) Or (Kings = 3 And Pet = 2) Or (Kings = 3 And Sest = 2) Or (Kings = 3 And Sedam = 2) Or (Kings = 3 And Osam = 2) Or (Kings = 3 And Devet = 2) Or (Kings = 3 And Deset = 2) Or (Kings = 3 And Jacks = 2) Or (Kings = 3 And Quins = 2) Then
     Earn = Ulog * 8
     Timer1.Enabled = True
   Else
   If Aces = 3 Or Dva = 3 Or Tri = 3 Or Cetiri = 3 Or Pet = 3 Or Sest = 3 Or Sedam = 3 Or Osam = 3 Or Devet = 3 Or Deset = 3 Or Jacks = 3 Or Quins = 3 Or Kings = 3 Then
     Earn = Ulog * 3
     Timer1.Enabled = True
   Else
   If (13 >= (Int(List4.List(0))) And (Int(List4.List(0))) >= 1) And (13 >= (Int(List4.List(1))) And (Int(List4.List(1)) >= 1)) And (13 >= (Int(List4.List(2))) And (Int(List4.List(2)) >= 1)) And (13 >= (Int(List4.List(3)) And (Int(List4.List(3)) >= 1)) And (13 >= (Int(List4.List(4))) And (Int(List4.List(4)) >= 1))) Then
     Earn = Ulog * 5
     Timer1.Enabled = True
   Else
   If (26 >= (Int(List4.List(0))) And (Int(List4.List(0))) >= 14) And (26 >= (Int(List4.List(1))) And (Int(List4.List(1)) >= 14)) And (26 >= (Int(List4.List(2))) And (Int(List4.List(2)) >= 14)) And (26 >= (Int(List4.List(3)) And (Int(List4.List(3)) >= 14)) And (26 >= (Int(List4.List(4))) And (Int(List4.List(4)) >= 14))) Then
     Earn = Ulog * 5
     Timer1.Enabled = True
   Else
   If (39 >= (Int(List4.List(0))) And (Int(List4.List(0))) >= 27) And (39 >= (Int(List4.List(1))) And (Int(List4.List(1)) >= 27)) And (39 >= (Int(List4.List(2))) And (Int(List4.List(2)) >= 27)) And (39 >= (Int(List4.List(3)) And (Int(List4.List(3)) >= 27)) And (39 >= (Int(List4.List(4))) And (Int(List4.List(4)) >= 27))) Then
     Earn = Ulog * 5
     Timer1.Enabled = True
   Else
   If (52 >= (Int(List4.List(0))) And (Int(List4.List(0))) >= 40) And (52 >= (Int(List4.List(1))) And (Int(List4.List(1)) >= 40)) And (52 >= (Int(List4.List(2))) And (Int(List4.List(2)) >= 40)) And (52 >= (Int(List4.List(3)) And (Int(List4.List(3)) >= 40)) And (52 >= (Int(List4.List(4))) And (Int(List4.List(4)) >= 40))) Then
     Earn = Ulog * 5
     Timer1.Enabled = True
   Else
   If Aces = 2 Or Jacks = 2 Or Quins = 2 Or Kings = 2 Then
     Earn = Ulog
     Timer1.Enabled = True
   End If
   End If
   End If
   End If
   End If
   End If
   End If
   End If
   End If
   End If
   End If
   End If
   End If
   End If
   End If
   End If
   End If
   End If
   End If
   End If
   End If
   End If
   End If
   End If
   End If
   End If
   End If
   End If
   End If
   End If
End If
End Sub


Private Sub Form_Load()
List1.Clear
List3.Clear
Randomize
Ulog = 5
Label6.Caption = "Straith: " & Ulog * 4
Label7.Caption = "Three of a kind: " & Ulog * 3
Label8.Caption = "Two pairs: " & Ulog * 2
Label9.Caption = "Jacks or better: " & Ulog
Label10.Caption = "Flush royal: " & Ulog * 100
Label11.Caption = "Four of a kind: " & Ulog * 25
Label12.Caption = "Full house: " & Ulog * 8
Label13.Caption = "Flush: " & Ulog * 5
Label14.Caption = "Straight flush: " & Ulog * 50
Picture1.Picture = card(55).Picture
Picture2.Picture = card(55).Picture
Picture3.Picture = card(55).Picture
Picture4.Picture = card(55).Picture
Picture5.Picture = card(55).Picture
Picture1.Enabled = False
Picture2.Enabled = False
Picture3.Enabled = False
Picture4.Enabled = False
Picture5.Enabled = False
MediaPlayer1.FileName = App.Path & "\" & "hold.wav"
MediaPlayer2.FileName = App.Path & "\" & "unhold.wav"
MediaPlayer3.FileName = App.Path & "\" & "coin.wav"
Text1.Text = Ulog
RichTextBox1.LoadFile (App.Path & "\" & "Cash.txt")
Label16.Caption = RichTextBox1.Text
Cash = Int(Label16.Caption)
Jacks = 0
Quins = 0
Kings = 0
Aces = 0
Dva = 0
Tri = 0
Cetiri = 0
Pet = 0
Sest = 0
Sedam = 0
Osam = 0
Devet = 0
Deset = 0
Earn = 0
i = 0
End Sub

Private Sub Form_Terminate()
  RichTextBox1.Text = Cash
  RichTextBox1.SaveFile (App.Path & "\" & "Cash.txt")
End Sub

Private Sub mnuAbout_Click()
  frmAbout.Show vbModal
End Sub

Private Sub mnuExit_Click()
  RichTextBox1.Text = Cash
  RichTextBox1.SaveFile (App.Path & "\" & "Cash.txt")
  End
End Sub

Private Sub mnuNew_Click()
  Cash = 100
  Label16.Caption = 100
  RichTextBox1.Text = Label16.Caption
  RichTextBox1.SaveFile (App.Path & "\" & "Cash.txt")
  Picture1.Picture = card(55).Picture
  Picture2.Picture = card(55).Picture
  Picture3.Picture = card(55).Picture
  Picture4.Picture = card(55).Picture
  Picture5.Picture = card(55).Picture
  Command1.Caption = "Deal"
  Label1.Visible = False
  Label2.Visible = False
  Label3.Visible = False
  Label4.Visible = False
  Label5.Visible = False
  UpDown1.Enabled = True
  Jacks = 0
  Quins = 0
  Kings = 0
  Aces = 0
  Dva = 0
  Tri = 0
  Cetiri = 0
  Pet = 0
  Sest = 0
  Sedam = 0
  Osam = 0
  Devet = 0
  Deset = 0
  i = 0
End Sub

Private Sub Picture1_Click()
If Command1.Caption = "Deal" Then
  If Label1.Visible = False Then
    Label1.Visible = True
  Else
    Label1.Visible = False
  End If
Else
  If Label1.Visible = False Then
    Label1.Visible = True
    If MediaPlayer1.PlayState <> mpPlaying Then
      MediaPlayer1.Play
    End If
  Else
    Label1.Visible = False
    If MediaPlayer2.PlayState <> mpPlaying Then
      MediaPlayer2.Play
    End If
  End If
End If
End Sub

Private Sub Picture2_Click()
If Command1.Caption = "Deal" Then
  If Label2.Visible = False Then
    Label2.Visible = True
  Else
    Label2.Visible = False
  End If
Else
  If Label2.Visible = False Then
    Label2.Visible = True
    If MediaPlayer1.PlayState <> mpPlaying Then
      MediaPlayer1.Play
    End If
  Else
    Label2.Visible = False
    If MediaPlayer2.PlayState <> mpPlaying Then
      MediaPlayer2.Play
    End If
  End If
End If
End Sub

Private Sub Picture3_Click()
If Command1.Caption = "Deal" Then
  If Label3.Visible = False Then
    Label3.Visible = True
  Else
    Label3.Visible = False
  End If
Else
  If Label3.Visible = False Then
    Label3.Visible = True
    If MediaPlayer1.PlayState <> mpPlaying Then
      MediaPlayer1.Play
    End If
  Else
    Label3.Visible = False
    If MediaPlayer2.PlayState <> mpPlaying Then
      MediaPlayer2.Play
    End If
  End If
End If
End Sub

Private Sub Picture4_Click()
If Command1.Caption = "Deal" Then
  If Label4.Visible = False Then
  Else
    Label4.Visible = False
  End If
Else
  If Label4.Visible = False Then
    Label4.Visible = True
    If MediaPlayer1.PlayState <> mpPlaying Then
      MediaPlayer1.Play
    End If
    Label4.Visible = True
  Else
    Label4.Visible = False
    If MediaPlayer2.PlayState <> mpPlaying Then
      MediaPlayer2.Play
    End If
  End If
End If
End Sub

Private Sub Picture5_Click()
If Command1.Caption = "Deal" Then
  If Label5.Visible = False Then
    Label5.Visible = True
  Else
    Label5.Visible = False
  End If
Else
  If Label5.Visible = False Then
    Label5.Visible = True
    If MediaPlayer1.PlayState <> mpPlaying Then
      MediaPlayer1.Play
    End If
  Else
    Label5.Visible = False
    If MediaPlayer2.PlayState <> mpPlaying Then
      MediaPlayer2.Play
    End If
  End If
End If
End Sub

Private Sub Timer1_Timer()
  i = i + 1
  UpDown1.Enabled = False
  Cash = Cash + 1
  Label16.Caption = Cash
  If MediaPlayer3.PlayState <> mpPlaying Then
    MediaPlayer3.Play
  End If
 End Sub


Private Sub Timer2_Timer()
  If i = Earn And i <> 0 Then
    Timer1.Enabled = False
    Command1.Enabled = True
    UpDown1.Enabled = True
  End If
  If Timer1.Enabled = True Then
    Command1.Enabled = False
  Else
    Command1.Enabled = True
  End If
End Sub

Private Sub UpDown1_DownClick()
Label6.Caption = "Straith: " & Ulog * 4
Label7.Caption = "Three of a kind: " & Ulog * 3
Label8.Caption = "Two pairs: " & Ulog * 2
Label9.Caption = "Jacks or better: " & Ulog
Label10.Caption = "Flush royal: " & Ulog * 100
Label11.Caption = "Four of a kind: " & Ulog * 25
Label12.Caption = "Full house: " & Ulog * 8
Label13.Caption = "Flush: " & Ulog * 5
Label14.Caption = "Straight flush: " & Ulog * 50
Select Case Ulog
    Case 1
      Ulog = 1
      Text1.Text = Ulog
      Label6.Caption = "Straith: " & Ulog * 4
      Label7.Caption = "Three of a kind: " & Ulog * 3
      Label8.Caption = "Two pairs: " & Ulog * 2
      Label9.Caption = "Jacks or better: " & Ulog
      Label10.Caption = "Flush royal: " & Ulog * 100
      Label11.Caption = "Four of a kind: " & Ulog * 25
      Label12.Caption = "Full house: " & Ulog * 8
      Label13.Caption = "Flush: " & Ulog * 5
      Label14.Caption = "Straight flush: " & Ulog * 50
    Case 2
      Ulog = Ulog - 1
      Text1.Text = Ulog
      Label6.Caption = "Straith: " & Ulog * 4
      Label7.Caption = "Three of a kind: " & Ulog * 3
      Label8.Caption = "Two pairs: " & Ulog * 2
      Label9.Caption = "Jacks or better: " & Ulog
      Label10.Caption = "Flush royal: " & Ulog * 100
      Label11.Caption = "Four of a kind: " & Ulog * 25
      Label12.Caption = "Full house: " & Ulog * 8
      Label13.Caption = "Flush: " & Ulog * 5
      Label14.Caption = "Straight flush: " & Ulog * 50
    Case 3
      Ulog = Ulog - 1
      Text1.Text = Ulog
      Label6.Caption = "Straith: " & Ulog * 4
      Label7.Caption = "Three of a kind: " & Ulog * 3
      Label8.Caption = "Two pairs: " & Ulog * 2
      Label9.Caption = "Jacks or better: " & Ulog
      Label10.Caption = "Flush royal: " & Ulog * 100
      Label11.Caption = "Four of a kind: " & Ulog * 25
      Label12.Caption = "Full house: " & Ulog * 8
      Label13.Caption = "Flush: " & Ulog * 5
      Label14.Caption = "Straight flush: " & Ulog * 50
    Case 4
      Ulog = Ulog - 1
      Text1.Text = Ulog
      Label6.Caption = "Straith: " & Ulog * 4
      Label7.Caption = "Three of a kind: " & Ulog * 3
      Label8.Caption = "Two pairs: " & Ulog * 2
      Label9.Caption = "Jacks or better: " & Ulog
      Label10.Caption = "Flush royal: " & Ulog * 100
      Label11.Caption = "Four of a kind: " & Ulog * 25
      Label12.Caption = "Full house: " & Ulog * 8
      Label13.Caption = "Flush: " & Ulog * 5
      Label14.Caption = "Straight flush: " & Ulog * 50
    Case 5
      Ulog = Ulog - 1
      Text1.Text = Ulog
      Label6.Caption = "Straith: " & Ulog * 4
      Label7.Caption = "Three of a kind: " & Ulog * 3
      Label8.Caption = "Two pairs: " & Ulog * 2
      Label9.Caption = "Jacks or better: " & Ulog
      Label10.Caption = "Flush royal: " & Ulog * 100
      Label11.Caption = "Four of a kind: " & Ulog * 25
      Label12.Caption = "Full house: " & Ulog * 8
      Label13.Caption = "Flush: " & Ulog * 5
      Label14.Caption = "Straight flush: " & Ulog * 50
    Case Else
      Ulog = 5
      Text1.Text = Ulog
      Label6.Caption = "Straith: " & Ulog * 4
      Label7.Caption = "Three of a kind: " & Ulog * 3
      Label8.Caption = "Two pairs: " & Ulog * 2
      Label9.Caption = "Jacks or better: " & Ulog
      Label10.Caption = "Flush royal: " & Ulog * 100
      Label11.Caption = "Four of a kind: " & Ulog * 25
      Label12.Caption = "Full house: " & Ulog * 8
      Label13.Caption = "Flush: " & Ulog * 5
      Label14.Caption = "Straight flush: " & Ulog * 50
  End Select
End Sub

Private Sub UpDown1_UpClick()
  Select Case Ulog
    Case 1
      Ulog = Ulog + 1
      Text1.Text = Ulog
      Label6.Caption = "Straith: " & Ulog * 4
      Label7.Caption = "Three of a kind: " & Ulog * 3
      Label8.Caption = "Two pairs: " & Ulog * 2
      Label9.Caption = "Jacks or better: " & Ulog
      Label10.Caption = "Flush royal: " & Ulog * 100
      Label11.Caption = "Four of a kind: " & Ulog * 25
      Label12.Caption = "Full house: " & Ulog * 8
      Label13.Caption = "Flush: " & Ulog * 5
      Label14.Caption = "Straight flush: " & Ulog * 50
    Case 2
      Ulog = Ulog + 1
      Text1.Text = Ulog
      Label6.Caption = "Straith: " & Ulog * 4
      Label7.Caption = "Three of a kind: " & Ulog * 3
      Label8.Caption = "Two pairs: " & Ulog * 2
      Label9.Caption = "Jacks or better: " & Ulog
      Label10.Caption = "Flush royal: " & Ulog * 100
      Label11.Caption = "Four of a kind: " & Ulog * 25
      Label12.Caption = "Full house: " & Ulog * 8
      Label13.Caption = "Flush: " & Ulog * 5
      Label14.Caption = "Straight flush: " & Ulog * 50
    Case 3
      Ulog = Ulog + 1
      Text1.Text = Ulog
      Label6.Caption = "Straith: " & Ulog * 4
      Label7.Caption = "Three of a kind: " & Ulog * 3
      Label8.Caption = "Two pairs: " & Ulog * 2
      Label9.Caption = "Jacks or better: " & Ulog
      Label10.Caption = "Flush royal: " & Ulog * 100
      Label11.Caption = "Four of a kind: " & Ulog * 25
      Label12.Caption = "Full house: " & Ulog * 8
      Label13.Caption = "Flush: " & Ulog * 5
      Label14.Caption = "Straight flush: " & Ulog * 50
    Case 4
      Ulog = Ulog + 1
      Text1.Text = Ulog
      Label6.Caption = "Straith: " & Ulog * 4
      Label7.Caption = "Three of a kind: " & Ulog * 3
      Label8.Caption = "Two pairs: " & Ulog * 2
      Label9.Caption = "Jacks or better: " & Ulog
      Label10.Caption = "Flush royal: " & Ulog * 100
      Label11.Caption = "Four of a kind: " & Ulog * 25
      Label12.Caption = "Full house: " & Ulog * 8
      Label13.Caption = "Flush: " & Ulog * 5
      Label14.Caption = "Straight flush: " & Ulog * 50
    Case 5
      Ulog = 5
      Text1.Text = Ulog
      Label6.Caption = "Straith: " & Ulog * 4
      Label7.Caption = "Three of a kind: " & Ulog * 3
      Label8.Caption = "Two pairs: " & Ulog * 2
      Label9.Caption = "Jacks or better: " & Ulog
      Label10.Caption = "Flush royal: " & Ulog * 100
      Label11.Caption = "Four of a kind: " & Ulog * 25
      Label12.Caption = "Full house: " & Ulog * 8
      Label13.Caption = "Flush: " & Ulog * 5
      Label14.Caption = "Straight flush: " & Ulog * 50
    Case Else
      Ulog = 5
      Text1.Text = Ulog
      Label6.Caption = "Straith: " & Ulog * 4
      Label7.Caption = "Three of a kind: " & Ulog * 3
      Label8.Caption = "Two pairs: " & Ulog * 2
      Label9.Caption = "Jacks or better: " & Ulog
      Label10.Caption = "Flush royal: " & Ulog * 100
      Label11.Caption = "Four of a kind: " & Ulog * 25
      Label12.Caption = "Full house: " & Ulog * 8
      Label13.Caption = "Flush: " & Ulog * 5
      Label14.Caption = "Straight flush: " & Ulog * 50
  End Select
End Sub
