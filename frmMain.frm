VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Diablo II Trainer - Beta"
   ClientHeight    =   6540
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   8430
   MaxButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   8430
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRunDiablo 
      BackColor       =   &H8000000C&
      Height          =   885
      Left            =   60
      Picture         =   "frmMain.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   63
      ToolTipText     =   "Run Diablo II"
      Top             =   5580
      Width           =   615
   End
   Begin VB.Frame fraAttr 
      BackColor       =   &H80000008&
      Caption         =   "Edit Character Attributes"
      ForeColor       =   &H0000FFFF&
      Height          =   6405
      Left            =   4710
      TabIndex        =   48
      Top             =   60
      Width           =   3645
      Begin VB.ComboBox cboTitle 
         BackColor       =   &H80000007&
         ForeColor       =   &H80000005&
         Height          =   315
         ItemData        =   "frmMain.frx":030A
         Left            =   480
         List            =   "frmMain.frx":031A
         TabIndex        =   64
         Text            =   "Select A Character"
         Top             =   4050
         Width           =   3075
      End
      Begin VB.CommandButton cmdApply 
         Caption         =   "&Apply"
         Height          =   435
         Left            =   1170
         TabIndex        =   62
         Top             =   5310
         Width           =   1515
      End
      Begin VB.TextBox txtMaxL 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   255
         Left            =   2610
         TabIndex        =   61
         Top             =   3120
         Width           =   675
      End
      Begin VB.TextBox txtCurM 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000005&
         Height          =   255
         Left            =   1830
         TabIndex        =   60
         Top             =   3540
         Width           =   675
      End
      Begin VB.TextBox txtMaxM 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   255
         Left            =   2610
         TabIndex        =   59
         Top             =   3540
         Width           =   675
      End
      Begin VB.TextBox txtCurL 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000005&
         Height          =   255
         Left            =   1830
         TabIndex        =   58
         Top             =   3090
         Width           =   675
      End
      Begin VB.TextBox txtMaxSt 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   255
         Left            =   2610
         TabIndex        =   57
         Top             =   2640
         Width           =   675
      End
      Begin VB.TextBox txtCurSt 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000005&
         Height          =   255
         Left            =   1830
         TabIndex        =   56
         Top             =   2640
         Width           =   675
      End
      Begin VB.PictureBox pic 
         BackColor       =   &H80000008&
         BorderStyle     =   0  'None
         Height          =   1755
         Left            =   2910
         ScaleHeight     =   1755
         ScaleWidth      =   315
         TabIndex        =   54
         Top             =   390
         Width           =   315
      End
      Begin VB.TextBox txtEnergy 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000005&
         Height          =   255
         Left            =   2190
         TabIndex        =   53
         Top             =   1770
         Width           =   675
      End
      Begin VB.TextBox txtVit 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000005&
         Height          =   255
         Left            =   2190
         TabIndex        =   52
         Top             =   1380
         Width           =   675
      End
      Begin VB.TextBox txtDex 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000005&
         Height          =   255
         Left            =   2190
         TabIndex        =   51
         Top             =   900
         Width           =   675
      End
      Begin VB.TextBox txtStr 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000005&
         Height          =   255
         Left            =   2190
         TabIndex        =   50
         Top             =   480
         Width           =   675
      End
      Begin VB.CheckBox chkMakeBK3 
         BackColor       =   &H80000008&
         Caption         =   "Make A Backup"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1230
         TabIndex        =   49
         ToolTipText     =   "(Highly Recommend!)"
         Top             =   5910
         Value           =   1  'Checked
         Width           =   1485
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Title:"
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   7
         Left            =   90
         TabIndex        =   65
         Top             =   4110
         Width           =   345
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current:    Maximum:"
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   8
         Left            =   1860
         TabIndex        =   55
         Top             =   2310
         Width           =   1440
      End
      Begin VB.Image Image1 
         Height          =   1305
         Left            =   450
         Picture         =   "frmMain.frx":0388
         Top             =   2580
         Width           =   2880
      End
      Begin VB.Image imgEnergy 
         Height          =   405
         Left            =   870
         Picture         =   "frmMain.frx":224D
         Top             =   1740
         Width           =   2175
      End
      Begin VB.Image imgVit 
         Height          =   390
         Left            =   870
         Picture         =   "frmMain.frx":2CCE
         Top             =   1290
         Width           =   2190
      End
      Begin VB.Image imgDex 
         Height          =   405
         Left            =   870
         Picture         =   "frmMain.frx":3698
         Top             =   840
         Width           =   2205
      End
      Begin VB.Image imgStr 
         Height          =   405
         Left            =   870
         Picture         =   "frmMain.frx":41E1
         Top             =   390
         Width           =   2235
      End
   End
   Begin VB.Frame fraSocket 
      BackColor       =   &H80000008&
      Caption         =   "Socket Items"
      ForeColor       =   &H0000FFFF&
      Height          =   6405
      Left            =   4710
      TabIndex        =   28
      Top             =   60
      Width           =   3645
      Begin VB.CheckBox chkMakeBK2 
         BackColor       =   &H80000008&
         Caption         =   "Make A Backup"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1050
         TabIndex        =   45
         ToolTipText     =   "(Highly Recommend!)"
         Top             =   5910
         Value           =   1  'Checked
         Width           =   1485
      End
      Begin VB.CommandButton cmdSocketAll 
         Caption         =   "&Socket All"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   900
         TabIndex        =   41
         ToolTipText     =   "Will socket all of the items at once (Alt+S)"
         Top             =   5100
         Width           =   1845
      End
      Begin VB.CommandButton cmdInv 
         BackColor       =   &H80000008&
         Height          =   1335
         Index           =   0
         Left            =   1170
         Picture         =   "frmMain.frx":4D28
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   840
         Width           =   1275
      End
      Begin VB.CommandButton cmdInv 
         BackColor       =   &H80000008&
         Height          =   2265
         Index           =   1
         Left            =   120
         Picture         =   "frmMain.frx":5F08
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   2490
         Width           =   1215
      End
      Begin VB.CommandButton cmdInv 
         BackColor       =   &H80000008&
         Height          =   2265
         Index           =   2
         Left            =   2340
         Picture         =   "frmMain.frx":8203
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   2490
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Note: Items must be equipped to be socketed"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   1845
         Index           =   5
         Left            =   1440
         TabIndex        =   47
         Top             =   2640
         Width           =   825
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Right Hand Item:"
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   4
         Left            =   2340
         TabIndex        =   44
         Top             =   2250
         Width           =   1200
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Left Hand Item:"
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   3
         Left            =   150
         TabIndex        =   43
         Top             =   2250
         Width           =   1095
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Helm:"
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   2
         Left            =   1560
         TabIndex        =   42
         Top             =   570
         Width           =   405
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select items to be socketed:"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   6
         Left            =   825
         TabIndex        =   32
         Top             =   330
         Width           =   2025
      End
   End
   Begin VB.CommandButton cmdChangePath 
      BackColor       =   &H8000000C&
      Caption         =   "&Change Path"
      Height          =   345
      Left            =   60
      TabIndex        =   24
      ToolTipText     =   $"frmMain.frx":A4FE
      Top             =   90
      Width           =   1455
   End
   Begin VB.TextBox txtPath 
      BackColor       =   &H80000007&
      Enabled         =   0   'False
      ForeColor       =   &H80000005&
      Height          =   345
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   90
      Width           =   3105
   End
   Begin VB.Timer tmrUpdate 
      Interval        =   1
      Left            =   90
      Top             =   1680
   End
   Begin VB.ListBox lstChars 
      BackColor       =   &H80000007&
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   390
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.ListBox lstSaves 
      BackColor       =   &H80000007&
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   210
      TabIndex        =   21
      Top             =   360
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Frame fraCharList 
      BackColor       =   &H80000012&
      Caption         =   "Character List:"
      ForeColor       =   &H0000FFFF&
      Height          =   3135
      Left            =   180
      TabIndex        =   19
      Top             =   2250
      Width           =   4335
      Begin VB.ListBox lstChars2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
         ForeColor       =   &H80000005&
         Height          =   2760
         Left            =   30
         TabIndex        =   20
         Top             =   180
         Width           =   4275
      End
   End
   Begin VB.Frame fraGems 
      BackColor       =   &H80000015&
      Caption         =   "Convert Potions To Gems"
      ForeColor       =   &H0000FFFF&
      Height          =   6405
      Left            =   4710
      TabIndex        =   5
      Top             =   60
      Width           =   3645
      Begin VB.CheckBox chkMakeBK 
         BackColor       =   &H80000008&
         Caption         =   "Make A Backup"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   900
         TabIndex        =   46
         ToolTipText     =   "(Highly Recommend!)"
         Top             =   3810
         Value           =   1  'Checked
         Width           =   1485
      End
      Begin VB.CommandButton cmdUpgrade 
         BackColor       =   &H8000000C&
         Caption         =   "&Upgrade All To Perfect"
         Height          =   390
         Left            =   120
         TabIndex        =   26
         Top             =   4710
         Width           =   3465
      End
      Begin VB.CommandButton cmdConvert 
         BackColor       =   &H8000000C&
         Caption         =   "&Convert Potions to Gems"
         Height          =   390
         Left            =   120
         TabIndex        =   25
         Top             =   4260
         Width           =   3465
      End
      Begin VB.Frame fraPotions 
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Height          =   1875
         Left            =   90
         TabIndex        =   15
         Top             =   420
         Width           =   885
         Begin VB.OptionButton optP 
            BackColor       =   &H80000012&
            Caption         =   "Health"
            ForeColor       =   &H8000000E&
            Height          =   915
            Index           =   0
            Left            =   0
            Picture         =   "frmMain.frx":A590
            Style           =   1  'Graphical
            TabIndex        =   17
            Tag             =   "Health"
            Top             =   30
            Width           =   855
         End
         Begin VB.OptionButton optP 
            BackColor       =   &H80000012&
            Caption         =   "Mana"
            ForeColor       =   &H8000000E&
            Height          =   915
            Index           =   1
            Left            =   0
            Picture         =   "frmMain.frx":AB78
            Style           =   1  'Graphical
            TabIndex        =   16
            Tag             =   "Mana"
            Top             =   960
            Width           =   855
         End
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H80000012&
         Caption         =   "Topazes"
         ForeColor       =   &H8000000E&
         Height          =   915
         Index           =   6
         Left            =   1890
         Picture         =   "frmMain.frx":B1A2
         Style           =   1  'Graphical
         TabIndex        =   10
         Tag             =   "Topazes"
         Top             =   2310
         Width           =   855
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H80000012&
         Caption         =   "Skulls"
         ForeColor       =   &H8000000E&
         Height          =   915
         Index           =   5
         Left            =   1890
         Picture         =   "frmMain.frx":B774
         Style           =   1  'Graphical
         TabIndex        =   12
         Tag             =   "Skulls"
         Top             =   1380
         Width           =   855
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H80000012&
         Caption         =   "Sapphires"
         ForeColor       =   &H8000000E&
         Height          =   915
         Index           =   4
         Left            =   1890
         Picture         =   "frmMain.frx":BDC6
         Style           =   1  'Graphical
         TabIndex        =   11
         Tag             =   "Sapphires"
         Top             =   450
         Width           =   855
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H80000012&
         Caption         =   "Rubys"
         ForeColor       =   &H8000000E&
         Height          =   915
         Index           =   3
         Left            =   2730
         Picture         =   "frmMain.frx":C374
         Style           =   1  'Graphical
         TabIndex        =   9
         Tag             =   "Rubys"
         Top             =   3240
         Width           =   855
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H80000012&
         Caption         =   "Emeralds"
         ForeColor       =   &H8000000E&
         Height          =   915
         Index           =   2
         Left            =   2730
         Picture         =   "frmMain.frx":C922
         Style           =   1  'Graphical
         TabIndex        =   8
         Tag             =   "Emeralds"
         Top             =   2310
         Width           =   855
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H80000012&
         Caption         =   "Diamonds"
         ForeColor       =   &H8000000E&
         Height          =   915
         Index           =   1
         Left            =   2730
         Picture         =   "frmMain.frx":CECA
         Style           =   1  'Graphical
         TabIndex        =   7
         Tag             =   "Diamonds"
         Top             =   1380
         Width           =   855
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H80000012&
         Caption         =   "Amethysts"
         ForeColor       =   &H8000000E&
         Height          =   915
         Index           =   0
         Left            =   2730
         Picture         =   "frmMain.frx":D418
         Style           =   1  'Graphical
         TabIndex        =   6
         Tag             =   "Amethysts"
         Top             =   450
         Width           =   855
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Hover the mouse over a gem to get a description"
         ForeColor       =   &H8000000E&
         Height          =   855
         Left            =   60
         TabIndex        =   18
         Top             =   5490
         Width           =   3495
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To:"
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   1
         Left            =   2640
         TabIndex        =   14
         Top             =   210
         Width           =   240
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From:"
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   13
         Top             =   210
         Width           =   390
      End
   End
   Begin VB.CommandButton cmdGems 
      BackColor       =   &H80000008&
      Height          =   435
      Left            =   2850
      Picture         =   "frmMain.frx":D9C4
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Convert Potions To Gems"
      Top             =   5550
      Width           =   1305
   End
   Begin VB.CommandButton cmdSocket 
      BackColor       =   &H80000008&
      Height          =   435
      Left            =   2070
      Picture         =   "frmMain.frx":EC2A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Socket Equipped Items"
      Top             =   6030
      Width           =   2085
   End
   Begin VB.CommandButton cmdAttrib 
      BackColor       =   &H80000008&
      Height          =   435
      Left            =   720
      Picture         =   "frmMain.frx":1102C
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Edit Character Attributes"
      Top             =   5550
      Width           =   2085
   End
   Begin VB.CommandButton cmdMovies 
      BackColor       =   &H80000008&
      Height          =   435
      Left            =   720
      Picture         =   "frmMain.frx":133AE
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Enable Movies"
      Top             =   6030
      Width           =   1305
   End
   Begin VB.PictureBox picLogo 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1635
      Left            =   540
      Picture         =   "frmMain.frx":14CB8
      ScaleHeight     =   1635
      ScaleWidth      =   3675
      TabIndex        =   27
      Top             =   420
      Width           =   3675
   End
   Begin VB.Frame fraMovies 
      BackColor       =   &H80000008&
      Caption         =   "Movies"
      ForeColor       =   &H0000FFFF&
      Height          =   6405
      Left            =   4710
      TabIndex        =   33
      Top             =   30
      Width           =   3645
      Begin VB.OptionButton optMov 
         BackColor       =   &H80000012&
         Caption         =   "All Movies (5)"
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   4
         Left            =   120
         TabIndex        =   39
         Top             =   2100
         Width           =   1335
      End
      Begin VB.OptionButton optMov 
         BackColor       =   &H80000012&
         Caption         =   "4 Movies"
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   3
         Left            =   120
         TabIndex        =   38
         Top             =   1800
         Width           =   1335
      End
      Begin VB.OptionButton optMov 
         BackColor       =   &H80000012&
         Caption         =   "3 Movies"
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   37
         Top             =   1500
         Width           =   1335
      End
      Begin VB.OptionButton optMov 
         BackColor       =   &H80000012&
         Caption         =   "2 Movies"
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   36
         Top             =   1170
         Width           =   1335
      End
      Begin VB.OptionButton optMov 
         BackColor       =   &H80000012&
         Caption         =   "1 Movie"
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   35
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton cmdSetMovies 
         Caption         =   "&Set Movies"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   660
         TabIndex        =   34
         Top             =   2700
         Width           =   2295
      End
      Begin VB.Label lblSelect 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Select the amount of movies you want to be enabled:"
         ForeColor       =   &H0000FFFF&
         Height          =   465
         Left            =   690
         TabIndex        =   40
         Top             =   330
         Width           =   2235
      End
   End
   Begin VB.Label lblTrn 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TRAINER"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   1770
      TabIndex        =   2
      Top             =   2010
      Width           =   1215
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sINIFile As String, sMSetting(1 To 5) As String, sBKFile As String, sFile As String, sActiveFrame As String, nPrevIdx As Integer
'Diablo II Trainer, by Max Raskin
'The Socket and the Gems source code were made by Disk2 (Will) , and can be found on PSC:
'http://www.planet-source-code.com/vb/

Private Const Ver = "1.0b"
' Used to find the beginning and end of the inventory section...
Private Type ItemHeader
    szFirstJM As String * 2
    iItemCount As Byte
    iEmpty As Byte
    szLastJM As String * 2
End Type

' This isn't complete :) It does what I need it to though.
Private Type Item
    iSubType As Byte
    iType As Byte
End Type

' This is obviously not a complete item type :)
' This only holds the inventory position and the Equipped position.
' It's all I need for this purpose...
Private Type ItemSocket
    iSocketed As Byte
    iInvPos As Byte
End Type

' Holds the filename and declares the item header...
Dim sGemDest As String, sGemSrc As String
Dim ItemHead As ItemHeader

Private Sub SetMovieVars()
'Set movie settings:
    sMSetting(1) = "216.148.246.34" '1 movie
    sMSetting(2) = "216.148.246.38" '2 movies etc..
    sMSetting(3) = "216.148.246.98"
    sMSetting(4) = "216.148.246.40"
    sMSetting(5) = "216.148.246.50"
End Sub

Private Sub cboTitle_GotFocus()
    cboTitle.Refresh
End Sub

Private Sub cmdApply_Click()
    If lstChars2.ListIndex = -1 Then
        MsgNoChar
        Exit Sub
    End If
    If chkMakeBK3.Value = 1 Then BackUp
    SetAttrib sFile, CurLife, txtCurL
    SetAttrib sFile, CurMana, txtCurM
    SetAttrib sFile, CurStamina, txtCurSt
    SetAttrib sFile, Dexterity, txtDex
    SetAttrib sFile, Energy, txtEnergy
    SetAttrib sFile, Life, txtMaxL
    SetAttrib sFile, Mana, txtMaxM
    SetAttrib sFile, Stamina, txtMaxSt
    SetAttrib sFile, Strength, txtStr
    SetAttrib sFile, Vitality, txtVit
    ApplyTitle
    GetAttributes
    MsgBox "Changes Applied!", vbInformation, "Applied"
End Sub

Private Sub cmdAttrib_Click()
    SetFrames fraAttr
End Sub

Private Sub cmdConvert_Click()
    On Error Resume Next
    If chkMakeBK2.Value = 1 Then BackUp
    ' Check to see if the user select a source and destination type...
    If lstChars2.ListIndex <> -1 Then
        If sGemSrc = "" Or sGemDest = "" Then
            MsgBox "You must select a source and destination type before converting!", vbCritical + vbOKOnly, "Error"
        Else
            ' If the user DID, then convert the items.
            Convert txtPath.Text & "Save\" & lstChars.List(lstChars.ListIndex) & ".d2s"
        End If
    Else
        MsgNoChar
    End If
End Sub

Private Sub cmdGems_Click()
    SetFrames fraGems
End Sub

Private Sub cmdInv_Click(Index As Integer)
    On Error Resume Next
    If lstChars2.ListIndex = -1 Then
        MsgNoChar
        Exit Sub
    End If
    SetFiles
    'Copy backup file
    If chkMakeBK.Value = 1 Then
        
    End If
    Select Case Index
    Case 0 'Helm
        Socket &H1, "Your helm is now socketed.", sFile
    Case 1 'Left Hand
        Socket &H4, "Your left-hand item is now socketed.", sFile
    Case 2 'Right Hand
        Socket &H5, "Your right-hand item is now socketed.", sFile
    End Select
End Sub

Private Sub cmdMovies_Click()
    SetFrames fraMovies
End Sub

Private Sub cmdRunDiablo_Click()
    ChDir txtPath.Text
    If Dir(txtPath.Text & "DLoad.exe") <> "" Then
        Shell txtPath.Text & "DLoad.exe", vbNormalFocus 'Incase you are to lazy to insert the cd everytime, and you have the crack (like i do =) this is a usefull line
    Else
        Shell txtPath.Text & "Diablo II.exe", vbNormalFocus 'Anyway, if not ...
    End If
End Sub

Private Sub cmdSetMovies_Click()
    'Check selection and done! :)
    If optMov(0).Value = False Then If optMov(1).Value = False Then If optMov(2).Value = False Then If optMov(3).Value = False Then If optMov(4).Value = False Then MsgBox "Select Number Of Movies First!", vbInformation, ""
    For i = 0 To 4
        If optMov(i).Value = True Then
            UpdateKey HKEY_CURRENT_USER, "Software\Blizzard Entertainment\Diablo II", "Aux Battle.Net", sMSetting(i + 1)
            MsgBox "You have now enabled " & optMov(i).Caption, vbInformation, "Movies Enabled!"
        End If
    Next
End Sub

Private Sub cmdSocket_Click()
    SetFrames fraSocket
End Sub

Private Sub cmdSocketAll_Click()
    On Error Resume Next
    SetFiles
    If lstChars2.ListIndex = -1 Then
        MsgNoChar
        Exit Sub
    End If
    If chkMakeBK.Value = 1 Then BackUp
    ' Socket all the items and show the message
    Socket &H0, "All of your equipped socketable items are now socketed!", sFile
End Sub

Private Sub cmdUpgrade_Click()
    On Error Resume Next
    SetFiles
    ' Fix the user's gems... This is done in the FixGems function
    If lstChars2.ListIndex <> -1 Then
        FixGems sFile
    Else
        MsgNoChar
    End If
    'Copy backup file
    If chkMakeBK.Value = 1 Then BackUp
End Sub

Private Sub mnuAbout_Click()
    MsgBox "Diablo II Trainer " & Ver & vbCrLf & "By Max Raskin" & vbCrLf & vbCrLf & "Credits:" & vbCrLf & "~~~~~" & vbCrLf & "Socket and Gems Conversion source codes by Disk2" & vbCrLf & vbCrLf & "EnJoY! =)", vbInformation, "About Diablo II Trainer"
End Sub

Private Sub opt_Click(Index As Integer)
    sGemDest = opt(Index).Tag
End Sub

Private Sub opt_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case opt(Index).Tag
    Case "Amethysts"
        lblInfo.Caption = "Weapon: Adds To Attack Rating" & vbCrLf & "Shield: Adds To Shield Defense Rating" & vbCrLf & "Helm: Adds To Strength"
    Case "Diamonds"
        lblInfo.Caption = "Weapon: Adds To Damage VS. Undead" & vbCrLf & "Shield: Adds To All Resistences" & vbCrLf & "Helm: Addss To Attack Rating"
    Case "Emeralds"
        lblInfo.Caption = "Weapon: Adds Poison Damage To Attack" & vbCrLf & "Shield: Adds To All Resistences" & vbCrLf & "Helm: Addss To Attack Rating"
    Case "Rubys"
        lblInfo.Caption = "Weapon: Adds Fire Damage To Attack" & vbCrLf & "Shield: Adds Resistence To Fire" & vbCrLf & "Helm: Addss To Maximum Life"
    Case "Sapphires"
        lblInfo.Caption = "Weapon: Adds Cold Damage To Attack" & vbCrLf & "Shield: Adds Resistence To Cold" & vbCrLf & "Helm: Addss To Maximum Mana"
    Case "Skulls"
        lblInfo.Caption = "Weapon: Adds Mana And Life Steal To Attack" & vbCrLf & "Shield: Adds Attacker Takes Damage" & vbCrLf & "Helm: Addss Mana And Life Regeneration"
    Case "Topazes"
        lblInfo.Caption = "Weapon: Adds Lightning Damage To Attack" & vbCrLf & "Shield: Adds Resistence To Lightning" & vbCrLf & "Helm: Addss Chance To Find Magic Items"
    End Select
End Sub

Private Sub EnumChars()
    On Error Resume Next
    Dim i As Integer, l As Integer, Stats As String, CharClass As String
    tmrUpdate.Enabled = False 'to make sure, the timer wont work when its not suppost to
    nPrevIdx = lstChars2.ListIndex
    lstChars2.Clear
    lstChars.Clear
    lstSaves.Clear 'Clear up the invisible text box
    EnumFilesByExt txtPath.Text & "Save", lstSaves, "d2s" 'Enum files by extension, diablo2's saves extension is "d2s"
    For i = 0 To lstSaves.ListCount - 1 'Remove the '.d2s' strings from the list items
        l = Len(lstSaves.List(i))
        lstChars.AddItem Left(lstSaves.List(i), l - 4)
        Stats = GetStatus(txtPath.Text & "Save\" & lstSaves.List(i))
        CharClass = GetClass(txtPath.Text & "Save\" & lstSaves.List(i))
        If Stats = "" Then
            lstChars2.AddItem Left(lstSaves.List(i), l - 4) & " (" & CharClass & " Level " & GetLevel(txtPath.Text & "Save\" & lstSaves.List(i)) & ")"
        Else
            lstChars2.AddItem Stats & " " & Left(lstSaves.List(i), l - 4) & " (" & CharClass & " Level " & GetLevel(txtPath.Text & "Save\" & lstSaves.List(i)) & ")"
        End If
    Next
    tmrUpdate.Enabled = True
    lstChars2.ListIndex = nPrevIdx
End Sub

Private Sub lstChars2_Click()
    If lstChars2.ListIndex <> -1 Then
        lstChars.ListIndex = lstChars2.ListIndex
        If fraAttr.Visible = True Then GetAttributes
        If fraAttr.Visible = True Then cboTitle.Text = GetStatus(sFile)
        SetTitle
    End If
End Sub

Private Sub optP_Click(Index As Integer)
    sGemSrc = optP(Index).Tag
End Sub

'This timer will instantly update the characters in the list if the number of characters is changed
Private Sub tmrUpdate_Timer()
    lstSaves.Clear
    EnumFilesByExt txtPath.Text & "Save", lstSaves, "d2s"  'Enum files by extension, diablo2's saves extension is "d2s"
    If lstSaves.ListCount <> lstChars.ListCount Then
        EnumChars
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    sINIFile = App.Path & "Settings.ini" 'Set default INI file name
    If Right(App.Path, 1) <> "\" Then sINIFile = App.Path & "\" & "Settings.ini"
    LoadSettings
    sRetVal = GetINI(sINIFile, "Settings", "DiabloPath", "") 'Attempt to get path from INI
    'BrowseForFolder if no path saved in the INI
Browse:     If sRetVal = "" Then sRetVal = BrowseForFolder(Me.hwnd, "Browse For Diablo II's Folder:", ReturnFileSystemFoldersOnly)
    If sRetVal = "" Then
        'Make sure user selects the path
        msgResult = MsgBox("You Must Select A Path, Browse Again?", vbYesNo Or vbQuestion, "No Path Selected")
        If msgResult = vbYes Then
            GoTo Browse
        Else
            End
        End If
    Else
        txtPath.Text = sRetVal
        If Right(txtPath.Text, 1) <> "\" Then txtPath.Text = txtPath.Text & "\"
    End If
    EnumChars 'Get all characters from Diablo2Dir\Save directory
    SetMovieVars 'Set movies values settings
    Text2Numeric txtStr
    Text2Numeric txtDex
    Text2Numeric txtVit
    Text2Numeric txtEnergy
End Sub

Private Sub SaveSettings()
    writeINI sINIFile, "Settings", "Top", Me.Top 'Remember form's Y
    writeINI sINIFile, "Settings", "Left", Me.Left 'Remember form's X
    writeINI sINIFile, "Settings", "DiabloPath", txtPath.Text 'Remember Diablo II's Path dir
    writeINI sINIFile, "Settings", "ActiveFrame", sActiveFrame 'Remember current active frame
End Sub

Private Sub LoadSettings()
    Me.Top = GetINI(sINIFile, "Settings", "Top", Me.Top) 'Get Form's Y
    Me.Left = GetINI(sINIFile, "Settings", "Left", Me.Left) 'Get Form's X
    sActiveFrame = GetINI(sINIFile, "Settings", "ActiveFrame", "") 'Get previously activated frame
    If sActiveFrame <> "" Then
        Select Case sActiveFrame
        Case "fraGems"
            SetFrames fraGems
        Case "fraMovies"
            SetFrames fraMovies
        Case "fraSocket"
            SetFrames fraSocket
        Case "fraAttr"
            SetFrames fraAttr
        End Select
    End If
End Sub

Private Sub cmdChangePath_Click()
    Dim sPath As String
    sPath = BrowseForFolder(Me.hwnd, "Browse For Diablo II's Folder:", ReturnFileSystemFoldersOnly)
    If sPath <> "" Then txtPath = sPath
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSettings 'Remember settings on exit
    End
End Sub


'Disk2's SRC: (FixGems and Convert)
Private Function FixGems(strFileName As String)
    On Error Resume Next
    
    Dim iPos As Integer ' Holds the position in the file
    Dim xItem As Item ' Temp item
    Dim TheString As String * 4 ' will be JMJM if we're at the end of the items...
    Dim TheEnd As ItemHeader ' Used to find the end of the items

    ' Reset ItemHead
    ItemHead.szFirstJM = ""
    ItemHead.szLastJM = ""
    ItemHead.iItemCount = 0
    ItemHead.iEmpty = 0

    iPos = &H1 ' Start at the beginning of the file

    ' Open the file
    Open strFileName For Binary As #1
        ' Read from the file until we find the "JM  JM". This means we've found the beginning of the item data...
        Do Until ItemHead.szFirstJM = "JM" And ItemHead.szLastJM = "JM"
            Get #1, iPos, ItemHead
            
            iPos = iPos + 1 ' Increase the position
        Loop
    
        iPos = iPos + 3 ' Go to the REAL start of the item information.

        ' If the user has no items, don't continue...
        If ItemHead.iItemCount = 0 Then
            MsgBox "This character doesn't appear to have any items! If this is an error please email me at cregistry@yahoo.com and attach the saved game file. Thanks!", vbOKOnly + vbInformation, "Notice"
            
            Close #1 ' Close the file
            Exit Function
        End If

        ' Read items. Compare them with known gem codes. Convert them if they aren't perfect.
        Do Until TheString = "JMJM"
            Get #1, iPos, TheEnd
            
            TheString = TheEnd.szFirstJM & TheEnd.szLastJM
            
            iPos = iPos + 2
        
            Get #1, iPos + 6, xItem.iSubType
            Get #1, iPos + 7, xItem.iType
            
            ' Hehe. This is confusing.
            ' The actual concept isn't. It's just the way I coded it :)
            ' Here's a table of all the gem codes...
            '
            '          | Chipped | Flawed | Regular | Flawless | Perfect |
            ' ---------|--------------------------------------------------
            ' Diamond  | 5015    | 6015   | 7015    | 8015     | 9015    |
            ' ---------|--------------------------------------------------
            ' Ruby     | 1015    | 0015   | 2015    | 3015     | 4015    |
            ' ---------|--------------------------------------------------
            ' Topaz    | 1014    | 2014   | 3014    | 4014     | 5014    |
            ' ---------|--------------------------------------------------
            ' Sapphire | 6014    | 7014   | 8014    | 9014     | A014    |
            ' ---------|--------------------------------------------------
            ' Amethyst | C013    | D013   | E013    | F013     | 0014    |
            ' ---------|--------------------------------------------------
            ' Emerald  | B014    | C014   | D014    | E014     | F014    |
            ' ---------|--------------------------------------------------
            ' Skull    | 4016    | 5016   | 6016    | 7016     | 8016    |
            ' ---------|--------------------------------------------------
            
            ' With that in mind, you can figure out this code.
            Select Case xItem.iType
                Case &H13
                    Select Case xItem.iSubType
                        Case &HD0
                            Put #1, iPos + 6, &H0
                            Put #1, iPos + 7, &H14
                        Case &HC0
                            Put #1, iPos + 6, &H0
                            Put #1, iPos + 7, &H14
                        Case &HF0
                            Put #1, iPos + 6, &H0
                            Put #1, iPos + 7, &H14
                        Case &HE0
                            Put #1, iPos + 6, &H0
                            Put #1, iPos + 7, &H14
                    End Select
                Case &H14
                    Select Case xItem.iSubType
                        Case &H20
                            Put #1, iPos + 6, &H50
                            Put #1, iPos + 7, &H14
                        Case &H10
                            Put #1, iPos + 6, &H50
                            Put #1, iPos + 7, &H14
                        Case &H40
                            Put #1, iPos + 6, &H50
                            Put #1, iPos + 7, &H14
                        Case &H30
                            Put #1, iPos + 6, &H50
                            Put #1, iPos + 7, &H14
                        Case &H60
                            Put #1, iPos + 6, &HA0
                            Put #1, iPos + 7, &H14
                        Case &H70
                            Put #1, iPos + 6, &HA0
                            Put #1, iPos + 7, &H14
                        Case &H80
                            Put #1, iPos + 6, &HA0
                            Put #1, iPos + 7, &H14
                        Case &H90
                            Put #1, iPos + 6, &HA0
                            Put #1, iPos + 7, &H14
                        Case &HC0
                            Put #1, iPos + 6, &HF0
                            Put #1, iPos + 7, &H14
                        Case &HB0
                            Put #1, iPos + 6, &HF0
                            Put #1, iPos + 7, &H14
                        Case &HE0
                            Put #1, iPos + 6, &HF0
                            Put #1, iPos + 7, &H14
                        Case &HD0
                            Put #1, iPos + 6, &HF0
                            Put #1, iPos + 7, &H14
                    End Select
                    Case &H15
                    Select Case xItem.iSubType
                        Case &H60
                            Put #1, iPos + 6, &H90
                            Put #1, iPos + 7, &H15
                        Case &H50
                            Put #1, iPos + 6, &H90
                            Put #1, iPos + 7, &H15
                        Case &H80
                            Put #1, iPos + 6, &H90
                            Put #1, iPos + 7, &H15
                        Case &H70
                            Put #1, iPos + 6, &H90
                            Put #1, iPos + 7, &H15
                        Case &H0
                            Put #1, iPos + 6, &H40
                            Put #1, iPos + 7, &H15
                        Case &H10
                            Put #1, iPos + 6, &H40
                            Put #1, iPos + 7, &H15
                        Case &H20
                            Put #1, iPos + 6, &H40
                            Put #1, iPos + 7, &H15
                        Case &H30
                            Put #1, iPos + 6, &H40
                            Put #1, iPos + 7, &H15
                    End Select
                    Case &H16
                    Select Case xItem.iSubType
                        Case &H40
                            Put #1, iPos + 6, &H80
                            Put #1, iPos + 7, &H16
                        Case &H50
                            Put #1, iPos + 6, &H80
                            Put #1, iPos + 7, &H16
                        Case &H60
                            Put #1, iPos + 6, &H80
                            Put #1, iPos + 7, &H16
                        Case &H70
                            Put #1, iPos + 6, &H80
                            Put #1, iPos + 7, &H16
                    End Select
            End Select
            
            ' Increase the position so we can read the next item.
            iPos = iPos + 25
        Loop
    Close #1
    
    ' Tell the user the gems were perfected
    MsgBox "All your gems are now perfect.", vbOKOnly + vbInformation, "Done"
End Function

Private Sub Convert(strFileName As String)
    On Error Resume Next

    Dim iPos As Integer
    Dim xItem As Item
    Dim dItem As Item
    Dim TheEnd As ItemHeader
    Dim TheString As String * 4
    ' Depending of the destination type, set the temp item type to a perfect gem.
    ' Refer to the table in FixGems for the gem codes...
    Select Case sGemDest
        Case "Diamonds"
            dItem.iType = &H15
            dItem.iSubType = &H90
        Case "Rubys"
            dItem.iType = &H15
            dItem.iSubType = &H40
        Case "Topazes"
            dItem.iType = &H14
            dItem.iSubType = &H50
        Case "Sapphires"
            dItem.iType = &H14
            dItem.iSubType = &HA0
        Case "Amethysts"
            dItem.iType = &H14
            dItem.iSubType = &H0
        Case "Emeralds"
            dItem.iType = &H14
            dItem.iSubType = &HF0
        Case "Skulls"
            dItem.iType = &H16
            dItem.iSubType = &H80
    End Select

    ' Reset ItemHead
    ItemHead.szFirstJM = ""
    ItemHead.szLastJM = ""
    ItemHead.iItemCount = 0
    ItemHead.iEmpty = 0

    ' Start at the beginning of the file
    iPos = &H1

    ' Open the file
    Open strFileName For Binary As #1
        Do Until ItemHead.szFirstJM = "JM" And ItemHead.szLastJM = "JM"
            Get #1, iPos, ItemHead
            
            iPos = iPos + 1
        Loop
    
         ' Go to the REAL item data start
        iPos = iPos + 3

        ' If the user has no items, there's no point in continuing.
        If ItemHead.iItemCount = 0 Then
            MsgBox "This character doesn't appear to have any items! If this is an error please email me at cregistry@yahoo.com and attach the saved game file. Thanks!", vbOKOnly + vbInformation, "Notice"
            
            Close #1
            Exit Sub
        End If

        ' Read items until we reach the end of the file.
        Do Until TheString = "JMJM"
            Get #1, iPos, TheEnd
            
            TheString = TheEnd.szFirstJM & TheEnd.szLastJM
            
            iPos = iPos + 2
            
            ' Get the item type
            Get #1, iPos + 6, xItem.iSubType
            Get #1, iPos + 7, xItem.iType
            
            ' Depending on the source type, look for health or mana potions.
            ' You can figure out the potion codes by looking at the code below.
            If optP(0).Value = True Then
                If xItem.iType = &H15 Then
                    If xItem.iSubType = &HA0 Or xItem.iSubType = &HB0 Or xItem.iSubType = &HC0 Or xItem.iSubType = &HD0 Or xItem.iSubType = &HE0 Then
                        Put #1, iPos + 6, dItem
                    End If
                End If
            End If
            If optP(1).Value = True Then
                If xItem.iType = &H16 Then
                    If xItem.iSubType = &H0 Or xItem.iSubType = &H10 Or xItem.iSubType = &H20 Or xItem.iSubType = &H30 Then
                        Put #1, iPos + 6, dItem
                    End If
                ElseIf xItem.iType = &H15 Then
                    If xItem.iSubType = &HF0 Then
                        Put #1, iPos + 6, dItem
                    End If
                End If
            End If
            
            ' Increase the position so we can read the next file
            iPos = iPos + 25
        Loop
    Close #1

    ' Show the message...
    If optP(0).Value = True Then MsgBox "All your " & optP(0).Tag & " are now " & sGemDest & ".", vbOKOnly + vbInformation, "Success"
    If optP(1).Value = True Then MsgBox "All your " & optP(1).Tag & " are now " & sGemDest & ".", vbOKOnly + vbInformation, "Success"
End Sub


Private Sub Socket(Position As Integer, Message As String, strFileName As String)
    On Error Resume Next ' If we encounter an error, resume next :)
    
    Dim iPos As Integer ' IMPORTANT: This holds our position in the file...
    Dim xItem As ItemSocket ' The temp item. Used to compare item positions, etc...
    Dim TheEnd As ItemHeader ' Used to check if we're at the end of the inventory
    Dim TheString As String * 4 ' Should be "JMJM" if we're at the end of the inventory

    ' See declaration of ItemHeader type (at top of code)
    TheEnd.iEmpty = &H0
    TheEnd.iItemCount = &H0
    TheEnd.szFirstJM = ""
    TheEnd.szLastJM = ""

    ' Clear ItemHead (it's a global variable so we need to clear it each time...)
    ItemHead.szFirstJM = ""
    ItemHead.szLastJM = ""
    ItemHead.iItemCount = 0
    ItemHead.iEmpty = 0

    ' Start at the beginning of the file
    iPos = &H1

    ' Open the filename (strFileName)
    Open strFileName For Binary As #1
        ' Get the position of the start of the inventory data
        Do Until ItemHead.szFirstJM = "JM" And ItemHead.szLastJM = "JM"
            Get #1, iPos, ItemHead
            
            iPos = iPos + 1
        Loop
    
        ' OK. We found it. Now we have to increase our position by 3 to get to the first item
        iPos = iPos + 3

        ' If the item count is zero then there's no point in continuing :)
        If ItemHead.iItemCount = 0 Then
            MsgBox "This character doesn't appear to have any items! If this is an error please email me at cregistry@yahoo.com and attach the saved game file. Thanks!", vbOKOnly + vbInformation, "Notice"
            
            ' Close the file and exit the sub...
            Close #1
            Exit Sub
        End If

        ' The ItemHead.iItemCount is a fake value for the number of items (I guess).
        ' The number doesn't account for gems that are in socketed items.
        ' So now we have to read items until we find the closing "JM  JM" in the file...
        Do Until TheString = "JMJM"
            ' First of all, make sure we aren't at the end of the inventory
            Get #1, iPos, TheEnd
            ' If TheString equals "JMJM" then we are at the end
            TheString = TheEnd.szFirstJM & TheEnd.szLastJM
            
            ' Increase our position by 2 to get to the item data
            ' BTW: Each item is 25 bytes long...
            iPos = iPos + 2
            
            ' Read the position of the item.
            Get #1, iPos + 4, xItem.iInvPos
            
            ' Depending on the value of Position when the function was called,
            ' socket the appropriate item(s).
            ' BTW: &H0 means socket all the items that are equipped...
            Select Case Position
                Case &H0
                    If xItem.iInvPos = &H1 Or xItem.iInvPos = &H4 Or xItem.iInvPos = &H5 Then
                        Put #1, iPos + 1, &H8
                    End If
                Case &H1
                    If xItem.iInvPos = &H1 Then
                        Put #1, iPos + 1, &H8
                    End If
                Case &H4
                    If xItem.iInvPos = &H4 Then
                        Put #1, iPos + 1, &H8
                    End If
                Case &H5
                    If xItem.iInvPos = &H5 Then
                        Put #1, iPos + 1, &H8
                    End If
            End Select
            
            ' Increase the position by 25 so that we can read the next item...
            iPos = iPos + 25
            
            ' Then loop back to the beginning :)
        Loop
        
    ' Close the file
    Close #1
    
    ' Show the message (this is shown regardless of wether the item was socketed
    ' successfully :)
    MsgBox "Done!" & Message, vbOKOnly + vbInformation, "Done"
End Sub

Private Sub SetFiles()
    sFile = txtPath.Text & "Save\" & lstChars.List(lstChars.ListIndex) & ".d2s"
    sBKFile = txtPath.Text & "Save\" & lstChars.List(lstChars.ListIndex) & ".d2s.bak"
End Sub

Private Sub SetFrames(Optional CallingFrame As Frame)
    If Me.Width <> 8520 Then Me.Width = 8520
    sActiveFrame = CallingFrame.Name
    Select Case CallingFrame.Name
    Case "fraSocket"
        fraGems.Visible = False
        fraMovies.Visible = False
        fraSocket.Visible = True
        fraAttr.Visible = False
    Case "fraMovies"
        fraGems.Visible = False
        fraMovies.Visible = True
        fraSocket.Visible = False
        fraAttr.Visible = False
    Case "fraGems"
        fraGems.Visible = True
        fraMovies.Visible = False
        fraSocket.Visible = False
        fraAttr.Visible = False
    Case "fraAttr"
       fraGems.Visible = False
       fraMovies.Visible = False
       fraSocket.Visible = False
       fraAttr.Visible = True
       If lstChars2.ListIndex <> -1 Then
           GetAttributes
           cboTitle.Text = GetStatus(sFile)
           SetTitle
       End If
    End Select
End Sub

Private Sub txtStr_GotFocus()
    txtStr.Refresh
End Sub

Private Sub txtDex_GotFocus()
    txtDex.Refresh
End Sub

Private Sub txtCurSt_GotFocus()
    txtCurSt.Refresh
End Sub

Private Sub txtMaxSt_GotFocus()
    txtMaxSt.Refresh
End Sub

Private Sub txtMaxL_GotFocus()
    txtMaxL.Refresh
End Sub

Private Sub txtCurL_GotFocus()
    txtCurL.Refresh
End Sub

Private Sub txtMaxM_GotFocus()
    txtMaxM.Refresh
End Sub

Private Sub txtCurM_GotFocus()
    txtCurM.Refresh
End Sub

Private Sub txtCurM_LostFocus()
    If lstChars2.ListIndex <> -1 Then If Trim(txtCurM.Text) = "" Then txtCurM.Text = "10"
End Sub

Private Sub txtMaxM_LostFocus()
    If lstChars2.ListIndex <> -1 Then If Trim(txtMaxM.Text) = "" Then txtMaxM.Text = "10"
End Sub

Private Sub txtMaxL_LostFocus()
    If lstChars2.ListIndex <> -1 Then If Trim(txtMaxL.Text) = "" Then txtMaxL.Text = "10"
End Sub

Private Sub txtCurL_LostFocus()
    If lstChars2.ListIndex <> -1 Then If Trim(txtCurL.Text) = "" Then txtCurL.Text = "10"
End Sub

Private Sub txtCurSt_LostFocus()
    If lstChars2.ListIndex <> -1 Then If Trim(txtCurSt.Text) = "" Then txtCurSt.Text = "10"
End Sub

Private Sub txtMaxSt_LostFocus()
    If lstChars2.ListIndex <> -1 Then If Trim(txtMaxSt.Text) = "" Then txtMaxSt.Text = "10"
End Sub

Private Sub txtStr_LostFocus()
    If lstChars2.ListIndex <> -1 Then If Trim(txtStr.Text) = "" Then txtStr.Text = "10"
End Sub

Private Sub txtDex_LostFocus()
    If lstChars2.ListIndex <> -1 Then If Trim(txtDex.Text) = "" Then txtDex.Text = "10"
End Sub

'Private Sub txtVit_Change()
    'txtmaxst = txtvit
'End Sub

Private Sub txtVit_LostFocus()
    If lstChars2.ListIndex <> -1 Then If Trim(txtVit.Text) = "" Then txtVit.Text = "10"
End Sub

Private Sub txtEnergy_LostFocus()
    If lstChars2.ListIndex <> -1 Then If Trim(txtEnergy.Text) = "" Then txtEnergy.Text = "10"
End Sub

Private Sub txtVit_GotFocus()
    txtVit.Refresh
End Sub

Private Sub txtEnergy_GotFocus()
    txtEnergy.Refresh
End Sub

Private Sub GetAttributes()
    SetFiles
    txtStr.Text = GetAttrib(sFile, Strength)
    txtDex.Text = GetAttrib(sFile, Dexterity)
    txtVit.Text = GetAttrib(sFile, Vitality)
    txtEnergy.Text = GetAttrib(sFile, Energy)
    txtMaxSt = GetAttrib(sFile, Stamina)
    txtMaxM = GetAttrib(sFile, Mana)
    txtMaxL = GetAttrib(sFile, Life)
    txtCurSt = GetAttrib(sFile, CurStamina)
    txtCurL = GetAttrib(sFile, CurLife)
    txtCurM = GetAttrib(sFile, CurMana)
End Sub

Private Sub MsgNoChar()
    MsgBox "Select A Character First!", vbExclamation, ""
End Sub

Private Sub BackUp()
On Error Resume Next
    SetFiles
    Kill sBKFile 'Get rid of previous backup files
    FileCopy sFile, sBKFile 'Do backup
End Sub

Private Sub SetTitle()
    Dim sStr As String
    SetFiles
    sStr = GetStatus(sFile)
    Select Case sStr
    Case ""
        sStr = "None (Not killed Diablo yet)"
    Case "Sir"
        sStr = "Sir/Dame (Nightmare)"
    Case "Dame"
        sStr = "Sir/Dame (Nightmare)"
    Case "Lord"
        sStr = "Lord/Lady (Hell)"
    Case "Lady"
        sStr = "Lord/Lady (Hell)"
    Case "Baron"
        sStr = "Baron/Baroness (Passed all levels)"
    Case "Baroness"
        sStr = "Baron/Baroness (Passed all levels)"
    End Select
    If fraAttr.Visible = True Then cboTitle.Text = sStr
End Sub

Private Sub ApplyTitle()
    SetFiles
    Select Case Left(cboTitle.Text, 3)
    Case "No"
        SetStatus sFile, Normal
    Case "Sir"
        SetStatus sFile, Sir
    Case "Lor"
        SetStatus sFile, Lord
    Case "Bar"
        SetStatus sFile, Baron
    End Select
    EnumChars
End Sub
