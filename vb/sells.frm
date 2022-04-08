VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form sells 
   BackColor       =   &H80000004&
   Caption         =   "Form1"
   ClientHeight    =   8700
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16380
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8700
   ScaleWidth      =   16380
   WindowState     =   2  'Maximized
   Begin VB.CommandButton insert 
      BackColor       =   &H0080FF80&
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   8760
      Width           =   1695
   End
   Begin VB.CommandButton new 
      BackColor       =   &H00C0C0C0&
      Caption         =   " NEW"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      MaskColor       =   &H00C0FFFF&
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   8760
      Width           =   1695
   End
   Begin VB.CommandButton delete 
      BackColor       =   &H008080FF&
      Caption         =   " DELETE"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   8760
      Width           =   1695
   End
   Begin VB.CommandButton view 
      BackColor       =   &H00FF8080&
      Caption         =   " VIEW"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14040
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   8760
      Width           =   1695
   End
   Begin VB.CommandButton update 
      BackColor       =   &H0080C0FF&
      Caption         =   " UPDATE"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   8760
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFF80&
      Caption         =   "SEARCH"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11640
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   8760
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   720
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000004&
      Height          =   8295
      Left            =   1800
      TabIndex        =   0
      Top             =   360
      Width           =   14175
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   11880
         TabIndex        =   92
         Text            =   "Combo2"
         Top             =   960
         Width           =   1815
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0080FF80&
         Caption         =   " ADD"
         Height          =   375
         Left            =   12720
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   4920
         Width           =   1095
      End
      Begin VB.ListBox List3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2430
         ItemData        =   "sells.frx":0000
         Left            =   6840
         List            =   "sells.frx":0002
         TabIndex        =   78
         Top             =   4800
         Width           =   1110
      End
      Begin VB.ListBox List4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2430
         ItemData        =   "sells.frx":0004
         Left            =   7920
         List            =   "sells.frx":0006
         TabIndex        =   77
         Top             =   4800
         Width           =   1470
      End
      Begin VB.ListBox tax 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2430
         ItemData        =   "sells.frx":0008
         Left            =   9360
         List            =   "sells.frx":000A
         TabIndex        =   76
         Top             =   4800
         Width           =   870
      End
      Begin VB.ListBox List7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2430
         ItemData        =   "sells.frx":000C
         Left            =   10200
         List            =   "sells.frx":000E
         TabIndex        =   75
         Top             =   4800
         Width           =   1110
      End
      Begin VB.ListBox List8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2430
         ItemData        =   "sells.frx":0010
         Left            =   11280
         List            =   "sells.frx":0012
         TabIndex        =   74
         Top             =   4800
         Width           =   1095
      End
      Begin VB.ListBox List9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2430
         ItemData        =   "sells.frx":0014
         Left            =   4440
         List            =   "sells.frx":0016
         TabIndex        =   72
         Top             =   4800
         Width           =   1590
      End
      Begin VB.ListBox List6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2430
         ItemData        =   "sells.frx":0018
         Left            =   3480
         List            =   "sells.frx":001A
         TabIndex        =   69
         Top             =   4800
         Width           =   990
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   10320
         MaxLength       =   4
         TabIndex        =   67
         Top             =   3840
         Width           =   1080
      End
      Begin VB.TextBox qty 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   12360
         MaxLength       =   5
         TabIndex        =   65
         Top             =   3840
         Width           =   1080
      End
      Begin VB.ComboBox Combo5 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   7440
         Style           =   2  'Dropdown List
         TabIndex        =   64
         Top             =   3840
         Width           =   1815
      End
      Begin VB.ComboBox Combo4 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4920
         TabIndex        =   62
         Text            =   "Combo4"
         Top             =   3120
         Width           =   1695
      End
      Begin VB.TextBox order_date 
         BackColor       =   &H00C0C0C0&
         Height          =   360
         Left            =   8520
         TabIndex        =   61
         Top             =   3120
         Width           =   1695
      End
      Begin VB.TextBox address 
         Height          =   405
         Left            =   11880
         TabIndex        =   60
         Top             =   2400
         Width           =   1815
      End
      Begin VB.TextBox phone_no 
         Height          =   375
         Left            =   11880
         TabIndex        =   59
         Top             =   1920
         Width           =   1815
      End
      Begin VB.TextBox customer_name 
         Height          =   405
         Left            =   11880
         TabIndex        =   58
         Top             =   1440
         Width           =   1815
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H80000004&
         Caption         =   "Cash"
         Height          =   255
         Left            =   2520
         TabIndex        =   55
         Top             =   7920
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   5760
         TabIndex        =   50
         Text            =   "0"
         Top             =   7800
         Width           =   1695
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   8520
         TabIndex        =   49
         Text            =   "0"
         Top             =   7800
         Width           =   1695
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   11280
         TabIndex        =   48
         Text            =   "0"
         Top             =   7800
         Width           =   1695
      End
      Begin VB.ComboBox Combo3 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   4920
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   3840
         Width           =   1455
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2430
         ItemData        =   "sells.frx":001C
         Left            =   6000
         List            =   "sells.frx":001E
         TabIndex        =   34
         Top             =   4800
         Width           =   870
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Caption         =   " DELETE"
         Height          =   375
         Left            =   12720
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   5640
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   3840
         Width           =   1815
      End
      Begin VB.ListBox p_id 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2430
         ItemData        =   "sells.frx":0020
         Left            =   2520
         List            =   "sells.frx":0022
         TabIndex        =   71
         Top             =   4800
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2430
         ItemData        =   "sells.frx":0024
         Left            =   1440
         List            =   "sells.frx":0026
         TabIndex        =   33
         Top             =   4800
         Width           =   2055
      End
      Begin VB.ComboBox Combo6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4200
         Style           =   2  'Dropdown List
         TabIndex        =   73
         Top             =   360
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H80000004&
         Caption         =   "Check"
         Height          =   255
         Left            =   2520
         TabIndex        =   47
         Top             =   7920
         Width           =   855
      End
      Begin VB.ListBox List5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2430
         ItemData        =   "sells.frx":0028
         Left            =   480
         List            =   "sells.frx":002A
         TabIndex        =   56
         Top             =   4800
         Width           =   990
      End
      Begin VB.ListBox sn 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2430
         ItemData        =   "sells.frx":002C
         Left            =   480
         List            =   "sells.frx":002E
         TabIndex        =   42
         Top             =   4800
         Width           =   990
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   5400
         TabIndex        =   91
         Top             =   7920
         Width           =   120
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   12120
         TabIndex        =   90
         Top             =   3840
         Width           =   120
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   10080
         TabIndex        =   89
         Top             =   3840
         Width           =   120
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   7080
         TabIndex        =   88
         Top             =   3840
         Width           =   120
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   4680
         TabIndex        =   87
         Top             =   3840
         Width           =   120
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   1680
         TabIndex        =   86
         Top             =   3840
         Width           =   120
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   1800
         TabIndex        =   85
         Top             =   480
         Width           =   120
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   10680
         TabIndex        =   84
         Top             =   960
         Width           =   120
      End
      Begin VB.Line Line14 
         BorderStyle     =   2  'Dash
         X1              =   1320
         X2              =   13440
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Line Line13 
         BorderStyle     =   4  'Dash-Dot
         X1              =   7200
         X2              =   7200
         Y1              =   120
         Y2              =   3000
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Brand"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3600
         TabIndex        =   70
         Top             =   4440
         Width           =   510
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   9480
         TabIndex        =   66
         Top             =   3840
         Width           =   510
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6600
         TabIndex        =   63
         Top             =   3840
         Width           =   420
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Unit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5040
         TabIndex        =   57
         Top             =   4440
         Width           =   420
      End
      Begin VB.Line Line12 
         X1              =   3480
         X2              =   3480
         Y1              =   4320
         Y2              =   4800
      End
      Begin VB.Line Line2 
         X1              =   1460
         X2              =   1460
         Y1              =   4800
         Y2              =   4320
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackColor       =   &H80000007&
         BackStyle       =   0  'Transparent
         Caption         =   "Mode of Payment :-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   240
         TabIndex        =   54
         Top             =   7920
         Width           =   2010
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackColor       =   &H80000007&
         BackStyle       =   0  'Transparent
         Caption         =   "Advance Pay"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   3840
         TabIndex        =   53
         Top             =   7875
         Width           =   1395
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H80000007&
         BackStyle       =   0  'Transparent
         Caption         =   "Dues :-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   7680
         TabIndex        =   52
         Top             =   7875
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000007&
         BackStyle       =   0  'Transparent
         Caption         =   "Total :-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   10440
         TabIndex        =   51
         Top             =   7875
         Width           =   750
      End
      Begin VB.Shape Shape11 
         Height          =   375
         Left            =   2280
         Top             =   7860
         Width           =   1215
      End
      Begin VB.Label invoice_no 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   2400
         TabIndex        =   45
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label tax_total 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "tax"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   10575
         TabIndex        =   44
         Top             =   7200
         Width           =   345
      End
      Begin VB.Label pay_total 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   11580
         TabIndex        =   43
         Top             =   7200
         Width           =   495
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Qty"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   11640
         TabIndex        =   35
         Top             =   3840
         Width           =   420
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   480
         X2              =   13800
         Y1              =   3600
         Y2              =   3600
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " S.N"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   720
         TabIndex        =   31
         Top             =   4440
         Width           =   390
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2160
         TabIndex        =   30
         Top             =   4440
         Width           =   975
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL :-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   720
         TabIndex        =   29
         Top             =   7305
         Width           =   795
      End
      Begin VB.Line Line11 
         X1              =   480
         X2              =   12360
         Y1              =   7170
         Y2              =   7200
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   11520
         TabIndex        =   28
         Top             =   4440
         Width           =   510
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Tax Amount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   10200
         TabIndex        =   27
         Top             =   4440
         Width           =   1080
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Tax"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   9480
         TabIndex        =   26
         Top             =   4440
         Width           =   390
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Net Amount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   8040
         TabIndex        =   25
         Top             =   4440
         Width           =   1065
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6960
         TabIndex        =   24
         Top             =   4440
         Width           =   720
      End
      Begin VB.Line Line10 
         X1              =   480
         X2              =   12360
         Y1              =   4800
         Y2              =   4800
      End
      Begin VB.Line Line9 
         X1              =   11290
         X2              =   11290
         Y1              =   4320
         Y2              =   7560
      End
      Begin VB.Line Line8 
         X1              =   10220
         X2              =   10220
         Y1              =   4320
         Y2              =   7560
      End
      Begin VB.Line Line7 
         X1              =   9380
         X2              =   9380
         Y1              =   4320
         Y2              =   7200
      End
      Begin VB.Line Line6 
         X1              =   7940
         X2              =   7940
         Y1              =   4320
         Y2              =   7200
      End
      Begin VB.Line Line5 
         X1              =   6840
         X2              =   6840
         Y1              =   4320
         Y2              =   7200
      End
      Begin VB.Line Line4 
         X1              =   6020
         X2              =   6020
         Y1              =   4320
         Y2              =   7200
      End
      Begin VB.Line Line3 
         X1              =   4440
         X2              =   4440
         Y1              =   4320
         Y2              =   7200
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   " Order Date :-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   7080
         TabIndex        =   22
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Order Number :-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   3240
         TabIndex        =   21
         Top             =   3120
         Width           =   1425
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Marufganj,  Patna City    (bihar)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2760
         TabIndex        =   20
         Top             =   2400
         Width           =   3240
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " 7764946866"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2760
         TabIndex        =   19
         Top             =   1920
         Width           =   1275
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " 06AAKFG9614J"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3000
         TabIndex        =   18
         Top             =   1440
         Width           =   1650
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Laxmi Narayana Traders"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3000
         TabIndex        =   17
         Top             =   960
         Width           =   2640
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Address :-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   840
         TabIndex        =   16
         Top             =   2400
         Width           =   1140
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Phone No :-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   840
         TabIndex        =   15
         Top             =   1920
         Width           =   1290
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " GST No :-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   840
         TabIndex        =   14
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Company Name  :-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   720
         TabIndex        =   13
         Top             =   960
         Width           =   1980
      End
      Begin VB.Label invoice_date 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   12840
         TabIndex        =   12
         Top             =   360
         Width           =   45
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Customer Name  :-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   9240
         TabIndex        =   11
         Top             =   1440
         Width           =   1980
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Customer ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   9240
         TabIndex        =   10
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Address  :-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   9480
         TabIndex        =   9
         Top             =   2400
         Width           =   1200
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Phone No  :-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   9360
         TabIndex        =   8
         Top             =   1920
         Width           =   1350
      End
      Begin VB.Label Label16 
         Caption         =   " Quantity  *"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         TabIndex        =   6
         Top             =   8880
         Width           =   1575
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6240
         TabIndex        =   5
         Top             =   4440
         Width           =   420
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Brand"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3960
         TabIndex        =   4
         Top             =   3840
         Width           =   690
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Product Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   3
         Top             =   3840
         Width           =   1275
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice Date  *"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   10920
         TabIndex        =   2
         Top             =   360
         Width           =   1530
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   480
         TabIndex        =   1
         Top             =   480
         Width           =   1125
      End
      Begin VB.Shape Shape1 
         Height          =   3255
         Left            =   480
         Top             =   4320
         Width           =   11895
      End
      Begin VB.Label product_idfe 
         BackStyle       =   0  'Transparent
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   32
         Top             =   3840
         Visible         =   0   'False
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   8295
      Left            =   1800
      TabIndex        =   80
      Top             =   360
      Width           =   14175
      Begin VB.CommandButton Command4 
         Caption         =   "CLOSE"
         Height          =   375
         Left            =   6720
         TabIndex        =   81
         Top             =   7440
         Width           =   1095
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "sells.frx":0030
         Height          =   2895
         Left            =   120
         TabIndex        =   82
         Top             =   3960
         Width           =   13815
         _ExtentX        =   24368
         _ExtentY        =   5106
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   11
         BeginProperty Column00 
            DataField       =   "SN_NO"
            Caption         =   "SN_NO"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16393
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "DESCRIPTION"
            Caption         =   "DESCRIPTION"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16393
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "UNIT_PRICE"
            Caption         =   "UNIT_PRICE"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16393
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "QTY"
            Caption         =   "QTY"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16393
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "NET_AMOUNT"
            Caption         =   "NET_AMOUNT"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16393
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "TAX_AMOUNT"
            Caption         =   "TAX_AMOUNT"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16393
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "TOT_AMOUNT"
            Caption         =   "TOT_AMOUNT"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16393
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "INVOICE_NO"
            Caption         =   "INVOICE_NO"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16393
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "PRODUCT_ID"
            Caption         =   "PRODUCT_ID"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16393
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column09 
            DataField       =   "BRAND"
            Caption         =   "BRAND"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16393
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column10 
            DataField       =   "UNIT"
            Caption         =   "UNIT"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16393
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   840.189
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   989.858
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1289.764
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1289.764
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   1739.906
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "sells.frx":0045
         Height          =   2295
         Left            =   720
         TabIndex        =   83
         Top             =   720
         Width           =   12135
         _ExtentX        =   21405
         _ExtentY        =   4048
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   9
         BeginProperty Column00 
            DataField       =   "INVOICE_NO"
            Caption         =   "INVOICE_NO"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16393
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "INVOICE_DATE"
            Caption         =   "INVOICE_DATE"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16393
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "ORDER_DATE"
            Caption         =   "ORDER_DATE"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16393
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "MODE_OF_PAYMENT"
            Caption         =   "MODE_OF_PAYMENT"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16393
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "ADVANCE_PAY"
            Caption         =   "ADVANCE_PAY"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16393
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "DUES"
            Caption         =   "DUES"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16393
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "PAY_AMOUNT"
            Caption         =   "PAY_AMOUNT"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16393
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "CUSTOMER_ID"
            Caption         =   "CUSTOMER_ID"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16393
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "ORDER_NO"
            Caption         =   "ORDER_NO"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16393
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   989.858
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1665.071
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1184.882
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   915.024
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1124.787
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1184.882
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1289.764
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   495
         Left            =   480
         Top             =   4320
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   873
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
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
         Connect         =   "Provider=MSDAORA.1;Password=lnt123;User ID=lnt;Persist Security Info=True"
         OLEDBString     =   "Provider=MSDAORA.1;Password=lnt123;User ID=lnt;Persist Security Info=True"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   "lnt"
         Password        =   "lnt123"
         RecordSource    =   "select * from sell_product"
         Caption         =   "Adodc2"
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
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   495
         Left            =   5520
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
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
         Connect         =   "Provider=MSDAORA.1;Password=lnt123;User ID=lnt;Persist Security Info=True"
         OLEDBString     =   "Provider=MSDAORA.1;Password=lnt123;User ID=lnt;Persist Security Info=True"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   "lnt"
         Password        =   "lnt123"
         RecordSource    =   "select * from invoice_detail"
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
   End
   Begin VB.Label Label38 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "SELL PRODUCT"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   4920
      TabIndex        =   68
      Top             =   0
      Width           =   9045
   End
End
Attribute VB_Name = "sells"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Dim opt As String
Dim sql As String
Dim net As Integer
Dim no As Integer
Dim i As Integer
Dim z As Integer
Dim b As Integer
Public ind As Integer
Public ind1 As Integer
Public invv As String
Dim qtt As Integer


Private Sub Combo1_Click()
If Combo4.Text = "" Then

If Combo2.Text = "Select Customer ID" Or Combo2.Text = "" Then
MsgBox "chose customer Id First"
Combo2.SetFocus
Else
sql = "select product_id from product_detail where product_name='" + Combo1.Text + "'"
Set r = c.Execute(sql)
product_idfe.Caption = r.Fields(0)
End If

sql = "select distinct(brand) from ordered_product where product_name='" + Combo1.Text + "'"
Set r = c.Execute(sql)
Combo3.clear
Do While Not r.EOF
Combo3.AddItem r!brand
r.MoveNext
Loop

Else

Set r = c.Execute("select * from customer_ordered_product ")

Combo3.Text = r.Fields(2)
Combo3.Enabled = False
qty.Text = r.Fields(5)
'product_name.Caption = r.Fields(1)

End If
End Sub



Private Sub Combo2_Click()
Combo1.Enabled = True
customer_combo
End Sub



Private Sub Combo3_Click()
Set r = c.Execute("select unit from ordered_product where brand='" + Combo3.Text + "' and product_name='" + Combo1.Text + "'")
Combo5.clear
Do While Not r.EOF
Combo5.AddItem r!unit
r.MoveNext
Loop
End Sub

Private Sub Combo4_Change()
If Combo4.Text <> blank Or Combo4.Text = "NULL" Then
 Combo1.Enabled = False
 Combo3.Enabled = False
 Combo5.Enabled = False
 Text1.Enabled = False
 qty.Enabled = False
 Command2.Enabled = False
 Command1.Enabled = False
Else
  Combo1.Enabled = True
 Combo3.Enabled = True
 Combo5.Enabled = True
 Text1.Enabled = True
 qty.Enabled = True
 Command1.Enabled = True
 Command2.Enabled = True
End If
End Sub

Private Sub Combo4_click()

If Combo4.Text <> blank Then
 Combo1.Enabled = False
 Combo3.Enabled = False
 Combo5.Enabled = False
 Text1.Enabled = False
 qty.Enabled = False
 Command2.Enabled = False
 Command1.Enabled = False
Else
  Combo1.Enabled = True
 Combo3.Enabled = True
 Combo5.Enabled = True
 Text1.Enabled = True
 qty.Enabled = True
 Command2.Enabled = True
 Command1.Enabled = True
End If

Combo2.Enabled = False



List1.clear
List2.clear
List3.clear
List4.clear
List5.clear
List6.clear
List7.clear
List8.clear
List9.clear
p_id.clear
tax.clear
sn.clear

Set r = c.Execute("select * from customer_ordered_product where ord_no='" + Combo4.Text + "'")

While r.EOF = False
i = 0
sn.AddItem b + 1
List1.AddItem r.Fields(1)
p_id.AddItem r.Fields(8)
List9.AddItem r.Fields(3)
List2.AddItem r.Fields(4)
List3.AddItem r.Fields(5)
List4.AddItem Val(List3.List(i)) * Val(List2.List(i))
tax.AddItem r.Fields(6)
List6.AddItem r.Fields(2)
List7.AddItem List4.List(i) * (tax.List(i) / 100)
List8.AddItem r.Fields(7)


If List5.List(0) = "" Then
sql = "select max(sn_no) from sell_product"
Set r = c.Execute(sql)
If IsNull(r.Fields(0)) Then
no = 0
Else
no = r.Fields(0) + 1
End If
List5.AddItem no
Else
List5.AddItem no + 1
no = no + 1
End If

r.MoveNext
Wend

For j = 0 To List8.ListCount - 1
t = t + Val(List7.List(j))
tax_total.Caption = t
tot = tot + Val(List8.List(j))
pay_total.Caption = tot
Next
Text4.Text = pay_total.Caption
Set r = New ADODB.Recordset

Set r = New ADODB.Recordset
Set r = c.Execute("select * from customer_order_detail where order_number='" + Combo4.Text + "'")
Combo2.Text = r.Fields(11)

customer_name.Text = r.Fields(8)
Text2.Text = r.Fields(5)
order_date.Text = r.Fields(1)
address.Text = r.Fields(9)
Text3.Text = r.Fields(6)
If r.Fields(4) = "check" Then
Option1.Value = True
Else
Option2.Value = True
End If
r.close
Set r = c.Execute("select phone_no from customer_detail where customer_id='" + Combo2.Text + "' ")
phone_no.Text = r.Fields(0)

b = b + 1
End Sub

Private Sub Combo5_Click()


Set r = c.Execute("select rate,quantity from ordered_product where unit='" & Combo5.Text & "' and brand='" & Combo3.Text & "' and p_id='" & product_idfe.Caption & "'")
qtt = r.Fields(1)
Text1.Text = r.Fields(0)
End Sub

Private Sub Combo6_Click()
Dim i As Integer
ind1 = Combo6.ListIndex
b = 0
Set r = c.Execute("select * from invoice_detail where invoice_no='" + Combo6.Text + "'")
 invoice_date.Caption = r.Fields(1)
 Combo2.Text = r.Fields(7)
 
 If IsNull(r.Fields(8)) Then
  Combo4.Text = "None"
 Else
  Combo4.Text = r.Fields(8)
 End If
  If IsNull(r.Fields(2)) Then
  order_date.Text = "None"
 Else
 
  order_date.Text = r.Fields(2)
 End If
 If r.Fields(3) = "cash" Then
  Option2.Value = True
 Else
  Option1.Value = True
 End If
 
 Text2.Text = r.Fields(4)
 Text2.Enabled = False

 List1.clear
 List2.clear
 List3.clear
 List4.clear
 List5.clear
 sn.clear
 tax.clear
 List6.clear
 List7.clear
 List8.clear
 List9.clear
 p_id.clear
 Set r = c.Execute("select * from sell_product where invoice_no='" + Combo6.Text + "'")
 Do While Not r.EOF
 
   sn.AddItem b + 1
   List5.AddItem r.Fields(0)
   List1.AddItem r.Fields(1)
   List9.AddItem r.Fields(10)
   List2.AddItem r.Fields(2)
   List3.AddItem r.Fields(3)
   p_id.AddItem r.Fields(8)
   List6.AddItem r.Fields(9)
   List4.AddItem Val(List3.List(i)) * Val(List2.List(i))
   
  
    Set r = c.Execute("select gst from customer_ordered_product where p_id='" + p_id.List(i) + "' and brand='" + List6.List(i) + "' and unit='" + List9.List(i) + "'")
    tax.AddItem r.Fields(0)
 
   List7.AddItem List4.List(i) * (tax.List(i) / 100)
   List8.AddItem Val(List4.List(i)) + Val(List7.List(i))
  
 r.MoveNext
 Loop
 
 For j = 0 To List8.ListCount - 1
  t = t + Val(List7.List(j))
  tax_total.Caption = t
  tot = tot + Val(List8.List(j))
  pay_total.Caption = tot
 Next
 
 If pay_total.Caption = "" Or pay_total.Caption = "total" Then
pay = 0
Else
pay = pay_total.Caption
End If

If Text2.Text = "" Then
t2 = 0
Else
t2 = Text2.Text
End If
 
 Text3.Text = Val(pay - Text2.Text)
Text4.Text = pay

Set r = c.Execute("select * from customer_detail where customer_id='" + Combo2.Text + "'")
customer_name.Text = r.Fields(1)
phone_no.Text = r.Fields(6)
address.Text = r.Fields(3)

End Sub

Private Sub Command1_Click()
Dim l, l1, l2, l3, l4, l5, l6, l7, l8, l9 As Integer
l = sn.ListIndex
l1 = List1.ListIndex
l2 = List2.ListIndex
l3 = List3.ListIndex
l4 = List4.ListIndex
l5 = tax.ListIndex
'l6 = List1.ListIndex
l7 = List7.ListIndex
l8 = List8.ListIndex
If (l >= 0) Or (l1 >= 0) Or (l2 >= 0) Or (l3 >= 0) Or (l4 >= 0) Or (l5 >= 0) Or (l7 >= 0) Or (l8 >= 0) Then
sn.RemoveItem l
List1.RemoveItem l1
List2.RemoveItem l2
List3.RemoveItem l3
List4.RemoveItem l4
tax.RemoveItem l5
'List6.RemoveItem l6
List7.RemoveItem l7
List8.RemoveItem l8
End If
End Sub

Private Sub Command2_Click()
Dim j As Integer
Dim k As Integer
If Combo1.Text = blank Or Combo3.Text = "" Or Combo5.Text = "" Or Text1.Text = "" Or qty.Text = "" Then
 a = MsgBox("Some fields are Blank", vbOKOnly + vbCritical, "Warning")
Else

 List1.BackColor = vbWhite
List1.BackColor = vbWhite
List2.BackColor = vbWhite
List3.BackColor = vbWhite
List4.BackColor = vbWhite
List5.BackColor = vbWhite
List6.BackColor = vbWhite
List7.BackColor = vbWhite
List8.BackColor = vbWhite
List9.BackColor = vbWhite
tax.BackColor = vbWhite
sn.BackColor = vbWhite

 If Combo4.Text = "" Then
  sn.AddItem b + 1
  List1.AddItem Combo1.Text
  List9.AddItem Combo5.Text
  List2.AddItem Text1.Text
  List3.AddItem qty.Text
  p_id.AddItem product_idfe.Caption

  List4.AddItem Val(List3.List(i)) * Val(List2.List(i))

  Set r = c.Execute("select gst from customer_ordered_product where p_id='" + product_idfe.Caption + "'")
  If IsNull(r.Fields(0)) Then
   tx = 18
  Else
   tx = r.Fields(0)
  End If
  tax.AddItem tx
  List7.AddItem List4.List(i) * (tax.List(i) / 100)
  List6.AddItem Combo3.Text
  List8.AddItem Val(List4.List(i)) + Val(List7.List(i))

  If List5.List(0) = "" Then
    sql = "select max(sn_no) from sell_product"
    Set r = c.Execute(sql)
    If IsNull(r.Fields(0)) Then
        no = 1
    Else
     no = r.Fields(0) + 1
    End If
     List5.AddItem no
  Else
    List5.AddItem no + 1
    no = no + 1
  End If

 For j = 0 To List8.ListCount - 1
  t = t + Val(List7.List(j))
  tax_total.Caption = t
  tot = tot + Val(List8.List(j))
  pay_total.Caption = tot
 Next

 If pay_total.Caption = "" Or pay_total.Caption = "total" Then
  pay = 0
 Else
  pay = pay_total.Caption
 End If

 If Text2.Text = "" Then
  t2 = 0
 Else
  t2 = Text2.Text
 End If
 Text3.Text = Val(pay - t2)
 Text4.Text = pay_total.Caption
 b = b + 1
 
ElseIf Combo6.Visible = True Then
 List1.RemoveItem (ind)
 List6.RemoveItem (ind)
 List9.RemoveItem (ind)
 List2.RemoveItem (ind)
 List3.RemoveItem (ind)
 List4.RemoveItem (ind)
 tax.RemoveItem (ind)
 List7.RemoveItem (ind)
 List8.RemoveItem (ind)
 p_id.RemoveItem (ind)

 List1.AddItem Combo1.Text
 List9.AddItem Combo5.Text
 List2.AddItem Text1.Text
 List3.AddItem qty.Text
 p_id.AddItem product_idfe.Caption

 List4.AddItem Val(List3.List(i)) * Val(List2.List(i))

 Set r = c.Execute("select gst from customer_ordered_product where p_id='" + product_idfe.Caption + "'")
 tax.AddItem r.Fields(0)
 List7.AddItem List4.List(i) * (tax.List(i) / 100)

 List6.AddItem Combo3.Text
 List8.AddItem Val(List4.List(i)) + Val(List7.List(i))

 For j = 0 To List8.ListCount - 1
 t = t + Val(List7.List(j))
 tax_total.Caption = t
 tot = tot + Val(List8.List(j))
 pay_total.Caption = tot
 Next

 If pay_total.Caption = "" Or pay_total.Caption = "total" Then
  pay = 0
 Else
  pay = pay_total.Caption
 End If

 If Text2.Text = "" Then
  t2 = 0
 Else
  t2 = Text2.Text
 End If
 Text3.Text = Val(pay - t2)
 Text4.Text = pay_total.Caption
 End If
End If
End Sub



Private Sub Command3_Click()
Combo6.Visible = True
invoice_no.Visible = False

insert.Enabled = False
update.Enabled = True
delete.Enabled = True

Set r = c.Execute("select distinct(invoice_no) from invoice_detail")
Combo6.clear
Do While Not r.EOF
 Combo6.AddItem r.Fields(0)
r.MoveNext
Loop
End Sub



Private Sub Command4_Click()
Frame1.Visible = True
Frame2.Visible = False
End Sub


Private Sub delete_Click()
ans = MsgBox("Do you want to Delete", vbYesNo + vbInformation, "Update")
If ans = vbYes Then
 
 Set r = c.Execute("delete sell_product where  INVOICE_NO='" + Combo6.Text + "'")
 Set r = c.Execute("delete invoice_detail where  INVOICE_NO='" + Combo6.Text + "'")
End If
ans = MsgBox("record deleted", vbOKOnly + vbInformation, "Delete")
clear
Combo6.RemoveItem (ind1)
End Sub

Private Sub Form_Load()
MDIForm1.Picture2.Visible = True

Connection
auto_inv_no
invoice_date.Caption = Date

combo2_item

If Option1.Value = True Then
opt = Option1.Caption
Else
opt = Option2.Caption
End If


Set r = c.Execute("select order_number from customer_order_detail")
Combo4.clear
Do While Not r.EOF
Combo4.AddItem r.Fields(0)
r.MoveNext
Loop

Set r = c.Execute("select distinct(product_name) from ordered_product")
Combo1.clear
Do While Not r.EOF
Combo1.AddItem r!product_name
r.MoveNext
Loop

sells.Caption = "sale"
Text2.Text = 0
End Sub
Public Function combo2_item()
Set r = New ADODB.Recordset
sql = "select customer_id from customer_detail "
Set r = c.Execute(sql)
Do While Not r.EOF
Combo2.AddItem r!customer_id
r.MoveNext
Loop
End Function

Public Function customer_combo()
Set r = New ADODB.Recordset
sql = "select customer_name,phone_no,address from customer_detail where customer_id='" + Combo2.Text + "'"
Set r = c.Execute(sql)

customer_name.Text = r.Fields(0)
phone_no.Text = r.Fields(1)
address.Text = r.Fields(2)

End Function



Private Sub insert_Click()

ans = MsgBox("Do you Want to Save", vbYesNo + vbInformation, "For Save")
If ans = vbYes Then
If Combo2.Text = "" Then
 Combo2.BackColor = &HC0C0FF
ElseIf Combo2.Text = "Select Customer ID" Then
 Combo2.BackColor = &HC0C0FF
ElseIf List1.List(0) = "" Or List2.List(0) = "" Or List3.List(0) = "" Or List4.List(0) = "" Or List7.List(0) = "" Or List5.List(0) = "" Or List8.List(0) = "" Or p_id.List(0) = "" Or sn.List(0) = "" Or tax.List(0) = "" Or List6.List(0) = "" Then
 List1.BackColor = &HC0C0FF
List1.BackColor = &HC0C0FF
List2.BackColor = &HC0C0FF
List3.BackColor = &HC0C0FF
List4.BackColor = &HC0C0FF
List5.BackColor = &HC0C0FF
List6.BackColor = &HC0C0FF
List7.BackColor = &HC0C0FF
List8.BackColor = &HC0C0FF
List9.BackColor = &HC0C0FF
tax.BackColor = &HC0C0FF
sn.BackColor = &HC0C0FF
MsgBox "Add Product in Listbox ", vbCritical, "Warining"
Else
 Set r = New ADODB.Recordset
 If Combo4.Text = "" Then
  ord = "NULL"
 Else
  ord = Combo4.Text
 End If

 If order_date.Text = "" Then
  ord_d = Null
 Else
  ord_d = order_date.Text
 End If

 sql = "insert into invoice_detail values('" + invoice_no.Caption + "','" + Format(invoice_date.Caption, "dd/mmm/yyyy") + "','" + Format(ord_d, "dd/mmm/yyyy") + "','" + opt + "'," + Text2.Text + "," + Text3.Text + "," + Text4.Text + ",'" + Combo2.Text + "','" & ord & "')"
 Set r = c.Execute(sql)

 For z = 0 To List2.ListCount - 1
  sql = "insert into sell_product values(" + List5.List(z) + ",'" + List1.List(z) + "'," + List2.List(z) + "," + List3.List(z) + "," + List4.List(z) + "," + List7.List(z) + "," + List8.List(z) + ",'" + invoice_no.Caption + "','" + p_id.List(z) + "','" + List6.List(i) + "','" + List9.List(i) + "') "
  Set r = c.Execute(sql)
 Next
 MsgBox "record saved"
 Set r = c.Execute("update customer_order_detail set status='yes' where order_number='" + Combo4.Text + "'")
 
 For i = 0 To List3.ListCount - 1
  Set r = c.Execute("select avl_quantity from stock_detail where product_nm='" + List1.List(i) + "' and unit='" + List9.List(i) + "'")
  qty1 = Val(r.Fields(0)) - Val(List3.List(i))
  Set r = c.Execute("update stock_detail set avl_quantity= " & qty1 & " where product_nm='" + List1.List(i) + "' and unit='" + List9.List(i) + "'")
 Next
  
  Set r = c.Execute("commit")
  
 ans = MsgBox("Do you Want to Print Bill", vbYesNo + vbInformation, "Bill Print")
 If ans = vbYes Then
  Set r = c.Execute(" select invoice_no from invoice_detail where invoice_no='" + invoice_no.Caption + "'")
 abc = r.Fields(0)
 Set r = c.Execute(" select * from customer_detail where customer_id='" + Combo2.Text + "' ")
  sell_bill_data.sale_bill_cmd abc
   
   sale_bill_report.Sections("section1").Controls("c_name").Caption = r.Fields(1)
   sale_bill_report.Sections("section1").Controls("c_phone_no").Caption = r.Fields(6)
   sale_bill_report.Sections("section1").Controls("c_address").Caption = r.Fields(4) & r.Fields(5) & r.Fields(3)
   
   sale_bill_report.Sections("section1").Controls("label9").Caption = r.Fields(5)
   If IsNull(r.Fields(11)) Then
   r11 = ""
   Else
   r11 = r.Fields(11)
   End If
   
   sale_bill_report.Sections("section1").Controls("bank").Caption = r11
   If IsNull(r.Fields(10)) Then
   r10 = ""
   Else
   r10 = r.Fields(10)
   End If
   sale_bill_report.Sections("section1").Controls("ifsc").Caption = r10
   If IsNull(r.Fields(8)) Then
   r8 = ""
   Else
   r8 = r.Fields(8)
   End If
   sale_bill_report.Sections("section1").Controls("ac").Caption = r8
   If IsNull(r.Fields(9)) Then
   r9 = ""
   Else
   r9 = r.Fields(9)
   End If
   sale_bill_report.Sections("section1").Controls("holder_name").Caption = r9
  
sale_bill_report.Show
sale_bill_report.Refresh
sell_bill_data.rssale_bill_cmd.close

End If

clear
auto_inv_no
Combo2.Enabled = True
End If
End If
End Sub


Private Sub List1_Click()
If Combo6.Visible = True Then
ind = List1.ListIndex

product_idfe.Caption = p_id.List(ind)
Combo1.Text = List1.List(ind)
Combo3.Text = List6.List(ind)
Combo5.Text = List9.List(ind)
Text1.Text = List2.List(ind)
qty.Text = List3.List(ind)

  Combo1.Enabled = True
 Combo3.Enabled = True
 Combo5.Enabled = True
 Text1.Enabled = True
 qty.Enabled = True
 Command1.Enabled = True
 Command2.Enabled = True
 End If
End Sub

Private Sub List2_Click()
If Combo6.Visible = True Then
ind = List2.ListIndex
product_idfe.Caption = p_id.List(ind)
Combo1.Text = List1.List(ind)
Combo3.Text = List6.List(ind)
Combo5.Text = List9.List(ind)
Text1.Text = List2.List(ind)
qty.Text = List3.List(ind)

  Combo1.Enabled = True
 Combo3.Enabled = True
 Combo5.Enabled = True
 Text1.Enabled = True
 qty.Enabled = True
 Command1.Enabled = True
 Command2.Enabled = True
 End If
End Sub

Private Sub List3_Click()
If Combo6.Visible = True Then
ind = List3.ListIndex
product_idfe.Caption = p_id.List(ind)
Combo1.Text = List1.List(ind)
Combo3.Text = List6.List(ind)
Combo5.Text = List9.List(ind)
Text1.Text = List2.List(ind)
qty.Text = List3.List(ind)

  Combo1.Enabled = True
 Combo3.Enabled = True
 Combo5.Enabled = True
 Text1.Enabled = True
 qty.Enabled = True
 Command1.Enabled = True
 Command2.Enabled = True
 End If
End Sub

Private Sub List4_Click()
If Combo6.Visible = True Then
ind = List4.ListIndex
product_idfe.Caption = p_id.List(ind)
Combo1.Text = List1.List(ind)
Combo3.Text = List6.List(ind)
Combo5.Text = List9.List(ind)
Text1.Text = List2.List(ind)
qty.Text = List3.List(ind)

  Combo1.Enabled = True
 Combo3.Enabled = True
 Combo5.Enabled = True
 Text1.Enabled = True
 qty.Enabled = True
 Command1.Enabled = True
 Command2.Enabled = True
 End If
End Sub

Private Sub List6_Click()
If Combo6.Visible = True Then
ind = List6.ListIndex
product_idfe.Caption = p_id.List(ind)
Combo1.Text = List1.List(ind)
Combo3.Text = List6.List(ind)
Combo5.Text = List9.List(ind)
Text1.Text = List2.List(ind)
qty.Text = List3.List(ind)
  Combo1.Enabled = True
 Combo3.Enabled = True
 Combo5.Enabled = True
 Text1.Enabled = True
 qty.Enabled = True
 Command1.Enabled = True
 Command2.Enabled = True
 End If
End Sub

Private Sub List7_Click()
If Combo6.Visible = True Then
ind = List7.ListIndex
product_idfe.Caption = p_id.List(ind)
Combo1.Text = List1.List(ind)
Combo3.Text = List6.List(ind)
Combo5.Text = List9.List(ind)
Text1.Text = List2.List(ind)
qty.Text = List3.List(ind)

  Combo1.Enabled = True
 Combo3.Enabled = True
 Combo5.Enabled = True
 Text1.Enabled = True
 qty.Enabled = True
 Command1.Enabled = True
 Command2.Enabled = True
 End If
End Sub

Private Sub List8_Click()
If Combo6.Visible = True Then
ind = List8.ListIndex
product_idfe.Caption = p_id.List(ind)
Combo1.Text = List1.List(ind)
Combo3.Text = List6.List(ind)
Combo5.Text = List9.List(ind)
Text1.Text = List2.List(ind)
qty.Text = List3.List(ind)

  Combo1.Enabled = True
 Combo3.Enabled = True
 Combo5.Enabled = True
 Text1.Enabled = True
 qty.Enabled = True
 Command1.Enabled = True
 Command2.Enabled = True
 End If
End Sub

Private Sub List9_Click()
If Combo6.Visible = True Then
ind = List9.ListIndex
product_idfe.Caption = p_id.List(ind)
Combo1.Text = List1.List(ind)
Combo3.Text = List6.List(ind)
Combo5.Text = List9.List(ind)
Text1.Text = List2.List(ind)
qty.Text = List3.List(ind)

  Combo1.Enabled = True
 Combo3.Enabled = True
 Combo5.Enabled = True
 Text1.Enabled = True
 qty.Enabled = True
 Command1.Enabled = True
 Command2.Enabled = True
 End If
End Sub

Private Sub new_Click()

auto_inv_no
clear

insert.Enabled = True
update.Enabled = False
delete.Enabled = False

Combo6.Visible = False
invoice_no.Visible = True
Combo2.Enabled = True

End Sub

Public Function auto_inv_no()

Dim a As String
Set r = New ADODB.Recordset
sql = "select max(to_number(substr(invoice_no,6,length(invoice_no)))) from invoice_detail"
Set r = c.Execute(sql)
If IsNull(r.Fields(0)) Then
invoice_no.Caption = "INV" & "00" & 1
Else
invoice_no.Caption = "INV" & "00" & r.Fields(0) + 1
End If
a = invoice_no.Caption
If (a = "INV" & "001" & "0") Then
sql = "select max(to_number(substr(invoice_no,5,length(invoice_no)))) from invoice_detail"
Set r = c.Execute(sql)
invoice_no.Caption = "INV" & "0" & r.Fields(0) + 1
End If

End Function

Public Function clear()
customer_name.Text = ""
phone_no.Text = ""
address.Text = ""
Text1.Text = ""
qty.Text = ""
sn.clear
List5.clear
List1.clear
List2.clear
List3.clear
List4.clear
tax.clear
List6.clear
List7.clear
List8.clear
List9.clear
Combo5.clear
sn.clear

Combo3.clear
p_id.clear
sn.clear

pay_total.Caption = ""
tax_total.Caption = ""
order_date.Text = ""
End Function



Private Sub qty_Change()
If qty.Text = "" Then
ElseIf qty.Text > qtt Then
MsgBox "Invaild Quantity . max Quantity is  =" & qtt, vbCritical
qty.Text = ""
End If
End Sub

Private Sub qty_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn Then
  If KeyAscii = 13 Then
  Command2.SetFocus
  End If
Else
KeyAscii = 0
MsgBox "Enter only number"
End If
End Sub

Private Sub sn_Click()
Dim t As Integer
t = List1.ListIndex
If Combo6.Visible = True Then
ind = sn.ListIndex
product_idfe.Caption = p_id.List(ind)
Combo1.Text = List1.List(ind)
Combo3.Text = List6.List(ind)
Combo5.Text = List9.List(ind)
Text1.Text = List2.List(ind)
qty.Text = List3.List(ind)

  Combo1.Enabled = True
 Combo3.Enabled = True
 Combo5.Enabled = True
 Text1.Enabled = True
 qty.Enabled = True
 Command1.Enabled = True
 Command2.Enabled = True
 End If
End Sub



Private Sub tax_Click()
If Combo6.Visible = True Then
ind = tax.ListIndex

Combo1.Text = List1.List(ind)
Combo3.Text = List6.List(ind)
Combo5.Text = List9.List(ind)
Text1.Text = List2.List(ind)
qty.Text = List3.List(ind)

Combo1.Enabled = True
 Combo3.Enabled = True
 Combo5.Enabled = True
 Text1.Enabled = True
 qty.Enabled = True
 Command1.Enabled = True
 Command2.Enabled = True
 End If
End Sub

Private Sub Text2_change()
If pay_total.Caption = "" Or pay_total.Caption = "total" Then
pay = 0
Else
pay = pay_total.Caption
End If

If Text2.Text = "" Then
t2 = 0
Else
t2 = Text2.Text
End If
Text3.Text = Val(pay - t2)
Text4.Text = pay_total.Caption
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn Then

Else
KeyAscii = 0
MsgBox "Enter only number", vbCritical
End If
End Sub

Private Sub update_Click()
ans = MsgBox("Do you want to update", vbYesNo + vbInformation, "Update")
If ans = vbYes Then
 Set r = c.Execute(" update invoice_detail set customer_id='" + Combo2.Text + "' where invoice_no='" + Combo6.Text + "' ")
 For z = 0 To List1.ListCount - 1
  Set r = c.Execute(" update sell_product set DESCRIPTION='" + List1.List(z) + "',UNIT_PRICE=" + List2.List(z) + ",QTY=" + List3.List(z) + ",NET_AMOUNT=" + List4.List(z) + ",TAX_AMOUNT=" + List7.List(z) + ",TOT_AMOUNT=" + List8.List(z) + ",PRODUCT_ID='" + p_id.List(z) + "',BRAND='" + List6.List(i) + "',UNIT='" + List9.List(i) + "' ")
 Next
 MsgBox "Update Completed"
 clear
End If
End Sub

Private Sub view_Click()
Frame2.Visible = True
Frame1.Visible = False
Adodc1.Refresh
Adodc2.Refresh
End Sub
