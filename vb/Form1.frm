VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form report 
   Caption         =   "Form1"
   ClientHeight    =   7905
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15495
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   7905
   ScaleWidth      =   15495
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command8 
      Caption         =   "Return  Report"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   15330
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   360
      Width           =   2000
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Supplier  Report"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   13260
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   360
      Width           =   2000
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Customer Report"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   11190
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   360
      Width           =   2000
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Product  Report"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   360
      Width           =   2000
   End
   Begin VB.CommandButton Command4 
      Caption         =   " Stock  Report"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   7050
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   360
      Width           =   2000
   End
   Begin VB.CommandButton Command3 
      Caption         =   "  Sell Report"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   4980
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   360
      Width           =   2000
   End
   Begin VB.CommandButton Command2 
      Caption         =   " Purchase  Report"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   2910
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   2000
   End
   Begin VB.CommandButton Command1 
      Caption         =   " Order  Report"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   2000
   End
   Begin VB.Frame sell_report 
      Caption         =   "Sell Report"
      Height          =   5175
      Left            =   4560
      TabIndex        =   10
      Top             =   1680
      Width           =   8175
      Begin VB.CommandButton Command14 
         BackColor       =   &H00FFFF80&
         Caption         =   "REPORT"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   3240
         Width           =   2055
      End
      Begin VB.ComboBox Combo12 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3600
         TabIndex        =   44
         Top             =   2160
         Width           =   2295
      End
      Begin VB.ComboBox Combo11 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "Form1.frx":25CE1
         Left            =   3600
         List            =   "Form1.frx":25CF1
         TabIndex        =   42
         Top             =   1200
         Width           =   2295
      End
      Begin MSComCtl2.DTPicker DTPicker7 
         Height          =   375
         Left            =   2880
         TabIndex        =   70
         Top             =   2160
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   125501441
         CurrentDate     =   43615
      End
      Begin MSComCtl2.DTPicker DTPicker8 
         Height          =   375
         Left            =   5160
         TabIndex        =   71
         Top             =   2160
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   125501441
         CurrentDate     =   43615
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "To"
         Height          =   195
         Left            =   4560
         TabIndex        =   72
         Top             =   2280
         Width           =   195
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1560
         TabIndex        =   45
         Top             =   2160
         Width           =   780
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search By"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1560
         TabIndex        =   43
         Top             =   1200
         Width           =   1245
      End
   End
   Begin VB.Frame purchase_report 
      Caption         =   "purchase Report"
      Height          =   5175
      Left            =   4560
      TabIndex        =   9
      Top             =   1680
      Width           =   8175
      Begin MSComCtl2.DTPicker DTPicker6 
         Height          =   375
         Left            =   5160
         TabIndex        =   68
         Top             =   2400
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   125501441
         CurrentDate     =   43614
      End
      Begin MSComCtl2.DTPicker DTPicker5 
         Height          =   375
         Left            =   2880
         TabIndex        =   67
         Top             =   2400
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   125501441
         CurrentDate     =   43614
      End
      Begin VB.CommandButton Command15 
         BackColor       =   &H00FFFF80&
         Caption         =   "REPORT"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   3600
         Width           =   2055
      End
      Begin VB.ComboBox Combo14 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3840
         TabIndex        =   49
         Top             =   2400
         Width           =   2295
      End
      Begin VB.ComboBox Combo13 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "Form1.frx":25D28
         Left            =   3840
         List            =   "Form1.frx":25D38
         TabIndex        =   47
         Top             =   1440
         Width           =   2295
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "TO"
         Height          =   195
         Left            =   4680
         TabIndex        =   69
         Top             =   2520
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1800
         TabIndex        =   50
         Top             =   2400
         Width           =   780
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search By"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1800
         TabIndex        =   48
         Top             =   1440
         Width           =   1245
      End
   End
   Begin VB.Frame order_report 
      Caption         =   "order Report"
      Height          =   5175
      Left            =   4560
      TabIndex        =   8
      Top             =   1680
      Width           =   8175
      Begin VB.OptionButton Option4 
         Caption         =   "Supplier Order"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1320
         TabIndex        =   53
         Top             =   480
         Width           =   2535
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Customer Order"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4680
         TabIndex        =   52
         Top             =   480
         Width           =   2295
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   5040
         TabIndex        =   61
         Top             =   2400
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   125501441
         CurrentDate     =   43585
      End
      Begin VB.ComboBox Combo16 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "Form1.frx":25D73
         Left            =   3840
         List            =   "Form1.frx":25D75
         TabIndex        =   56
         Top             =   2400
         Width           =   2415
      End
      Begin VB.ComboBox Combo15 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "Form1.frx":25D77
         Left            =   3840
         List            =   "Form1.frx":25D79
         TabIndex        =   55
         Text            =   " "
         Top             =   1320
         Width           =   2415
      End
      Begin VB.ComboBox Combo17 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "Form1.frx":25D7B
         Left            =   3840
         List            =   "Form1.frx":25D7D
         TabIndex        =   59
         Text            =   " "
         Top             =   1320
         Width           =   2415
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2400
         TabIndex        =   60
         Top             =   2400
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         Format          =   125501441
         CurrentDate     =   43585
      End
      Begin VB.CommandButton Command16 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Report"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   3480
         Width           =   2295
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Select"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2160
         TabIndex        =   57
         Top             =   2400
         Width           =   780
      End
      Begin VB.Image Image1 
         Height          =   375
         Left            =   6360
         Picture         =   "Form1.frx":25D7F
         Top             =   2400
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Search By"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         TabIndex        =   58
         Top             =   1440
         Width           =   1155
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TO"
         Height          =   195
         Left            =   4560
         TabIndex        =   62
         Top             =   2520
         Visible         =   0   'False
         Width           =   225
      End
   End
   Begin VB.Frame product_report 
      Caption         =   "Product Report"
      Height          =   5175
      Left            =   4560
      TabIndex        =   12
      Top             =   1680
      Width           =   8175
      Begin VB.CommandButton Command13 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Report"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   3360
         Width           =   2295
      End
      Begin VB.ComboBox Combo9 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "Form1.frx":26302
         Left            =   3720
         List            =   "Form1.frx":26304
         TabIndex        =   33
         Text            =   " "
         Top             =   1920
         Width           =   2415
      End
      Begin VB.ComboBox Combo8 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "Form1.frx":26306
         Left            =   3720
         List            =   "Form1.frx":26313
         TabIndex        =   32
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Select"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2040
         TabIndex        =   36
         Top             =   1920
         Width           =   780
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Search By"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         TabIndex        =   35
         Top             =   960
         Width           =   1155
      End
   End
   Begin VB.Frame customer_report 
      Caption         =   "Customer Report"
      Height          =   5175
      Left            =   4560
      TabIndex        =   13
      Top             =   1680
      Width           =   8175
      Begin VB.CommandButton Command12 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Report"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   3840
         Width           =   2295
      End
      Begin VB.ComboBox Combo7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "Form1.frx":2633F
         Left            =   3480
         List            =   "Form1.frx":26341
         TabIndex        =   28
         Text            =   " "
         Top             =   1800
         Width           =   2415
      End
      Begin VB.ComboBox Combo6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "Form1.frx":26343
         Left            =   3480
         List            =   "Form1.frx":26353
         TabIndex        =   27
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Select"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1800
         TabIndex        =   31
         Top             =   1800
         Width           =   780
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Search By"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         TabIndex        =   30
         Top             =   840
         Width           =   1155
      End
   End
   Begin VB.Frame supplier_report 
      Caption         =   "Supplier Report"
      Height          =   5175
      Left            =   4560
      TabIndex        =   14
      Top             =   1680
      Width           =   8175
      Begin VB.OptionButton Option7 
         Caption         =   "Supplier Product"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5400
         TabIndex        =   75
         Top             =   3240
         Width           =   2175
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Supplier Account"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   74
         Top             =   3240
         Width           =   2175
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Supplier Detail"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   73
         Top             =   3240
         Width           =   2055
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "Form1.frx":2638A
         Left            =   3360
         List            =   "Form1.frx":26397
         TabIndex        =   18
         Top             =   1080
         Width           =   2415
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "Form1.frx":263C7
         Left            =   3360
         List            =   "Form1.frx":263C9
         TabIndex        =   17
         Text            =   " "
         Top             =   2040
         Width           =   2415
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Report"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   4080
         Width           =   2295
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Search By"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   20
         Top             =   1080
         Width           =   1155
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Select"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1680
         TabIndex        =   19
         Top             =   2040
         Width           =   780
      End
   End
   Begin VB.Frame return_report 
      Caption         =   "Return Report"
      Height          =   5175
      Left            =   4560
      TabIndex        =   15
      Top             =   1680
      Width           =   8175
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   375
         Left            =   1920
         TabIndex        =   64
         Top             =   3120
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         Format          =   125501441
         CurrentDate     =   43585
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00FFFF80&
         Caption         =   "REPORT"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   3720
         Width           =   2055
      End
      Begin VB.ComboBox Combo5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "Form1.frx":263CB
         Left            =   3600
         List            =   "Form1.frx":263CD
         TabIndex        =   24
         Top             =   2640
         Width           =   2295
      End
      Begin VB.OptionButton Option2 
         Caption         =   " Sale Return"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4680
         TabIndex        =   22
         Top             =   480
         Width           =   2295
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Purchase Return"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1320
         TabIndex        =   21
         Top             =   480
         Width           =   2535
      End
      Begin MSComCtl2.DTPicker DTPicker4 
         Height          =   375
         Left            =   4560
         TabIndex        =   65
         Top             =   3120
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   125501441
         CurrentDate     =   43585
      End
      Begin VB.ComboBox Combo18 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "Form1.frx":263CF
         Left            =   3600
         List            =   "Form1.frx":263D1
         TabIndex        =   63
         Top             =   1680
         Width           =   2295
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4080
         TabIndex        =   66
         Top             =   3240
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1560
         TabIndex        =   25
         Top             =   2640
         Width           =   780
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search By"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1560
         TabIndex        =   23
         Top             =   1680
         Width           =   1245
      End
   End
   Begin VB.Frame stock_report 
      Caption         =   "StockReport"
      Height          =   5175
      Left            =   4560
      TabIndex        =   11
      Top             =   1680
      Width           =   8175
      Begin VB.CommandButton Command9 
         BackColor       =   &H00FFFF80&
         Caption         =   "REPORT"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   3600
         Width           =   2055
      End
      Begin VB.ComboBox Combo10 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3720
         TabIndex        =   39
         Top             =   2520
         Width           =   2295
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "Form1.frx":263D3
         Left            =   3720
         List            =   "Form1.frx":263E0
         TabIndex        =   37
         Top             =   1560
         Width           =   2295
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1680
         TabIndex        =   40
         Top             =   2520
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search By"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1680
         TabIndex        =   38
         Top             =   1560
         Width           =   1245
      End
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      BorderStyle     =   5  'Dash-Dot-Dot
      FillColor       =   &H00FFC0FF&
      FillStyle       =   0  'Solid
      Height          =   5415
      Left            =   4440
      Top             =   1560
      Width           =   8415
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00C0FFC0&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   720
      Top             =   240
      Width           =   16695
   End
End
Attribute VB_Name = "report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim var As String


Private Sub Check1_Click()
Combo2.Text = ""
Combo3.Text = ""
End Sub

Private Sub Check2_Click()
 Combo2.Text = ""
Combo3.Text = ""
End Sub

Private Sub Check3_Click()
 Combo2.Text = ""
Combo3.Text = ""
End Sub

Private Sub Combo1_Click()

If Combo1.Text = "All stock" Then
 Combo10.Enabled = False
ElseIf Combo1.Text = "Stock No" Then
 Combo10.clear
 Set r = c.Execute("select stock_no from stock_detail")
 Do While Not r.EOF
  Combo10.AddItem r.Fields(0)
  r.MoveNext
 Loop
ElseIf Combo1.Text = "Product Name" Then
 Combo10.clear
 Set r = c.Execute("select distinct(product_nm) from stock_detail")
 Do While Not r.EOF
  Combo10.AddItem r.Fields(0)
  r.MoveNext
 Loop
End If
End Sub

Private Sub Combo11_Click()
If Combo11.Text = "All Sells" Then
 Combo12.Enabled = False
   Combo12.Visible = True
 DTPicker7.Visible = False
 DTPicker8.Visible = False
 Label20.Visible = False
 
ElseIf Combo11.Text = "Invoice No" Then
 Combo12.Enabled = True
   Combo12.Visible = True
 DTPicker7.Visible = False
 DTPicker8.Visible = False
 Label20.Visible = False
 Combo12.clear
 
 Set r = c.Execute("select invoice_no from invoice_detail")
 Do While Not r.EOF
   Combo12.AddItem r.Fields(0)
   r.MoveNext
 Loop
ElseIf Combo11.Text = "Invoice Date" Then
 Combo12.Enabled = True
  Combo12.Visible = True
 DTPicker7.Visible = False
 DTPicker8.Visible = False
 Label20.Visible = False
 Combo12.clear
 Set r = c.Execute("select invoice_date from invoice_detail")
 Do While Not r.EOF
   Combo12.AddItem r.Fields(0)
   r.MoveNext
 Loop
 
ElseIf Combo11.Text = "B/W Two Date" Then
 Combo12.Visible = False
 DTPicker7.Visible = True
 DTPicker8.Visible = True
 Label20.Visible = True
End If
End Sub





Private Sub Command14_Click()
If Combo11.Text = "All Sells" Then
  sell_all_report.Show
  sell_all_report.Refresh
  
ElseIf Combo11.Text = "Invoice No" And Combo12.Text <> "" Then
  sell_data.sell_inv_no Combo12.Text, Combo12.Text
  sell_id_report.Show
  sell_id_report.Refresh
  sell_data.rssell_inv_no.close
  
ElseIf Combo11.Text = "Invoice Date" And Combo12.Text <> "" Then
   sell_data.sell_inv_date Combo12.Text, Combo12.Text
   sell_date_report.Show
   sell_date_report.Refresh
   sell_data.rssell_inv_date.close
   
 ElseIf Combo11.Text = "B/W Two Date" Then
   Set r = c.Execute("select distinct(INVOICE_NO) from invoice_detail where invoice_date between '" + Format(DTPicker7.Value, "dd/mmm/yyyy") + "' AND '" + Format(DTPicker8.Value, "dd/mmm/yyyy") + "' ")

   If IsNull(r.Fields(0)) Then
       MsgBox "Record Not Found", vbCritical
  Else
   dt = r.Fields(0)
   
   sell_data.bw_two_date dt, dt
   sell_bw_two_report.Show
   sell_bw_two_report.Refresh
   sell_data.rsbw_two_date.close
  End If
End If
End Sub

Private Sub Command15_Click()
If Combo13.Text = "All Purchase" And Combo14.Enabled = False Then
  all_purchase_report.Show
ElseIf Combo13.Text = "Invoice ID" And Combo14.Text <> blank Then
  purchase_data.purchase_id_cmd Combo14.Text
  purchase_id_report.Show
  purchase_id_report.Refresh
  purchase_data.rspurchase_id_cmd.close
  
ElseIf Combo13.Text = "Purchase Date" And Combo14.Text <> blank Then
Set r = c.Execute("select invoice_no from purchase_invoice where invoice_date='" + Format(Combo14.Text, "dd/mmm/yyyy") + "'")
var = r.Fields(0)
  purchase_data.purchase_date_cmd var
  purchase_date_report.Show
  purchase_date_report.Refresh
  purchase_data.rspurchase_date_cmd.close

ElseIf Combo13.Text = "B/w Two Date" Then
  Set r = c.Execute("select invoice_no from purchase_invoice where invoice_date between '" + Format(DTPicker5.Value, "dd/mmm/yyyy") + "' AND '" + Format(DTPicker6.Value, "dd/mmm/yyyy") + "' ")
  dt = r.Fields(0)
  
  purchase_data.between_date_cmd dt, dt, dt
  p_between_date_report.Show
  p_between_date_report.Refresh
  purchase_data.rsbetween_date_cmd.close

End If
End Sub
Private Sub Combo13_Click()
If Combo13.Text = "All Purchase" Then
  Combo14.Visible = True
  Combo14.Enabled = False
  DTPicker5.Visible = False
  Label19.Visible = False
  DTPicker6.Visible = False
ElseIf Combo13.Text = "Invoice ID" Then
  Combo14.Visible = True
  Combo14.Enabled = True
    DTPicker5.Visible = False
  Label19.Visible = False
  DTPicker6.Visible = False
  Combo14.clear
  Set r = c.Execute("select invoice_no from purchase_invoice")
  Do While Not r.EOF
   Combo14.AddItem r.Fields(0)
   r.MoveNext
  Loop
  
ElseIf Combo13.Text = "Purchase Date" Then
  Combo14.Visible = True
  Combo14.Enabled = True
   DTPicker5.Visible = False
  Label19.Visible = False
  DTPicker6.Visible = False
  Combo14.clear
  Set r = c.Execute("select INVOICE_DATE from purchase_invoice")
  Do While Not r.EOF
   Combo14.AddItem r.Fields(0)
   r.MoveNext
  Loop
  
 ElseIf Combo13.Text = "B/w Two Date" Then
 Combo14.Visible = False
 DTPicker5.Visible = True
  Label19.Visible = True
  DTPicker6.Visible = True
  Combo14.clear
  
 End If
End Sub

Private Sub Combo15_Click()
If Combo15.Text = "Customer Order No" Then
Label15.Visible = True
Combo16.Visible = True
Label17.Visible = False
DTPicker1.Visible = False
DTPicker2.Visible = False
Image1.Visible = False

Combo16.clear
Set r = c.Execute("select order_number from customer_order_detail")
Do While Not r.EOF
Combo16.AddItem r.Fields(0)
r.MoveNext
Loop

ElseIf Combo15.Text = "Order Date" Then
Label15.Visible = True
Combo16.Visible = True

Label17.Visible = False
DTPicker1.Visible = False
DTPicker2.Visible = False
Image1.Visible = False

Combo16.clear
Set r = c.Execute("select order_date from customer_order_detail")
Do While Not r.EOF
Combo16.AddItem r.Fields(0)
r.MoveNext
Loop

ElseIf Combo15.Text = "B/W Two Date" Then
Label17.Visible = True
Label15.Visible = False
Combo16.Visible = False
DTPicker1.Visible = True
DTPicker2.Visible = True
Image1.Visible = False

ElseIf Combo15.Text = "Monthly" Then
Label15.Visible = True
Combo16.Visible = True
Label17.Visible = False
DTPicker1.Visible = False
DTPicker2.Visible = False
Image1.Visible = False

Combo16.clear
Combo16.AddItem "January"
Combo16.AddItem "February"
Combo16.AddItem "March"
Combo16.AddItem "April"
Combo16.AddItem "May"
Combo16.AddItem "June"
Combo16.AddItem "July"
Combo16.AddItem "August"
Combo16.AddItem "September"
Combo16.AddItem "October"
Combo16.AddItem "November"
Combo16.AddItem "December"

ElseIf Combo15.Text = "Yearly" Then
Label15.Visible = True
Combo16.Visible = True
Label17.Visible = False
DTPicker1.Visible = False
DTPicker2.Visible = False
Image1.Visible = True

Combo16.clear
Combo16.AddItem "2019"
Combo16.AddItem "2020"

ElseIf Combo15.Text = "Pending Order" Then
Label15.Visible = True
Combo16.Visible = True
Label17.Visible = False
DTPicker1.Visible = False
DTPicker2.Visible = False
Image1.Visible = False

Combo16.clear
Combo16.AddItem "All Pending Order"
Set r = c.Execute("select Order_number from customer_order_detail where status='no'")
Do While Not r.EOF
Combo16.AddItem r.Fields(0)
r.MoveNext
Loop

ElseIf Combo15.Text = "Delivered Order" Then
Label15.Visible = True
Combo16.Visible = True
Label17.Visible = False
DTPicker1.Visible = False
DTPicker2.Visible = False
Image1.Visible = False

Combo16.clear
Combo16.AddItem "All Delivered Order"
Set r = c.Execute("select Order_number from order_detail where inv_status='yes'")
Do While Not r.EOF
Combo16.AddItem r.Fields(0)
r.MoveNext
Loop

End If
End Sub





Private Sub Combo15_GotFocus()
'If Combo15.Text = "" Or Combo15.ListIndex = -1 Then
'MsgBox "Select Suppler Order Or Customer Order First"
'End If
End Sub

Private Sub Combo2_Click()
If Combo2.Text = "Supplier ID" Then
 Set r = c.Execute("select supplier_id from supplier_detail")
 Combo3.clear
 Do While Not r.EOF
  Combo3.AddItem r!supplier_id
  r.MoveNext
 Loop
ElseIf Combo2.Text = "Supplier Name" Then
 Set r = c.Execute("select supplier_name from supplier_detail")
 Combo3.clear
 Do While Not r.EOF
  Combo3.AddItem r!supplier_name
  r.MoveNext
 Loop
ElseIf Combo2.Text = "Supplier Brand" Then
 Set r = c.Execute("select brand from supplier_product_brand")
 Combo3.clear
 Do While Not r.EOF
  Combo3.AddItem r!brand
  r.MoveNext
 Loop
 End If
End Sub



Private Sub Combo3_Click()
If Combo6.Text = "Customer Name" Then
Set r = c.Execute("select supplier_id from supplier_detail where supplier_name='" + Combo3.Text + "'")
var = r.Fields(0)
End If
End Sub

Private Sub Combo6_Click()
If Combo6.Text = "All Customer" Then
Combo7.Enabled = False
'Else
'Combo7.Enabled = True
ElseIf Combo6.Text = "Customer ID" Then
Combo7.Enabled = True
Combo7.clear
Set r = c.Execute("select customer_id from customer_detail")
Do While Not r.EOF
Combo7.AddItem r!customer_id
r.MoveNext
Loop
ElseIf Combo6.Text = "Customer Name" Then
Combo7.Enabled = True
Combo7.clear
Set r = c.Execute("select customer_name from customer_detail")
Do While Not r.EOF
Combo7.AddItem r!customer_name
r.MoveNext
Loop
ElseIf Combo6.Text = "Address" Then
Combo7.Enabled = True
Combo7.clear
Set r = c.Execute("select address from customer_detail")
Do While Not r.EOF
Combo7.AddItem r!address
r.MoveNext
Loop
End If
End Sub

Private Sub Combo8_Click()
If Combo8.Text = "All Product" Then
Combo9.Enabled = False

ElseIf Combo8.Text = "Productr ID" Then
Combo9.Enabled = True
Set r = c.Execute("select product_id from product_detail")
Combo9.clear
Do While Not r.EOF
Combo9.AddItem r.Fields(0)
r.MoveNext
Loop

ElseIf Combo8.Text = "Product Name" Then
Combo9.Enabled = True
Set r = c.Execute("select product_name from product_detail")
Combo9.clear
Do While Not r.EOF
Combo9.AddItem r.Fields(0)
r.MoveNext
Loop

End If
End Sub

Private Sub Combo9_Click()
If Combo8.Text = "Product Name" Then
Set r = c.Execute("select product_id from product_detail where product_name='" + Combo9.Text + "'")
var = r.Fields(0)
End If
End Sub

Private Sub Command1_Click()
order_report_function

End Sub

Private Sub Command10_Click()
If Combo2.Text = "Supplier ID" And Combo3.Text <> "" Then
 DataEnvironment11.supp_id Combo3.Text, Combo3.Text, Combo3.Text, Combo3.Text

 Supplier_id_report.Show
 Supplier_id_report.Refresh
 DataEnvironment11.rssupp_id.close
 
 ElseIf Combo2.Text = "Supplier Name" And Combo3.Text <> "" Then
 Set r = c.Execute("select supplier_id from supplier_detail where supplier_name='" + Combo3.Text + "'")
    var = r.Fields(0)
 DataEnvironment11.supp_name var, var, var, var
 Supplier_name_report.Show
 Supplier_name_report.Refresh
DataEnvironment11.rssupp_name.close

  ElseIf Combo2.Text = "Supplier Brand" And Combo3.Text <> "" Then
   Set r = c.Execute("select sup_id from supplier_product_brand where brand='" + Combo3.Text + "'")
    var = r.Fields(0)
    
 DataEnvironment11.supp_brand var, var, var, var
 supplier_brand_report.Show
 supplier_brand_report.Refresh
    DataEnvironment11.rssupp_brand.close
    
ElseIf Option5.Value = True Then

supplier_report1.Show
 supplier_report1.Refresh
 
 ElseIf Option6.Value = True Then

 sup_account.Show
 sup_account.Refresh
 
  ElseIf Option7.Value = True Then

 sup_product.Show
 sup_product.Refresh


End If
End Sub

Private Sub Command12_Click()
If Combo6.Text = "All Customer" Then
 customer_all_report.Show
 customer_all_report.Refresh
 

ElseIf Combo6.Text = "Customer ID" And Combo7.Text <> blank Then
customer_data.customer_id_command Combo7.Text
 customer_id_report.Show
 customer_id_report.Refresh
 customer_data.rscustomer_id_command.close
 
ElseIf Combo6.Text = "Customer Name" And Combo7.Text <> blank Then
 customer_data.customer_name_command Combo7.Text
 customer_name_report.Show
  customer_name_report.Refresh
 customer_data.rscustomer_name_command.close
ElseIf Combo6.Text = "Address" And Combo7.Text <> blank Then
 customer_data.cust_address_command Combo7.Text
 cust_address_report.Show
 cust_address_report.Refresh
 customer_data.rscust_address_command.close
End If
End Sub

Private Sub Command13_Click()
If Combo8.Text = "All Product" Then
  all_product_report.Show
  all_product_report.Refresh
  
ElseIf Combo8.Text = "Productr ID" And Combo9.Text <> "" Then
    product_data.product_id Combo9.Text, Combo9.Text
    DataReport1.Show
    DataReport1.Refresh
    product_data.rsproduct_id.close
ElseIf Combo8.Text = "Product Name" And Combo9.Text <> "" Then
Set r = c.Execute("select distinct(product_id) from product_detail where product_name='" + Combo9.Text + "'")
var = r.Fields(0)
    product_data.product_name var, var
    product_name_report.Show
    product_name_report.Refresh
    product_data.rsproduct_name.close
End If

End Sub



Private Sub Command16_Click()
Dim dt As String
If Combo17.Text = "Supplier Order No" And Combo16.Text <> blank Then
  supplier_order_data.supplier_orderno_command Combo16.Text, Combo16.Text, Combo16.Text
  supplier_order_report.Show
  supplier_order_report.Refresh
  supplier_order_data.rssupplier_orderno_command.close

ElseIf Combo17.Text = "Order Date" And Combo16.Text <> blank Then
  Set r = c.Execute(" select order_number from order_detail where order_date='" + Format(Combo16.Text, "dd/mmm/yyyy") + "'")
  dt = r.Fields(0)
  supplier_order_data.order_date_command dt, dt, dt
  sup_ord_date_report.Show
  sup_ord_date_report.Refresh
  supplier_order_data.rsorder_date_command.close

ElseIf Combo17.Text = "B/w Two Date" Then
  Set r = c.Execute(" select order_number from order_detail where delivery_date between '" + Format(DTPicker1.Value, "dd/mmm/yyyy") + "' AND '" + Format(DTPicker2.Value, "dd/mmm/yyyy") + "' ")
  dt = r.Fields(0)
  supplier_order_data.Date dt, dt, dt
  sup_date_report.Show
  sup_date_report.Refresh
  supplier_order_data.rsdate.close
 

ElseIf Combo17.Text = "Pending Order" And Combo16.Text <> blank Then
  supplier_order_data.pen_ord_command Combo16.Text, Combo16.Text, Combo16.Text
  sup_ord_pending_report.Show
  sup_ord_pending_report.Refresh
  supplier_order_data.rspen_ord_command.close
  
ElseIf Combo17.Text = "Delivered Order" And Combo16.Text <> blank Then
  supplier_order_data.del_ord_command Combo16.Text, Combo16.Text, Combo16.Text
  DataReport3.Show
  DataReport3.Refresh
  supplier_order_data.rsdel_ord_command.close
  
  
ElseIf Combo15.Text = "ALL Customer Order" Then
    customer_order_all.Show
    customer_order_all.Refresh
    
ElseIf Combo15.Text = "Customer Order No" And Combo16.Text <> blank Then
customer_order_data.cust_ordno Combo16.Text, Combo16.Text
customer_order_no.Show
customer_order_no.Refresh
customer_order_data.rscust_ordno.close


ElseIf Combo15.Text = "Order Date" And Combo16.Text <> blank Then
Set r = c.Execute("select order_number from customer_order_detail where order_date='" + Format(Combo16.Text, "dd/mmm/yyyy") + "'")
 var = r.Fields(0)
    customer_order_data.cust_orddate var, var
    customer_order_date.Show
    customer_order_date.Refresh
    customer_order_data.rscust_orddate.close
    
ElseIf Combo15.Text = "B/W Two Date" And Combo16.Text <> blank Then
 Set r = c.Execute(" select order_number from order_detail where delivery_date between '" + Format(DTPicker1.Value, "dd/mmm/yyyy") + "' AND '" + Format(DTPicker2.Value, "dd/mmm/yyyy") + "' ")
 var = r.Fields(0)
    customer_order_data.between_date var, var
    customer_order_bdate.Show
    customer_order_bdate.Refresh
    customer_order_data.rsbetween_date.close

End If
End Sub

Private Sub Command2_Click()
purchase_report_function

End Sub

Private Sub Command3_Click()
sale_report_function
End Sub

Private Sub Command4_Click()
stock_report_function

End Sub

Private Sub Command5_Click()
product_report_function

End Sub

Private Sub Command6_Click()
customer_report_function

End Sub

Private Sub Command7_Click()
supplier_report_function

End Sub

Private Sub Command8_Click()
return_report_function

End Sub

Private Sub Command9_Click()
If Combo1.Text = "All Stock" Then
 stk_all_report.Show
 stk_all_report.Refresh
ElseIf Combo1.Text = "Stock No" And Combo10.Text <> blank Then
 stock_data.stock_no Combo10.Text
 stk_no_report.Show
 stk_no_report.Refresh
 stock_data.rsstock_no.close
ElseIf Combo1.Text = "Product Name" And Combo10.Text <> blank Then
 stock_data.p_name Combo10.Text
 stk_p_name_report.Show
  stk_p_name_report.Refresh
 stock_data.rsp_name.close
End If
End Sub

Private Sub Form_Load()

Connection

report.Caption = "report"
MDIForm1.Picture2.Visible = True

order_report_function
Combo3.clear
End Sub


Private Sub Option1_Click()
Option1.Value = True
Option2.Value = False

DTPicker3.Visible = False
DTPicker4.Visible = False
Label18.Visible = False
Combo18.clear
Combo18.AddItem "All Purchase Return"
Combo18.AddItem "Purchase Return No"
Combo18.AddItem "Return Date"
Combo18.AddItem "Return B/w Two Date"

End Sub

Private Sub Combo18_Click()
If Option1.Value = True And Combo18.Text = "All Purchase Return" Then
  
  Combo5.Enabled = True
  Combo5.clear
  Combo5.Enabled = False
ElseIf Option1.Value = True And Combo18.Text = "Purchase Return No" Then
    Combo5.Visible = True
  Combo5.Enabled = True
  Combo5.clear
  Set r = c.Execute("select return_no from purchase_return")
  Do While Not r.EOF
    Combo5.AddItem r.Fields(0)
  r.MoveNext
  Loop
ElseIf Option1.Value = True And Combo18.Text = "Return Date" Then
  Combo5.Visible = True
  Combo5.Enabled = True
  Combo5.clear
  Set r = c.Execute("select return_date from purchase_return")
  Do While Not r.EOF
    Combo5.AddItem r.Fields(0)
  r.MoveNext
  Loop
ElseIf Option1.Value = True And Combo18.Text = "Return B/w Two Date" Then

 DTPicker3.Visible = True
 DTPicker4.Visible = True
 Label18.Visible = True
 DTPicker4.Value = Date
 DTPicker3.Value = Date - 30
 
 
ElseIf Option2.Value = True And Combo18.Text = "All Sell Return" Then
  Combo5.Visible = True
  Combo5.Enabled = True
  
  Combo5.clear
  Combo5.Enabled = False
ElseIf Option2.Value = True And Combo18.Text = "Sale Return No" Then
  Combo5.Visible = True
  Combo5.Enabled = True
  
  Combo5.clear
  Set r = c.Execute("select return_no from sell_return")
  Do While Not r.EOF
    Combo5.AddItem r.Fields(0)
  r.MoveNext
  Loop
ElseIf Option2.Value = True And Combo18.Text = "Return date" Then
  Combo5.Visible = True
  Combo5.Enabled = True
  Combo5.clear
  Set r = c.Execute("select return_date from sell_return")
  Do While Not r.EOF
    Combo5.AddItem r.Fields(0)
  r.MoveNext
  Loop
ElseIf Option2.Value = True And Combo18.Text = "S Return B/w Two date" Then
 DTPicker3.Visible = True
 DTPicker4.Visible = True
 Label18.Visible = True
 Combo5.Visible = False
 Label6.Visible = False
 DTPicker4.Value = Date
 DTPicker3.Value = Date - 30

End If
End Sub
Private Sub Command11_Click()
If Combo18.Text = "All Purchase Return" And Combo5.Enabled = False Then
  p_r_all_report.Show
  p_r_all_report.Refresh
  
ElseIf Combo18.Text = "Purchase Return No" And Combo5.Text <> blank Then
  purchase_return_data.p_return_no Combo5.Text, Combo5.Text
  p_r_no_report.Show
  p_r_no_report.Refresh
  purchase_return_data.rsp_return_no.close
  
ElseIf Combo18.Text = "Return Date" And Combo5.Text <> blank Then
  Set r = c.Execute(" select distinct(return_no) from purchase_return where return_date='" + Format(Combo5.Text, "dd/mmm/yyyy") + "' ")
  noo = r.Fields(0)
  purchase_return_data.p_return_date noo, noo
  p_r_date_report.Show
  p_r_date_report.Refresh
  purchase_return_data.rsp_return_date.close
  
ElseIf Combo18.Text = "Return B/w Two Date" And Combo5.Text <> blank Then
  Set r = c.Execute(" select distinct(return_no) from purchase_return where return_date between '" + Format(DTPicker3.Value, "dd/mmm/yyyy") + "' AND '" + Format(DTPicker4.Value, "dd/mmm/yyyy") + "' ")
  noo = r.Fields(0)
  purchase_return_data.p_return_bdate noo, noo
  p_r_bdate_report.Show
  p_r_bdate_report.Refresh
  purchase_return_data.rsp_return_bdate.close
  
  
  
  
ElseIf Combo18.Text = "All Sell Return" And Combo5.Enabled = False Then
  s_all_report.Show
  s_all_report.Refresh
  
ElseIf Combo18.Text = "Sale Return No" And Combo5.Text <> blank Then
  s_return_data.s_no_return Combo5.Text, Combo5.Text
  s_no_report.Show
  s_no_report.Refresh
  s_return_data.rss_no_return.close
  
ElseIf Combo18.Text = "Return date" And Combo5.Text <> blank Then
  Set r = c.Execute(" select distinct(return_no) from sell_return where return_date='" + Format(Combo5.Text, "dd/mmm/yyyy") + "' ")
  no = r.Fields(0)
  s_return_data.s_date_return no, no
  s_date_report.Show
  s_date_report.Refresh
  s_return_data.rss_date_return.close
  
ElseIf Combo18.Text = "S Return B/w Two date" And Combo5.Text <> blank Then
  Set r = c.Execute(" select distinct(return_no) from sell_return where return_date between '" + Format(DTPicker3.Value, "dd/mmm/yyyy") + "' AND '" + Format(DTPicker4.Value, "dd/mmm/yyyy") + "' ")
  no = r.Fields(0)
  s_return_data.s_bdate_return no, no
  s_bdate_report.Show
  s_bdate_report.Refresh
  s_return_data.rss_bdate_return.close
End If
End Sub
Private Sub Option2_Click()
Option1.Value = False
Option2.Value = True
DTPicker3.Visible = False
DTPicker4.Visible = False
Label18.Visible = False

Combo18.clear
Combo18.AddItem "All Sell Return"
Combo18.AddItem "Sale Return No"
Combo18.AddItem "Return date"
Combo18.AddItem "S Return B/w Two date"


End Sub

Private Sub Option3_Click()
If Option3.Value = True Then
 Option4.Value = False
 Combo15.Visible = True
 Combo17.Visible = False
 Combo15.clear
 Combo15.AddItem "ALL Customer Order"
 Combo15.AddItem "Customer Order No"
 Combo15.AddItem "Order Date"
 Combo15.AddItem "B/W Two Date"

End If
End Sub

Private Sub Option4_Click()
If Option4.Value = True Then
 Option3.Value = False
 Combo15.Visible = False
 Combo17.Visible = True
 Combo17.clear
 Combo17.AddItem "Supplier Order No"
 Combo17.AddItem "Order Date"
 Combo17.AddItem "B/w Two Date"
End If
End Sub
Private Sub Combo17_gotfocus()
If Combo17.ListIndex = -1 Then
 MsgBox "Select Suppler Order Or Customer Order First"
End If
End Sub
Private Sub Combo17_Click()
If Combo17.Text = "Supplier Order No" Then
 Label15.Visible = True
 Combo16.Visible = True
 Label17.Visible = False
 DTPicker1.Visible = False
 DTPicker2.Visible = False
 Image1.Visible = False

 Combo16.clear
 Set r = c.Execute("select order_number from order_detail")
 Do While Not r.EOF
  Combo16.AddItem r.Fields(0)
  r.MoveNext
 Loop

ElseIf Combo17.Text = "Order Date" Then
 Label15.Visible = True
 Combo16.Visible = True

 Label17.Visible = False
 DTPicker1.Visible = False
 DTPicker2.Visible = False
 Image1.Visible = False

Combo16.clear
Set r = c.Execute("select order_date from order_detail")
Do While Not r.EOF
Combo16.AddItem r.Fields(0)
r.MoveNext
Loop

ElseIf Combo17.Text = "B/w Two Date" Then
Label17.Visible = True
Label15.Visible = False
Combo16.Visible = False
DTPicker1.Visible = True
DTPicker2.Visible = True
Image1.Visible = False

ElseIf Combo17.Text = "Monthly" Then
Label15.Visible = True
Combo16.Visible = True
Label17.Visible = False
DTPicker1.Visible = False
DTPicker2.Visible = False
Image1.Visible = False

Combo16.clear
Combo16.AddItem "January"
Combo16.AddItem "February"
Combo16.AddItem "March"
Combo16.AddItem "April"
Combo16.AddItem "May"
Combo16.AddItem "June"
Combo16.AddItem "July"
Combo16.AddItem "August"
Combo16.AddItem "September"
Combo16.AddItem "October"
Combo16.AddItem "November"
Combo16.AddItem "December"

ElseIf Combo17.Text = "Yearly" Then
Label15.Visible = True
Combo16.Visible = True
Label17.Visible = False
DTPicker1.Visible = False
DTPicker2.Visible = False
Image1.Visible = True

Combo16.clear
Combo16.AddItem "2019"
Combo16.AddItem "2020"

ElseIf Combo17.Text = "Pending Order" Then
Label15.Visible = True
Combo16.Visible = True
Label17.Visible = False
DTPicker1.Visible = False
DTPicker2.Visible = False
Image1.Visible = False

Combo16.clear
Combo16.AddItem "All Pending Order"
Set r = c.Execute("select Order_number from order_detail where inv_status='no'")
Do While Not r.EOF
Combo16.AddItem r.Fields(0)
r.MoveNext
Loop

ElseIf Combo17.Text = "Delivered Order" Then
Label15.Visible = True
Combo16.Visible = True
Label17.Visible = False
DTPicker1.Visible = False
DTPicker2.Visible = False
Image1.Visible = False

Combo16.clear
Combo16.AddItem "All Delivered Order"
Set r = c.Execute("select Order_number from order_detail where inv_status='yes'")
Do While Not r.EOF
Combo16.AddItem r.Fields(0)
r.MoveNext
Loop

End If
End Sub
