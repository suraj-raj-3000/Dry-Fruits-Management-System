VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form CUSTOMER_ORDER 
   BackColor       =   &H80000004&
   Caption         =   "customer order"
   ClientHeight    =   8850
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18150
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8850
   ScaleWidth      =   18150
   WindowState     =   2  'Maximized
   Begin VB.CommandButton update 
      BackColor       =   &H0080C0FF&
      Caption         =   " UPDATE"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   8880
      Width           =   1695
   End
   Begin VB.CommandButton insert 
      BackColor       =   &H0080FF80&
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   8880
      Width           =   1695
   End
   Begin VB.CommandButton search 
      BackColor       =   &H00FFFF80&
      Caption         =   "SEARCH"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11640
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   8880
      Width           =   1695
   End
   Begin VB.CommandButton new 
      BackColor       =   &H00C0C0C0&
      Caption         =   " NEW"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   8880
      Width           =   1695
   End
   Begin VB.CommandButton delete 
      BackColor       =   &H008080FF&
      Caption         =   " DELETE"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   8880
      Width           =   1695
   End
   Begin VB.CommandButton view 
      BackColor       =   &H00FF8080&
      Caption         =   " VIEW"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13920
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   8880
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000004&
      Height          =   6015
      Left            =   2400
      TabIndex        =   0
      Top             =   2520
      Width           =   13335
      Begin VB.ComboBox Combo2 
         BackColor       =   &H00FFFFFF&
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
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   69
         Top             =   960
         Width           =   2550
      End
      Begin VB.TextBox product_idfe 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2400
         TabIndex        =   68
         Text            =   " "
         Top             =   960
         Width           =   2535
      End
      Begin VB.ComboBox unit 
         BackColor       =   &H00FFFFFF&
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
         Left            =   6600
         Style           =   2  'Dropdown List
         TabIndex        =   49
         Top             =   360
         Width           =   2190
      End
      Begin VB.ListBox List10 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   2565
         ItemData        =   "CUSTOMER_ORDER.frx":0000
         Left            =   1080
         List            =   "CUSTOMER_ORDER.frx":0002
         TabIndex        =   47
         Top             =   2280
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H80000004&
         Caption         =   "Cash"
         Height          =   255
         Left            =   2400
         TabIndex        =   45
         Top             =   5520
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   11040
         MaxLength       =   6
         TabIndex        =   44
         Text            =   " "
         Top             =   5400
         Width           =   1695
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   8280
         MaxLength       =   5
         TabIndex        =   42
         Text            =   " "
         Top             =   5400
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   5400
         MaxLength       =   5
         TabIndex        =   40
         Text            =   " "
         Top             =   5400
         Width           =   1695
      End
      Begin VB.ListBox List8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   2565
         ItemData        =   "CUSTOMER_ORDER.frx":0004
         Left            =   10080
         List            =   "CUSTOMER_ORDER.frx":0006
         TabIndex        =   33
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   6600
         MaxLength       =   5
         TabIndex        =   31
         Top             =   1080
         Width           =   2175
      End
      Begin VB.ListBox List3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   2565
         ItemData        =   "CUSTOMER_ORDER.frx":0008
         Left            =   2880
         List            =   "CUSTOMER_ORDER.frx":000A
         TabIndex        =   11
         Top             =   2280
         Width           =   1455
      End
      Begin VB.ListBox List4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   2565
         ItemData        =   "CUSTOMER_ORDER.frx":000C
         Left            =   4320
         List            =   "CUSTOMER_ORDER.frx":000E
         TabIndex        =   10
         Top             =   2280
         Width           =   1575
      End
      Begin VB.ListBox List5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   2565
         ItemData        =   "CUSTOMER_ORDER.frx":0010
         Left            =   5880
         List            =   "CUSTOMER_ORDER.frx":0012
         TabIndex        =   9
         Top             =   2280
         Width           =   1455
      End
      Begin VB.ListBox List6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   2565
         ItemData        =   "CUSTOMER_ORDER.frx":0014
         Left            =   8760
         List            =   "CUSTOMER_ORDER.frx":0016
         TabIndex        =   8
         Top             =   2280
         Width           =   1335
      End
      Begin VB.ListBox List7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   2565
         ItemData        =   "CUSTOMER_ORDER.frx":0018
         Left            =   7320
         List            =   "CUSTOMER_ORDER.frx":001A
         TabIndex        =   7
         Top             =   2280
         Width           =   1455
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Height          =   2565
         ItemData        =   "CUSTOMER_ORDER.frx":001C
         Left            =   1080
         List            =   "CUSTOMER_ORDER.frx":001E
         TabIndex        =   6
         Top             =   2280
         Width           =   1815
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H008080FF&
         Caption         =   "DELETE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   11760
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3360
         Width           =   1100
      End
      Begin VB.CommandButton add 
         BackColor       =   &H0080FF80&
         Caption         =   "ADD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   11760
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2520
         Width           =   1100
      End
      Begin VB.TextBox quantity 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   10800
         MaxLength       =   5
         TabIndex        =   2
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox gst 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   10800
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   1
         Top             =   360
         Width           =   2055
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00FFFFFF&
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
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   360
         Width           =   2550
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   2565
         ItemData        =   "CUSTOMER_ORDER.frx":0020
         Left            =   360
         List            =   "CUSTOMER_ORDER.frx":0022
         TabIndex        =   12
         Top             =   2280
         Width           =   735
      End
      Begin VB.ListBox List9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   2565
         ItemData        =   "CUSTOMER_ORDER.frx":0024
         Left            =   360
         List            =   "CUSTOMER_ORDER.frx":0026
         TabIndex        =   34
         Top             =   2280
         Width           =   735
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H80000004&
         Caption         =   "Check"
         Height          =   255
         Left            =   2400
         TabIndex        =   46
         Top             =   5520
         Visible         =   0   'False
         Width           =   735
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
         Left            =   5880
         TabIndex        =   89
         Top             =   5520
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
         Left            =   10320
         TabIndex        =   88
         Top             =   1080
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
         Left            =   10200
         TabIndex        =   87
         Top             =   480
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
         Left            =   6120
         TabIndex        =   86
         Top             =   1080
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
         Left            =   6120
         TabIndex        =   85
         Top             =   480
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
         Left            =   1440
         TabIndex        =   84
         Top             =   1080
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
         Left            =   2040
         TabIndex        =   83
         Top             =   480
         Width           =   120
      End
      Begin VB.Shape Shape11 
         Height          =   375
         Left            =   2280
         Top             =   5460
         Width           =   975
      End
      Begin VB.Label Label28 
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
         Left            =   10200
         TabIndex        =   43
         Top             =   5475
         Width           =   750
      End
      Begin VB.Label Label27 
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
         Left            =   7440
         TabIndex        =   41
         Top             =   5475
         Width           =   750
      End
      Begin VB.Label Label26 
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
         Left            =   3600
         TabIndex        =   39
         Top             =   5475
         Width           =   1395
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
         TabIndex        =   38
         Top             =   5520
         Width           =   2010
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " = "
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
         Left            =   10200
         TabIndex        =   37
         Top             =   4800
         Width           =   255
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pay  Total :-"
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
         Left            =   480
         TabIndex        =   36
         Top             =   4875
         Width           =   1065
      End
      Begin VB.Label pay_total 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   10560
         TabIndex        =   35
         Top             =   4845
         Width           =   555
      End
      Begin VB.Line Line8 
         X1              =   10080
         X2              =   10080
         Y1              =   4680
         Y2              =   5040
      End
      Begin VB.Shape Shape2 
         Height          =   300
         Left            =   360
         Top             =   4800
         Width           =   11055
      End
      Begin VB.Line Line7 
         X1              =   10080
         X2              =   10080
         Y1              =   2040
         Y2              =   2400
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Tot Amount"
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
         TabIndex        =   22
         Top             =   2065
         Width           =   1050
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H80000007&
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
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   5520
         TabIndex        =   32
         Top             =   1080
         Width           =   510
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Unit Price"
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
         Left            =   6120
         TabIndex        =   25
         Top             =   2065
         Width           =   915
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Quantity / unit"
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
         Left            =   7320
         TabIndex        =   24
         Top             =   2070
         Width           =   1305
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Gst"
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
         Left            =   9360
         TabIndex        =   23
         Top             =   2065
         Width           =   360
      End
      Begin VB.Label Label11 
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
         Left            =   4800
         TabIndex        =   21
         Top             =   2065
         Width           =   420
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Brand"
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
         Left            =   3240
         TabIndex        =   20
         Top             =   2065
         Width           =   570
      End
      Begin VB.Label Label8 
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
         Left            =   1320
         TabIndex        =   19
         Top             =   2065
         Width           =   1275
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "S no"
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
         Left            =   480
         TabIndex        =   18
         Top             =   2065
         Width           =   405
      End
      Begin VB.Line Line6 
         X1              =   8760
         X2              =   8760
         Y1              =   2040
         Y2              =   2400
      End
      Begin VB.Line Line5 
         X1              =   7320
         X2              =   7320
         Y1              =   2040
         Y2              =   2400
      End
      Begin VB.Line Line4 
         X1              =   5880
         X2              =   5880
         Y1              =   2040
         Y2              =   2400
      End
      Begin VB.Line Line3 
         X1              =   4320
         X2              =   4320
         Y1              =   2040
         Y2              =   2400
      End
      Begin VB.Line Line2 
         X1              =   2880
         X2              =   2880
         Y1              =   2040
         Y2              =   2400
      End
      Begin VB.Line Line1 
         X1              =   1080
         X2              =   1080
         Y1              =   2040
         Y2              =   2400
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H80000007&
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
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   600
         TabIndex        =   17
         Top             =   1080
         Width           =   690
      End
      Begin VB.Label Label19 
         BackColor       =   &H80000007&
         BackStyle       =   0  'Transparent
         Caption         =   " Product Name"
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
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H80000007&
         BackStyle       =   0  'Transparent
         Caption         =   " Unit"
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
         Left            =   5520
         TabIndex        =   15
         Top             =   480
         Width           =   480
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H80000007&
         BackStyle       =   0  'Transparent
         Caption         =   " Quantity"
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
         Left            =   9360
         TabIndex        =   14
         Top             =   1080
         Width           =   915
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H80000007&
         BackStyle       =   0  'Transparent
         Caption         =   "Gst"
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
         Left            =   9720
         TabIndex        =   13
         Top             =   480
         Width           =   360
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000004&
         FillColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   360
         Top             =   2040
         Width           =   11055
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2295
      Left            =   2400
      TabIndex        =   53
      Top             =   360
      Width           =   13335
      Begin VB.TextBox customer_name 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   10080
         TabIndex        =   66
         Top             =   960
         Width           =   2535
      End
      Begin VB.TextBox address 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   10080
         TabIndex        =   65
         Top             =   1560
         Width           =   2535
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   360
         Left            =   3000
         TabIndex        =   57
         Top             =   1560
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   -2147483635
         Format          =   126025729
         CurrentDate     =   43556
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   360
         Left            =   3000
         TabIndex        =   58
         Top             =   960
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   -2147483635
         Format          =   126025729
         CurrentDate     =   43556
      End
      Begin VB.ComboBox Combo3 
         BackColor       =   &H00C0C0FF&
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
         Left            =   10080
         Style           =   2  'Dropdown List
         TabIndex        =   64
         Top             =   360
         Width           =   2550
      End
      Begin VB.TextBox order_no 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   3000
         TabIndex        =   60
         Text            =   " "
         Top             =   360
         Width           =   2415
      End
      Begin VB.ComboBox Combo4 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Left            =   3000
         Style           =   2  'Dropdown List
         TabIndex        =   59
         Top             =   360
         Visible         =   0   'False
         Width           =   2430
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
         Left            =   9000
         TabIndex        =   82
         Top             =   1800
         Width           =   120
      End
      Begin VB.Label Label38 
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
         Left            =   9480
         TabIndex        =   81
         Top             =   1200
         Width           =   120
      End
      Begin VB.Label Label37 
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
         Left            =   9240
         TabIndex        =   80
         Top             =   600
         Width           =   120
      End
      Begin VB.Label Label36 
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
         Left            =   2400
         TabIndex        =   79
         Top             =   1680
         Width           =   120
      End
      Begin VB.Label Label35 
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
         Left            =   2160
         TabIndex        =   78
         Top             =   960
         Width           =   120
      End
      Begin VB.Label Label34 
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
         Left            =   2040
         TabIndex        =   77
         Top             =   360
         Width           =   120
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00400000&
         BorderStyle     =   2  'Dash
         X1              =   6600
         X2              =   6600
         Y1              =   120
         Y2              =   2280
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000007&
         BackStyle       =   0  'Transparent
         Caption         =   "Customer ID"
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
         Left            =   7920
         TabIndex        =   63
         Top             =   600
         Width           =   1275
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H80000001&
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
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
         Left            =   8040
         TabIndex        =   62
         Top             =   1800
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000001&
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name"
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
         TabIndex        =   61
         Top             =   1200
         Width           =   1665
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackColor       =   &H80000007&
         BackStyle       =   0  'Transparent
         Caption         =   "Order No"
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
         Left            =   960
         TabIndex        =   56
         Top             =   360
         Width           =   960
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackColor       =   &H80000007&
         BackStyle       =   0  'Transparent
         Caption         =   " Order Date"
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
         Left            =   840
         TabIndex        =   55
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   " Delivery Date"
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
         Left            =   840
         TabIndex        =   54
         Top             =   1680
         Width           =   1500
      End
   End
   Begin VB.Frame Frame1 
      Height          =   7215
      Left            =   2400
      TabIndex        =   50
      Top             =   720
      Width           =   13335
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "CLOSE"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5880
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   6720
         Width           =   1095
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2655
         Left            =   480
         TabIndex        =   51
         Top             =   600
         Width           =   12375
         _ExtentX        =   21828
         _ExtentY        =   4683
         View            =   3
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ORDER NO"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ORDER DATE"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "DELIVERY DATE"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "PRODUCT ID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "PAYMENT MODE"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "STATUS"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "CUSTOMER NAME"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "PHONE"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "ADDRESS"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   2655
         Left            =   480
         TabIndex        =   52
         Top             =   3960
         Width           =   12375
         _ExtentX        =   21828
         _ExtentY        =   4683
         View            =   3
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S.N"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "PRODUCT NAME"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "BRAND"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "UNIT"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "GST"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "RATE"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "QTY"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "PAID AMOUNT"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "TOTAL"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "!!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   300
         Left            =   9120
         TabIndex        =   76
         Top             =   3600
         Width           =   165
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "!!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   300
         Left            =   4560
         TabIndex        =   75
         Top             =   3600
         Width           =   165
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Customer Ordered Product Information"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   285
         Left            =   4785
         TabIndex        =   74
         Top             =   3600
         Width           =   4320
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "!!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   300
         Left            =   7800
         TabIndex        =   73
         Top             =   240
         Width           =   165
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "!!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   300
         Left            =   4440
         TabIndex        =   72
         Top             =   240
         Width           =   165
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Customer Order Information"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   285
         Left            =   4650
         TabIndex        =   71
         Top             =   240
         Width           =   3150
      End
      Begin VB.Line Line10 
         BorderStyle     =   3  'Dot
         X1              =   120
         X2              =   13150
         Y1              =   3360
         Y2              =   3360
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "CUSTOMER  ORDER"
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
      TabIndex        =   70
      Top             =   0
      Width           =   9045
   End
End
Attribute VB_Name = "CUSTOMER_ORDER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim sn As Integer
Dim no As Integer
Dim opt As String
Public ind As Integer
Public status As String
Dim qtt As Integer

Private Sub add_Click()
If Combo1.Text = "Select Product ID" Or Combo1.Text = "" Then
 Combo1.BackColor = &HC0C0FF
 MsgBox "Select Product Name", vbCritical
ElseIf Combo2.Text = "Select Product Brand" Or Combo2.Text = "" Then
 Combo2.BackColor = &HC0C0FF
 MsgBox "Select Brand", vbCritical
ElseIf unit.Text = "" Then
 unit.BackColor = &HC0C0FF
 MsgBox "Select Unit", vbCritical
ElseIf Text1.Text = "" Then
 Text1.BackColor = &HC0C0FF
 MsgBox "Rate Fields is Blank", vbCritical
ElseIf gst.Text = "" Then
 gst.BackColor = &HC0C0FF
 MsgBox "Gst Fields is Blank", vbCritical
ElseIf quantity.Text = "" Then
 quantity.BackColor = &HC0C0FF
 MsgBox "Quantity Fields is Blank", vbCritical
Else

List1.BackColor = vbWhite
 List3.BackColor = vbWhite
 List4.BackColor = vbWhite
 List5.BackColor = vbWhite
 List6.BackColor = vbWhite
 List7.BackColor = vbWhite
 List8.BackColor = vbWhite
 List10.BackColor = vbWhite

List1.AddItem i + 1
List3.AddItem Combo2.Text
List4.AddItem unit.Text
List5.AddItem Text1.Text
List7.AddItem quantity.Text
List6.AddItem gst.Text
List10.AddItem Combo1.Text
List8.AddItem (Val(List5.List(i)) * Val(List7.List(i))) + (Val(List5.List(i)) * Val(List7.List(i))) * Val(List6.List(i)) / 100
i = i + 1

If List9.List(0) = "" Then
sql = "select count(s_no) from customer_ordered_product"
Set r = c.Execute(sql)
no = r.Fields(0) + 1
List9.AddItem no
Else
List9.AddItem no + 1
no = no + 1
End If

total

Text3.Text = Val(pay_total.Caption - Text2.Text)
Text4.Text = pay_total.Caption

gst.Text = ""
quantity.Text = ""
Text1.Text = ""

End If
'If Combo1.Text <> blank Or Combo2.Text <> blank Or unit.Text <> blank And Text1.Text <> blank And quantity.Text <> blank And gst.Text <> blank Then
'List10.RemoveItem (ind)
'List3.RemoveItem (ind)
'List4.RemoveItem (ind)
'List5.RemoveItem (ind)
'List6.RemoveItem (ind)
'List7.RemoveItem (ind)
'List8.RemoveItem (ind)
'List1.RemoveItem (ind)
'
'List1.AddItem i + 1
'List3.AddItem Combo2.Text
'List4.AddItem unit.Text
'List5.AddItem Text1.Text
'List7.AddItem quantity.Text
'List6.AddItem gst.Text
'List10.AddItem Combo1.Text
'List8.AddItem (Val(List5.List(i)) * Val(List7.List(i))) + (Val(List5.List(i)) * Val(List7.List(i))) * Val(List6.List(i)) / 100
'i = i + 1
'
'If List9.List(0) = "" Then
'sql = "select count(s_no) from customer_ordered_product"
'Set r = c.Execute(sql)
'no = r.Fields(0) + 1
'List9.AddItem no
'Else
'List9.AddItem no + 1
'no = no + 1
'End If
'total
'
'Text3.Text = Val(pay_total.Caption - Text2.Text)
'Text4.Text = pay_total.Caption




End Sub



Private Sub Combo1_Click()
Combo1.BackColor = vbWhite
If Combo3.Text = "Select Customer ID" Then
MsgBox "Select Customer Id"
Combo1.Text = ""
Combo3.SetFocus
Else

Set r = c.Execute("select product_id from product_detail where product_name='" + Combo1.Text + "'")
product_idfe.Text = r.Fields(0)

Combo2.clear
sql = "select distinct(brand) from ordered_product where p_id='" + product_idfe.Text + "'"
Set r = c.Execute(sql)
Do While Not r.EOF
Combo2.AddItem r!brand
r.MoveNext
Loop
End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
Combo1.BackColor = vbWhite
If Combo3.Text = "Select Customer ID" Or Combo3.Text <> blank Then
MsgBox "Select Customer Id"
Combo3.SetFocus
End If
End Sub



Private Sub Combo2_Click()
Combo2.BackColor = vbWhite
sql = "select unit from ordered_product where brand='" + Combo2.Text + "' and p_id='" + product_idfe.Text + "' "
Set r = c.Execute(sql)
unit.clear
Do While Not r.EOF
unit.AddItem r.Fields(0)
r.MoveNext
Loop

End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
Combo2.BackColor = vbWhite
End Sub

Private Sub Combo3_Change()
If Combo3.Text = "" Then
customer_name.Enabled = True
address.Enabled = True
End If
End Sub



Private Sub Combo3_KeyPress(KeyAscii As Integer)
Combo3.BackColor = vbWhite
End Sub

Private Sub Combo4_click()
sql = "select * from customer_order_detail where order_number='" + Combo4.Text + "'"
Set r = c.Execute(sql)

DTPicker2.Value = r.Fields(1)
DTPicker1.Value = r.Fields(2)
Combo3.Text = r.Fields(11)

If r.Fields(5) = "cash" Then
Option1.Value = True
Else
Option2.Value = True
End If



If Text2.Text = "" Then
adpay = "0"
Else
adpay = r.Fields(5)

End If
Text2.Text = adpay

sql = "select c_id,customer_name,address from customer_order_detail where c_id='" + Combo3.Text + "' "
Set r = c.Execute(sql)

customer_name.Text = r.Fields(1)
address.Text = r.Fields(2)

Combo3.Enabled = False



sql = "select * from customer_ordered_product where ord_no='" + Combo4.Text + "'"
Set r = c.Execute(sql)
List1.clear
List3.clear
List4.clear
List5.clear
List6.clear
List7.clear
List8.clear
List10.clear
Do While Not r.EOF
List1.AddItem r.Fields(0)
List3.AddItem r.Fields(2)
List4.AddItem r.Fields(3)
List5.AddItem r.Fields(4)
List6.AddItem r.Fields(6)
List7.AddItem r.Fields(5)
List8.AddItem r.Fields(7)
List10.AddItem r.Fields(1)
r.MoveNext
total
Loop
If Text2.Text = "" Then
adpay = "0"
End If
Text3.Text = Val(pay_total.Caption - adpay)
Text4.Text = pay_total.Caption
End Sub

Private Sub Command1_Click()
Frame1.Visible = False
Frame3.Visible = True
Frame2.Visible = True
End Sub

Private Sub Command3_Click()
If List1.ListIndex = -1 Then
Else
List1.RemoveItem ind
List3.RemoveItem ind
List4.RemoveItem ind
List5.RemoveItem ind
List6.RemoveItem ind
List7.RemoveItem ind
List8.RemoveItem ind
List9.RemoveItem ind
List10.RemoveItem ind
total
End If
End Sub

Private Sub delete_Click()
ans = MsgBox("do you want to delete", vbYesNo + vbInformation)
If ans = 1 Then
c.Execute ("delete  customer_ordered_product where ord_no='" + Combo4.Text + "'")
c.Execute ("delete  customer_order_detail where order_number='" + Combo4.Text + "'")
MsgBox "Order deleted"
End If
End Sub

Private Sub Form_Load()

Connection
autogenerate
MDIForm1.Picture2.Visible = True
Set r = New ADODB.Recordset
sql = "select customer_id from customer_detail "
Set r = c.Execute(sql)
Do While Not r.EOF
Combo3.AddItem r!customer_id
r.MoveNext
Loop
product_id


If Option1.Value = True Then
Option2.Value = False
opt = "cash"
Else
opt = "check"
End If
CUSTOMER_ORDER.Caption = "Customer Order"
pay_total.Caption = "0"
Text2.Text = "0"
End Sub

Private Sub Combo3_Click()
Combo3.BackColor = vbWhite
sql = "select * from customer_detail where customer_id='" + Combo3.Text + "'"
Set r = c.Execute(sql)

customer_name.Text = r.Fields(1)
customer_name.Enabled = False
address.Text = r.Fields(3)
address.Enabled = False


End Sub

Public Function product_id()
Set r = New ADODB.Recordset
sql = "select distinct(product_name) from ordered_product"
Set r = c.Execute(sql)
Combo1.clear
While r.EOF = False
Combo1.AddItem r.Fields(0)
r.MoveNext
Wend
End Function
Public Function additem_combo()
Set r = New ADODB.Recordset
sql = "select * from product_detail where p_id='" + Combo1.Text + "'"
Set r = c.Execute(sql)
product_idfe.Text = r.Fields(3)
'unit.Text = r.Fields(1)
unit.Text = r.Fields(5)
gst.Text = r.Fields(8)
Text1.Text = r.Fields(6)


End Function



Private Sub gst_KeyPress(KeyAscii As Integer)
gst.BackColor = vbWhite
 If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn Then

Else
KeyAscii = 0
MsgBox "Enter only number", vbCritical
End If
End Sub

Private Sub insert_Click()
Dim j As Integer
ans = MsgBox("Do you Want to Save", vbOKCancel + vbInformation, "Save")
If ans = 1 Then

If Combo3.Text = "" Then
 Combo3.BackColor = &HC0C0FF
 MsgBox "Select Customer ID", vbCritical
ElseIf List1.List(0) = "" Or List3.List(0) = "" Or List4.List(0) = "" Or List5.List(0) = "" Or List6.List(0) = "" Or List7.List(0) = "" Or List8.List(0) = "" Or List10.List(0) = "" Then
 List1.BackColor = &HC0C0FF
 List3.BackColor = &HC0C0FF
 List4.BackColor = &HC0C0FF
 List5.BackColor = &HC0C0FF
 List6.BackColor = &HC0C0FF
 List7.BackColor = &HC0C0FF
 List8.BackColor = &HC0C0FF
 List10.BackColor = &HC0C0FF
 MsgBox "Add product details in list box", vbCritical
ElseIf Text2.Text = 0 Or Text2.Text = "" Then
 Text2.BackColor = &HC0C0FF
 MsgBox "Advance Fields is Empty", vbCritical
ElseIf Text3.Text = 0 Or Text3.Text = "" Then
 Text3.BackColor = &HC0C0FF
Else
sql = "insert into customer_order_detail values('" + order_no.Text + "','" + Format(DTPicker2.Value, "dd/mmm/yyyy") + "','" + Format(DTPicker1.Value, "dd/mmm/yyyy") + "','" + product_idfe.Text + "','" + opt + "'," + Text2.Text + "," + Text3.Text + "," + Text4.Text + ",'" + customer_name.Text + "','" + address.Text + "','" + "no" + "','" + Combo3.Text + "')"
Set r = c.Execute(sql)
For j = 0 To List1.ListCount - 1
sql = "insert into customer_ordered_product values('" + List9.List(j) + "','" + List10.List(j) + "','" + List3.List(j) + "','" + List4.List(j) + "'," + List5.List(j) + "," + List7.List(j) + "," + List6.List(j) + "," + List8.List(j) + ",'" + product_idfe.Text + "','" + order_no.Text + "')"
Set r = c.Execute(sql)
Next
MsgBox "Order Placed Successfully"
blank_fields
autogenerate
End If
End If

End Sub






Private Sub List10_Click()
ind = List10.ListIndex
Combo1.Text = List10.List(ind)
Combo2.Text = List3.List(ind)
unit.Text = List4.List(ind)
Text1.Text = List5.List(ind)
quantity.Text = List7.List(ind)
gst.Text = List6.List(ind)
End Sub

Private Sub List3_Click()
ind = List3.ListIndex
Combo1.Text = List10.List(ind)
Combo2.Text = List3.List(ind)
unit.Text = List4.List(ind)
Text1.Text = List5.List(ind)
quantity.Text = List7.List(ind)
gst.Text = List6.List(ind)
End Sub

Private Sub List4_Click()
ind = List4.ListIndex
Combo1.Text = List10.List(ind)
Combo2.Text = List3.List(ind)
unit.Text = List4.List(ind)
Text1.Text = List5.List(ind)
quantity.Text = List7.List(ind)
gst.Text = List6.List(ind)
End Sub

Private Sub List5_Click()
ind = List5.ListIndex
Combo1.Text = List10.List(ind)
Combo2.Text = List3.List(ind)
unit.Text = List4.List(ind)
Text1.Text = List5.List(ind)
quantity.Text = List7.List(ind)
gst.Text = List6.List(ind)
End Sub

Private Sub List6_Click()
ind = List6.ListIndex
Combo1.Text = List10.List(ind)
Combo2.Text = List3.List(ind)
unit.Text = List4.List(ind)
Text1.Text = List5.List(ind)
quantity.Text = List7.List(ind)
gst.Text = List6.List(ind)
End Sub

Private Sub List7_Click()
ind = List7.ListIndex
Combo1.Text = List10.List(ind)
Combo2.Text = List3.List(ind)
unit.Text = List4.List(ind)
Text1.Text = List5.List(ind)
quantity.Text = List7.List(ind)
gst.Text = List6.List(ind)
End Sub

Private Sub quantity_Change()
If quantity.Text = "" Then
ElseIf quantity.Text > qtt Then
MsgBox "Invaild Quantity . max Quantity is  =" & qtt, vbCritical
quantity.Text = ""
End If
End Sub

Private Sub quantity_KeyPress(KeyAscii As Integer)
quantity.BackColor = vbWhite
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn Then
  If KeyAscii = 13 Then
   quantity.SetFocus
  End If
Else
KeyAscii = 0
MsgBox "Enter only number"
End If
End Sub

Private Sub search_Click()
update.Enabled = True
delete.Enabled = True
insert.Enabled = False

Combo4.Enabled = True
order_no.Visible = False
Combo4.Visible = True
add_order_id

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Text1.BackColor = vbWhite
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn Then
  If KeyAscii = 13 Then
   Text1.SetFocus
  End If
Else
KeyAscii = 0
MsgBox "Enter only number"
End If
End Sub


Private Sub new_Click()
autogenerate
blank_fields

update.Enabled = True
delete.Enabled = True
insert.Enabled = False
Combo4.Visible = False
End Sub


Public Function autogenerate()
Dim a As String
Set r = New ADODB.Recordset
sql = "select max(to_number(substr(order_number,5,length(order_number)))) from customer_order_detail"
Set r = c.Execute(sql)
If IsNull(r.Fields(0)) Then
order_no.Text = "ord" & "0" & 1
Else
order_no.Text = "ord" & "0" & r.Fields(0) + 1
End If
a = order_no.Text
If (a = "ord" & "01" & "0") Then
sql = "select max(to_number(substr(order_number,4,length(order_number)))) from customer_order_detail"
Set r = c.Execute(sql)
order_no.Text = "ord" & r.Fields(0) + 1
End If

End Function



Public Function total()
For j = 0 To List8.ListCount - 1
tot = tot + Val(List8.List(j))
pay_total.Caption = tot
Next
End Function



Public Function blank_fields()
order_no.Visible = True
DTPicker1.Value = Date
DTPicker2.Value = Date

Text1.Text = ""
Text2.Text = "0"
Text3.Text = "0"
Text4.Text = "0"
List1.clear
List2.clear
List3.clear
List4.clear
List5.clear
List6.clear
List7.clear
List8.clear
List9.clear
List10.clear
pay_total.Caption = "0"
customer_name.Text = ""
address.Text = ""
gst.Text = ""
quantity.Text = ""
product_idfe.Text = ""
End Function


Public Function add_order_id()
sql = "select order_number from customer_order_detail"
Set r = c.Execute(sql)
Combo4.clear
Do While Not r.EOF
Combo4.AddItem r!order_number
r.MoveNext
Loop
End Function

Private Sub Text2_change()
If Text2.Text = "" Then
 adpay = "0"
Else
 adpay = Text2.Text
End If

If adpay <= Val(pay_total.Caption) Then

Text3.Text = Val(pay_total.Caption - adpay)
Text4.Text = pay_total.Caption

Else
MsgBox "Maximum Advance payment  =" & pay_total.Caption, vbCritical
Text2.Text = ""
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
Text2.BackColor = vbWhite
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn Then

Else
KeyAscii = 0
MsgBox "Enter only number", vbCritical
End If
End Sub



Private Sub Text3_KeyPress(KeyAscii As Integer)
Text3.BackColor = vbWhite
End Sub

Private Sub unit_Click()
unit.BackColor = vbWhite
Set r = c.Execute("select igst,rate,quantity from ordered_product where p_id='" + product_idfe.Text + "' and unit='" + unit.Text + "' AND brand='" + Combo2.Text + "'")
gst.Text = r.Fields(2)
qtt = r.Fields(0)
If IsNull(r.Fields(1)) Then
Text1.Text = 0
Else
cal = r.Fields(1) + ((r.Fields(1) * 20) / 100)
Text1.Text = cal
End If
End Sub

Private Sub unit_KeyPress(KeyAscii As Integer)
unit.BackColor = vbWhite
End Sub

Private Sub update_Click()
ans = MsgBox("do you want to update", vbYesNo + vbInformation)
If ans = vbYes Then

 sql = " UPDATE customer_order_detail set order_date='" + Format(DTPicker2.Value, "dd/mmm/yyyy") + "', delivery_date='" + Format(DTPicker1.Value, "dd/mmm/yyyy") + "' , ADVANCE_PAYMENT=" + Text2.Text + ",DUES=" + Text3.Text + ",total=" + Text4.Text + ", CUSTOMER_NAME='" + customer_name.Text + "',address='" + address.Text + "',c_id='" + Combo3.Text + "' where order_number='" + Combo4.Text + "' "
 MsgBox sql
 Set r = c.Execute(sql)
 
For j = 0 To List10.ListCount - 1
 Set r = c.Execute("update customer_ordered_product set s_no='" + List9.List(j) + "',PRODUCT_NAME='" + List10.List(j) + "',BRAND='" + List3.List(j) + "',unit='" + List4.List(j) + "',UNIT_PRICE=" + List5.List(j) + ",QUANTITY=" + List7.List(j) + ", GST=" + List6.List(j) + ", TOT_AMOUNT=" + List8.List(j) + ",P_ID='" + product_idfe.Text + "' where ORD_NO='" + Combo4.Text + "' ")
 Next
 MsgBox "Order Updated"
End If

End Sub

Private Sub view_Click()
Frame1.Visible = True
Frame3.Visible = False
Frame2.Visible = False

End Sub
