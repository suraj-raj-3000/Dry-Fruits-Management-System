VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form supplier_order 
   BackColor       =   &H80000004&
   Caption         =   "purchase"
   ClientHeight    =   8655
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11280
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8655
   ScaleWidth      =   11280
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   1815
      Left            =   2400
      TabIndex        =   43
      Top             =   480
      Width           =   13215
      Begin VB.TextBox delivery_date 
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
         Left            =   2760
         TabIndex        =   44
         Text            =   " "
         Top             =   1320
         Width           =   2175
      End
      Begin VB.TextBox supplier_id 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   360
         Left            =   9840
         TabIndex        =   57
         Top             =   360
         Width           =   2535
      End
      Begin VB.TextBox company_name 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   360
         Left            =   9840
         TabIndex        =   56
         Top             =   840
         Width           =   2535
      End
      Begin VB.TextBox gstin_num 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   360
         Left            =   9840
         TabIndex        =   55
         Top             =   1320
         Width           =   2535
      End
      Begin VB.ComboBox Combo4 
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
         ItemData        =   "supplier_order.frx":0000
         Left            =   4800
         List            =   "supplier_order.frx":001C
         Style           =   2  'Dropdown List
         TabIndex        =   50
         Top             =   1320
         Width           =   390
      End
      Begin VB.TextBox order_date 
         Enabled         =   0   'False
         Height          =   375
         Left            =   2760
         TabIndex        =   45
         Top             =   840
         Width           =   2415
      End
      Begin VB.TextBox order_no 
         Height          =   375
         Left            =   2760
         TabIndex        =   51
         Text            =   " "
         Top             =   360
         Width           =   2415
      End
      Begin VB.ComboBox Combo3 
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
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   360
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Label Label29 
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
         Left            =   8640
         TabIndex        =   71
         Top             =   1440
         Width           =   120
      End
      Begin VB.Label Label28 
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
         TabIndex        =   70
         Top             =   960
         Width           =   120
      End
      Begin VB.Label Label27 
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
         Left            =   8760
         TabIndex        =   69
         Top             =   480
         Width           =   120
      End
      Begin VB.Label Label24 
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
         TabIndex        =   68
         Top             =   1440
         Width           =   120
      End
      Begin VB.Label Label23 
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
         TabIndex        =   67
         Top             =   960
         Width           =   120
      End
      Begin VB.Label Label31 
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
         Left            =   1920
         TabIndex        =   66
         Top             =   480
         Width           =   120
      End
      Begin VB.Line Line10 
         BorderStyle     =   4  'Dash-Dot
         BorderWidth     =   2
         X1              =   6240
         X2              =   6240
         Y1              =   135
         Y2              =   1700
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000007&
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier ID"
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
         TabIndex        =   54
         Top             =   480
         Width           =   1170
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H80000001&
         BackStyle       =   0  'Transparent
         Caption         =   " Gstin No"
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
         Left            =   7560
         TabIndex        =   53
         Top             =   1440
         Width           =   960
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000001&
         BackStyle       =   0  'Transparent
         Caption         =   " Company Name"
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
         Left            =   7200
         TabIndex        =   52
         Top             =   960
         Width           =   1725
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
         Left            =   840
         TabIndex        =   49
         Top             =   480
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
         Left            =   720
         TabIndex        =   48
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H80000007&
         BackStyle       =   0  'Transparent
         Caption         =   "Delivery Date"
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
         TabIndex        =   47
         Top             =   1440
         Width           =   1440
      End
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
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   8520
      Width           =   1695
   End
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
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   8520
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
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   8520
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
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   8520
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
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   8520
      Width           =   1695
   End
   Begin VB.CommandButton report 
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
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   8520
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000004&
      Height          =   6135
      Left            =   2400
      TabIndex        =   0
      Top             =   2160
      Width           =   13215
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   10080
         MaxLength       =   6
         TabIndex        =   35
         Top             =   5400
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   6360
         MaxLength       =   5
         TabIndex        =   33
         Top             =   5400
         Width           =   1695
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2160
         MaxLength       =   5
         TabIndex        =   31
         Top             =   5400
         Width           =   1815
      End
      Begin VB.TextBox unit_price 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   10440
         MaxLength       =   5
         TabIndex        =   22
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox gst 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   6360
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   14
         Top             =   1080
         Width           =   2415
      End
      Begin VB.TextBox quantity 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   10440
         MaxLength       =   5
         TabIndex        =   13
         Top             =   360
         Width           =   1575
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
         Left            =   11880
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2160
         Width           =   975
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
         Left            =   11880
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   3120
         Width           =   975
      End
      Begin VB.ListBox List9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   2760
         ItemData        =   "supplier_order.frx":0038
         Left            =   1080
         List            =   "supplier_order.frx":003A
         TabIndex        =   9
         Top             =   2040
         Width           =   1455
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   2760
         ItemData        =   "supplier_order.frx":003C
         Left            =   2520
         List            =   "supplier_order.frx":003E
         TabIndex        =   8
         Top             =   2040
         Width           =   1455
      End
      Begin VB.ListBox List8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   2760
         ItemData        =   "supplier_order.frx":0040
         Left            =   10080
         List            =   "supplier_order.frx":0042
         TabIndex        =   7
         Top             =   2040
         Width           =   1455
      End
      Begin VB.ListBox List7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   2760
         ItemData        =   "supplier_order.frx":0044
         Left            =   8280
         List            =   "supplier_order.frx":0046
         TabIndex        =   6
         Top             =   2040
         Width           =   975
      End
      Begin VB.ListBox List6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   2760
         ItemData        =   "supplier_order.frx":0048
         Left            =   9240
         List            =   "supplier_order.frx":004A
         TabIndex        =   5
         Top             =   2040
         Width           =   855
      End
      Begin VB.ListBox List5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   2760
         ItemData        =   "supplier_order.frx":004C
         Left            =   6960
         List            =   "supplier_order.frx":004E
         TabIndex        =   4
         Top             =   2040
         Width           =   1335
      End
      Begin VB.ListBox List4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   2760
         ItemData        =   "supplier_order.frx":0050
         Left            =   5400
         List            =   "supplier_order.frx":0052
         TabIndex        =   3
         Top             =   2040
         Width           =   1575
      End
      Begin VB.ListBox List3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   2760
         ItemData        =   "supplier_order.frx":0054
         Left            =   3960
         List            =   "supplier_order.frx":0056
         TabIndex        =   2
         Top             =   2040
         Width           =   1455
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   2760
         ItemData        =   "supplier_order.frx":0058
         Left            =   360
         List            =   "supplier_order.frx":005A
         TabIndex        =   1
         Top             =   2040
         Width           =   735
      End
      Begin VB.ComboBox brand 
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
         Left            =   6360
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   360
         Width           =   2430
      End
      Begin VB.ComboBox Combo5 
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
         TabIndex        =   39
         Top             =   1080
         Width           =   2535
      End
      Begin VB.TextBox prod_id 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   " "
         Top             =   1080
         Width           =   2535
      End
      Begin VB.ComboBox product_name 
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
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   360
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
         Left            =   9720
         TabIndex        =   80
         Top             =   5400
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
         Left            =   6120
         TabIndex        =   79
         Top             =   5400
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
         Left            =   1800
         TabIndex        =   78
         Top             =   5400
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
         Left            =   10080
         TabIndex        =   77
         Top             =   1080
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
         Left            =   10080
         TabIndex        =   76
         Top             =   480
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
         Left            =   5880
         TabIndex        =   75
         Top             =   1080
         Width           =   120
      End
      Begin VB.Label Label33 
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
         Left            =   6000
         TabIndex        =   74
         Top             =   480
         Width           =   120
      End
      Begin VB.Label Label32 
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
         Left            =   1560
         TabIndex        =   73
         Top             =   1080
         Width           =   120
      End
      Begin VB.Label Label30 
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
         TabIndex        =   72
         Top             =   480
         Width           =   120
      End
      Begin VB.Label total 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "total"
         Height          =   255
         Left            =   10080
         TabIndex        =   38
         Top             =   4800
         Width           =   1455
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL  :-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4800
         TabIndex        =   37
         Top             =   4825
         Width           =   1575
      End
      Begin VB.Line Line9 
         X1              =   10080
         X2              =   10080
         Y1              =   4680
         Y2              =   5040
      End
      Begin VB.Shape Shape2 
         Height          =   300
         Left            =   360
         Top             =   4770
         Width           =   11175
      End
      Begin VB.Line Line8 
         X1              =   3960
         X2              =   3960
         Y1              =   1680
         Y2              =   2040
      End
      Begin VB.Line Line7 
         X1              =   5400
         X2              =   5400
         Y1              =   1680
         Y2              =   2040
      End
      Begin VB.Line Line6 
         X1              =   6960
         X2              =   6960
         Y1              =   1680
         Y2              =   2040
      End
      Begin VB.Line Line5 
         X1              =   8280
         X2              =   8280
         Y1              =   1680
         Y2              =   2040
      End
      Begin VB.Line Line4 
         X1              =   9240
         X2              =   9240
         Y1              =   1680
         Y2              =   2040
      End
      Begin VB.Line Line3 
         X1              =   10080
         X2              =   10080
         Y1              =   1680
         Y2              =   2040
      End
      Begin VB.Line Line2 
         X1              =   2520
         X2              =   2520
         Y1              =   1680
         Y2              =   2040
      End
      Begin VB.Line Line1 
         X1              =   1080
         X2              =   1080
         Y1              =   1680
         Y2              =   2040
      End
      Begin VB.Shape Shape1 
         Height          =   375
         Left            =   360
         Top             =   1680
         Width           =   11175
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H80000007&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Amount"
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
         Left            =   8280
         TabIndex        =   36
         Top             =   5400
         Width           =   1380
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H80000007&
         BackStyle       =   0  'Transparent
         Caption         =   "Balance Amount"
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
         Left            =   4320
         TabIndex        =   34
         Top             =   5400
         Width           =   1695
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H80000007&
         BackStyle       =   0  'Transparent
         Caption         =   "Advance"
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
         Left            =   720
         TabIndex        =   32
         Top             =   5400
         Width           =   930
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H80000007&
         BackStyle       =   0  'Transparent
         Caption         =   " Rate"
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
         TabIndex        =   24
         Top             =   1080
         Width           =   570
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H80000007&
         BackStyle       =   0  'Transparent
         Caption         =   " Gst"
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
         Left            =   5280
         TabIndex        =   20
         Top             =   1080
         Width           =   420
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H80000007&
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
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
         Left            =   9120
         TabIndex        =   19
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H80000007&
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
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   960
         TabIndex        =   18
         Top             =   1080
         Width           =   420
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H80000007&
         BackStyle       =   0  'Transparent
         Caption         =   "Product Name"
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
         Left            =   480
         TabIndex        =   17
         Top             =   480
         Width           =   1485
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
         Left            =   5160
         TabIndex        =   16
         Top             =   480
         Width           =   690
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         Caption         =   $"supplier_order.frx":005C
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   1745
         Width           =   11055
      End
   End
   Begin VB.Frame Frame1 
      Height          =   7815
      Left            =   2040
      TabIndex        =   40
      Top             =   480
      Visible         =   0   'False
      Width           =   13935
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "CLOSE"
         Height          =   375
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   7200
         Width           =   1335
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2895
         Left            =   480
         TabIndex        =   41
         Top             =   4200
         Width           =   12975
         _ExtentX        =   22886
         _ExtentY        =   5106
         View            =   3
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         MousePointer    =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   12
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S.N"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "PRODUCT ID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "PRODUCT NAME"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "BRAND"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "UNIT"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "GST"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "QTY"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "RATE"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "AMOUNT"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "PAID AMOUNT"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "BAL. AMOUNT"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "TOTAL"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   2655
         Left            =   600
         TabIndex        =   58
         Top             =   720
         Width           =   12855
         _ExtentX        =   22675
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
         MousePointer    =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   7
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
            Text            =   "ORDER STATUS"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "SUPPLIER NAME"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "PHONE"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "ADDRESS"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Line Line11 
         BorderStyle     =   3  'Dot
         X1              =   120
         X2              =   13800
         Y1              =   3600
         Y2              =   3600
      End
      Begin VB.Label Label21 
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
         Left            =   8880
         TabIndex        =   65
         Top             =   3840
         Width           =   165
      End
      Begin VB.Label Label18 
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
         Left            =   4680
         TabIndex        =   64
         Top             =   3840
         Width           =   165
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   " Supplier Order Product Information"
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
         Left            =   4815
         TabIndex        =   63
         Top             =   3840
         Width           =   4020
      End
      Begin VB.Label Label26 
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
         Left            =   8760
         TabIndex        =   62
         Top             =   360
         Width           =   165
      End
      Begin VB.Label Label14 
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
         Left            =   5280
         TabIndex        =   61
         Top             =   360
         Width           =   165
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   " Supplier Order Information"
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
         Left            =   5520
         TabIndex        =   60
         Top             =   360
         Width           =   3090
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "SUPPLIER  ORDER"
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
      TabIndex        =   59
      Top             =   120
      Width           =   9045
   End
End
Attribute VB_Name = "supplier_order"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim item As ListItem
Public ind As Integer

Private Sub add_Click()
Dim a As String
Dim j As Integer

If Combo5.Text = "" And product_name.Text = "" Or brand.Text = "" Or gst.Text = "" Or quantity.Text = "" Or unit_price.Text = "" Then
 MsgBox "Product Detail Fields are blank", vbCritical
Else

j = List1.ListCount
Set r = New ADODB.Recordset
sql = "select max(to_number(substr(s_no,5,length(s_no)))) from ordersno"
Set r = c.Execute(sql)
If IsNull(r.Fields(0)) Then
List1.AddItem "sn" & "00" & 1
Else
List1.AddItem "sn" & "00" & r.Fields(0) + 1
End If
a = List1.List(j)
If (a = "sn" & "001" & "0") Then
sql = "select max(to_number(substr(s_no,4,length(s_no)))) from ordersno"
Set r = c.Execute(sql)
List1.AddItem "sn" & "0" & r.Fields(0) + 1
End If
sql = "insert into ordersno values('" + List1.List(j) + "')"
Set r = c.Execute(sql)

List9.AddItem prod_id.Text
List2.AddItem product_name.Text
List3.AddItem brand.Text
List4.AddItem Combo5.Text
List5.AddItem unit_price.Text
List6.AddItem gst.Text
List7.AddItem quantity.Text
List8.AddItem ((Val(unit_price.Text) * Val(quantity.Text)) - ((Val(gst.Text / 100)) * ((Val(unit_price.Text) * Val(quantity.Text)))))

For i = 0 To List8.ListCount - 1
tot = Val(tot) + Val(List8.List(i))
Next
total.Caption = tot
End If
End Sub

Private Sub brand_Click()
Dim i As Integer
Dim a As String
Set r = New ADODB.Recordset
sql = "select  *from supplier_product_brand where brand='" + brand.Text + "'"
Set r = c.Execute(sql)
a = r.Fields(3)

sql = "select *from supplier_product where s_no='" + a + "' "
Set r = c.Execute(sql)
supplier_id.Text = r.Fields(1)
sql = "select * from supplier_detail where supplier_id='" + supplier_id.Text + "'"
Set r = c.Execute(sql)
company_name.Text = r.Fields(3)
gstin_num.Text = r.Fields(4)

'sql = "select product_name from supplier_product where s_id='" + supplier_id.Text + "'"
'Set r = c.Execute(sql)
'Do While Not r.EOF
'product_name.AddItem r!product_name
'r.MoveNext
'Loop
End Sub

Private Sub Combo3_Click()
List1.clear
List2.clear
List3.clear
List4.clear
List5.clear
List6.clear
List7.clear
List8.clear
List9.clear
sql = "select *from order_detail"
Set r = c.Execute(sql)
order_date.Text = r.Fields(1)
supplier_id.Text = r.Fields(2)
delivery_date.Text = r.Fields(3)
Text4.Text = r.Fields(5)
Text2.Text = r.Fields(6)
Text1.Text = r.Fields(7)
sql = "select company_name,gstin_no from supplier_detail where supplier_id='" + supplier_id.Text + "'"
Set r = c.Execute(sql)
company_name.Text = r.Fields(0)
gstin_num.Text = r.Fields(1)
sql = "select *from ordered_product where order_no='" + Combo3.Text + "'"
Set r = c.Execute(sql)
While r.EOF = False
List1.AddItem r.Fields(0)
List9.AddItem r.Fields(7)
List2.AddItem r.Fields(1)
List3.AddItem r.Fields(2)
List4.AddItem r.Fields(3)
List5.AddItem r.Fields(9)
List7.AddItem r.Fields(4)
List6.AddItem r.Fields(5)
List8.AddItem r.Fields(10)
r.MoveNext
Wend

End Sub

Private Sub Combo4_click()
delivery_date.BackColor = vbWhite
delivery_date.Text = Date + Combo4.Text

End Sub



Private Sub Combo5_Click()
sql = "select distinct(brand) from supplier_product_brand where unit='" + Combo5.Text + "'"
Set r = c.Execute(sql)
brand.clear
Do While Not r.EOF
brand.AddItem r!brand
r.MoveNext
Loop
End Sub

Private Sub Command1_Click()
Frame1.Visible = False
Frame2.Visible = True
Frame3.Visible = True
End Sub

Private Sub Command3_Click()
If List2.List(ind) = "" Then
Else
 List1.RemoveItem ind
 List2.RemoveItem ind
 List3.RemoveItem ind
 List4.RemoveItem ind
 List5.RemoveItem ind
 List6.RemoveItem ind
 List7.RemoveItem ind
 List8.RemoveItem ind
 List9.RemoveItem ind
End If
End Sub





Private Sub delete_Click()
ans = MsgBox("Do you Want to Delete", vbOKCancel + vbInformation)
If ans = 1 Then

Set r = c.Execute("delete ordered_product where order_no='" + Combo3.Text + "'")

Set r = c.Execute("delete ordered_product_amount where order_number='" + Combo3.Text + "'")

Set r = c.Execute("delete order_detail where order_number='" + Combo3.Text + "'")

MsgBox "order_deleted"
End If
clear
End Sub



Private Sub delivery_date_KeyPress(KeyAscii As Integer)
delivery_date.BackColor = vbWhite
End Sub

Private Sub Form_Load()
Connection
'order_view
supplier_order.Caption = "Supplier Order"

autogenerate
Set r = New ADODB.Recordset
sql = "select distinct(product_name) from supplier_product"
product_name.clear
Set r = c.Execute(sql)
Do While Not r.EOF
product_name.AddItem r!product_name
r.MoveNext
Loop
order_date.Text = Date
Text1.Text = 0
MDIForm1.Picture2.Visible = True
'Command1.Visible = False

End Sub


Public Function autogenerate()
Dim a As String
Set r = New ADODB.Recordset
sql = "select max(to_number(substr(order_number,5,length(order_number)))) from order_detail"
Set r = c.Execute(sql)
If IsNull(r.Fields(0)) Then
order_no.Text = "or" & "00" & 1
Else
order_no.Text = "or" & "00" & r.Fields(0) + 1
End If
a = order_no.Text
If (a = "or" & "001" & "0") Then
sql = "select max(to_number(substr(order_number,4,length(order_number)))) from order_detail"
Set r = c.Execute(sql)
order_no.Text = "or" & "0" & r.Fields(0) + 1
End If

End Function



Private Sub gst_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn Then

Else
KeyAscii = 0
MsgBox "Must be a Number", vbCritical
End If
End Sub

Private Sub insert_Click()
Dim i As Integer
Dim sp As Integer
ans = MsgBox("Do you Want to Save", vbOKCancel + vbInformation)
If ans = 1 Then

If delivery_date.Text = "" Then
 delivery_date.BackColor = &HC0C0FF
 MsgBox "Select Delivey Day", vbCritical
ElseIf List1.List(0) = "" Or List2.List(0) = "" Or List3.List(0) = "" Or List4.List(0) = "" Or List5.List(0) = "" Or List6.List(0) = "" Or List7.List(0) = "" Or List8.List(0) = "" Or List9.List(0) = "" Then
 List1.BackColor = &HC0C0FF
 List2.BackColor = &HC0C0FF
 List3.BackColor = &HC0C0FF
 List4.BackColor = &HC0C0FF
 List5.BackColor = &HC0C0FF
 List6.BackColor = &HC0C0FF
 List7.BackColor = &HC0C0FF
 List8.BackColor = &HC0C0FF
 List9.BackColor = &HC0C0FF
 MsgBox "Add Product Details in listbox", vbCritical
ElseIf Text4.Text = "" Then
 Text4.BackColor = &HC0C0FF
 MsgBox "Advance Fields is empty", vbCritical
ElseIf Text2.Text = "" Then
 Text2.BackColor = &HC0C0FF
 MsgBox "Balance Amount Fields is empty", vbCritical
ElseIf Text1.Text = "" Then
 Text1.BackColor = &HC0C0FF
 MsgBox "Total amount Fields is empty", vbCritical

Else
Set r = New ADODB.Recordset
sql = "insert into order_detail values('" + order_no.Text + "','" + Format(order_date, "dd/mmm/yyyy") + "','" + supplier_id.Text + "','" + Format(delivery_date, "dd/mmm/yyyy") + "', 'no'," + Text4.Text + "," + Text2.Text + "," + Text1.Text + " )"

Set r = c.Execute(sql)
For i = 0 To List1.ListCount - 1
sql = "insert into ordered_product values('" + List1.List(i) + "','" + List2.List(i) + "','" + List3.List(i) + "','" + List4.List(i) + "'," + List6.List(i) + "," + List7.List(i) + ",'" + supplier_id.Text + "','" + List9.List(i) + "','" + order_no.Text + "'," + List5.List(i) + "," + List8.List(i) + ")"

Set r = c.Execute(sql)
Next i

sql = "insert into ordered_product_amount values('" + order_no.Text + "','" + Text1.Text + "','" + Text4.Text + "','" + Text2.Text + "')"
Set r = c.Execute(sql)
MsgBox "Supplier Order Placed"
End If

For f = 0 To List4.ListCount
sp = Val(List5.List(f)) + (Val(List5.List(f)) * 20 / 100)
sql = "update product_brand set cost_price='" & List5.List(f) & "',selling_price= " & sp & " where product_id='" & prod_id.Text & "' and brand='" + List3.List(f) + "' and unit='" + List4.List(f) + "'"
Set r = c.Execute(sql)
Next
End If
End Sub

Private Sub List1_Click()
Dim i As Integer
ind = List1.ListIndex
i = List1.ListIndex
Command1.Visible = True
product_name.Text = List2.List(i)
prod_id.Text = List9.List(i)
brand.Text = List3.List(i)
quantity.Text = List7.List(i)
gst.Text = List6.List(i)
unit_price.Text = List5.List(i)
Combo5.AddItem List4.List(i)
End Sub

Private Sub List2_Click()
ind = List2.ListIndex
End Sub

Private Sub List3_Click()
ind = List3.ListIndex
End Sub

Private Sub List4_Click()
ind = List4.ListIndex
End Sub

Private Sub List5_Click()
ind = List5.ListIndex
End Sub

Private Sub List6_Click()
ind = List6.ListIndex
End Sub

Private Sub List7_Click()
ind = List7.ListIndex
End Sub

Private Sub List8_Click()
ind = List8.ListIndex
End Sub

Private Sub List9_Click()
ind = List9.ListIndex
End Sub

Private Sub new_Click()
clear
autogenerate
Set r = New ADODB.Recordset
sql = "select distinct(product_name) from supplier_product"
Set r = c.Execute(sql)
product_name.clear
Do While Not r.EOF
product_name.AddItem r!product_name
r.MoveNext
Loop
order_date.Text = Date
order_no.Visible = True
Combo3.Visible = False
insert.Enabled = True
update.Enabled = False
delete.Enabled = False
End Sub

Public Function clear()
order_no.Text = ""
order_date.Text = ""
supplier_id.Text = ""
delivery_date.Text = ""
company_name.Text = ""
gstin_num.Text = ""
gst.Text = ""
Text2.Text = ""
prod_id.Text = ""
total.Caption = total
order_date.Text = ""
Text4.Text = ""
Text2.Text = ""
Text1.Text = ""
unit_price.Text = ""
unit_price.Text = ""
quantity.Text = ""
Combo5.clear
List1.clear
List2.clear
List3.clear
List4.clear
List5.clear
List6.clear
List7.clear
List8.clear
List9.clear
brand.clear
End Function


Private Sub product_name_Click()
Dim i As Integer
Dim a As String
Set r = New ADODB.Recordset
sql = "select s_no from supplier_product where product_name='" + product_name.Text + "'"
Set r = c.Execute(sql)
a = r.Fields(0)

sql = "select * from supplier_product where product_name='" + product_name.Text + "'"
Set r = c.Execute(sql)
gst.Text = r.Fields(4)
prod_id.Text = r.Fields(2)

Set r = c.Execute("select unit from product_brand where product_id='" + prod_id.Text + "'")
Combo5.clear
brand.clear
Do While Not r.EOF
Combo5.AddItem r.Fields(0)
r.MoveNext
Loop
End Sub

Private Sub quantity_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn Then

Else
KeyAscii = 0
MsgBox "Enter only number", vbCritical
End If
End Sub

Private Sub report_Click()
order_no.Visible = False
Combo3.Visible = True
insert.Enabled = False

update.Enabled = True
delete.Enabled = True
insert.Enabled = False
Combo3.clear

Set r = c.Execute("select order_number from order_detail")
Do While Not r.EOF
Combo3.AddItem r!order_number
r.MoveNext
Loop


End Sub

Private Sub Text2_change()
Text2.Text = Val(Text1.Text) - Val(Text4.Text)
End Sub

Private Sub Text2_click()
Text2.Text = Val(Text1.Text) - Val(Text4.Text)
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
 If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn Then

Else
KeyAscii = 0
MsgBox "Enter only number", vbCritical
End If
End Sub

Private Sub Text4_lostfocus()
If total.Caption = "total" Or total.Caption = "" Then
 tt = 0
Else
tt = total.Caption
End If
If Text4.Text = "" Then
 t4 = 0
Else
 t4 = Text4.Text
End If
Text2.Text = tt - t2
Text1.Text = tt
End Sub



Private Sub unit_price_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn Then
Else
 KeyAscii = 0
 MsgBox "Enter only number", vbCritical
End If
End Sub

Private Sub update_Click()
sql = "update order_detail set delivery_date='" + Format(delivery_date.Text, "dd/mmm/yyyy") + "' "
Set r = c.Execute(sql)
sql = "upda"
sql = "UPDATE ordered_product set product_name='" + product_name.Text + "',brand='" + brand.Text + "',quantity='" + quantity.Text + "',igst='" + gst.Text + "',rate='" + unit_price.Text + "',unit='" + Combo5.Text + "' where order_no='" + Combo3.Text + "'"

Set r = c.Execute(sql)
MsgBox "table ordered_product updated"
clear
End Sub

Private Sub view_Click()
Frame1.Visible = True
Frame2.Visible = False
Frame3.Visible = False
order_view
End Sub

Public Function order_view()

Set r = c.Execute("select * from ORDER_DETAIL,supplier_detail")
ListView2.ListItems.clear
While Not r.EOF

Set item = ListView2.ListItems.add(, , r.Fields(0))
item.SubItems(1) = r.Fields(1)
item.SubItems(2) = r.Fields(3)
item.SubItems(3) = r.Fields(4)
item.SubItems(4) = r.Fields(15)
item.SubItems(5) = r.Fields(9)
item.SubItems(6) = r.Fields(14)
r.MoveNext
Wend

Set r = c.Execute("select * from ordered_product,ordered_product_amount ")
ListView1.ListItems.clear
While Not r.EOF

Set item = ListView1.ListItems.add(, , r.Fields(0))
item.SubItems(1) = r.Fields(7)
item.SubItems(2) = r.Fields(1)
item.SubItems(3) = r.Fields(2)

item.SubItems(4) = r.Fields(3)
item.SubItems(5) = r.Fields(4)
item.SubItems(6) = r.Fields(5)
item.SubItems(7) = r.Fields(9)
item.SubItems(8) = r.Fields(10)
item.SubItems(9) = r.Fields(13)
item.SubItems(10) = r.Fields(14)
item.SubItems(11) = r.Fields(12)
r.MoveNext
Wend
End Function
