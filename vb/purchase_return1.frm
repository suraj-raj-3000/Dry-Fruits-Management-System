VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form purchase_return 
   BackColor       =   &H80000004&
   Caption         =   "purchase_prd"
   ClientHeight    =   9375
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17685
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9375
   ScaleWidth      =   17685
   WindowState     =   2  'Maximized
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
      TabIndex        =   6
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
      TabIndex        =   5
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
      TabIndex        =   4
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
      TabIndex        =   3
      Top             =   8880
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
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   2
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
      TabIndex        =   1
      Top             =   8880
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000004&
      Height          =   5655
      Left            =   2400
      TabIndex        =   30
      Top             =   3000
      Width           =   13455
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   2955
         ItemData        =   "purchase_return1.frx":0000
         Left            =   480
         List            =   "purchase_return1.frx":0002
         TabIndex        =   43
         Top             =   2160
         Width           =   735
      End
      Begin VB.ListBox List11 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   2955
         ItemData        =   "purchase_return1.frx":0004
         Left            =   1200
         List            =   "purchase_return1.frx":0006
         TabIndex        =   41
         Top             =   2160
         Width           =   1215
      End
      Begin VB.ListBox List10 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   2955
         ItemData        =   "purchase_return1.frx":0008
         Left            =   9720
         List            =   "purchase_return1.frx":000A
         TabIndex        =   40
         Top             =   2160
         Width           =   1215
      End
      Begin VB.ListBox List9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   2955
         ItemData        =   "purchase_return1.frx":000C
         Left            =   8640
         List            =   "purchase_return1.frx":000E
         TabIndex        =   39
         Top             =   2160
         Width           =   1095
      End
      Begin VB.ListBox List3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   2955
         ItemData        =   "purchase_return1.frx":0010
         Left            =   4200
         List            =   "purchase_return1.frx":0012
         TabIndex        =   38
         Top             =   2160
         Width           =   1455
      End
      Begin VB.ListBox List4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   2955
         ItemData        =   "purchase_return1.frx":0014
         Left            =   5640
         List            =   "purchase_return1.frx":0016
         TabIndex        =   37
         Top             =   2160
         Width           =   1335
      End
      Begin VB.ListBox List5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   2955
         ItemData        =   "purchase_return1.frx":0018
         Left            =   6960
         List            =   "purchase_return1.frx":001A
         TabIndex        =   36
         Top             =   2160
         Width           =   735
      End
      Begin VB.ListBox List6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   2955
         ItemData        =   "purchase_return1.frx":001C
         Left            =   7680
         List            =   "purchase_return1.frx":001E
         TabIndex        =   35
         Top             =   2160
         Width           =   975
      End
      Begin VB.ListBox List8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   2955
         ItemData        =   "purchase_return1.frx":0020
         Left            =   10920
         List            =   "purchase_return1.frx":0022
         TabIndex        =   34
         Top             =   2160
         Width           =   1095
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   2955
         ItemData        =   "purchase_return1.frx":0024
         Left            =   2400
         List            =   "purchase_return1.frx":0026
         TabIndex        =   33
         Top             =   2160
         Width           =   1815
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H008080FF&
         Caption         =   "REMOVE"
         Height          =   375
         Left            =   12240
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   3360
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0080FF80&
         Caption         =   "ADD"
         Height          =   375
         Left            =   12240
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   2520
         Width           =   975
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Height          =   1575
         Left            =   120
         TabIndex        =   48
         Top             =   120
         Width           =   13215
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
            Left            =   5880
            Style           =   2  'Dropdown List
            TabIndex        =   54
            Top             =   240
            Width           =   2175
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
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   53
            Top             =   840
            Width           =   2175
         End
         Begin VB.TextBox unit_price 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   5880
            Locked          =   -1  'True
            TabIndex        =   52
            Text            =   " "
            Top             =   840
            Width           =   2175
         End
         Begin VB.ComboBox description 
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
            ItemData        =   "purchase_return1.frx":0028
            Left            =   10560
            List            =   "purchase_return1.frx":0035
            TabIndex        =   51
            Top             =   840
            Width           =   1935
         End
         Begin VB.TextBox quantity 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   10560
            MaxLength       =   5
            TabIndex        =   50
            Text            =   " "
            Top             =   240
            Width           =   1935
         End
         Begin VB.ComboBox product_id 
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
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   49
            Top             =   240
            Width           =   2175
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   375
            Left            =   10560
            TabIndex        =   55
            Top             =   1200
            Visible         =   0   'False
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   661
            _Version        =   393216
            Format          =   120389633
            CurrentDate     =   43582
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
            Left            =   10080
            TabIndex        =   85
            Top             =   840
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
            Left            =   9960
            TabIndex        =   84
            Top             =   360
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
            Left            =   5160
            TabIndex        =   83
            Top             =   840
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
            Left            =   5160
            TabIndex        =   82
            Top             =   360
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
            Left            =   1200
            TabIndex        =   81
            Top             =   960
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
            Left            =   1680
            TabIndex        =   80
            Top             =   360
            Width           =   120
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
            Left            =   4560
            TabIndex        =   62
            Top             =   360
            Width           =   420
         End
         Begin VB.Label Label18 
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
            Height          =   375
            Left            =   120
            TabIndex        =   61
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackColor       =   &H80000007&
            BackStyle       =   0  'Transparent
            Caption         =   "Reasion "
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
            Left            =   9000
            TabIndex        =   60
            Top             =   840
            Width           =   945
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H80000007&
            BackStyle       =   0  'Transparent
            Caption         =   "Exp Date"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   9000
            TabIndex        =   59
            Top             =   1320
            Visible         =   0   'False
            Width           =   795
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
            Left            =   9000
            TabIndex        =   58
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label12 
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
            Left            =   4560
            TabIndex        =   57
            Top             =   840
            Width           =   510
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000007&
            BackStyle       =   0  'Transparent
            Caption         =   "Brand"
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
            Left            =   360
            TabIndex        =   56
            Top             =   960
            Width           =   630
         End
      End
      Begin VB.TextBox product_name 
         Height          =   285
         Left            =   2280
         TabIndex        =   42
         Top             =   360
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   $"purchase_return1.frx":006A
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
         Left            =   600
         TabIndex        =   46
         Top             =   1860
         Width           =   11415
      End
      Begin VB.Shape Shape1 
         Height          =   375
         Left            =   480
         Top             =   1800
         Width           =   11535
      End
      Begin VB.Line Line1 
         X1              =   1200
         X2              =   1200
         Y1              =   1800
         Y2              =   2160
      End
      Begin VB.Line Line3 
         X1              =   10920
         X2              =   10920
         Y1              =   1800
         Y2              =   2160
      End
      Begin VB.Line Line4 
         X1              =   9720
         X2              =   9720
         Y1              =   1800
         Y2              =   2160
      End
      Begin VB.Line Line5 
         X1              =   8640
         X2              =   8640
         Y1              =   1800
         Y2              =   2160
      End
      Begin VB.Line Line6 
         X1              =   7680
         X2              =   7680
         Y1              =   1800
         Y2              =   2160
      End
      Begin VB.Line Line7 
         X1              =   6960
         X2              =   6960
         Y1              =   1800
         Y2              =   2160
      End
      Begin VB.Line Line8 
         X1              =   4200
         X2              =   4200
         Y1              =   1800
         Y2              =   2160
      End
      Begin VB.Line Line9 
         X1              =   5640
         X2              =   5640
         Y1              =   1800
         Y2              =   2160
      End
      Begin VB.Line Line10 
         X1              =   2400
         X2              =   2400
         Y1              =   1800
         Y2              =   2160
      End
      Begin VB.Shape Shape2 
         Height          =   375
         Left            =   480
         Top             =   5040
         Width           =   11535
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Left            =   11280
         TabIndex        =   45
         Top             =   5160
         Width           =   105
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total  ="
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
         Left            =   10080
         TabIndex        =   44
         Top             =   5160
         Width           =   660
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000004&
      Height          =   2175
      Left            =   2400
      TabIndex        =   7
      Top             =   840
      Width           =   13455
      Begin VB.ComboBox invoice_no 
         BackColor       =   &H00FFC0C0&
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
         Left            =   2160
         TabIndex        =   25
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox supplier_id 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   11040
         TabIndex        =   23
         Text            =   " "
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox gstin_no 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   405
         Left            =   11040
         TabIndex        =   22
         Top             =   1560
         Width           =   2175
      End
      Begin VB.TextBox company_name 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   405
         Left            =   11040
         TabIndex        =   21
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox bill_no 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   360
         Left            =   6600
         TabIndex        =   17
         Text            =   " "
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox order_no 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   360
         Left            =   6600
         TabIndex        =   13
         Text            =   " "
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox invoice_date 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   360
         Left            =   2160
         TabIndex        =   12
         Text            =   " "
         Top             =   1560
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   390
         Left            =   6600
         TabIndex        =   16
         Top             =   1560
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   688
         _Version        =   393216
         Format          =   120389633
         CurrentDate     =   43582
      End
      Begin VB.ComboBox Combo1 
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
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   360
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox return_no 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2160
         TabIndex        =   47
         Text            =   " "
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox invoice_text 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2160
         TabIndex        =   71
         Text            =   " "
         Top             =   360
         Width           =   2175
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
         Left            =   10440
         TabIndex        =   79
         Top             =   1560
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
         Left            =   10800
         TabIndex        =   78
         Top             =   960
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
         Left            =   10560
         TabIndex        =   77
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
         Left            =   6240
         TabIndex        =   76
         Top             =   1560
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
         Left            =   5880
         TabIndex        =   75
         Top             =   480
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
         Left            =   1680
         TabIndex        =   74
         Top             =   1560
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
         Left            =   1440
         TabIndex        =   73
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
         Left            =   1440
         TabIndex        =   72
         Top             =   480
         Width           =   120
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
         Left            =   9240
         TabIndex        =   20
         Top             =   480
         Width           =   1170
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000001&
         BackStyle       =   0  'Transparent
         Caption         =   "Company Name"
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
         Left            =   9000
         TabIndex        =   19
         Top             =   960
         Width           =   1665
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H80000001&
         BackStyle       =   0  'Transparent
         Caption         =   "Gstin No"
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
         TabIndex        =   18
         Top             =   1560
         Width           =   900
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bill No"
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
         Left            =   4920
         TabIndex        =   15
         Top             =   960
         Width           =   705
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Return Date"
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
         Left            =   4800
         TabIndex        =   14
         Top             =   1560
         Width           =   1245
      End
      Begin VB.Label Label6 
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
         Left            =   4800
         TabIndex        =   11
         Top             =   480
         Width           =   960
      End
      Begin VB.Label Label7 
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
         Left            =   240
         TabIndex        =   10
         Top             =   960
         Width           =   1125
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice Date"
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
         Left            =   240
         TabIndex        =   9
         Top             =   1560
         Width           =   1320
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Return No"
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
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   1050
      End
   End
   Begin VB.Frame Frame2 
      Height          =   7095
      Left            =   3120
      TabIndex        =   26
      Top             =   1200
      Visible         =   0   'False
      Width           =   11775
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFF80&
         Caption         =   "CLOSE"
         Height          =   375
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   6480
         Width           =   1095
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "purchase_return1.frx":0101
         Height          =   2655
         Left            =   480
         TabIndex        =   28
         Top             =   3720
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   4683
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
            DataField       =   "S_NO"
            Caption         =   "S_NO"
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
            DataField       =   "RETURN_NO"
            Caption         =   "RETURN_NO"
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
         BeginProperty Column03 
            DataField       =   "PRODUCT_NM"
            Caption         =   "PRODUCT_NM"
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
         BeginProperty Column05 
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
         BeginProperty Column06 
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
         BeginProperty Column07 
            DataField       =   "QUANTITY"
            Caption         =   "QUANTITY"
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
            DataField       =   "EXP_DATE"
            Caption         =   "EXP_DATE"
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
         BeginProperty Column10 
            DataField       =   "AMOUNT"
            Caption         =   "AMOUNT"
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
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column03 
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1289.764
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   975.118
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   915.024
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   1289.764
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "purchase_return1.frx":0116
         Height          =   2175
         Left            =   480
         TabIndex        =   27
         Top             =   600
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   3836
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
         ColumnCount     =   8
         BeginProperty Column00 
            DataField       =   "RETURN_NO"
            Caption         =   "RETURN_NO"
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
            DataField       =   "BILL_NO"
            Caption         =   "BILL_NO"
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
         BeginProperty Column04 
            DataField       =   "SUPPLIER_ID"
            Caption         =   "SUPPLIER_ID"
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
            DataField       =   "RETURN_DATE"
            Caption         =   "RETURN_DATE"
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
         BeginProperty Column07 
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column07 
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
         RecordSource    =   "select * from purchase_return_product"
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
         Height          =   375
         Left            =   5400
         Top             =   960
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
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
         RecordSource    =   "select * from purchase_return"
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
      Begin VB.Line Line2 
         BorderStyle     =   5  'Dash-Dot-Dot
         X1              =   0
         X2              =   11760
         Y1              =   3120
         Y2              =   3120
      End
      Begin VB.Label Label24 
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
         Left            =   3480
         TabIndex        =   70
         Top             =   3360
         Width           =   165
      End
      Begin VB.Label Label23 
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
         TabIndex        =   69
         Top             =   3360
         Width           =   165
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Purchase Return ProductInformation"
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
         Left            =   3735
         TabIndex        =   68
         Top             =   3360
         Width           =   4020
      End
      Begin VB.Label Label25 
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
         Left            =   3720
         TabIndex        =   67
         Top             =   240
         Width           =   165
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
         Left            =   7320
         TabIndex        =   66
         Top             =   240
         Width           =   165
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Purchase Return Information"
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
         Left            =   4035
         TabIndex        =   65
         Top             =   240
         Width           =   3180
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
         Left            =   4080
         TabIndex        =   64
         Top             =   240
         Width           =   165
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
         Left            =   6960
         TabIndex        =   63
         Top             =   240
         Width           =   165
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "PURCHASE RETURN"
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
      Left            =   5685
      TabIndex        =   0
      Top             =   120
      Width           =   9045
   End
End
Attribute VB_Name = "purchase_return"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public del_qty As Integer
Public index1 As Integer
Public q As Integer



Private Sub brand_Click()
sql = "select  distinct(unit) from invoice_product_detail where product_nm='" + product_id.Text + "' and brand='" + brand.Text + "'"
Set r = c.Execute(sql)

unit.clear
Do While Not r.EOF
unit.AddItem r.Fields(0)
r.MoveNext
Loop
End Sub



Private Sub Combo1_Click()

Set r = c.Execute("select * from purchase_return where return_no='" + Combo1.Text + "'")

invoice_no.Text = r.Fields(6)
invoice_date.Text = r.Fields(1)
bill_no.Text = r.Fields(2)
order_no.Text = r.Fields(3)
supplier_id.Text = r.Fields(4)
DTPicker3.Value = r.Fields(5)

index1 = Combo1.ListIndex

Set r = c.Execute("select COMPANY_NAME,gstin_no from supplier_detail where  supplier_id='" + supplier_id.Text + "'")
company_name.Text = r.Fields(0)
gstin_no.Text = r.Fields(1)

''add item in list box
Set r = c.Execute("select * from purchase_return_product where return_no='" + Combo1.Text + "'")

List1.clear
List11.clear
List2.clear
List3.clear
List4.clear
List5.clear
List6.clear
List8.clear
List9.clear
List10.clear
Do While Not r.EOF

List1.AddItem r.Fields(0)
List11.AddItem r.Fields(2)
List2.AddItem r.Fields(3)
List3.AddItem r.Fields(4)
List4.AddItem r.Fields(5)
List5.AddItem r.Fields(6)
List6.AddItem r.Fields(7)
List9.AddItem r.Fields(8)
List10.AddItem r.Fields(9)
List8.AddItem r.Fields(10)

r.MoveNext
Loop

invoice_no.Enabled = False


End Sub

Private Sub Command1_Click()
Frame2.Visible = False
Frame1.Visible = True
Frame3.Visible = True
End Sub

Private Sub Command2_Click()

Dim i As Integer
 List1.BackColor = vbWhite
 List2.BackColor = vbWhite
 List3.BackColor = vbWhite
 List4.BackColor = vbWhite
 List5.BackColor = vbWhite
 List6.BackColor = vbWhite
 List8.BackColor = vbWhite
 List9.BackColor = vbWhite
 List10.BackColor = vbWhite
 List11.BackColor = vbWhite
If Combo1.Visible = False Then

 i = List1.ListCount
 auto
 Set r = New ADODB.Recordset
 sql = "insert into prpsno values('" + List1.List(i) + "')"
 Set r = c.Execute(sql)
 List11.AddItem product_name.Text
 List2.AddItem product_id.Text
 List3.AddItem brand.Text
 List4.AddItem unit.Text
 List5.AddItem unit_price.Text
 List6.AddItem quantity.Text
 List9.AddItem DTPicker2.Value
 List10.AddItem description.Text
 List8.AddItem Val(List5.List(i)) * Val(List6.List(i))
 tot = 0
 For j = 0 To List8.ListCount - 1
   tot = tot + Val(List8.List(i))
 Next
 Label5.Caption = tot

Else
 List1.RemoveItem (index1)
 List2.RemoveItem (index1)
 List3.RemoveItem (index1)
 List4.RemoveItem (index1)
 List5.RemoveItem (index1)
 List6.RemoveItem (index1)
 List9.RemoveItem (index1)
 List10.RemoveItem (index1)
 List8.RemoveItem (index1)

 i = List1.ListCount
 auto
 Set r = New ADODB.Recordset
 sql = "insert into prpsno values('" + List1.List(i) + "')"
 Set r = c.Execute(sql)
 List2.AddItem product_id.Text
 List3.AddItem brand.Text
 List4.AddItem unit.Text
 List5.AddItem unit_price.Text
 List6.AddItem quantity.Text
 List9.AddItem DTPicker2.Value
 List10.AddItem description.Text
 List8.AddItem Val(List5.List(i)) * Val(List6.List(i))
 tot = 0
 For j = 0 To List8.ListCount - 1
   tot = tot + Val(List8.List(i))
 Next
 Label5.Caption = tot
End If
End Sub

Private Sub Command3_Click()


List1.RemoveItem (q)
List2.RemoveItem (q)
List3.RemoveItem (q)
List4.RemoveItem (q)
List5.RemoveItem (q)
List6.RemoveItem (q)
List8.RemoveItem (q)
List9.RemoveItem (q)
List10.RemoveItem (q)
List11.RemoveItem (q)
End Sub

Private Sub delete_Click()
ans = MsgBox("Do you Want to Delete", vbOKCancel + vbInformation)
If ans = 1 Then

Set r = c.Execute("delete purchase_return_product where return_no='" + Combo1.Text + "'")
Set r = c.Execute("delete purchase_return where return_no='" + Combo1.Text + "'")

MsgBox "Record Deleted"
Combo1.RemoveItem (index1)
clear
End If

End Sub

Private Sub description_Click()
If description = "Expired Product" Then
DTPicker2.Visible = True
Label8.Visible = True
Else
DTPicker2.Visible = False
Label8.Visible = False
End If
End Sub

Private Sub Form_Load()
Connection
Set r = New ADODB.Recordset
sql = "select distinct(invoice_no) from purchase_invoice"
Set r = c.Execute(sql)
While r.EOF = False
invoice_no.AddItem r.Fields(0)
r.MoveNext
Wend
autogenerate

purchase_return.Caption = "Purchase Return"
End Sub
Public Function Connection()
Set c = New ADODB.Connection
c.Open "Provider=MSDAORA.1;Password=lnt123;User ID=LNT;Persist Security Info=True"

End Function



Private Sub insert_Click()
Dim i As Integer

ans = MsgBox("Do You want to save", vbYesNo + vbInformation, "save")
If ans = vbYes Then

If invoice_no.Text = "" Then
 invoice_no.BackColor = &HC0C0FF
 MsgBox "Invoice no fields is Empty", vbCritical
ElseIf List1.List(0) = "" Or List2.List(0) = "" Or List3.List(0) = "" Or List4.List(0) = "" Or List5.List(0) = "" Or List8.List(0) = "" Or List6.List(0) = "" Or List9.List(0) = "" Or List10.List(0) = "" Or List11.List(0) = "" Then
 List1.BackColor = &HC0C0FF
 List2.BackColor = &HC0C0FF
 List3.BackColor = &HC0C0FF
 List4.BackColor = &HC0C0FF
 List5.BackColor = &HC0C0FF
 List6.BackColor = &HC0C0FF
 List8.BackColor = &HC0C0FF
 List9.BackColor = &HC0C0FF
 List10.BackColor = &HC0C0FF
 List11.BackColor = &HC0C0FF
 MsgBox "Add product in Listbox", vbCritical

Else
Set r = New ADODB.Recordset
sql = "insert into purchase_return values('" + return_no.Text + "','" + Format(invoice_date, "dd/mmm/yyyy") + "','" + bill_no.Text + "','" + order_no.Text + "','" + supplier_id.Text + "','" + Format(DTPicker3.Value, "dd/mmm/yyyy") + "','" + invoice_no.Text + "'," + Label5.Caption + ")"
Set r = c.Execute(sql)

For i = 0 To List1.ListCount - 1
sql = "insert into purchase_return_product values('" + List1.List(i) + "','" + return_no.Text + "','" + List11.List(i) + "','" + List2.List(i) + "','" + List3.List(i) + "','" + List4.List(i) + "','" + List5.List(i) + "','" + List6.List(i) + "','" + Format(List9.List(i), "dd/mmm/yyyy") + "','" + List10.List(i) + "'," + List8.List(i) + ")"
Set r = c.Execute(sql)

Next
MsgBox "data inserted in purchase_return_product"


For i = 0 To List6.ListCount - 1
Set r = c.Execute("select avl_quantity from stock_detail where unit='" + List4.List(i) + "' and brand='" + List3.List(i) + "' and product_id='" + List11.List(i) + "' ")
avlqty = r.Fields(0) - Val(List6.List(i))
sql = "update stock_detail set avl_quantity= " & avlqty & " where unit='" + List4.List(i) + "' and brand='" + List3.List(i) + "' and product_id='" + List11.List(i) + "' "
Set r = c.Execute(sql)
Next
MsgBox "stock_updated"
Adodc1.Refresh
Adodc2.Refresh
End If
End If
End Sub

Private Sub invoice_no_Click()
Dim i As Integer
Dim invtext As String
invoice_no.BackColor = vbWhite
Set r = New ADODB.Recordset
sql = "select * from purchase_invoice where invoice_no='" + invoice_no.Text + "'"
Set r = c.Execute(sql)
invoice_date.Text = r.Fields(1)
bill_no.Text = r.Fields(2)
order_no.Text = r.Fields(3)
supplier_id.Text = r.Fields(4)

sql = "select company_name,gstin_no from supplier_detail where supplier_id='" + supplier_id.Text + "'"
Set r = c.Execute(sql)
company_name.Text = r.Fields(0)
gstin_no.Text = r.Fields(1)

sql = "select * from invoice_product_detail"
Set r = c.Execute(sql)

'i = invoice_no.ListIndex
invtext = invoice_no.Text
'invoice_no.RemoveItem i
invoice_no.Text = invtext


sql = "select distinct(product_nm) from invoice_product_detail where invoice_no='" & invtext & "'"
Set r = c.Execute(sql)
product_id.clear
While r.EOF = False
product_id.AddItem r.Fields(0)
r.MoveNext
Wend

End Sub

Public Function autogenerate()
Dim a As String
Set r = New ADODB.Recordset
sql = "select max(to_number(substr(return_no,5,length(return_no)))) from purchase_return"
Set r = c.Execute(sql)
If IsNull(r.Fields(0)) Then
return_no.Text = "re" & "00" & 1
Else
return_no.Text = "re" & "00" & r.Fields(0) + 1
End If
a = return_no.Text
If (a = "re" & "001" & "0") Then
sql = "select max(to_number(substr(return_no,4,length(return_no)))) from purchase_return"
Set r = c.Execute(sql)
return_no.Text = "re" & "0" & r.Fields(0) + 1
End If
End Function
Public Function auto()
Dim a As String
Dim i As Integer
i = List1.ListCount
Set r = New ADODB.Recordset
sql = "select max(to_number(substr(s_no,5,length(s_no)))) from prpsno"
Set r = c.Execute(sql)
If IsNull(r.Fields(0)) Then
List1.AddItem "sn" & "00" & 1
Else
List1.AddItem "sn" & "00" & r.Fields(0) + 1
End If
a = List1.List(i)
If (a = "sn" & "001" & "0") Then
sql = "select max(to_number(substr(s_no,4,length(s_no)))) from prpsno"
Set r = c.Execute(sql)
List1.AddItem "in" & "0" & r.Fields(0) + 1
End If
End Function

Private Sub List1_Click()
q = List1.ListIndex

index1 = List2.ListIndex
product_id.Text = List2.List(index1)
brand.Text = List3.List(index1)
unit.Text = List4.List(index1)
unit_price.Text = List5.List(index1)
quantity.Text = List6.List(index1)
description.Text = List10.List(index1)
'DTPicker2.Value = List9.List(index1)

Frame4.Enabled = True
End Sub

Private Sub List10_Click()
index1 = List2.ListIndex
product_id.Text = List2.List(index1)
brand.Text = List3.List(index1)
unit.Text = List4.List(index1)
unit_price.Text = List5.List(index1)
quantity.Text = List6.List(index1)
description.Text = List10.List(index1)
DTPicker2.Value = List9.List(index1)

Frame4.Enabled = True
End Sub

Private Sub List11_Click()
q = List11.ListIndex

index1 = List2.ListIndex
product_id.Text = List2.List(index1)
brand.Text = List3.List(index1)
unit.Text = List4.List(index1)
unit_price.Text = List5.List(index1)
quantity.Text = List6.List(index1)
description.Text = List10.List(index1)
DTPicker2.Value = List9.List(index1)

Frame4.Enabled = True
End Sub

Private Sub List2_Click()
q = List2.ListIndex

index1 = List2.ListIndex
product_id.Text = List2.List(index1)
brand.Text = List3.List(index1)
unit.Text = List4.List(index1)
unit_price.Text = List5.List(index1)
quantity.Text = List6.List(index1)
description.Text = List10.List(index1)


Frame4.Enabled = True
End Sub

Private Sub List3_Click()
q = List3.ListIndex

index1 = List2.ListIndex
product_id.Text = List2.List(index1)
brand.Text = List3.List(index1)
unit.Text = List4.List(index1)
unit_price.Text = List5.List(index1)
quantity.Text = List6.List(index1)
description.Text = List10.List(index1)
'DTPicker2.Value = List9.List(index1)

Frame4.Enabled = True
End Sub

Private Sub List4_Click()
q = List4.ListIndex
index1 = List2.ListIndex
product_id.Text = List2.List(index1)
brand.Text = List3.List(index1)
unit.Text = List4.List(index1)
unit_price.Text = List5.List(index1)
quantity.Text = List6.List(index1)
description.Text = List10.List(index1)
'DTPicker2.Value = List9.List(index1)

Frame4.Enabled = True
End Sub

Private Sub List5_Click()
q = List5.ListIndex
index1 = List2.ListIndex
product_id.Text = List2.List(index1)
brand.Text = List3.List(index1)
unit.Text = List4.List(index1)
unit_price.Text = List5.List(index1)
quantity.Text = List6.List(index1)
description.Text = List10.List(index1)
'DTPicker2.Value = List9.List(index1)

Frame4.Enabled = True
End Sub

Private Sub List6_Click()
index1 = List2.ListIndex
product_id.Text = List2.List(index1)
brand.Text = List3.List(index1)
unit.Text = List4.List(index1)
unit_price.Text = List5.List(index1)
quantity.Text = List6.List(index1)
description.Text = List10.List(index1)
'DTPicker2.Value = List9.List(index1)

Frame4.Enabled = True
End Sub

Private Sub List8_Click()
index1 = List2.ListIndex
product_id.Text = List2.List(index1)
brand.Text = List3.List(index1)
unit.Text = List4.List(index1)
unit_price.Text = List5.List(index1)
quantity.Text = List6.List(index1)
description.Text = List10.List(index1)
DTPicker2.Value = List9.List(index1)

Frame4.Enabled = True
End Sub

Private Sub List9_Click()
index1 = List2.ListIndex
product_id.Text = List2.List(index1)
brand.Text = List3.List(index1)
unit.Text = List4.List(index1)
unit_price.Text = List5.List(index1)
quantity.Text = List6.List(index1)
description.Text = List10.List(index1)
DTPicker2.Value = List9.List(index1)

Frame4.Enabled = True
End Sub

Private Sub new_Click()
clear
autogenerate

invoice_no.Enabled = True
update.Enabled = False
delete.Enabled = False
insert.Enabled = True
Combo1.Visible = False
return_no.Visible = True
End Sub

Private Sub product_id_Click()
Set r = New ADODB.Recordset
sql = "select  distinct brand,p_id from invoice_product_detail where product_nm='" + product_id.Text + "'"
Set r = c.Execute(sql)
product_name.Text = r.Fields(1)
brand.clear
Do While Not r.EOF
brand.AddItem r.Fields(0)
r.MoveNext
Loop
End Sub



Private Sub quantity_KeyPress(KeyAscii As Integer)
quantity.BackColor = vbWhite
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn Then
  If KeyAscii = 13 Then
   End If
Else
KeyAscii = 0
MsgBox "Enter only number"
End If
End Sub

Private Sub quantity_LostFocus()
If Val(quantity.Text) > del_qty Then
MsgBox "maximum delivered quantity is " & Val(del_qty)
End If
End Sub



Private Sub search_Click()
return_no.Visible = False
update.Enabled = True
delete.Enabled = True
insert.Enabled = False

Combo1.Visible = True
insert.Enabled = False

Frame4.Enabled = False

Set r = c.Execute("select return_no from purchase_return")
Combo1.clear

Do While Not r.EOF
Combo1.AddItem r.Fields(0)
r.MoveNext
Loop

End Sub

Private Sub unit_Click()
Set r = New ADODB.Recordset
sql = "select  * from invoice_product_detail where product_nm='" + product_id.Text + "' and brand='" + brand.Text + "' and unit='" + unit.Text + "'"
Set r = c.Execute(sql)
del_qty = r.Fields(10)
unit_price.Text = r.Fields(6)
End Sub


Public Function clear()
List1.clear
List11.clear
List2.clear
List3.clear
List4.clear
List5.clear
List6.clear
List8.clear
List9.clear
List10.clear
invoice_no.Text = "Select Invoice No"
invoice_date.Text = ""
order_no.Text = ""
bill_no.Text = ""
DTPicker3.Value = Date
gstin_no.Text = ""
supplier_id.Text = ""
company_name.Text = ""
Label5.Caption = 0

unit_price.Text = 0
quantity.Text = ""
description.Text = ""
End Function



Private Sub unit_price_KeyPress(KeyAscii As Integer)
unit_price.BackColor = vbWhite
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn Then
  If KeyAscii = 13 Then
   End If
Else
KeyAscii = 0
MsgBox "Enter only number"
End If
End Sub

Private Sub update_Click()
ans = MsgBox("do you want to Update", vbYesNo + vbInformation, "Update")
If ans = vbYes Then
  sql = "update purchase_return_product set PRODUCT_ID='" + List11.List(i) + "',PRODUCT_NM='" + List2.List(i) + "',BRAND='" + List3.List(i) + "',UNIT='" + List4.List(i) + "',UNIT_PRICE=" + List5.List(i) + ",QUANTITY='" + List6.List(i) + "',EXP_DATE='" + Format(List9.List(i), "dd/mmm/yyyy") + "',DESCRIPTION='" + List10.List(i) + "',AMOUNT=" + List8.List(i) + "   "
  Set r = c.Execute(sql)
  asd = MsgBox("Record Updated", vbOKOnly + vbInformation, "Update")
End If

End Sub

Private Sub view_Click()
Frame1.Visible = False
Frame3.Visible = False
Frame2.Visible = True
Adodc1.Refresh
Adodc2.Refresh
End Sub
