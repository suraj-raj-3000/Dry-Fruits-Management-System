VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form purchased_product 
   BackColor       =   &H80000004&
   Caption         =   "purchase_prd"
   ClientHeight    =   7995
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15765
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7995
   ScaleWidth      =   15765
   WindowState     =   2  'Maximized
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
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   8640
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
      Left            =   12480
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   8640
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
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   8640
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
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   8640
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
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   8640
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      Height          =   4935
      Left            =   1320
      TabIndex        =   5
      Top             =   3120
      Width           =   14655
      Begin VB.ListBox List6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   2955
         ItemData        =   "purchase_product.frx":0000
         Left            =   8160
         List            =   "purchase_product.frx":0002
         TabIndex        =   16
         Top             =   720
         Width           =   1095
      End
      Begin VB.ListBox List12 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   2955
         ItemData        =   "purchase_product.frx":0004
         Left            =   12120
         List            =   "purchase_product.frx":0006
         TabIndex        =   32
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H80000004&
         Caption         =   "Cash"
         Height          =   255
         Left            =   2400
         TabIndex        =   4
         Top             =   4240
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.CommandButton Command2 
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
         Left            =   13560
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   1800
         Width           =   855
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   2955
         ItemData        =   "purchase_product.frx":0008
         Left            =   2280
         List            =   "purchase_product.frx":000A
         TabIndex        =   19
         Top             =   720
         Width           =   1815
      End
      Begin VB.ListBox List8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   2955
         ItemData        =   "purchase_product.frx":000C
         Left            =   9240
         List            =   "purchase_product.frx":000E
         TabIndex        =   18
         Top             =   720
         Width           =   1215
      End
      Begin VB.ListBox List7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   2955
         ItemData        =   "purchase_product.frx":0010
         Left            =   7440
         List            =   "purchase_product.frx":0012
         TabIndex        =   17
         Top             =   720
         Width           =   735
      End
      Begin VB.ListBox List5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   2955
         ItemData        =   "purchase_product.frx":0014
         Left            =   6480
         List            =   "purchase_product.frx":0016
         TabIndex        =   15
         Top             =   720
         Width           =   975
      End
      Begin VB.ListBox List4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   2955
         ItemData        =   "purchase_product.frx":0018
         Left            =   5520
         List            =   "purchase_product.frx":001A
         TabIndex        =   14
         Top             =   720
         Width           =   975
      End
      Begin VB.ListBox List3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   2955
         ItemData        =   "purchase_product.frx":001C
         Left            =   4080
         List            =   "purchase_product.frx":001E
         TabIndex        =   13
         Top             =   720
         Width           =   1455
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   2955
         ItemData        =   "purchase_product.frx":0020
         Left            =   120
         List            =   "purchase_product.frx":0022
         TabIndex        =   12
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox paid_amount 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   5280
         MaxLength       =   5
         TabIndex        =   11
         Top             =   4200
         Width           =   1455
      End
      Begin VB.TextBox balance_amount 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   12840
         TabIndex        =   10
         Top             =   4200
         Width           =   1455
      End
      Begin VB.TextBox total_amount 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   9000
         TabIndex        =   9
         Top             =   4200
         Width           =   1575
      End
      Begin VB.ListBox List9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   2955
         ItemData        =   "purchase_product.frx":0024
         Left            =   11280
         List            =   "purchase_product.frx":0026
         TabIndex        =   8
         Top             =   720
         Width           =   855
      End
      Begin VB.ListBox List10 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   2955
         ItemData        =   "purchase_product.frx":0028
         Left            =   10440
         List            =   "purchase_product.frx":002A
         TabIndex        =   7
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox Text11 
         BackColor       =   &H00FFC0C0&
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
         Left            =   13440
         MaxLength       =   5
         TabIndex        =   3
         Text            =   " "
         Top             =   1080
         Width           =   1095
      End
      Begin VB.ListBox List11 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   2955
         ItemData        =   "purchase_product.frx":002C
         Left            =   840
         List            =   "purchase_product.frx":002E
         TabIndex        =   6
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H80000004&
         Caption         =   "Check"
         Height          =   255
         Left            =   2400
         TabIndex        =   2
         Top             =   4260
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Line Line12 
         X1              =   12120
         X2              =   12120
         Y1              =   3600
         Y2              =   3960
      End
      Begin VB.Label total 
         BackColor       =   &H80000007&
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
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
         Left            =   12360
         TabIndex        =   34
         Top             =   3705
         Width           =   690
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackColor       =   &H80000007&
         BackStyle       =   0  'Transparent
         Caption         =   "Total  = "
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
         Left            =   1080
         TabIndex        =   33
         Top             =   3720
         Width           =   735
      End
      Begin VB.Line Line11 
         X1              =   12120
         X2              =   12120
         Y1              =   360
         Y2              =   720
      End
      Begin VB.Shape Shape2 
         Height          =   255
         Left            =   120
         Top             =   3660
         Width           =   13095
      End
      Begin VB.Line Line10 
         X1              =   11280
         X2              =   11280
         Y1              =   360
         Y2              =   720
      End
      Begin VB.Line Line9 
         X1              =   10440
         X2              =   10440
         Y1              =   360
         Y2              =   720
      End
      Begin VB.Line Line8 
         X1              =   9240
         X2              =   9240
         Y1              =   360
         Y2              =   720
      End
      Begin VB.Line Line7 
         X1              =   8160
         X2              =   8160
         Y1              =   360
         Y2              =   720
      End
      Begin VB.Line Line6 
         X1              =   7440
         X2              =   7440
         Y1              =   360
         Y2              =   720
      End
      Begin VB.Line Line5 
         X1              =   6480
         X2              =   6480
         Y1              =   360
         Y2              =   720
      End
      Begin VB.Line Line4 
         X1              =   5520
         X2              =   5520
         Y1              =   360
         Y2              =   720
      End
      Begin VB.Line Line3 
         X1              =   4080
         X2              =   4080
         Y1              =   360
         Y2              =   720
      End
      Begin VB.Line Line2 
         X1              =   2280
         X2              =   2280
         Y1              =   360
         Y2              =   720
      End
      Begin VB.Line Line1 
         X1              =   840
         X2              =   840
         Y1              =   360
         Y2              =   720
      End
      Begin VB.Shape Shape1 
         Height          =   375
         Left            =   120
         Top             =   360
         Width           =   13095
      End
      Begin VB.Shape Shape11 
         Height          =   375
         Left            =   2160
         Top             =   4150
         Width           =   1215
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
         Left            =   120
         TabIndex        =   31
         Top             =   4200
         Width           =   2010
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   $"purchase_product.frx":0030
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   435
         Width           =   12975
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H80000007&
         BackStyle       =   0  'Transparent
         Caption         =   "Paid Amount :-"
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
         Left            =   3720
         TabIndex        =   23
         Top             =   4200
         Width           =   1515
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H80000007&
         BackStyle       =   0  'Transparent
         Caption         =   "Payable Amount  :-"
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
         Left            =   10680
         TabIndex        =   22
         Top             =   4200
         Width           =   1965
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H80000007&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Amount  :-"
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
         Left            =   7320
         TabIndex        =   21
         Top             =   4200
         Width           =   1635
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         BackStyle       =   0  'Transparent
         Caption         =   "DELIVERED QTY"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   555
         Left            =   13320
         TabIndex        =   20
         Top             =   360
         Width           =   1140
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1320
      TabIndex        =   35
      Text            =   "Text1"
      Top             =   3240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   1320
      TabIndex        =   36
      Top             =   720
      Width           =   14655
      Begin VB.TextBox gstin_no 
         Enabled         =   0   'False
         Height          =   405
         Left            =   12000
         TabIndex        =   54
         Top             =   1080
         Width           =   2055
      End
      Begin VB.TextBox order_date 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   7440
         Locked          =   -1  'True
         TabIndex        =   48
         Text            =   " "
         Top             =   1800
         Width           =   2175
      End
      Begin VB.TextBox bill_no 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2040
         TabIndex        =   0
         Top             =   1800
         Width           =   2535
      End
      Begin VB.TextBox company_name 
         Enabled         =   0   'False
         Height          =   405
         Left            =   7440
         TabIndex        =   47
         Top             =   1080
         Width           =   2175
      End
      Begin VB.ComboBox Combo5 
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
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1080
         Width           =   2535
      End
      Begin VB.TextBox invoice_no 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   2040
         TabIndex        =   42
         Text            =   " "
         Top             =   360
         Width           =   2535
      End
      Begin VB.TextBox supplier_id 
         Enabled         =   0   'False
         Height          =   375
         Left            =   7440
         TabIndex        =   41
         Text            =   " "
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox order_no 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2040
         TabIndex        =   40
         Text            =   " "
         Top             =   1080
         Width           =   2535
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   12000
         TabIndex        =   52
         Top             =   360
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         Format          =   125566977
         CurrentDate     =   43580
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   12000
         TabIndex        =   53
         Top             =   1800
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         Format          =   125566977
         CurrentDate     =   43579
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
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   360
         Visible         =   0   'False
         Width           =   2535
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
         Left            =   11280
         TabIndex        =   73
         Top             =   1080
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
         Left            =   11640
         TabIndex        =   72
         Top             =   1800
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
         Left            =   11640
         TabIndex        =   71
         Top             =   360
         Width           =   120
      End
      Begin VB.Label Label26 
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
         Left            =   6960
         TabIndex        =   70
         Top             =   1800
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
         Left            =   7080
         TabIndex        =   69
         Top             =   1080
         Width           =   120
      End
      Begin VB.Label Label20 
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
         Left            =   6960
         TabIndex        =   68
         Top             =   360
         Width           =   120
      End
      Begin VB.Label Label17 
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
         TabIndex        =   67
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
         Left            =   1440
         TabIndex        =   66
         Top             =   360
         Width           =   120
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
         Left            =   10080
         TabIndex        =   51
         Top             =   1800
         Width           =   1440
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
         Left            =   10200
         TabIndex        =   50
         Top             =   360
         Width           =   1320
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H80000001&
         BackStyle       =   0  'Transparent
         Caption         =   "Gst  No"
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
         TabIndex        =   49
         Top             =   1080
         Width           =   780
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackColor       =   &H80000007&
         BackStyle       =   0  'Transparent
         Caption         =   "Order Date"
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
         Left            =   5640
         TabIndex        =   46
         Top             =   1800
         Width           =   1155
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice  No"
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
         TabIndex        =   45
         Top             =   1800
         Width           =   1185
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
         Left            =   5280
         TabIndex        =   44
         Top             =   1080
         Width           =   1665
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
         Left            =   5640
         TabIndex        =   39
         Top             =   360
         Width           =   1170
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
         Left            =   480
         TabIndex        =   38
         Top             =   1080
         Width           =   960
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bill  No"
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
         Left            =   600
         TabIndex        =   37
         Top             =   360
         Width           =   765
      End
   End
   Begin VB.Frame Frame2 
      Height          =   7695
      Left            =   1320
      TabIndex        =   55
      Top             =   600
      Width           =   14655
      Begin VB.CommandButton close 
         BackColor       =   &H00FFFFC0&
         Caption         =   "CLOSE"
         Height          =   375
         Left            =   6840
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   7200
         Width           =   1095
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2655
         Left            =   840
         TabIndex        =   56
         Top             =   600
         Width           =   13215
         _ExtentX        =   23310
         _ExtentY        =   4683
         View            =   3
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "INVOICE NO"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "INVOICE DATE"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "BILL_NO"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "ORDER NO"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "SUPPLIER NAME"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "PHONE NO"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "ADDRESS"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "GSTIN NO"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   3015
         Left            =   720
         TabIndex        =   57
         Top             =   4080
         Width           =   13215
         _ExtentX        =   23310
         _ExtentY        =   5318
         View            =   3
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   14
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S.N"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "INVOICE NO"
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
            Text            =   "RATE"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "GST"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "TOTAL QTY"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "DELIVERED QTY"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "BALANCE QTY"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "AMOUNT"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "PAID AMOUNT"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "BALANCE AMT."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "TOTAL AMOUNT"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Line Line13 
         BorderStyle     =   3  'Dot
         X1              =   120
         X2              =   14520
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Label Label16 
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
         Left            =   9360
         TabIndex        =   65
         Top             =   3720
         Width           =   165
      End
      Begin VB.Label Label12 
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
         Left            =   5760
         TabIndex        =   64
         Top             =   3720
         Width           =   165
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Purchase Product Information"
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
         Left            =   6030
         TabIndex        =   63
         Top             =   3720
         Width           =   3270
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
         Left            =   8400
         TabIndex        =   62
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
         Left            =   5520
         TabIndex        =   61
         Top             =   240
         Width           =   165
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Purchase  Information"
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
         Left            =   5850
         TabIndex        =   60
         Top             =   240
         Width           =   2430
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "PURCHASE PRODUCT"
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
Attribute VB_Name = "purchased_product"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim qty As Integer
Dim item As ListItem

Private Sub balance_amount_Click()
balance_amount.Text = Val(total_amount.Text) - Val(paid_amount.Text)
End Sub




Private Sub bill_no_KeyPress(KeyAscii As Integer)
bill_no.BackColor = vbWhite
End Sub

Private Sub close_Click()
Frame2.Visible = False
Frame1.Visible = True
Frame3.Visible = True
End Sub

Private Sub Combo1_Click()
Set r = c.Execute("select * from purchase_invoice where invoice_no='" + Combo1.Text + "'")

DTPicker2.Value = r.Fields(1)
bill_no.Text = r.Fields(2)
order_no.Text = r.Fields(3)

Set r = c.Execute("select * from supplier_detail where supplier_id='" + r.Fields(4) + "'")

supplier_id.Text = r.Fields(0)
company_name.Text = r.Fields(3)
gstin_no.Text = r.Fields(4)

Set r = c.Execute("select * from order_detail where order_number='" + order_no.Text + "'")
order_date.Text = r.Fields(1)
DTPicker1.Value = r.Fields(3)

List1.clear
List2.clear
List3.clear
List4.clear
List5.clear
List6.clear
List7.clear
List8.clear
List10.clear
List9.clear
List12.clear
List11.clear
sql = "select *from invoice_product_detail where invoice_no='" + Combo1.Text + "'"
Set r = c.Execute(sql)
Do While Not r.EOF
List1.AddItem r!sno
List2.AddItem r!product_nm
List3.AddItem r!brand
List4.AddItem r!unit
List5.AddItem r!unit_price
List7.AddItem r!quantity
List6.AddItem r!igst
List8.AddItem r!amount
List10.AddItem r!deliverd_qty
List9.AddItem r!balance_qty
List11.AddItem r!p_id
List12.AddItem r!pur_amount
r.MoveNext
Loop

Set r = c.Execute("select * from invoice_product_amount where invpoice_no='" + Combo1.Text + "'")
If IsNull(r.Fields(3)) Then
 pd = 0
Else
pd = r.Fields(3)
End If
paid_amount.Text = pd
balance_amount.Text = r.Fields(4)
total_amount.Text = r.Fields(2)
total.Caption = r.Fields(2)

End Sub

Private Sub Combo5_Click()
Dim i As Integer
Combo5.BackColor = vbWhite
List1.BackColor = vbWhite
 List11.BackColor = vbWhite
 List2.BackColor = vbWhite
 List3.BackColor = vbWhite
 List4.BackColor = vbWhite
 List5.BackColor = vbWhite
 List7.BackColor = vbWhite
 List6.BackColor = vbWhite
 List8.BackColor = vbWhite

Set r = New ADODB.Recordset
sql = "select *from order_detail where order_number='" + Combo5.Text + "'"
Set r = c.Execute(sql)
order_no.Text = r.Fields(0)
order_date.Text = r.Fields(1)
supplier_id.Text = r.Fields(2)
DTPicker1.Value = r.Fields(3)
total_amount.Text = r.Fields(7)
paid_amount.Text = r.Fields(5)
balance_amount.Text = r.Fields(6)


Set r = New ADODB.Recordset
sql = "select *from supplier_detail where supplier_id='" + supplier_id.Text + "'"
Set r = c.Execute(sql)
company_name.Text = r.Fields(3)
gstin_no.Text = r.Fields(4)

List1.clear
List2.clear
List3.clear
List4.clear
List5.clear
List6.clear
List7.clear
List8.clear
List11.clear
sql = "select *from ordered_product where order_no='" + Combo5.Text + "'"
Set r = c.Execute(sql)
Do While Not r.EOF
List1.AddItem r!s_no
List2.AddItem r!product_name
List3.AddItem r!brand
List4.AddItem r!unit
List5.AddItem r!rate
List7.AddItem r!quantity
List6.AddItem r!igst
List8.AddItem r!total_amount
List11.AddItem r!p_id
r.MoveNext
Loop

'sql = "select *from ordered_product_amount where order_number='" + order_no.Text + "'"
'Set r = c.Execute(sql)
'total_amount.Text = r.Fields(1)
'paid_amount.Text = r.Fields(2)
'balance_amount.Text = r.Fields(3)
End Sub




Private Sub Command1_Click()

Set r = c.Execute("select avl_quantity from stock_detail where product_id='" + List11.List(i) + "'")
MsgBox r.Fields(0)
qty = Val(r.Fields(0)) + Val(Text11.Text)
Text1.Text = qty
For i = 0 To List1.ListCount - 1
sql = "update stock_detail set invoice_no='" + invoice_no.Text + "',invoice_dt='" + Format(DTPicker2.Value, "dd/mmm/yyyy") + "',avl_quantity=" + Text1.Text + " where product_id='" + List11.List(i) + "'"
Set r = c.Execute(sql)
Next i
End Sub


Private Sub Command2_Click()
Dim i As Integer

 List10.BackColor = vbWhite
 List9.BackColor = vbWhite
 List12.BackColor = vbWhite
 
i = List10.ListCount
j = List11.ListCount
If i = j Then

Else

sql = "select igst from ordered_product where order_no='" + Combo5.Text + "'"
Set r = c.Execute(sql)

If Val(Text11.Text) <= Val(r.Fields(0)) Then

List10.AddItem Text11.Text
List9.AddItem List6.List(i) - Text11.Text
List12.AddItem (List10.List(i) * List5.List(i)) + ((List10.List(i) * List5.List(i)) * List6.List(i) / 100)

For i = 0 To List12.ListCount - 1
tot = tot + Val(List12.List(i))
Next


total.Caption = tot
balance_amount.Text = Val(total.Caption) - (Val(paid_amount.Text))
total_amount.Text = total.Caption
Text11.Text = ""
Else
MsgBox "Invaild Quantity"
Text11.Text = ""
End If
End If
End Sub


Private Sub delete_Click()
an = MsgBox("Do You want to Delete", vbYesNo + vbQuestion, "For Update")
If an = vbYes Then

If Combo1.Text <> blank Then
 Set r = c.Execute("delete invoice_product_amount where invpoice_no='" + Combo1.Text + "'")
 Set r = c.Execute("delete invoice_product_detail where invoice_no='" + Combo1.Text + "'")
 Set r = c.Execute("delete purchase_invoice where invoice_no='" + Combo1.Text + "'")

 sql = "update order_detail set inv_status='no' where order_number='" + order_no.Text + "'"
 Set r = c.Execute(sql)
 Set r = c.Execute("commit")
 MsgBox "Delete Successed"
 clear

 Combo1.clear
 Set r = c.Execute("select invoice_no from purchase_invoice")
 Do While Not r.EOF
  Combo1.AddItem r!invoice_no
  r.MoveNext
 Loop
Else
MsgBox "Select Bill No First", vbCritical, "warning"
End If
End If
End Sub

Private Sub Form_Load()
Connection
Set r = New ADODB.Recordset
sql = "select order_number from order_detail where inv_status='no'"
Set r = c.Execute(sql)
Do While Not r.EOF
Combo5.AddItem r!order_number
r.MoveNext
Loop
autogenerate
MDIForm1.Picture2.Visible = True

purchased_product.Caption = "Purchase Product"

all_purchase1
all_purchase2
End Sub

Public Function clear()
invoice_no.Text = ""
DTPicker2.Value = Date
order_no.Text = ""
order_date.Text = ""
bill_no.Text = ""
supplier_id.Text = ""
DTPicker1.Value = Date
company_name.Text = ""
gstin_no.Text = ""
'unit.Text = ""
'gst.Text = ""
'unit_price.Text = ""
'brand.Text = ""
'quantity.Text = ""
'product_id.Text = ""
'bal_qty.Text = ""
'del_qty.Text = ""
'order_amt.Text = ""
'purchase_amt.Text = ""
'igst.Text = ""
paid_amount.Text = ""
balance_amount.Text = ""
total_amount.Text = ""
Text11.Text = ""
total.Caption = "Total"
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
List11.clear
List12.clear
End Function



Private Sub insert_Click()
 Dim i As Integer
 Dim j As Integer
ans = MsgBox("Do you Want to Save", vbOKCancel + vbInformation)
If ans = 1 Then
If Combo5.Text = "" Then
 Combo5.BackColor = &HC0C0FF
 MsgBox "Select Order No", vbCritical
ElseIf bill_no.Text = "" Then
 bill_no.BackColor = &HC0C0FF
ElseIf List10.List(0) = "" Or List9.List(0) = "" Or List12.List(0) = "" Then
 List10.BackColor = &HC0C0FF
 List9.BackColor = &HC0C0FF
 List12.BackColor = &HC0C0FF
 MsgBox "Enter Delivered Quantity", vbCritical
ElseIf List1.List(0) = "" Or List11.List(0) = "" Or List2.List(0) = "" Or List3.List(0) = "" Or List4.List(0) = "" Or List5.List(0) = "" Or List7.List(0) = "" Or List6.List(0) = "" Or List8.List(0) = "" Or List10.List(0) = "" Or List9.List(0) = "" Or List12.List(0) = "" Then
 List1.BackColor = &HC0C0FF
 List11.BackColor = &HC0C0FF
 List2.BackColor = &HC0C0FF
 List3.BackColor = &HC0C0FF
 List4.BackColor = &HC0C0FF
 List5.BackColor = &HC0C0FF
 List7.BackColor = &HC0C0FF
 List6.BackColor = &HC0C0FF
 List8.BackColor = &HC0C0FF
 List10.BackColor = &HC0C0FF
 List9.BackColor = &HC0C0FF
 List12.BackColor = &HC0C0FF
 MsgBox "Add Product In List Box"

Else
Set r = New ADODB.Recordset
sql = "insert into purchase_invoice values('" + invoice_no.Text + "','" + Format(DTPicker2.Value, "dd/mmm/yyyy") + "','" + bill_no.Text + "','" + order_no.Text + "','" + supplier_id.Text + "')"

Set r = c.Execute(sql)

For i = 0 To List1.ListCount - 1
sql = "insert into invoice_product_detail values('" + List1.List(i) + "','" + invoice_no.Text + "','" + List11.List(i) + "','" + List2.List(i) + "','" + List3.List(i) + "','" + List4.List(i) + "','" + List5.List(i) + "','" + List7.List(i) + "','" + List6.List(i) + "','" + List8.List(i) + "','" + List10.List(i) + "','" + List9.List(i) + "'," + List12.List(i) + ")"

Set r = c.Execute(sql)
Next i

sql = "insert into invoice_product_amount values('" + invoice_no.Text + "','" + order_no.Text + "','" + total_amount.Text + "','" + paid_amount.Text + "','" + balance_amount.Text + "','" + bill_no.Text + "','" + supplier_id.Text + "')"
Set r = c.Execute(sql)


' ' update product Quantity ''

For i = 0 To List1.ListCount - 1
sql = "update stock_detail set invoice_dt='" + Format(DTPicker2.Value, "dd/mmm/yyyy") + "' where product_id='" + List11.List(i) + "'"
Set r = c.Execute(sql)

sql = "update stock_detail set avl_quantity='" + List10.List(i) + "' where product_nm='" + List2.List(i) + "' and brand='" + List3.List(i) + "' and unit='" + List4.List(i) + "' "
Set r = c.Execute(sql)

Next i

sql = "update order_detail set inv_status='yes' where order_number='" + order_no.Text + "'"
Set r = c.Execute(sql)


MsgBox "Data saved"




sql = "select order_number from order_detail where inv_status='no'"
Set r = c.Execute(sql)
Combo5.clear
Do While Not r.EOF
Combo5.AddItem r!order_number
r.MoveNext
Loop

clear
autogenerate
End If
End If
End Sub

Private Sub List1_Click()
Dim i As Integer
i = List1.ListIndex
product_id.Text = List11.List(i)
product_name.Text = List2.List(i)
unit_price.Text = List5.List(i)
brand.Text = List3.List(i)
unit.Text = List4.List(i)
unit_price.Text = List5.List(i)
quantity.Text = List7.List(i)
gst.Text = List6.List(i)
order_amt.Text = List8.List(i)
igst.Text = List6.List(i)
purchase_amt.Text = List12.List(i)
del_qty.Text = List10.List(i)
bal_qty.Text = List9.List(i)



End Sub

Public Function autogenerate()

Dim a As String
Set r = New ADODB.Recordset
sql = "select max(to_number(substr(invoice_no,5,length(invoice_no)))) from purchase_invoice"
Set r = c.Execute(sql)
If IsNull(r.Fields(0)) Then
invoice_no.Text = "in" & "00" & 1
Else
invoice_no.Text = "in" & "00" & r.Fields(0) + 1
End If
a = invoice_no.Text
If (a = "in" & "001" & "0") Then
sql = "select max(to_number(substr(invoice_no,4,length(invoice_no)))) from purchase_invoice"
Set r = c.Execute(sql)
invoice_no.Text = "in" & "0" & r.Fields(0) + 1
End If

End Function

Private Sub List10_Click()
i = List10.ListIndex
Text11.Text = List10.List(i)
End Sub



Private Sub new_Click()

clear
Set r = New ADODB.Recordset
sql = "select order_number from order_detail where inv_status='no'"
Set r = c.Execute(sql)
Combo5.clear
Do While Not r.EOF
Combo5.AddItem r!order_number
r.MoveNext
Loop
autogenerate

Combo1.Visible = False
invoice_no.Visible = True
insert.Enabled = True
delete.Enabled = False

End Sub


Private Sub report_Click()
Combo1.Visible = True
invoice_no.Visible = False

Combo1.clear
Set r = c.Execute("select invoice_no from purchase_invoice")
Do While Not r.EOF
Combo1.AddItem r!invoice_no
r.MoveNext
Loop
delete.Enabled = True
insert.Enabled = False
End Sub







Private Sub Text11_Change()
If Text11.Text = "" Then
i = 0
ElseIf Val(Text11.Text) > List6.List(i) Then
 MsgBox "Invaild Quantity"
 i = i + 1
End If
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn Then

Else
KeyAscii = 0
MsgBox "Enter only number", vbCritical
End If
End Sub





Private Sub view_Click()
Frame2.Visible = True
Frame1.Visible = False
Frame3.Visible = False

all_purchase1
all_purchase2
End Sub

Public Function all_purchase1()


ListView1.ListItems.clear
Set r = c.Execute("select * from purchase_invoice,supplier_detail ")

While Not r.EOF

Set item = ListView1.ListItems.add(, , r.Fields(0))
item.SubItems(1) = r.Fields(1)
item.SubItems(2) = r.Fields(2)
item.SubItems(3) = rf3
item.SubItems(4) = r.Fields(12)
item.SubItems(5) = r.Fields(6)
item.SubItems(6) = r.Fields(11)
item.SubItems(7) = r.Fields(9)
r.MoveNext
Wend
r.close

End Function

Public Function all_purchase2()

ListView2.ListItems.clear
Set r = New ADODB.Recordset
Set r = c.Execute("select * from invoice_product_detail,invoice_product_amount")
While Not r.EOF

Set item = ListView2.ListItems.add(, , r.Fields(0))
item.SubItems(1) = r.Fields(1)
item.SubItems(2) = r.Fields(3)
item.SubItems(3) = r.Fields(4)
item.SubItems(4) = r.Fields(5)
item.SubItems(5) = r.Fields(6)
item.SubItems(6) = r.Fields(7)
item.SubItems(7) = r.Fields(8)
item.SubItems(8) = r.Fields(10)
item.SubItems(9) = r.Fields(11)
item.SubItems(10) = r.Fields(9)
item.SubItems(11) = r.Fields(16)
item.SubItems(12) = r.Fields(17)
item.SubItems(13) = r.Fields(15)
r.MoveNext
Wend

End Function
