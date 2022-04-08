VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form stock_form 
   Caption         =   " "
   ClientHeight    =   8535
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15525
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8535
   ScaleWidth      =   15525
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFC0FF&
      Caption         =   " View"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   8400
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFC0FF&
      Caption         =   " Status"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   8400
      Width           =   1935
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   960
      Top             =   4560
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1085
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
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from stock_detail"
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
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "<<  Manually Update Stock Quantity  >>"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8400
      Width           =   5415
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000004&
      Height          =   6015
      Left            =   4440
      TabIndex        =   22
      Top             =   1680
      Visible         =   0   'False
      Width           =   8415
      Begin VB.OptionButton Option2 
         BackColor       =   &H80000004&
         Caption         =   "Decrese Quantity"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4320
         TabIndex        =   43
         Top             =   4560
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H80000004&
         Caption         =   "Increse Quantity"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2280
         TabIndex        =   42
         Top             =   4560
         Width           =   2055
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
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   960
         Width           =   1815
      End
      Begin VB.ComboBox Combo3 
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
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   1680
         Width           =   1815
      End
      Begin VB.ComboBox Combo5 
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
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   2400
         Width           =   1815
      End
      Begin VB.TextBox qty 
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
         Left            =   3720
         MaxLength       =   4
         TabIndex        =   33
         Top             =   3840
         Width           =   1800
      End
      Begin VB.TextBox rate 
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
         Left            =   3720
         MaxLength       =   4
         TabIndex        =   32
         Top             =   3120
         Width           =   1800
      End
      Begin VB.CommandButton update 
         BackColor       =   &H0080C0FF&
         Caption         =   " UPDATE"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
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
         Top             =   5280
         Width           =   1095
      End
      Begin VB.CommandButton new 
         BackColor       =   &H00C0C0C0&
         Caption         =   " NEW"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   5280
         Width           =   1095
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   " Product Name  :-"
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
         Left            =   1800
         TabIndex        =   41
         Top             =   960
         Width           =   1800
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Brand  :-"
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
         Left            =   2040
         TabIndex        =   40
         Top             =   1800
         Width           =   945
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Quantity :-"
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
         Left            =   2160
         TabIndex        =   39
         Top             =   3960
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unit  :-"
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
         Left            =   2160
         TabIndex        =   38
         Top             =   2520
         Width           =   675
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Avilable Qty :-"
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
         Left            =   1920
         TabIndex        =   37
         Top             =   3240
         Width           =   1470
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000004&
      Caption         =   "Frame3"
      Height          =   6735
      Left            =   3360
      TabIndex        =   25
      Top             =   1560
      Visible         =   0   'False
      Width           =   10695
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   4200
         TabIndex        =   27
         Top             =   480
         Width           =   2415
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4575
         Left            =   1080
         TabIndex        =   26
         Top             =   1320
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   8070
         View            =   3
         Sorted          =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S no"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Product ID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Product Name"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Unit"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Brand"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Qty"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Select Product"
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
         Left            =   1920
         TabIndex        =   28
         Top             =   480
         Width           =   1860
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000004&
      Height          =   7815
      Left            =   3360
      TabIndex        =   1
      Top             =   480
      Width           =   10695
      Begin VB.CommandButton Command4 
         BackColor       =   &H008080FF&
         Caption         =   "REFRESH"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8520
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
         Width           =   1575
      End
      Begin VB.Image Image9 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1455
         Left            =   6960
         Top             =   5880
         Width           =   2895
      End
      Begin VB.Image Image8 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1455
         Left            =   3960
         Top             =   5880
         Width           =   2895
      End
      Begin VB.Image Image7 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1455
         Left            =   960
         Top             =   5880
         Width           =   2895
      End
      Begin VB.Image Image6 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1455
         Left            =   6960
         Top             =   3600
         Width           =   2895
      End
      Begin VB.Image Image5 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1455
         Left            =   3960
         Top             =   3600
         Width           =   2895
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1455
         Left            =   960
         Top             =   3600
         Width           =   2895
      End
      Begin VB.Image Image4 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1455
         Left            =   6960
         Top             =   1320
         Width           =   2895
      End
      Begin VB.Image Image3 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1455
         Left            =   3960
         Top             =   1320
         Width           =   2895
      End
      Begin VB.Image Image2 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1455
         Left            =   960
         Top             =   1320
         Width           =   2895
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "Total Supplier Order"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   975
         TabIndex        =   21
         Top             =   2385
         Width           =   2865
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "Total Customer Invoice"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3975
         TabIndex        =   20
         Top             =   6945
         Width           =   2865
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "Total Customer Order"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   975
         TabIndex        =   19
         Top             =   6945
         Width           =   2865
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "Total Supplier Invoice"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3975
         TabIndex        =   18
         Top             =   2385
         Width           =   2865
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4800
         TabIndex        =   17
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "Total Product"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   975
         TabIndex        =   16
         Top             =   4665
         Width           =   2865
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1800
         TabIndex        =   15
         Top             =   3840
         Width           =   1335
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "Dameged Product"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3975
         TabIndex        =   14
         Top             =   4665
         Width           =   2865
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4800
         TabIndex        =   13
         Top             =   3840
         Width           =   1095
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "Expired Product"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6975
         TabIndex        =   12
         Top             =   4665
         Width           =   2865
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7800
         TabIndex        =   11
         Top             =   3840
         Width           =   975
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "Supplier Pending Order"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6975
         TabIndex        =   10
         Top             =   2385
         Width           =   2865
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7680
         TabIndex        =   9
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "STATUS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   4680
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
      Begin VB.Shape Shape8 
         Height          =   615
         Left            =   4320
         Shape           =   2  'Oval
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "Customer Pending Order"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6975
         TabIndex        =   7
         Top             =   6945
         Width           =   2865
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4560
         TabIndex        =   6
         Top             =   6120
         Width           =   1455
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1800
         TabIndex        =   5
         Top             =   6120
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1800
         TabIndex        =   4
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7800
         TabIndex        =   3
         Top             =   6120
         Width           =   1215
      End
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "STOCK  DETAIL"
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
      Left            =   4320
      TabIndex        =   31
      Top             =   120
      Width           =   9045
   End
End
Attribute VB_Name = "stock_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim item As ListItem

Private Sub Combo1_Click()

 Combo3.clear
 Set r = c.Execute("select distinct(brand) from stock_detail where product_nm='" + Combo1.Text + "'")
 Do While Not r.EOF
 Combo3.AddItem r.Fields(0)
 r.MoveNext
 Loop

End Sub

Private Sub Combo3_Click()
 Combo5.clear
 Set r = c.Execute("select distinct(unit) from stock_detail where product_nm='" + Combo1.Text + "' and brand='" + Combo3.Text + "'")
 Do While Not r.EOF
 Combo5.AddItem r.Fields(0)
 r.MoveNext
 Loop
End Sub

Private Sub Combo4_click()
Set r = c.Execute("select * from stock_detail where product_nm='" + Combo4.Text + "'")
ListView1.ListItems.clear
If IsNull(r.Fields(5)) Then
qty = 0
Else
qty = r.Fields(5)
End If
While Not r.EOF
Set item = ListView1.ListItems.add(, , r.Fields(0))
item.SubItems(1) = r.Fields(2)
item.SubItems(2) = r.Fields(3)
item.SubItems(3) = r.Fields(4)
item.SubItems(4) = r.Fields(7)
item.SubItems(5) = qty
 r.MoveNext
Wend

End Sub

Private Sub Combo5_Click()

Set r = c.Execute("select distinct(avl_quantity)  from stock_detail where product_nm='" + Combo1.Text + "' and brand='" + Combo3.Text + "' and unit='" + Combo5.Text + "' ")
rate.Text = r.Fields(0)
End Sub

Private Sub Command2_Click()
If Frame1.Visible = False Then
ans = MsgBox("Do You Want To Update Manually", vbOKCancel + vbInformation)
If ans = 1 Then
Frame2.Visible = False
Frame3.Visible = False
Frame1.Visible = True

Combo1.clear
 Set r = c.Execute("select distinct(PRODUCT_NM) from stock_detail")
 Do While Not r.EOF
 Combo1.AddItem r.Fields(0)
 r.MoveNext
 Loop
End If
End If
End Sub

Private Sub Command4_Click()
Unload Me
Load stock_form
stock_form.Show
End Sub

Private Sub Command5_Click()
Frame3.Visible = False
Frame2.Visible = True
Frame1.Visible = False
End Sub

Private Sub Command6_Click()
Frame2.Visible = False
Frame1.Visible = False
Frame3.Visible = True

Combo4.clear
Set r = c.Execute("select distinct(product_nm) from stock_detail")
Do While Not r.EOF
Combo4.AddItem r.Fields(0)
r.MoveNext
Loop

End Sub

Private Sub Form_Load()
Dim qty As Integer
Connection

Image1.MousePointer = vbCustom
Image1.MouseIcon = LoadPicture("C:\Program Files x86\Microsoft Visual Studio\COMMON\Graphics\Cursors\H_POINT.cur")

Image8.MousePointer = vbCustom
Image8.MouseIcon = LoadPicture("C:\Program Files x86\Microsoft Visual Studio\COMMON\Graphics\Cursors\H_POINT.cur")

Image2.MousePointer = vbCustom
Image2.MouseIcon = LoadPicture("C:\Program Files x86\Microsoft Visual Studio\COMMON\Graphics\Cursors\H_POINT.cur")

Image3.MousePointer = vbCustom
Image3.MouseIcon = LoadPicture("C:\Program Files x86\Microsoft Visual Studio\COMMON\Graphics\Cursors\H_POINT.cur")

Image4.MousePointer = vbCustom
Image4.MouseIcon = LoadPicture("C:\Program Files x86\Microsoft Visual Studio\COMMON\Graphics\Cursors\H_POINT.cur")

Image5.MousePointer = vbCustom
Image5.MouseIcon = LoadPicture("C:\Program Files x86\Microsoft Visual Studio\COMMON\Graphics\Cursors\H_POINT.cur")

Image6.MousePointer = vbCustom
Image6.MouseIcon = LoadPicture("C:\Program Files x86\Microsoft Visual Studio\COMMON\Graphics\Cursors\H_POINT.cur")

Image7.MousePointer = vbCustom
Image7.MouseIcon = LoadPicture("C:\Program Files x86\Microsoft Visual Studio\COMMON\Graphics\Cursors\H_POINT.cur")

Image9.MousePointer = vbCustom
Image9.MouseIcon = LoadPicture("C:\Program Files x86\Microsoft Visual Studio\COMMON\Graphics\Cursors\H_POINT.cur")



stock_form.Caption = "Stock"
MDIForm1.Picture2.Visible = True

Set r = c.Execute("select count(order_number) from order_detail")
Label5.Caption = r.Fields(0)

Set r = c.Execute("select count(invoice_no) from purchase_invoice")
Label9.Caption = r.Fields(0)

Set r = c.Execute("select count(order_number) from order_detail where inv_status='no'")
Label19.Caption = r.Fields(0)

Set r = c.Execute("select sum(avl_quantity) from stock_detail")
If IsNull(r.Fields(0)) Then
qty = 0
Else
qty = r.Fields(0)
End If
Label11.Caption = qty

Label14.Caption = "0"
Label17.Caption = "0"

'Set r = New ADODB.Recordset
'sql = "select count(product_id) from product_detail where exp_date='" + Format(Date, "dd/mmm/yyyy") + "'"
'Set r = c.Execute(sql)
'Label17.Caption = r.Fields(0)




Set r = c.Execute("select count(order_number) from customer_order_detail ")
Label25.Caption = r.Fields(0)


Set r = c.Execute("select count(order_number) from customer_order_detail where status='no' ")
Label21.Caption = r.Fields(0)


Set r = c.Execute("select count(invoice_no) from invoice_detail ")
Label23.Caption = r.Fields(0)



Set r = c.Execute("select * from stock_detail")
While Not r.EOF

If IsNull(r.Fields(5)) Then
qty = 0
Else
qty = r.Fields(5)
End If
Set item = ListView1.ListItems.add(, , r.Fields(0))
item.SubItems(1) = r.Fields(2)
item.SubItems(2) = r.Fields(3)
item.SubItems(3) = r.Fields(4)
item.SubItems(4) = r.Fields(7)
item.SubItems(5) = qty
 

r.MoveNext
Wend

End Sub




Private Sub Image1_Click()
Frame2.Visible = False
Frame1.Visible = False
Frame3.Visible = True

Combo4.clear
Set r = c.Execute("select distinct(product_nm) from stock_detail")
Do While Not r.EOF
Combo4.AddItem r.Fields(0)
r.MoveNext
Loop
End Sub

Private Sub new_Click()
clear
End Sub

Public Function clear()
Combo1.Text = ""
Combo3.clear
Combo5.clear
qty.Text = ""
rate.Text = ""
End Function


Private Sub Option2_Click()
'If rate.Text < qty.Text Then
'MsgBox "Invailed Quantity" & "Maximum Quantity is " & rate.Text
'End If
End Sub

Private Sub qty_LostFocus()
If Val(qty.Text) > Val(rate.Text) Then
 ans = MsgBox("Invailed Quantity", vbOKOnly + vbInformation, "Warring")
End If
End Sub

Private Sub update_Click()
If Combo1.Text <> blank And Combo3.Text <> blank And Combo5.Text <> blank And qty.Text <> blank Then
  If Option1.Value = True Then
   sql = "update stock_detail set avl_quantity=avl_quantity +'" + qty.Text + "' where product_nm='" + Combo1.Text + "' and brand='" + Combo3.Text + "' and unit='" + Combo5.Text + "' "
   MsgBox sql
   Set r = c.Execute(sql)
   MsgBox "Quantity Increased"
  ElseIf Option2.Value = True Then
    sql = "update stock_detail set avl_quantity=avl_quantity - '" + qty.Text + "' where product_nm='" + Combo1.Text + "' and brand='" + Combo3.Text + "' and unit='" + Combo5.Text + "' "
    MsgBox sql
    Set r = c.Execute(sql)
    MsgBox "Quantity Decreased"
  End If

End If
End Sub
